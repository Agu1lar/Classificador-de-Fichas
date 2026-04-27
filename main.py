import json
import logging
import os
import subprocess
import re
import shutil
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional

import cv2
import numpy as np
import pytesseract
import tkinter as tk
from pdf2image import convert_from_path
from tkinter import filedialog, messagebox, ttk

try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None


FICHA_REGEX = re.compile(r"\b(\d{5,6}-\d{2})\b")
FICHA_SPLIT_REGEX = re.compile(r"\b(\d{5,6})\s*[-_/]?\s*(\d{2})\b")
FICHA_CONTEXT_REGEX = re.compile(
    r"\b(?:ficha|contrato)\D{0,20}(\d{5,6})(?:\D{0,6}(\d{2}))?\b",
    re.IGNORECASE,
)
FICHA_CONTEXT_OCR_REGEX = re.compile(
    r"\b(?:ficha|contrato)\D{0,20}([0-9OQDGSBILZ]{5,6})(?:\D{0,6}([0-9OQDGSBILZ]{2}))?\b",
    re.IGNORECASE,
)
CHECKLIST_CONTRACT_REGEX = re.compile(
    r"(?:check[\s-]*list|plataforma|libera[çc][aã]o).{0,160}?(\d{5,6})(?:\D{0,10}(\d{2}))?\b",
    re.IGNORECASE | re.DOTALL,
)
CHECKLIST_CONTRACT_OCR_REGEX = re.compile(
    r"(?:check[\s-]*list|plataforma|libera[çc][aã]o).{0,160}?([0-9OQDGSBILZ]{5,6})(?:\D{0,10}([0-9OQDGSBILZ]{2}))?\b",
    re.IGNORECASE | re.DOTALL,
)
STANDALONE_CONTRACT_DIGITS_REGEX = re.compile(r"\b(\d{5,6})\b")
CONTRACT_FOLDER_NAME_REGEX = re.compile(r"^\d{5,6}$")
CONTRACT_INPUT_REGEX = re.compile(r"^\d{5,6}$")
FICHA_HINT_REGEX = re.compile(r"\bficha\b", re.IGNORECASE)
FICHA_KEYWORDS = ("ficha", "contrato", "comprovante")
REJECTION_KEYWORDS = ("telefone", "fone", "fax", "cel", "cnpj", "cep", "patrimonio", "patrimônio")
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png"}
PDF_EXTENSIONS = {".pdf"}
OCR_DIGIT_MAP = str.maketrans({
    "O": "0",
    "Q": "0",
    "D": "0",
    "I": "1",
    "L": "1",
    "Z": "2",
    "S": "5",
    "B": "8",
    "G": "6",
})

def get_app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = get_app_base_dir()
INPUT_DIR = BASE_DIR / "entrada"
DEFAULT_OUTPUT_DIR = BASE_DIR / "saida"
CONFIG_FILE = BASE_DIR / "config.json"
LOG_DIR = BASE_DIR / "logs"
SCAN_DEBUG_DIR = BASE_DIR / "scan_debug"
APP_ICON_FILE = BASE_DIR / "LogoSetup_app.ico"
MANUAL_REVIEW_DIRNAME = "_revisar_manual"
OCR_LANGUAGE = "eng"
WIA_SCANNER_DEVICE_TYPE = 1
WIA_DOCUMENT_HANDLING_SELECT = 3088
WIA_DOCUMENT_HANDLING_STATUS = 3087
WIA_DOCUMENT_HANDLING_CAPABILITIES = 3086
WIA_SCAN_PAGES = 3096
WIA_FEEDER = 1
WIA_FLATBED = 2
WIA_DUPLEX = 4
WIA_FRONT_ONLY = 32
WIA_FEED_READY = 1
WIA_FORMAT_JPEG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
WIA_MAX_BATCH_PAGES = 200
FILE_READY_RETRIES = 12
FILE_READY_DELAY_SECONDS = 0.35
MONITOR_POLL_SECONDS = 1.0
MONITOR_STABLE_POLLS = 3


def configure_tesseract() -> None:
    global OCR_LANGUAGE

    env_cmd = os.getenv("TESSERACT_CMD")
    default_windows_cmd = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
    tessdata_dir = None

    if env_cmd:
        pytesseract.pytesseract.tesseract_cmd = env_cmd
    elif default_windows_cmd.exists():
        pytesseract.pytesseract.tesseract_cmd = str(default_windows_cmd)

    tesseract_executable = Path(pytesseract.pytesseract.tesseract_cmd)
    if tesseract_executable.exists():
        candidate_tessdata = tesseract_executable.parent / "tessdata"
        if candidate_tessdata.exists():
            tessdata_dir = candidate_tessdata
            os.environ["TESSDATA_PREFIX"] = str(candidate_tessdata)

    available_languages = set()
    if tessdata_dir:
        available_languages = {
            trained.stem
            for trained in tessdata_dir.glob("*.traineddata")
        }

    if {"por", "eng"}.issubset(available_languages):
        OCR_LANGUAGE = "por+eng"
    elif "por" in available_languages:
        OCR_LANGUAGE = "por"
    elif "eng" in available_languages:
        OCR_LANGUAGE = "eng"
    else:
        OCR_LANGUAGE = "eng"


def load_config() -> dict:
    default_config = {
        "matriz_path": str(DEFAULT_OUTPUT_DIR),
        "scanner_device_id": "",
        "input_path": str(INPUT_DIR),
    }
    if CONFIG_FILE.exists():
        with CONFIG_FILE.open("r", encoding="utf-8") as file:
            loaded = json.load(file)
        default_config.update(loaded)
    return default_config


def save_config(config: dict) -> None:
    with CONFIG_FILE.open("w", encoding="utf-8") as file:
        json.dump(config, file, ensure_ascii=True, indent=2)


def setup_logging() -> Path:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / "nao_processados.log"

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    handler = logging.FileHandler(log_file, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
    logger.addHandler(handler)
    return log_file


def write_scan_debug(message: str) -> None:
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        debug_file = LOG_DIR / "scanner_debug.log"
        with debug_file.open("a", encoding="utf-8") as file:
            file.write(f"{datetime.now():%Y-%m-%d %H:%M:%S.%f} | {message}\n")
    except Exception:
        pass


def preprocess_image(image: np.ndarray) -> np.ndarray:
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    _, thresholded = cv2.threshold(
        blurred,
        0,
        255,
        cv2.THRESH_BINARY + cv2.THRESH_OTSU,
    )
    return thresholded


def preprocess_image_adaptive(image: np.ndarray) -> np.ndarray:
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    enlarged = cv2.resize(gray, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)
    filtered = cv2.bilateralFilter(enlarged, 9, 75, 75)
    return cv2.adaptiveThreshold(
        filtered,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31,
        11,
    )


def image_to_text(
    image: np.ndarray,
    psm: int = 6,
    adaptive: bool = False,
    extra_config: str = "",
) -> str:
    processed = preprocess_image_adaptive(image) if adaptive else preprocess_image(image)
    config = f"--oem 3 --psm {psm} {extra_config}".strip()
    return pytesseract.image_to_string(processed, lang=OCR_LANGUAGE, config=config)


def get_image_regions(image: np.ndarray) -> list[tuple[str, np.ndarray, int, bool, str]]:
    height, width = image.shape[:2]
    top_end = max(1, int(height * 0.38))
    right_start = max(0, int(width * 0.68))
    center_start = max(0, int(width * 0.52))
    bottom_start = max(0, int(height * 0.55))
    contract_top_end = max(1, int(height * 0.22))
    contract_right_start = max(0, int(width * 0.72))
    contract_tight_top_end = max(1, int(height * 0.17))
    contract_tight_right_start = max(0, int(width * 0.77))

    return [
        ("full", image, 6, False, ""),
        ("full_adaptive", image, 6, True, ""),
        ("top", image[:top_end, :], 6, False, ""),
        ("top_adaptive", image[:top_end, :], 6, True, ""),
        ("top_right", image[:top_end, right_start:], 6, True, ""),
        ("top_right_single", image[:top_end, right_start:], 11, True, ""),
        (
            "top_right_contract",
            image[:contract_top_end, contract_right_start:],
            7,
            True,
            "-c tessedit_char_whitelist=0123456789-",
        ),
        (
            "top_right_contract_tight",
            image[:contract_tight_top_end, contract_tight_right_start:],
            8,
            True,
            "-c tessedit_char_whitelist=0123456789-",
        ),
        ("right_strip", image[:, center_start:], 6, True, ""),
        ("bottom", image[bottom_start:, :], 6, True, ""),
    ]


def generate_oriented_regions(image: np.ndarray) -> list[tuple[str, np.ndarray, int, bool, str]]:
    orientations = [
        ("r0", image),
        ("r180", cv2.rotate(image, cv2.ROTATE_180)),
        ("r90", cv2.rotate(image, cv2.ROTATE_90_CLOCKWISE)),
        ("r270", cv2.rotate(image, cv2.ROTATE_90_COUNTERCLOCKWISE)),
    ]

    regions: list[tuple[str, np.ndarray, int, bool, str]] = []
    for orientation_name, oriented_image in orientations:
        for region_name, region_image, psm, adaptive, extra_config in get_image_regions(oriented_image):
            regions.append((f"{orientation_name}_{region_name}", region_image, psm, adaptive, extra_config))
    return regions


def image_to_text_variants(image: np.ndarray) -> list[str]:
    variants: list[str] = []
    for region_name, region_image, psm, adaptive, extra_config in generate_oriented_regions(image):
        if region_image.size == 0:
            continue
        try:
            text = image_to_text(region_image, psm=psm, adaptive=adaptive, extra_config=extra_config)
        except Exception as exc:
            write_scan_debug(f"image_to_text_variants -> {region_name} error: {exc}")
            continue
        normalized = re.sub(r"\s+", " ", text).strip()
        if normalized:
            variants.append(f"[REGION:{region_name}] {normalized}")
    return variants


def ocr_image_file(image_path: Path) -> str:
    image_bytes = np.fromfile(str(image_path), dtype=np.uint8)
    image = cv2.imdecode(image_bytes, cv2.IMREAD_COLOR)
    if image is None:
        raise ValueError(f"Falha ao abrir a imagem: {image_path}")
    return "\n".join(image_to_text_variants(image))


def pil_to_cv2(pil_image) -> np.ndarray:
    rgb_image = pil_image.convert("RGB")
    return cv2.cvtColor(np.array(rgb_image), cv2.COLOR_RGB2BGR)


def ocr_pdf_file(pdf_path: Path) -> str:
    pages = convert_from_path(str(pdf_path), dpi=300)
    if not pages:
        raise ValueError(f"Nenhuma pagina foi convertida do PDF: {pdf_path}")

    texts = []
    for page in pages:
        page_image = pil_to_cv2(page)
        texts.extend(image_to_text_variants(page_image))
    return "\n".join(texts)


def score_ficha_candidate(text: str, start: int, end: int) -> int:
    window_start = max(0, start - 30)
    window_end = min(len(text), end + 30)
    context = text[window_start:window_end].lower()
    score = 0

    if any(keyword in context for keyword in FICHA_KEYWORDS):
        score += 5
    if "ficha" in context:
        score += 4
    if "contrato" in context:
        score += 3
    if "comprovante" in context:
        score += 2
    if any(keyword in context for keyword in REJECTION_KEYWORDS):
        score -= 6
    if start < max(80, int(len(text) * 0.35)):
        score += 2

    return score


def score_checklist_standalone_contract_candidate(text: str, start: int, end: int) -> int:
    window_start = max(0, start - 110)
    window_end = min(len(text), end + 110)
    context = text[window_start:window_end]
    ctx_lower = context.lower()

    score = 0
    if "check" in ctx_lower and "list" in ctx_lower:
        score += 9
    if "plataforma" in ctx_lower:
        score += 7
    if "liberação" in context or "liberacao" in ctx_lower:
        score += 6
    if "libera" in ctx_lower and "plataforma" in ctx_lower:
        score += 4
    if "cliente" in ctx_lower or "obra" in ctx_lower:
        score += 2
    if any(keyword in ctx_lower for keyword in FICHA_KEYWORDS):
        score += 4
    if any(keyword in ctx_lower for keyword in REJECTION_KEYWORDS):
        score -= 9
    if start < max(140, int(len(text) * 0.42)):
        score += 2
    return score


def normalize_ficha_candidate(number_a: str, number_b: Optional[str]) -> str:
    core = number_a
    if len(core) == 5 and core.isdigit():
        core = core.zfill(6)
    if number_b:
        return f"{core}-{number_b}"
    return f"{core}-00"


def normalize_ocr_digits(raw_value: str) -> str:
    return raw_value.upper().translate(OCR_DIGIT_MAP)


def generate_ficha_candidate_variants(number_a: str, number_b: Optional[str], context_score: int) -> list[tuple[int, str]]:
    variants: list[tuple[int, str]] = []
    normalized_a = normalize_ocr_digits(number_a)
    normalized_b = normalize_ocr_digits(number_b) if number_b else None

    if len(normalized_a) == 5 and normalized_a.isdigit():
        normalized_a = normalized_a.zfill(6)

    base_candidate = normalize_ficha_candidate(normalized_a, normalized_b)
    variants.append((context_score, base_candidate))

    if len(normalized_a) == 6 and normalized_a[0] in {"6", "8", "9"}:
        corrected_leading_zero = normalize_ficha_candidate(f"0{normalized_a[1:]}", normalized_b)
        variants.append((context_score + 3, corrected_leading_zero))

    if len(normalized_a) == 6 and normalized_a[:2] in {"90", "80", "60"}:
        corrected_double_zero = normalize_ficha_candidate(f"00{normalized_a[2:]}", normalized_b)
        variants.append((context_score + 2, corrected_double_zero))

    return variants


def extract_ficha_number(text: str) -> Optional[str]:
    normalized_text = re.sub(r"\s+", " ", text)
    candidates: list[tuple[int, str]] = []

    for match in FICHA_REGEX.finditer(normalized_text):
        candidate = match.group(1)
        score = score_ficha_candidate(normalized_text, match.start(1), match.end(1)) + 6
        candidates.append((score, candidate))

    for match in FICHA_CONTEXT_REGEX.finditer(normalized_text):
        score = score_ficha_candidate(normalized_text, match.start(1), match.end(1)) + 8
        candidates.extend(generate_ficha_candidate_variants(match.group(1), match.group(2), score))

    for match in FICHA_SPLIT_REGEX.finditer(normalized_text):
        score = score_ficha_candidate(normalized_text, match.start(1), match.end(2))
        candidates.extend(generate_ficha_candidate_variants(match.group(1), match.group(2), score))

    for match in FICHA_CONTEXT_OCR_REGEX.finditer(normalized_text):
        score = score_ficha_candidate(normalized_text, match.start(1), match.end(1)) + 7
        candidates.extend(generate_ficha_candidate_variants(match.group(1), match.group(2), score))

    for match in CHECKLIST_CONTRACT_REGEX.finditer(normalized_text):
        score = score_ficha_candidate(normalized_text, match.start(1), match.end(1)) + 9
        candidates.extend(generate_ficha_candidate_variants(match.group(1), match.group(2), score))

    for match in CHECKLIST_CONTRACT_OCR_REGEX.finditer(normalized_text):
        score = score_ficha_candidate(normalized_text, match.start(1), match.end(1)) + 8
        candidates.extend(generate_ficha_candidate_variants(match.group(1), match.group(2), score))

    for match in STANDALONE_CONTRACT_DIGITS_REGEX.finditer(normalized_text):
        score = score_checklist_standalone_contract_candidate(normalized_text, match.start(1), match.end(1))
        if score >= 5:
            candidates.extend(generate_ficha_candidate_variants(match.group(1), None, score))

    if not candidates:
        return None

    candidate_scores: dict[str, int] = {}
    for score, candidate in candidates:
        current = candidate_scores.get(candidate)
        if current is None or score > current:
            candidate_scores[candidate] = score

    ranked_candidates = sorted(candidate_scores.items(), key=lambda item: item[1], reverse=True)
    best_candidate, best_score = ranked_candidates[0]

    # Heuristica para confusao comum do OCR no primeiro digito (0 vs 6/9).
    if len(best_candidate) == 9 and best_candidate[0] in {"6", "9"}:
        zero_candidate = f"0{best_candidate[1:]}"
        zero_score = candidate_scores.get(zero_candidate)
        if zero_score is not None and zero_score >= best_score - 2:
            best_candidate = zero_candidate
            best_score = zero_score

    write_scan_debug(f"extract_ficha_number -> candidates={ranked_candidates[:6]}")
    if best_score < 3:
        return None
    return best_candidate


def has_ficha_hint(text: str) -> bool:
    normalized_text = re.sub(r"\s+", " ", text)
    return bool(FICHA_HINT_REGEX.search(normalized_text))


def get_contract_number(ficha_number: str) -> str:
    prefix = ficha_number.split("-", maxsplit=1)[0]
    if len(prefix) == 5 and prefix.isdigit():
        return prefix.zfill(6)
    return prefix


def same_contract_number(a: Optional[str], b: Optional[str]) -> bool:
    if a is None or b is None:
        return False
    if not a.isdigit() or not b.isdigit():
        return a == b
    if len(a) > 6 or len(b) > 6:
        return a == b
    return int(a) == int(b)


def resolve_contract_destination_dir(matriz_dir: Path, contract_number: str) -> tuple[Path, str]:
    """Usa pasta existente (5 ou 6 digitos) se ja houver; senao cria a forma com 6 digitos."""
    raw = contract_number.strip()
    if not raw.isdigit() or len(raw) < 5 or len(raw) > 6:
        p = matriz_dir / raw
        return p, raw

    canonical_six = raw.zfill(6)
    stripped = str(int(canonical_six))

    candidates: list[tuple[float, Path]] = []
    seen_names: set[str] = set()
    for name in (canonical_six, stripped):
        if name in seen_names:
            continue
        seen_names.add(name)
        p = matriz_dir / name
        if p.is_dir() and CONTRACT_FOLDER_NAME_REGEX.fullmatch(p.name):
            try:
                candidates.append((p.stat().st_mtime, p))
            except OSError:
                continue

    if candidates:
        _ts, chosen = max(candidates, key=lambda item: item[0])
        return chosen, chosen.name

    new_path = matriz_dir / canonical_six
    return new_path, canonical_six


def move_file_to_contract_folder(file_path: Path, contract_number: str, matriz_dir: Path) -> tuple[Path, str]:
    destination_dir, folder_name = resolve_contract_destination_dir(matriz_dir, contract_number)
    destination_dir.mkdir(parents=True, exist_ok=True)

    destination_path = destination_dir / file_path.name
    counter = 1
    while destination_path.exists():
        destination_path = destination_dir / f"{file_path.stem}_{counter}{file_path.suffix}"
        counter += 1

    last_error = None
    for _ in range(FILE_READY_RETRIES):
        try:
            shutil.move(str(file_path), str(destination_path))
            return destination_path, folder_name
        except PermissionError as exc:
            last_error = exc
            time.sleep(FILE_READY_DELAY_SECONDS)
        except OSError as exc:
            last_error = exc
            time.sleep(FILE_READY_DELAY_SECONDS)

    raise RuntimeError(f"Nao foi possivel mover o arquivo para {destination_path}: {last_error}") from last_error


def move_file_to_manual_review_folder(file_path: Path, matriz_dir: Path, reason: str) -> Path:
    review_dir = matriz_dir / MANUAL_REVIEW_DIRNAME
    review_dir.mkdir(parents=True, exist_ok=True)

    destination_path = review_dir / file_path.name
    counter = 1
    while destination_path.exists():
        destination_path = review_dir / f"{file_path.stem}_{counter}{file_path.suffix}"
        counter += 1

    last_error = None
    for _ in range(FILE_READY_RETRIES):
        try:
            shutil.move(str(file_path), str(destination_path))
            break
        except PermissionError as exc:
            last_error = exc
            time.sleep(FILE_READY_DELAY_SECONDS)
        except OSError as exc:
            last_error = exc
            time.sleep(FILE_READY_DELAY_SECONDS)
    else:
        raise RuntimeError(f"Nao foi possivel mover o arquivo para revisao manual: {last_error}") from last_error

    note_path = destination_path.with_suffix(destination_path.suffix + ".txt")
    note_path.write_text(reason, encoding="utf-8")
    return destination_path


def find_recent_contract_folder(matriz_dir: Path) -> Optional[str]:
    contract_dirs: list[tuple[float, str]] = []
    for child in matriz_dir.iterdir():
        if child.is_dir() and CONTRACT_FOLDER_NAME_REGEX.fullmatch(child.name):
            try:
                modified_at = child.stat().st_mtime
            except OSError:
                continue
            contract_dirs.append((modified_at, child.name))

    if not contract_dirs:
        return None

    contract_dirs.sort(reverse=True)
    return contract_dirs[0][1]


def supported_files(directory: Path) -> Iterable[Path]:
    sortable_files: list[tuple[float, str, Path]] = []
    for file_path in directory.iterdir():
        if not file_path.is_file():
            continue
        if file_path.suffix.lower() not in IMAGE_EXTENSIONS.union(PDF_EXTENSIONS):
            continue
        try:
            created_at = file_path.stat().st_ctime
        except OSError:
            created_at = float("inf")
        sortable_files.append((created_at, file_path.name.lower(), file_path))

    for _, _, file_path in sorted(sortable_files):
        yield file_path


def find_wia_property(properties, property_id: int):
    try:
        return properties.Item(property_id)
    except Exception:
        pass

    try:
        count = int(properties.Count)
    except Exception:
        return None

    for index in range(1, count + 1):
        try:
            current = properties.Item(index)
        except Exception:
            continue
        if getattr(current, "PropertyID", None) == property_id:
            return current
    return None


def get_wia_property_value(properties, property_id: int, default=None):
    prop = find_wia_property(properties, property_id)
    if prop is None:
        return default
    try:
        return prop.Value
    except Exception:
        return default


def set_wia_property_value(properties, property_id: int, value) -> bool:
    prop = find_wia_property(properties, property_id)
    if prop is None:
        return False
    try:
        prop.Value = value
        return True
    except Exception:
        return False


def list_wia_scanners() -> list[tuple[str, str]]:
    if win32com is None:
        return []

    device_manager = win32com.client.Dispatch("WIA.DeviceManager")
    scanners: list[tuple[str, str]] = []

    for index in range(1, device_manager.DeviceInfos.Count + 1):
        device_info = device_manager.DeviceInfos.Item(index)
        if int(getattr(device_info, "Type", 0)) != WIA_SCANNER_DEVICE_TYPE:
            continue

        name = get_wia_property_value(device_info.Properties, 7, getattr(device_info, "DeviceID", f"scanner_{index}"))
        device_id = str(getattr(device_info, "DeviceID", ""))
        scanners.append((device_id, str(name)))

    return scanners


def connect_wia_device(device_id: str):
    if win32com is None:
        raise RuntimeError("pywin32 nao esta instalado. Instale com: pip install pywin32")

    device_manager = win32com.client.Dispatch("WIA.DeviceManager")
    for index in range(1, device_manager.DeviceInfos.Count + 1):
        device_info = device_manager.DeviceInfos.Item(index)
        if str(getattr(device_info, "DeviceID", "")) == device_id:
            return device_info.Connect()

    raise RuntimeError("Scanner WIA selecionado nao foi encontrado.")


def connect_wia_device_with_retry(device_id: str, status_callback=None, attempts: int = 3, delay_seconds: float = 1.2):
    last_error = None
    for attempt in range(1, attempts + 1):
        try:
            return connect_wia_device(device_id)
        except Exception as exc:
            last_error = exc
            if not is_device_busy_error(exc) or attempt == attempts:
                break
            if status_callback is not None:
                status_callback(
                    f"[SCAN] Scanner WIA ocupado. Nova tentativa {attempt + 1}/{attempts} em {delay_seconds:.1f}s."
                )
            time.sleep(delay_seconds)

    raise RuntimeError(f"Nao foi possivel conectar ao scanner WIA: {last_error}") from last_error


def scanner_supports_feeder(device) -> bool:
    capabilities = get_wia_property_value(device.Properties, WIA_DOCUMENT_HANDLING_CAPABILITIES, 0)
    try:
        return bool(int(capabilities) & WIA_FEEDER)
    except Exception:
        return False


def feeder_has_more_pages(device) -> bool:
    status = get_wia_property_value(device.Properties, WIA_DOCUMENT_HANDLING_STATUS, 0)
    try:
        return bool(int(status) & WIA_FEED_READY)
    except Exception:
        return False


def configure_scan_source(device, item) -> bool:
    supports_feeder = scanner_supports_feeder(device)
    selected_mode = WIA_FEEDER | WIA_FRONT_ONLY if supports_feeder else WIA_FLATBED

    # Alguns drivers WIA falham com "Parametro incorreto" quando certas
    # propriedades sao aplicadas no item de captura ou quando WIA_SCAN_PAGES
    # e forcado. O app tenta apenas no dispositivo e segue com fallback.
    set_wia_property_value(device.Properties, WIA_DOCUMENT_HANDLING_SELECT, selected_mode)

    return supports_feeder


def get_configured_wia_item(device_id: str, status_callback=None):
    device = connect_wia_device_with_retry(device_id, status_callback=status_callback)

    if device.Items.Count < 1:
        raise RuntimeError("O scanner WIA nao retornou itens de captura.")

    item = device.Items.Item(1)
    using_feeder = configure_scan_source(device, item)
    return device, item, using_feeder


def is_feeder_empty_error(exc: Exception) -> bool:
    message = str(exc).lower()
    return any(
        marker in message
        for marker in (
            "feeder",
            "no documents",
            "paper empty",
            "paper jam",
            "0x80210003",
            "0x80210020",
            "0x80210002",
            "-2145320958",
        )
    )


def is_device_busy_error(exc: Exception) -> bool:
    message = str(exc).lower()
    return any(
        marker in message
        for marker in (
            "ocupado",
            "busy",
            "device busy",
            "wia esta ocupado",
            "-2145320954",
            "0x80210006",
        )
    )


def is_recoverable_wia_scan_error(exc: Exception) -> bool:
    message = str(exc).lower()
    return any(
        marker in message
        for marker in (
            "parametro incorreto",
            "parâmetro incorreto",
            "device busy",
            "ocupado",
            "busy",
            "0x80210006",
            "0x8021000c",
            "0x80210002",
            "0x80210020",
            "-2147024809",
            "-2147352567",
            "-2145320958",
        )
    )


def save_wia_image(image, filename: Path) -> None:
    write_scan_debug(f"save_wia_image -> {filename}")
    image.SaveFile(str(filename))


def save_wia_image_from_binary_data(image, filename: Path) -> None:
    file_data = getattr(image, "FileData", None)
    if file_data is None:
        raise RuntimeError("Imagem WIA sem FileData para gravacao alternativa.")

    binary_data = getattr(file_data, "BinaryData", None)
    if binary_data is None:
        raise RuntimeError("Imagem WIA sem BinaryData para gravacao alternativa.")

    payload = bytes(binary_data)
    if not payload:
        raise RuntimeError("Imagem WIA retornou dados vazios.")

    write_scan_debug(f"save_wia_image_from_binary_data -> {filename} ({len(payload)} bytes)")
    filename.write_bytes(payload)


def wait_for_file_ready(file_path: Path, retries: int = FILE_READY_RETRIES, delay_seconds: float = FILE_READY_DELAY_SECONDS) -> None:
    last_error = None
    for _ in range(retries):
        try:
            if file_path.exists() and file_path.stat().st_size > 0:
                with file_path.open("rb") as file:
                    file.read(1)
                write_scan_debug(f"wait_for_file_ready -> ok {file_path} ({file_path.stat().st_size} bytes)")
                return
        except OSError as exc:
            last_error = exc
        time.sleep(delay_seconds)

    if file_path.exists():
        raise RuntimeError(f"Arquivo gerado, mas ainda indisponivel para leitura: {file_path}") from last_error
    raise RuntimeError(f"Arquivo do scan nao foi criado: {file_path}") from last_error


def is_file_stable_for_processing(file_path: Path, observed_size: Optional[int]) -> tuple[bool, Optional[int]]:
    try:
        current_size = file_path.stat().st_size
    except OSError:
        return False, observed_size

    if current_size <= 0:
        return False, current_size

    if observed_size is None:
        return False, current_size

    if current_size != observed_size:
        return False, current_size

    return True, current_size


def persist_scanned_image(image, filename: Path, status_callback) -> None:
    save_errors: list[str] = []

    try:
        write_scan_debug(f"persist_scanned_image -> SaveFile start {filename}")
        save_wia_image(image, filename)
        wait_for_file_ready(filename)
        return
    except Exception as exc:
        write_scan_debug(f"persist_scanned_image -> SaveFile error {filename}: {exc}")
        save_errors.append(f"SaveFile: {exc}")
        try:
            if filename.exists():
                filename.unlink()
        except OSError:
            pass

    try:
        write_scan_debug(f"persist_scanned_image -> BinaryData start {filename}")
        save_wia_image_from_binary_data(image, filename)
        wait_for_file_ready(filename)
        status_callback(f"[SCAN] Gravacao alternativa aplicada com sucesso: {filename.name}")
        return
    except Exception as exc:
        write_scan_debug(f"persist_scanned_image -> BinaryData error {filename}: {exc}")
        save_errors.append(f"BinaryData: {exc}")
        try:
            if filename.exists():
                filename.unlink()
        except OSError:
            pass

    raise RuntimeError(" | ".join(save_errors))


def transfer_wia_image(item, status_callback):
    try:
        write_scan_debug("transfer_wia_image -> item.Transfer(JPEG)")
        return item.Transfer(WIA_FORMAT_JPEG)
    except Exception as exc:
        write_scan_debug(f"transfer_wia_image -> JPEG error: {exc}")
        try:
            write_scan_debug("transfer_wia_image -> item.Transfer(default)")
            image = item.Transfer()
            status_callback("[SCAN] Driver WIA rejeitou JPEG explicito; usando formato padrao do scanner.")
            return image
        except Exception:
            raise exc


def transfer_wia_image_with_common_dialog(item, status_callback):
    common_dialog = win32com.client.Dispatch("WIA.CommonDialog")
    try:
        write_scan_debug("transfer_wia_image_with_common_dialog -> ShowTransfer(JPEG)")
        return common_dialog.ShowTransfer(item, WIA_FORMAT_JPEG, False)
    except Exception as exc:
        write_scan_debug(f"transfer_wia_image_with_common_dialog -> JPEG error: {exc}")
        try:
            write_scan_debug("transfer_wia_image_with_common_dialog -> ShowTransfer(default)")
            image = common_dialog.ShowTransfer(item, "{}", False)
            status_callback("[SCAN] CommonDialog rejeitou JPEG explicito; usando formato padrao do scanner.")
            return image
        except Exception:
            raise exc


def build_scan_filename(target_dir: Path, page_number: int, compat_mode: bool = False) -> Path:
    suffix = "_compat" if compat_mode else ""
    return target_dir / f"scan_{datetime.now():%Y%m%d_%H%M%S_%f}{suffix}_{page_number:03d}.jpg"


def scan_pages_with_common_dialog(target_dir: Path, status_callback) -> list[Path]:
    common_dialog = win32com.client.Dispatch("WIA.CommonDialog")
    status_callback("[SCAN] Modo de compatibilidade ativado. O Windows pode abrir a interface nativa do scanner.")

    scanned_files: list[Path] = []
    page_number = 1

    while True:
        try:
            write_scan_debug(f"scan_pages_with_common_dialog -> page {page_number} acquire start")
            image = common_dialog.ShowAcquireImage(
                WIA_SCANNER_DEVICE_TYPE,
                0,
                0,
                WIA_FORMAT_JPEG,
                False,
                True,
                False,
            )
        except Exception as exc:
            write_scan_debug(f"scan_pages_with_common_dialog -> page {page_number} acquire error: {exc}")
            if scanned_files:
                status_callback("[SCAN] Captura manual encerrada. As paginas ja digitalizadas serao processadas.")
                break
            raise RuntimeError(f"Falha no modo de compatibilidade do Windows: {exc}") from exc

        if image is None:
            write_scan_debug(f"scan_pages_with_common_dialog -> page {page_number} acquire returned None")
            if scanned_files:
                status_callback("[SCAN] Captura manual encerrada. As paginas ja digitalizadas serao processadas.")
                break
            raise RuntimeError("A captura foi cancelada na interface nativa do Windows.")

        filename = build_scan_filename(target_dir, page_number, compat_mode=True)
        persist_scanned_image(image, filename, status_callback)
        scanned_files.append(filename)
        write_scan_debug(f"scan_pages_with_common_dialog -> page {page_number} saved {filename}")
        status_callback(f"[SCAN] Pagina capturada em modo de compatibilidade: {filename.name}")
        page_number += 1

    return scanned_files


def scan_until_stopped(device_id: str, target_dir: Path, status_callback, should_stop) -> list[Path]:
    target_dir.mkdir(parents=True, exist_ok=True)
    if should_stop():
        return []

    try:
        scanned_files = scan_pages_from_wia_device(device_id, target_dir, status_callback)
    except Exception as exc:
        raise RuntimeError(f"Falha no scanner: {exc}") from exc

    if scanned_files:
        status_callback(
            f"[SCAN] Lote concluido com {len(scanned_files)} pagina(s). O processamento sera iniciado agora."
        )

    return scanned_files


def scan_single_document_from_wia_device(device_id: str, target_dir: Path, status_callback) -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    device, item, using_feeder = get_configured_wia_item(device_id, status_callback=status_callback)
    write_scan_debug(f"scan_single_document_from_wia_device -> using_feeder={using_feeder}")

    try:
        if win32com is not None:
            image = transfer_wia_image_with_common_dialog(item, status_callback)
        else:
            image = transfer_wia_image(item, status_callback)
    except Exception as exc:
        write_scan_debug(f"scan_single_document_from_wia_device -> transfer error: {exc}")
        if is_device_busy_error(exc):
            time.sleep(1.0)
            device, item, _ = get_configured_wia_item(device_id, status_callback=status_callback)
            image = transfer_wia_image(item, status_callback)
        else:
            raise RuntimeError(f"Falha ao capturar o documento atual: {exc}") from exc

    filename = build_scan_filename(target_dir, 1)
    persist_scanned_image(image, filename, status_callback)
    write_scan_debug(f"scan_single_document_from_wia_device -> saved {filename}")
    return filename


def scan_documents_until_stopped(device_id: str, target_dir: Path, status_callback, should_stop) -> list[Path]:
    target_dir.mkdir(parents=True, exist_ok=True)
    scanned_files: list[Path] = []
    document_number = 1

    while not should_stop():
        status_callback(f"[SCAN] Capturando documento {document_number}...")
        write_scan_debug(f"scan_documents_until_stopped -> document {document_number} start")
        try:
            file_path = scan_single_document_from_wia_device(device_id, target_dir, status_callback)
        except Exception as exc:
            if scanned_files and (is_device_busy_error(exc) or is_recoverable_wia_scan_error(exc)):
                status_callback("[SCAN] Scanner ocupado ao iniciar novo ciclo. Aguardando nova tentativa...")
                time.sleep(1.1)
                continue
            if scanned_files:
                status_callback(
                    "[AVISO] A captura foi interrompida apos alguns documentos. O processamento seguira com o que ja foi capturado."
                )
                break
            raise RuntimeError(f"Falha no scanner: {exc}") from exc

        scanned_files.append(file_path)
        status_callback(
            f"[SCAN] Documento {document_number} capturado: {file_path.name}. Insira o proximo ou clique em 'Parar scanner'."
        )
        write_scan_debug(f"scan_documents_until_stopped -> document {document_number} saved {file_path}")
        document_number += 1
        time.sleep(0.9)

    return scanned_files


def scan_pages_from_wia_device(device_id: str, target_dir: Path, status_callback) -> list[Path]:
    target_dir.mkdir(parents=True, exist_ok=True)
    try:
        write_scan_debug(f"scan_pages_from_wia_device -> device_id={device_id}")
        device, item, using_feeder = get_configured_wia_item(device_id, status_callback=status_callback)
        optimistic_multi_page = True
        write_scan_debug(f"scan_pages_from_wia_device -> initial using_feeder={using_feeder}")

        scanned_files: list[Path] = []
        page_number = 1

        while page_number <= WIA_MAX_BATCH_PAGES:
            try:
                write_scan_debug(
                    f"scan_pages_from_wia_device -> page {page_number} transfer start using_feeder={using_feeder} optimistic={optimistic_multi_page}"
                )
                if (using_feeder or optimistic_multi_page) and win32com is not None:
                    image = transfer_wia_image_with_common_dialog(item, status_callback)
                else:
                    image = transfer_wia_image(item, status_callback)
            except Exception as exc:
                write_scan_debug(f"scan_pages_from_wia_device -> page {page_number} transfer error: {exc}")
                if scanned_files and (using_feeder or optimistic_multi_page) and is_feeder_empty_error(exc):
                    status_callback("[SCAN] Alimentador sem mais folhas. Encerrando lote.")
                    break
                if scanned_files and (using_feeder or optimistic_multi_page) and is_recoverable_wia_scan_error(exc):
                    status_callback(
                        "[AVISO] O lote do scanner falhou apos capturar algumas paginas. "
                        "As paginas ja capturadas serao aproveitadas; insira o restante manualmente e execute outro scan se necessario."
                    )
                    break
                raise RuntimeError(f"Falha ao scanear pelo dispositivo WIA: {exc}") from exc

            filename = build_scan_filename(target_dir, page_number)
            try:
                persist_scanned_image(image, filename, status_callback)
            except Exception as exc:
                write_scan_debug(f"scan_pages_from_wia_device -> page {page_number} save error: {exc}")
                if using_feeder and is_feeder_empty_error(exc):
                    status_callback("[SCAN] Scanner retornou fim de lote durante a gravacao da pagina.")
                    break
                raise RuntimeError(f"Falha ao gravar a pagina {page_number}: {exc}") from exc

            scanned_files.append(filename)
            write_scan_debug(f"scan_pages_from_wia_device -> page {page_number} saved {filename}")
            status_callback(
                f"[SCAN] Pagina {page_number} capturada: {filename.name} ({filename.stat().st_size} bytes)"
            )
            page_number += 1

            time.sleep(0.35)
            try:
                device, item, refreshed_using_feeder = get_configured_wia_item(
                    device_id,
                    status_callback=status_callback,
                )
                using_feeder = using_feeder or refreshed_using_feeder
                write_scan_debug(
                    f"scan_pages_from_wia_device -> page {page_number} prepared refreshed_using_feeder={refreshed_using_feeder}"
                )
                status_callback(f"[SCAN] Preparando captura da pagina {page_number} no alimentador...")
            except Exception as exc:
                write_scan_debug(f"scan_pages_from_wia_device -> page {page_number} prepare error: {exc}")
                if feeder_has_more_pages(device):
                    raise RuntimeError(f"Falha ao preparar a proxima pagina no ADF: {exc}") from exc
                status_callback("[SCAN] Alimentador sem mais folhas. Encerrando lote.")
                break

        if page_number > WIA_MAX_BATCH_PAGES:
            status_callback(
                f"[AVISO] Limite de seguranca atingido em {WIA_MAX_BATCH_PAGES} paginas. O lote sera encerrado."
            )

        return scanned_files
    except Exception as exc:
        if is_recoverable_wia_scan_error(exc):
            status_callback(
                "[AVISO] O driver WIA falhou no modo em lote. Tentando modo de compatibilidade para captura manual."
            )
            try:
                return scan_pages_with_common_dialog(target_dir, status_callback)
            except Exception:
                raise RuntimeError(
                    "Falha no scanner em modo lote. Tente alimentar manualmente e executar nova captura."
                ) from exc
        if is_device_busy_error(exc):
            status_callback(
                "[AVISO] O scanner retornou 'dispositivo ocupado'. Tentando o modo de compatibilidade do Windows."
            )
            return scan_pages_with_common_dialog(target_dir, status_callback)
        raise


def extract_text_from_file(file_path: Path) -> str:
    wait_for_file_ready(file_path)
    suffix = file_path.suffix.lower()
    if suffix in IMAGE_EXTENSIONS:
        return ocr_image_file(file_path)
    if suffix in PDF_EXTENSIONS:
        return ocr_pdf_file(file_path)
    raise ValueError(f"Formato nao suportado: {file_path.suffix}")


def process_file(file_path: Path, matriz_dir: Path, last_contract_number: Optional[str]) -> tuple[bool, str, Optional[str]]:
    try:
        extracted_text = extract_text_from_file(file_path)
        ficha_number = extract_ficha_number(extracted_text)

        if ficha_number:
            contract_number = get_contract_number(ficha_number)
            destination, folder_name = move_file_to_contract_folder(file_path, contract_number, matriz_dir)
            return True, f"[OK] {file_path.name} -> contrato {folder_name} -> {destination}", folder_name

        if has_ficha_hint(extracted_text):
            logging.warning("Ficha detectada, mas numero nao lido no arquivo: %s", file_path.name)
            review_reason = "Ficha detectada, mas o numero nao foi lido pelo OCR."
            destination = move_file_to_manual_review_folder(file_path, matriz_dir, review_reason)
            return (
                False,
                f"[REVISAR] {file_path.name} -> ficha detectada sem numero legivel -> {destination}",
                last_contract_number,
            )

        if last_contract_number:
            destination, folder_name = move_file_to_contract_folder(file_path, last_contract_number, matriz_dir)
            return (
                True,
                f"[COMPLEMENTO] {file_path.name} -> contrato {folder_name} -> {destination}",
                folder_name,
            )

        text_preview = re.sub(r"\s+", " ", extracted_text).strip()[:140] or "sem texto legivel"
        review_reason = (
            "Documento sem ficha valida e sem contrato valido imediatamente anterior no lote atual.\n"
            f"Trecho lido pelo OCR: {text_preview}"
        )
        logging.warning("Documento sem ficha valida e sem contrato anterior para vincular: %s", file_path.name)
        destination = move_file_to_manual_review_folder(file_path, matriz_dir, review_reason)
        return (
            False,
            f"[REVISAR] {file_path.name} -> sem ficha valida e sem contrato imediatamente anterior -> {destination}",
            last_contract_number,
        )
    except Exception as exc:
        logging.exception("Erro ao processar %s: %s", file_path.name, exc)
        destination = move_file_to_manual_review_folder(file_path, matriz_dir, f"Erro ao processar arquivo:\n{exc}")
        return False, f"[ERRO] {file_path.name}: {exc} -> enviado para {destination}", last_contract_number


def process_scanned_session_files(files: list[Path], matriz_dir: Path, status_callback, finish_callback) -> None:
    try:
        configure_tesseract()
        ensure_directories(matriz_dir)
        log_file = setup_logging()

        if not files:
            finish_callback("Nenhum documento foi capturado.")
            return

        processed_count = 0
        failed_count = 0
        failure_messages: list[str] = []
        current_contract_number: Optional[str] = None
        current_contract_document_count = 0

        for file_path in files:
            try:
                extracted_text = extract_text_from_file(file_path)
                ficha_number = extract_ficha_number(extracted_text)

                if ficha_number:
                    contract_number = get_contract_number(ficha_number)
                    if not same_contract_number(contract_number, current_contract_number):
                        current_contract_document_count = 0

                    current_contract_document_count += 1

                    destination, folder_name = move_file_to_contract_folder(file_path, contract_number, matriz_dir)
                    current_contract_number = folder_name
                    message = f"[OK] {file_path.name} -> contrato {folder_name} -> {destination}"
                    processed_count += 1
                    status_callback(message)
                    continue

                if current_contract_number:
                    current_contract_document_count += 1
                    destination, folder_name = move_file_to_contract_folder(file_path, current_contract_number, matriz_dir)
                    current_contract_number = folder_name
                    message = (
                        f"[COMPLEMENTO] {file_path.name} -> contrato {folder_name} -> {destination}"
                    )
                    processed_count += 1
                    status_callback(message)
                    continue

                text_preview = re.sub(r"\s+", " ", extracted_text).strip()[:140] or "sem texto legivel"
                review_reason = (
                    "Documento sem ficha valida e sem contrato aberto na sessao.\n"
                    f"Trecho lido pelo OCR: {text_preview}"
                )
                destination = move_file_to_manual_review_folder(file_path, matriz_dir, review_reason)
                failed_count += 1
                message = (
                    f"[REVISAR] {file_path.name} -> sem ficha valida e sem contrato aberto na sessao -> {destination}"
                )
                failure_messages.append(message)
                status_callback(message)
            except Exception as exc:
                logging.exception("Erro ao processar o arquivo da sessao %s: %s", file_path.name, exc)
                destination = move_file_to_manual_review_folder(
                    file_path,
                    matriz_dir,
                    f"Erro ao processar arquivo da sessao:\n{exc}",
                )
                failed_count += 1
                message = f"[ERRO] {file_path.name}: {exc} -> enviado para {destination}"
                failure_messages.append(message)
                status_callback(message)

        failure_summary = ""
        if failure_messages:
            failure_summary = "\nPendencias:\n" + "\n".join(failure_messages[:5])
            if len(failure_messages) > 5:
                failure_summary += f"\n... e mais {len(failure_messages) - 5} item(ns)."

        finish_callback(
            "Processamento da sessao concluido.\n"
            f"Sucesso: {processed_count}\n"
            f"Falhas: {failed_count}\n"
            f"Log: {log_file}"
            f"{failure_summary}"
        )
    except Exception as exc:
        finish_callback(f"Falha geral no processamento da sessao: {exc}")


def process_incoming_file(
    file_path: Path,
    matriz_dir: Path,
    current_contract_number: Optional[str],
    current_contract_document_count: int,
    current_fallback_count: int,
    status_callback,
) -> tuple[bool, str, Optional[str], int, int, bool]:
    extracted_text = extract_text_from_file(file_path)
    ficha_number = extract_ficha_number(extracted_text)

    if ficha_number:
        contract_number = get_contract_number(ficha_number)
        if not same_contract_number(contract_number, current_contract_number):
            current_contract_document_count = 0
        current_fallback_count = 0
        current_contract_document_count += 1

        destination, folder_name = move_file_to_contract_folder(file_path, contract_number, matriz_dir)
        message = f"[OK] {file_path.name} -> contrato {folder_name} -> {destination}"
        return True, message, folder_name, current_contract_document_count, current_fallback_count, False

    if current_contract_number:
        current_fallback_count += 1
        current_contract_document_count += 1

        destination, folder_name = move_file_to_contract_folder(file_path, current_contract_number, matriz_dir)
        message = (
            f"[FALLBACK] {file_path.name} -> ficha nao lida; enviado para o ultimo contrato {folder_name} "
            f"({current_fallback_count} consecutivo(s) sem ficha) -> {destination}"
        )
        return True, message, folder_name, current_contract_document_count, current_fallback_count, False

    text_preview = re.sub(r"\s+", " ", extracted_text).strip()[:140] or "sem texto legivel"
    review_reason = (
        "Documento sem ficha valida e sem ultima pasta de contrato disponivel no monitoramento.\n"
        f"Trecho lido pelo OCR: {text_preview}"
    )
    destination = move_file_to_manual_review_folder(file_path, matriz_dir, review_reason)
    message = f"[REVISAR] {file_path.name} -> sem ficha valida e sem ultimo contrato disponivel -> {destination}"
    return False, message, current_contract_number, current_contract_document_count, current_fallback_count, False


def run_realtime_monitor(input_dir: Path, matriz_dir: Path, status_callback, finish_callback, should_stop, metrics_callback) -> None:
    try:
        configure_tesseract()
        input_dir.mkdir(parents=True, exist_ok=True)
        ensure_directories(matriz_dir)
        log_file = setup_logging()

        current_contract_number: Optional[str] = None
        current_contract_document_count = 0
        current_fallback_count = 0
        processed_count = 0
        failed_count = 0
        file_stability: dict[str, tuple[Optional[int], int]] = {}

        status_callback(f"[MONITOR] Pasta monitorada: {input_dir}")

        while True:
            files = list(supported_files(input_dir))
            pending_count = len(files)
            metrics_callback(processed_count, pending_count)

            if should_stop() and pending_count == 0 and not file_stability:
                break

            if not files:
                time.sleep(MONITOR_POLL_SECONDS)
                continue

            processed_any_file = False
            for file_path in files:
                file_key = str(file_path)
                previous_size, stable_polls = file_stability.get(file_key, (None, 0))
                is_stable, current_size = is_file_stable_for_processing(file_path, previous_size)

                if not is_stable:
                    if current_size is not None and current_size == previous_size and current_size > 0:
                        stable_polls += 1
                    else:
                        stable_polls = 1 if current_size and current_size > 0 else 0
                    file_stability[file_key] = (current_size, stable_polls)
                    continue

                stable_polls += 1
                file_stability[file_key] = (current_size, stable_polls)
                if stable_polls < MONITOR_STABLE_POLLS:
                    continue

                try:
                    success, message, current_contract_number, current_contract_document_count, current_fallback_count, should_finish = process_incoming_file(
                        file_path,
                        matriz_dir,
                        current_contract_number,
                        current_contract_document_count,
                        current_fallback_count,
                        status_callback,
                    )
                    status_callback(message)
                    processed_any_file = True
                    if success:
                        processed_count += 1
                    else:
                        failed_count += 1
                        logging.warning(message)

                    if should_finish:
                        metrics_callback(processed_count, len(list(supported_files(input_dir))))
                        finish_callback(
                            "Monitoramento encerrado por limite de seguranca.\n"
                            f"Sucesso: {processed_count}\n"
                            f"Falhas: {failed_count}\n"
                            f"Log: {log_file}\n"
                            f"{message}"
                        )
                        return
                    file_stability.pop(file_key, None)
                    metrics_callback(processed_count, max(len(files) - 1, 0))
                except Exception as exc:
                    transient_error = isinstance(exc, (RuntimeError, ValueError)) and (
                        "indisponivel para leitura" in str(exc).lower()
                        or "falha ao abrir a imagem" in str(exc).lower()
                    )
                    if transient_error:
                        write_scan_debug(f"run_realtime_monitor -> arquivo ainda instavel {file_path}: {exc}")
                        continue

                    logging.exception("Erro no monitoramento ao processar %s: %s", file_path.name, exc)
                    destination = move_file_to_manual_review_folder(
                        file_path,
                        matriz_dir,
                        f"Erro no monitoramento em tempo real:\n{exc}",
                    )
                    file_stability.pop(file_key, None)
                    failed_count += 1
                    processed_any_file = True
                    status_callback(f"[ERRO] {file_path.name}: {exc} -> enviado para {destination}")
                    metrics_callback(processed_count, max(len(files) - 1, 0))

            if not processed_any_file:
                time.sleep(MONITOR_POLL_SECONDS)
            else:
                time.sleep(0.2)

        finish_callback(
            "Monitoramento finalizado.\n"
            f"Sucesso: {processed_count}\n"
            f"Falhas: {failed_count}\n"
            f"Log: {log_file}"
        )
    except Exception as exc:
        finish_callback(f"Falha no monitoramento: {exc}")


def ensure_directories(matriz_dir: Path) -> None:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    SCAN_DEBUG_DIR.mkdir(parents=True, exist_ok=True)


def archive_scanned_files_for_debug(files: list[Path], status_callback) -> None:
    if not files:
        return

    batch_dir = SCAN_DEBUG_DIR / f"batch_{datetime.now():%Y%m%d_%H%M%S_%f}"
    batch_dir.mkdir(parents=True, exist_ok=True)

    for file_path in files:
        if not file_path.exists():
            write_scan_debug(f"archive_scanned_files_for_debug -> missing {file_path}")
            continue
        debug_copy = batch_dir / file_path.name
        shutil.copy2(file_path, debug_copy)
        write_scan_debug(f"archive_scanned_files_for_debug -> copied {file_path} to {debug_copy}")

    status_callback(f"[SCAN] Copia bruta do lote salva em: {batch_dir}")


def process_files(files: list[Path], matriz_dir: Path, status_callback, finish_callback) -> None:
    try:
        configure_tesseract()
        ensure_directories(matriz_dir)
        log_file = setup_logging()

        if not files:
            finish_callback(
                "Nenhum arquivo PDF/JPG/PNG encontrado em "
                f"{INPUT_DIR}\n\nLog: {log_file}"
            )
            return

        processed_count = 0
        failed_count = 0
        failure_messages: list[str] = []

        last_contract_number: Optional[str] = None

        for file_path in files:
            success, message, detected_contract_number = process_file(file_path, matriz_dir, last_contract_number)
            status_callback(message)
            if success:
                processed_count += 1
                last_contract_number = detected_contract_number
            else:
                failed_count += 1
                failure_messages.append(message)

        failure_summary = ""
        if failure_messages:
            failure_summary = "\nPendencias:\n" + "\n".join(failure_messages[:5])
            if len(failure_messages) > 5:
                failure_summary += f"\n... e mais {len(failure_messages) - 5} item(ns)."

        finish_callback(
            "Processamento concluido.\n"
            f"Sucesso: {processed_count}\n"
            f"Falhas: {failed_count}\n"
            f"Log: {log_file}"
            f"{failure_summary}"
        )
    except Exception as exc:
        finish_callback(f"Falha geral no processamento: {exc}")


def run_processing(matriz_dir: Path, status_callback, finish_callback) -> None:
    process_files(list(supported_files(INPUT_DIR)), matriz_dir, status_callback, finish_callback)


def run_scanning_session(device_id: str, matriz_dir: Path, status_callback, finish_callback, should_stop) -> None:
    initialized_com = False
    try:
        if pythoncom is not None:
            pythoncom.CoInitialize()
            initialized_com = True
        ensure_directories(matriz_dir)
        scanned_files = scan_documents_until_stopped(device_id, INPUT_DIR, status_callback, should_stop)
        if not scanned_files:
            finish_callback("Nenhuma pagina foi digitalizada.")
            return
        write_scan_debug(f"run_scanning_and_processing -> scanned_files={len(scanned_files)}")
        archive_scanned_files_for_debug(scanned_files, status_callback)
        status_callback("[SCAN] Captura encerrada. Iniciando processamento da sessao na ordem de scaneamento...")
        process_scanned_session_files(scanned_files, matriz_dir, status_callback, finish_callback)
    except Exception as exc:
        finish_callback(f"Falha no scanner: {exc}")
    finally:
        if initialized_com:
            pythoncom.CoUninitialize()


class OrganizerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Organizador de Fichas")
        self.root.geometry("860x700")
        self.root.minsize(760, 520)
        if APP_ICON_FILE.exists():
            try:
                self.root.iconbitmap(str(APP_ICON_FILE))
            except Exception:
                pass

        self.config = load_config()
        self.matriz_var = tk.StringVar(value=self.config.get("matriz_path", str(DEFAULT_OUTPUT_DIR)))
        self.input_var = tk.StringVar(value=self.config.get("input_path", str(INPUT_DIR)))
        self.scanner_var = tk.StringVar()
        self.scanner_info_var = tk.StringVar(value="Atualize a lista de scanners para verificar os recursos do dispositivo.")
        self.contract_search_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Selecione as pastas e inicie o monitoramento.")
        self.processed_count_var = tk.StringVar(value="Processados: 0")
        self.pending_var = tk.StringVar(value="Pendentes: 0")
        self.open_matriz_on_finish_var = tk.BooleanVar(value=True)
        self.processing = False
        self.scanning = False
        self.scan_stop_event = threading.Event()
        self.stop_when_safe = False
        self.scanner_devices: list[tuple[str, str]] = []

        self.configure_style()
        self.build_ui()
        self.refresh_scanner_list(initial=True)

    def configure_style(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")

        self.root.configure(bg="#f4efe7")
        style.configure("App.TFrame", background="#f4efe7")
        style.configure("Card.TFrame", background="#fffaf3", relief="flat")
        style.configure("Title.TLabel", background="#f4efe7", foreground="#2f241f", font=("Segoe UI Semibold", 22))
        style.configure("Subtitle.TLabel", background="#f4efe7", foreground="#6d5f57", font=("Segoe UI", 10))
        style.configure("Section.TLabel", background="#fffaf3", foreground="#2f241f", font=("Segoe UI Semibold", 11))
        style.configure("Body.TLabel", background="#fffaf3", foreground="#473a34", font=("Segoe UI", 10))
        style.configure("Primary.TButton", font=("Segoe UI Semibold", 10), padding=(16, 10), background="#b6653a", foreground="#ffffff")
        style.map("Primary.TButton", background=[("active", "#8f4f2d"), ("disabled", "#d9b7a4")])
        style.configure("Secondary.TButton", font=("Segoe UI", 10), padding=(12, 10), background="#e8d8c7", foreground="#2f241f")
        style.map("Secondary.TButton", background=[("active", "#dac3ae")])
        style.configure("Path.TEntry", fieldbackground="#fffdf9", padding=8)

    def build_ui(self) -> None:
        outer_frame = ttk.Frame(self.root, style="App.TFrame")
        outer_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(
            outer_frame,
            bg="#f4efe7",
            highlightthickness=0,
            bd=0,
        )
        scrollbar = ttk.Scrollbar(outer_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style="App.TFrame", padding=24)

        scrollable_frame.bind(
            "<Configure>",
            lambda event: canvas.configure(scrollregion=canvas.bbox("all")),
        )

        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        def resize_scrollable_frame(event) -> None:
            canvas.itemconfigure(canvas_window, width=event.width)

        canvas.bind("<Configure>", resize_scrollable_frame)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def on_mousewheel(event) -> None:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", on_mousewheel)

        main_frame = scrollable_frame

        header = ttk.Frame(main_frame, style="App.TFrame")
        header.pack(fill="x", pady=(0, 18))

        ttk.Label(header, text="Organizador de documentos", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="OCR para PDFs e imagens com organizacao automatica por contrato.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        card = ttk.Frame(main_frame, style="Card.TFrame", padding=22)
        card.pack(fill="x")

        ttk.Label(card, text="Pasta matriz de saida", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            card,
            text="As pastas dos contratos serao criadas dentro deste caminho. A escolha fica salva para os proximos usos.",
            style="Body.TLabel",
            wraplength=760,
        ).pack(anchor="w", pady=(4, 14))

        path_row = ttk.Frame(card, style="Card.TFrame")
        path_row.pack(fill="x")

        path_entry = ttk.Entry(path_row, textvariable=self.matriz_var, state="readonly", style="Path.TEntry")
        path_entry.pack(side="left", fill="x", expand=True)

        ttk.Button(
            path_row,
            text="Definir pasta",
            style="Secondary.TButton",
            command=self.choose_matriz_path,
        ).pack(side="left", padx=(12, 0))

        ttk.Button(
            path_row,
            text="Alterar pasta",
            style="Secondary.TButton",
            command=self.change_matriz_path,
        ).pack(side="left", padx=(12, 0))

        input_card = ttk.Frame(main_frame, style="Card.TFrame", padding=22)
        input_card.pack(fill="x", pady=(18, 0))

        ttk.Label(input_card, text="Pasta monitorada", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            input_card,
            text="A aplicacao acompanha esta pasta em tempo real. Sempre que um novo arquivo chegar, ele sera lido, movido para a pasta do contrato e removido da entrada.",
            style="Body.TLabel",
            wraplength=760,
        ).pack(anchor="w", pady=(4, 14))

        input_row = ttk.Frame(input_card, style="Card.TFrame")
        input_row.pack(fill="x")

        input_entry = ttk.Entry(input_row, textvariable=self.input_var, state="readonly", style="Path.TEntry")
        input_entry.pack(side="left", fill="x", expand=True)

        ttk.Button(
            input_row,
            text="Definir entrada",
            style="Secondary.TButton",
            command=self.choose_input_path,
        ).pack(side="left", padx=(12, 0))

        scanner_card = ttk.Frame(main_frame, style="Card.TFrame", padding=22)
        scanner_card.pack(fill="x", pady=(18, 0))

        ttk.Label(scanner_card, text="Scanner", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            scanner_card,
            text="Selecione um scanner WIA para digitalizar direto no aplicativo. Em scanners com ADF, o app tenta capturar em lote sem confirmar pagina por pagina.",
            style="Body.TLabel",
            wraplength=760,
        ).pack(anchor="w", pady=(4, 14))

        scanner_row = ttk.Frame(scanner_card, style="Card.TFrame")
        scanner_row.pack(fill="x")

        self.scanner_combo = ttk.Combobox(scanner_row, textvariable=self.scanner_var, state="readonly")
        self.scanner_combo.pack(side="left", fill="x", expand=True)

        ttk.Button(
            scanner_row,
            text="Atualizar scanners",
            style="Secondary.TButton",
            command=self.refresh_scanner_list,
        ).pack(side="left", padx=(12, 0))

        ttk.Button(
            scanner_row,
            text="Diagnosticar scanner",
            style="Secondary.TButton",
            command=self.show_scanner_diagnostics,
        ).pack(side="left", padx=(12, 0))

        ttk.Label(
            scanner_card,
            textvariable=self.scanner_info_var,
            style="Body.TLabel",
            wraplength=760,
            justify="left",
        ).pack(anchor="w", pady=(12, 0))

        search_card = ttk.Frame(main_frame, style="Card.TFrame", padding=22)
        search_card.pack(fill="x", pady=(18, 0))

        ttk.Label(search_card, text="Pesquisar contrato", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            search_card,
            text="Digite o numero do contrato para localizar e abrir a pasta correspondente dentro da matriz.",
            style="Body.TLabel",
            wraplength=760,
        ).pack(anchor="w", pady=(4, 14))

        search_row = ttk.Frame(search_card, style="Card.TFrame")
        search_row.pack(fill="x")

        self.contract_search_entry = ttk.Entry(search_row, textvariable=self.contract_search_var, style="Path.TEntry")
        self.contract_search_entry.pack(side="left", fill="x", expand=True)
        self.contract_search_entry.bind("<Return>", lambda event: self.search_contract_folder())

        ttk.Button(
            search_row,
            text="Pesquisar",
            style="Secondary.TButton",
            command=self.search_contract_folder,
        ).pack(side="left", padx=(12, 0))

        info_card = ttk.Frame(main_frame, style="Card.TFrame", padding=22)
        info_card.pack(fill="x", padx=24, pady=(18, 0))

        ttk.Label(info_card, text="Fluxo de trabalho", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            info_card,
            text=(
                "1. Defina a pasta matriz e a pasta monitorada antes de iniciar\n"
                "2. Clique em Iniciar monitoramento\n"
                "3. Envie os arquivos escaneados para a pasta monitorada\n"
                "4. O app identifica a ficha, cria a pasta do contrato e move os proximos documentos em tempo real"
            ),
            style="Body.TLabel",
            justify="left",
        ).pack(anchor="w", pady=(6, 0))

        action_row = ttk.Frame(main_frame, style="App.TFrame")
        action_row.pack(fill="x", padx=24, pady=(18, 0))

        self.process_button = ttk.Button(
            action_row,
            text="Iniciar monitoramento",
            style="Primary.TButton",
            command=self.start_monitoring,
        )
        self.process_button.pack(side="left")

        self.scan_button = ttk.Button(
            action_row,
            text="Scanear no scanner",
            style="Secondary.TButton",
            command=self.start_scanning,
        )
        self.scan_button.pack(side="left", padx=(12, 0))

        self.stop_scan_button = ttk.Button(
            action_row,
            text="Parar monitoramento",
            style="Secondary.TButton",
            command=self.stop_monitoring,
        )
        self.stop_scan_button.pack(side="left", padx=(12, 0))
        self.stop_scan_button.state(["disabled"])

        ttk.Checkbutton(
            action_row,
            text="Abrir pasta matriz ao concluir",
            variable=self.open_matriz_on_finish_var,
        ).pack(side="left", padx=(16, 0))

        ttk.Label(action_row, textvariable=self.status_var, style="Subtitle.TLabel").pack(side="left", padx=(16, 0))
        ttk.Label(action_row, textvariable=self.processed_count_var, style="Subtitle.TLabel").pack(side="left", padx=(16, 0))
        ttk.Label(action_row, textvariable=self.pending_var, style="Subtitle.TLabel").pack(side="left", padx=(16, 0))

        log_card = ttk.Frame(main_frame, style="Card.TFrame", padding=0)
        log_card.pack(fill="both", expand=True, padx=24, pady=(18, 24))

        ttk.Label(log_card, text="Status", style="Section.TLabel").pack(anchor="w", padx=22, pady=(18, 10))

        self.log_text = tk.Text(
            log_card,
            height=18,
            wrap="word",
            bg="#fffdf9",
            fg="#2f241f",
            relief="flat",
            font=("Consolas", 10),
            padx=16,
            pady=14,
        )
        self.log_text.pack(fill="both", expand=True, padx=22, pady=(0, 22))
        self.log_text.insert("end", "Aguardando processamento...\n")
        self.log_text.configure(state="disabled")

    def append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def choose_directory_dialog(self, title: str, initial_dir: str) -> Optional[str]:
        selected = filedialog.askdirectory(
            title=title,
            initialdir=initial_dir or str(BASE_DIR),
        )
        return selected or None

    def persist_matriz_path(self, selected_path: str) -> None:
        self.config["matriz_path"] = selected_path
        save_config(self.config)
        self.matriz_var.set(selected_path)
        self.status_var.set("Pasta matriz salva.")

    def persist_input_path(self, selected_path: str) -> None:
        self.config["input_path"] = selected_path
        save_config(self.config)
        self.input_var.set(selected_path)
        self.status_var.set("Pasta monitorada salva.")

    def persist_scanner_device(self, device_id: str) -> None:
        self.config["scanner_device_id"] = device_id
        save_config(self.config)

    def show_scanner_diagnostics(self) -> None:
        diagnostics = [
            f"Python atual: {sys.executable}",
            f"Versao do Python: {sys.version.split()[0]}",
            f"pywin32 carregado: {'sim' if win32com is not None else 'nao'}",
        ]

        if win32com is None:
            diagnostics.append("Instale o pywin32 no mesmo Python usado para abrir o app.")
        else:
            try:
                devices = list_wia_scanners()
                diagnostics.append(f"Scanners WIA encontrados: {len(devices)}")
                diagnostics.extend([f"- {name}" for _, name in devices] or ["- nenhum"])
            except Exception as exc:
                diagnostics.append(f"Falha ao listar scanners WIA: {exc}")

        messagebox.showinfo("Diagnostico do scanner", "\n".join(diagnostics))

    def describe_selected_scanner(self) -> None:
        current_index = self.scanner_combo.current()
        if current_index < 0 or current_index >= len(self.scanner_devices):
            self.scanner_info_var.set("Nenhum scanner WIA selecionado.")
            return

        _, device_name = self.scanner_devices[current_index]
        self.scanner_info_var.set(
            f"{device_name}: pronto para captura. Se o driver WIA falhar, o app tenta um modo de compatibilidade do Windows."
        )

    def refresh_scanner_list(self, initial: bool = False) -> None:
        try:
            self.scanner_devices = list_wia_scanners()
        except Exception as exc:
            self.scanner_devices = []
            self.scanner_combo["values"] = []
            if not initial:
                messagebox.showerror("Scanner", f"Nao foi possivel listar os scanners WIA: {exc}")
            return

        labels = [name for _, name in self.scanner_devices]
        self.scanner_combo["values"] = labels

        if not labels:
            self.scanner_var.set("Nenhum scanner WIA encontrado")
            self.scanner_info_var.set("Nenhum scanner WIA foi encontrado no Windows.")
            return

        saved_device_id = self.config.get("scanner_device_id", "")
        saved_index = next((index for index, (device_id, _) in enumerate(self.scanner_devices) if device_id == saved_device_id), 0)
        self.scanner_combo.current(saved_index)
        self.persist_scanner_device(self.scanner_devices[saved_index][0])
        self.scanner_combo.bind("<<ComboboxSelected>>", self.on_scanner_selected)
        self.describe_selected_scanner()

    def on_scanner_selected(self, _event=None) -> None:
        current_index = self.scanner_combo.current()
        if current_index < 0 or current_index >= len(self.scanner_devices):
            return
        device_id, device_name = self.scanner_devices[current_index]
        self.persist_scanner_device(device_id)
        self.describe_selected_scanner()
        self.append_log(f"[SCANNER] Dispositivo selecionado: {device_name}")

    def choose_matriz_path(self) -> None:
        selected_path = self.choose_directory_dialog(
            "Selecione a pasta matriz onde os contratos serao criados",
            self.matriz_var.get() or str(BASE_DIR),
        )
        if not selected_path:
            return
        self.persist_matriz_path(selected_path)

    def choose_input_path(self) -> None:
        selected_path = self.choose_directory_dialog(
            "Selecione a pasta monitorada de entrada",
            self.input_var.get() or str(BASE_DIR),
        )
        if not selected_path:
            return
        self.persist_input_path(selected_path)

    def change_matriz_path(self) -> None:
        current_path = self.matriz_var.get().strip()
        if not current_path:
            self.choose_matriz_path()
            return

        first_confirm = messagebox.askyesno(
            "Confirmar alteracao",
            "A pasta matriz ja esta definida. Deseja iniciar a alteracao?",
        )
        if not first_confirm:
            return

        second_confirm = messagebox.askyesno(
            "Segunda confirmacao",
            "Alterar a pasta matriz pode mudar o destino dos proximos documentos. Confirmar alteracao?",
        )
        if not second_confirm:
            return

        selected_path = self.choose_directory_dialog(
            "Selecione a pasta matriz onde os contratos serao criados",
            self.matriz_var.get() or str(BASE_DIR),
        )
        if not selected_path:
            return

        self.persist_matriz_path(selected_path)
        self.append_log(f"[CONFIG] Pasta matriz alterada para: {selected_path}")

    def start_monitoring(self) -> None:
        if self.processing or self.scanning:
            return

        matriz_path = self.matriz_var.get().strip()
        input_path = self.input_var.get().strip()
        if not matriz_path:
            messagebox.showwarning("Pasta matriz", "Defina a pasta matriz antes de iniciar o monitoramento.")
            return
        if not input_path:
            messagebox.showwarning("Pasta monitorada", "Defina a pasta monitorada antes de iniciar.")
            return

        matriz_dir = Path(matriz_path)
        input_dir = Path(input_path)
        self.processing = True
        self.scanning = False
        self.scan_stop_event.clear()
        self.stop_when_safe = False
        self.update_metrics(0, 0)
        self.process_button.state(["disabled"])
        self.scan_button.state(["disabled"])
        self.stop_scan_button.state(["!disabled"])
        self.status_var.set("Monitoramento em tempo real ativo.")
        self.append_log(f"[INICIO] Pasta matriz atual: {matriz_dir}")
        self.append_log(f"[MONITOR] Pasta monitorada: {input_dir}")
        self.append_log("[MONITOR] O app vai processar os arquivos assim que eles chegarem na pasta.")

        worker = threading.Thread(
            target=run_realtime_monitor,
            args=(
                input_dir,
                matriz_dir,
                self.thread_safe_log,
                self.thread_safe_finish,
                self.should_stop_scanning,
                self.thread_safe_metrics,
            ),
            daemon=True,
        )
        worker.start()

    def start_scanning(self) -> None:
        if self.processing or self.scanning:
            return

        if win32com is None:
            messagebox.showwarning(
                "Scanner",
                "O Python que esta executando o app nao carregou o pywin32.\n\n"
                f"Python atual: {sys.executable}\n\n"
                "Instale o pywin32 nesse mesmo Python ou gere o executavel com essa dependencia incluida.",
            )
            return

        matriz_path = self.matriz_var.get().strip()
        if not matriz_path:
            messagebox.showwarning("Pasta matriz", "Defina a pasta matriz antes de scanear.")
            return

        current_index = self.scanner_combo.current()
        if current_index < 0 or current_index >= len(self.scanner_devices):
            messagebox.showwarning("Scanner", "Selecione um scanner WIA valido.")
            return

        device_id, device_name = self.scanner_devices[current_index]
        self.persist_scanner_device(device_id)

        messagebox.showwarning(
            "Scanner",
            "Alguns scanners WIA falham em modo direto. Se isso acontecer, o app vai tentar um modo de compatibilidade do Windows.",
        )

        matriz_dir = Path(matriz_path)
        self.processing = True
        self.scanning = True
        self.scan_stop_event.clear()
        self.stop_when_safe = False
        self.update_metrics(0, 0)
        self.process_button.state(["disabled"])
        self.scan_button.state(["disabled"])
        self.stop_scan_button.state(["!disabled"])
        self.status_var.set("Sessao de scanner em andamento. Clique em 'Parar scanner' para processar.")
        self.append_log(f"[SCAN] Iniciando captura no scanner: {device_name}")
        self.append_log(f"[INICIO] Pasta matriz atual: {matriz_dir}")
        self.append_log("[SCAN] Cada documento sera capturado em um ciclo separado e acumulado ate voce clicar em 'Parar scanner'.")

        worker = threading.Thread(
            target=run_scanning_session,
            args=(device_id, matriz_dir, self.thread_safe_log, self.thread_safe_finish, self.should_stop_scanning),
            daemon=True,
        )
        worker.start()

    def stop_monitoring(self) -> None:
        if not self.processing and not self.scanning:
            return
        self.stop_when_safe = True
        self.stop_scan_button.state(["disabled"])
        if self.scanning:
            self.scan_stop_event.set()
            self.status_var.set("Encerrando sessao de captura e preparando processamento...")
            self.append_log("[SCAN] Parada solicitada pelo usuario. O app vai processar os documentos na ordem de scaneamento.")
            return

        input_path = self.input_var.get().strip()
        input_dir = Path(input_path) if input_path else INPUT_DIR
        pending_files = len(list(supported_files(input_dir)))
        if pending_files > 0:
            self.status_var.set(f"Parada solicitada. O app vai concluir {pending_files} arquivo(s) pendente(s) antes de encerrar.")
            self.append_log(
                f"[MONITOR] Parada solicitada. Restam {pending_files} arquivo(s) na pasta monitorada; o encerramento ocorrera quando a fila ficar vazia."
            )
            return

        self.status_var.set("Parada solicitada. Nenhum arquivo pendente encontrado; o monitoramento sera encerrado em seguida.")
        self.append_log("[MONITOR] Parada solicitada. Nenhum arquivo pendente encontrado na pasta monitorada.")

    def should_stop_scanning(self) -> bool:
        return self.scan_stop_event.is_set() or self.stop_when_safe

    def thread_safe_log(self, message: str) -> None:
        self.root.after(0, lambda: self.append_log(message))

    def thread_safe_finish(self, message: str) -> None:
        self.root.after(0, lambda: self.finish_processing(message))

    def thread_safe_metrics(self, processed_count: int, pending_count: int) -> None:
        self.root.after(0, lambda: self.update_metrics(processed_count, pending_count))

    def update_metrics(self, processed_count: int, pending_count: int) -> None:
        self.processed_count_var.set(f"Processados: {processed_count}")
        self.pending_var.set(f"Pendentes: {pending_count}")

    def open_matriz_folder(self) -> None:
        matriz_path = self.matriz_var.get().strip()
        if not matriz_path:
            return

        matriz_dir = Path(matriz_path)
        if not matriz_dir.exists():
            self.append_log(f"[AVISO] Pasta matriz nao encontrada para abrir: {matriz_dir}")
            return

        try:
            if os.name == "nt":
                os.startfile(str(matriz_dir))
            else:
                subprocess.Popen(["xdg-open", str(matriz_dir)])
        except Exception as exc:
            self.append_log(f"[ERRO] Nao foi possivel abrir a pasta matriz: {exc}")

    def open_folder(self, target_dir: Path) -> None:
        try:
            if os.name == "nt":
                os.startfile(str(target_dir))
            else:
                subprocess.Popen(["xdg-open", str(target_dir)])
        except Exception as exc:
            self.append_log(f"[ERRO] Nao foi possivel abrir a pasta: {exc}")

    def search_contract_folder(self) -> None:
        matriz_path = self.matriz_var.get().strip()
        contract_number = self.contract_search_var.get().strip()

        if not matriz_path:
            messagebox.showwarning("Pasta matriz", "Defina a pasta matriz antes de pesquisar.")
            return

        if not CONTRACT_INPUT_REGEX.fullmatch(contract_number):
            messagebox.showwarning("Contrato invalido", "Informe um numero de contrato com 5 ou 6 digitos.")
            return

        matriz_dir = Path(matriz_path)
        normalized_six = contract_number.zfill(6)
        stripped = str(int(normalized_six))
        contract_dir = None
        resolved_label = normalized_six
        for candidate_name in (normalized_six, stripped):
            candidate_path = matriz_dir / candidate_name
            if candidate_path.is_dir() and CONTRACT_FOLDER_NAME_REGEX.fullmatch(candidate_name):
                contract_dir = candidate_path
                resolved_label = candidate_name
                break

        if contract_dir is None:
            self.status_var.set(f"Contrato {normalized_six} nao encontrado.")
            self.append_log(f"[PESQUISA] Contrato nao encontrado: {matriz_dir / normalized_six}")
            messagebox.showinfo(
                "Pesquisa",
                f"Nenhuma pasta encontrada para o contrato {normalized_six} (nem {stripped}).",
            )
            return

        self.status_var.set(f"Contrato {resolved_label} localizado.")
        self.append_log(f"[PESQUISA] Abrindo pasta do contrato: {contract_dir}")
        self.open_folder(contract_dir)

    def finish_processing(self, message: str) -> None:
        self.processing = False
        self.scanning = False
        self.scan_stop_event.clear()
        self.stop_when_safe = False
        self.process_button.state(["!disabled"])
        self.scan_button.state(["!disabled"])
        self.stop_scan_button.state(["disabled"])
        self.status_var.set("Processamento finalizado.")
        self.update_metrics(0, 0)
        self.append_log(message)
        if self.open_matriz_on_finish_var.get():
            self.open_matriz_folder()
        messagebox.showinfo("Resultado", message)


def main() -> None:
    root = tk.Tk()
    OrganizerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

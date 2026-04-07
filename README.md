# Organizador de Fichas

Aplicativo desktop em Python para organizar documentos por contrato a partir da leitura da ficha no padrao `123456-78`.

O sistema foi feito para um cenario misto: parte dos equipamentos consegue digitalizar em fluxo pelo ADF, mas algumas impressoras antigas nao suportam esse modo de forma confiavel pelo driver. Por isso o projeto precisou aceitar mais de uma entrada:

- pasta monitorada ou pasta de entrada para arquivos ja digitalizados
- captura direta por scanner WIA quando o equipamento permite
- modo de compatibilidade quando o driver WIA falha no fluxo normal

Essa abordagem reduz dependencia do hardware. Se o scanner nao aceita captura em lote no app, o processo continua funcionando pelo envio dos arquivos para a pasta de entrada.

## Objetivo

Automatizar a separacao de fichas e anexos em pastas de contrato, evitando organizacao manual apos a digitalizacao.

## Como funciona

1. O usuario define a pasta matriz onde os contratos serao organizados.
2. O sistema recebe arquivos PDF ou imagem pela pasta de entrada configurada ou pelo scanner WIA.
3. O app aplica OCR quando necessario.
4. A ficha e extraida no formato `XXXXXX-XX`.
5. O documento e movido para `<pasta_matriz>/<contrato>`.
6. Se um arquivo nao trouxer ficha, ele pode ser tratado como complemento do ultimo contrato identificado.

## Por que o projeto foi feito assim

O principal motivo e compatibilidade com o parque de impressoras.

- Algumas impressoras antigas nao aceitam scan em fluxo de forma estavel.
- Em certos modelos, o driver WIA exposto ao Windows e limitado ou inconsistente.
- Em varios casos o ADF existe fisicamente, mas o driver nao entrega o comportamento esperado para captura continua.
- Quando isso acontece, depender apenas da digitalizacao direta no aplicativo tornaria o processo fragil.

Por isso o sistema foi desenhado com duas estrategias:

- `entrada monitorada`: funciona mesmo quando a digitalizacao precisa ser feita pelo software da propria impressora
- `scanner WIA no app`: melhora a operacao nos equipamentos mais novos ou melhor suportados

## Recursos

- OCR de PDFs e imagens
- extracao automatica de ficha por regex
- criacao automatica da pasta do contrato
- reaproveitamento da ultima pasta valida para anexos sem ficha
- selecao de scanner WIA no Windows
- tentativa de captura em lote quando o scanner suporta ADF
- fallback para modo de compatibilidade do Windows quando o driver falha
- persistencia de configuracao local em `config.json`
- logs de falha e diagnostico de scanner

## Estrutura principal

- `main.py`: aplicacao principal com interface e regras de processamento
- `config.json`: configuracao local da maquina
- `requirements.txt`: dependencias Python
- `OrganizadorFichas.spec`: arquivo de build do PyInstaller
- `entrada/`: arquivos aguardando processamento
- `logs/`: logs de erro e diagnostico
- `scan_debug/`: copias de apoio para diagnostico de capturas do scanner

## Requisitos

- Windows
- Python 3.10+
- Tesseract OCR instalado
- Poppler disponivel no `PATH`
- Scanner com driver WIA para usar a captura direta no aplicativo

## Instalacao

```powershell
pip install -r requirements.txt
```

Se necessario:

```powershell
$env:TESSERACT_CMD="C:\Program Files\Tesseract-OCR\tesseract.exe"
```

## Uso

```powershell
python main.py
```

No aplicativo, o fluxo esperado e:

1. Definir a pasta matriz.
2. Definir a pasta de entrada, se quiser usar arquivos gerados fora do app.
3. Atualizar e selecionar o scanner, se quiser capturar direto pelo Windows.
4. Processar os arquivos da entrada ou iniciar a captura pelo scanner.

## Build do executavel

```powershell
pyinstaller --noconsole --onefile --name OrganizadorFichas main.py
```

O executavel e gerado em `dist/`.

## Observacoes

- `config.json` guarda caminhos locais e identificador do scanner da maquina. Nao deve ir para o repositorio.
- Pastas como `build/`, `dist/`, `logs/`, `scan_debug/` e `__pycache__/` sao artefatos locais.
- Se o scanner falhar em modo direto, isso nao bloqueia o processo: a alternativa e digitalizar pelo software do equipamento e deixar os arquivos na pasta de entrada.

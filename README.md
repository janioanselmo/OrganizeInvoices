# OrganizeInvoices

## 🇧🇷 PT-BR

Utilitário em Python para organizar notas fiscais em PDF por data de emissão, gerar relatório Excel e registrar arquivos que precisam de revisão.

### Visão Geral

O projeto lê PDFs de notas fiscais, extrai texto com `pdfplumber`, identifica data de emissão, número da nota, emitente e CNPJ, e copia os arquivos para uma estrutura cronológica.

### Funcionalidades

- Leitura de arquivos PDF em lote.
- Extração automática de texto.
- Identificação heurística de data, número da nota, emitente e CNPJ.
- Renomeação cronológica dos arquivos.
- Organização em pastas por ano e mês.
- Geração de relatório `.xlsx`.
- Geração de log para arquivos com erro ou sem data identificada.
- Empacotamento com PyInstaller.

### Estrutura Esperada

```text
notas/
├── entrada/
├── organizadas/
└── relatorios/
```

Coloque os PDFs de entrada em:

```text
notas/entrada/
```

### Arquivos do Projeto

| Arquivo | Descrição |
| --- | --- |
| `OrganizeInvoices.py` | Script principal de organização |
| `requirements.txt` | Dependências Python |
| `OrganizeInvoices.spec` | Configuração PyInstaller para empacotamento |

### Instalação

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### Execução

```powershell
python OrganizeInvoices.py
```

### Saídas Geradas

| Caminho | Conteúdo |
| --- | --- |
| `notas/organizadas/AAAA/AAAA-MM/` | PDFs copiados e renomeados |
| `notas/organizadas/SEM_DATA/` | PDFs sem data identificada |
| `notas/relatorios/relatorio_notas.xlsx` | Relatório consolidado |
| `notas/relatorios/log_erros.txt` | Log de erros e revisões |

### Empacotamento

```powershell
pyinstaller OrganizeInvoices.spec
```

O executável será gerado em `dist/`.

### Observações Técnicas

- PDFs escaneados como imagem podem não ter texto extraível.
- A extração usa expressões regulares e heurísticas, pois layouts de notas fiscais variam bastante.
- O script copia os PDFs para a pasta organizada e não remove os arquivos originais.
- Os PDFs de entrada, relatórios e arquivos organizados são ignorados pelo Git para evitar publicação acidental de documentos fiscais.
- Contato: `anselmo.janio@gmail.com`.

### Licença

Projeto público pessoal de Jânio Anselmo.

---

## 🇺🇸 English

Python utility for organizing invoice PDFs by issue date, generating an Excel report, and logging files that require review.

### Overview

The project reads invoice PDFs, extracts text with `pdfplumber`, identifies issue date, invoice number, issuer and CNPJ, and copies files into a chronological folder structure.

### Features

- Batch PDF reading.
- Automatic text extraction.
- Heuristic identification of date, invoice number, issuer and CNPJ.
- Chronological file renaming.
- Organization by year and month.
- `.xlsx` report generation.
- Error/review log generation.
- Packaging with PyInstaller.

### Expected Structure

```text
notas/
├── entrada/
├── organizadas/
└── relatorios/
```

Place input PDFs in:

```text
notas/entrada/
```

### Project Files

| File | Description |
| --- | --- |
| `OrganizeInvoices.py` | Main organization script |
| `requirements.txt` | Python dependencies |
| `OrganizeInvoices.spec` | PyInstaller packaging configuration |

### Installation

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### Run

```powershell
python OrganizeInvoices.py
```

### Generated Outputs

| Path | Content |
| --- | --- |
| `notas/organizadas/YYYY/YYYY-MM/` | Copied and renamed PDFs |
| `notas/organizadas/SEM_DATA/` | PDFs without identified date |
| `notas/relatorios/relatorio_notas.xlsx` | Consolidated report |
| `notas/relatorios/log_erros.txt` | Error and review log |

### Packaging

```powershell
pyinstaller OrganizeInvoices.spec
```

The executable will be generated in `dist/`.

### Technical Notes

- Scanned image-only PDFs may not contain extractable text.
- Extraction uses regular expressions and heuristics because invoice layouts vary significantly.
- The script copies PDFs to the organized folder and does not delete original files.
- Input PDFs, reports and organized files are ignored by Git to avoid accidental publication of fiscal documents.
- Contact: `anselmo.janio@gmail.com`.

### License

Personal public project by Jânio Anselmo.

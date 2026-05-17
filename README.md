# OrganizeInvoices

Utilitário CLI em Python para organizar notas fiscais em PDF por data de emissão, gerar relatório Excel e registrar arquivos que precisam de revisão.

> Histórico completo de versões: [`CHANGELOG.md`](./CHANGELOG.md)
> Full release history: [`CHANGELOG.md`](./CHANGELOG.md)

---

## 🇧🇷 PT-BR

### Visão geral

O script lê PDFs de notas fiscais, extrai texto com `pdfplumber`, identifica data de emissão, número da nota, emitente e CNPJ, valida o dígito verificador do CNPJ e copia os arquivos para uma estrutura cronológica em `Notas/Organizadas/`. Também gera um relatório `.xlsx` e um log de erros.

### Funcionalidades

- Leitura em lote de PDFs.
- Extração heurística de data, número, emitente e CNPJ.
- **Validação do dígito verificador** do CNPJ.
- **Dedupe por MD5**: rodar 2× com a mesma entrada não duplica arquivos.
- Organização em pastas por ano/mês.
- Renomeação cronológica.
- Relatório `.xlsx` consolidado; **fallback automático** se o arquivo estiver aberto no Excel.
- **Append** no log de erros (preserva histórico entre execuções).
- CLI com `argparse` (`--base`, `--input`, `--output`, `--report`, `--verbose`).
- Empacotamento com PyInstaller.

### Estrutura padrão

```text
Notas/
├── Entrada/        # coloque seus PDFs aqui
├── Organizadas/    # PDFs organizados por ano/mês (gerado)
└── Relatorios/     # relatorio_notas.xlsx + log_erros.txt (gerado)
```

A pasta `Notas/` é criada ao lado do script (`OrganizeInvoices.py`) ou do executável empacotado. Para usar outra base, passe `--base CAMINHO`.

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

Opções:

```powershell
python OrganizeInvoices.py --base "D:\Notas" --verbose
python OrganizeInvoices.py --input "C:\Downloads\NFs" --output "D:\Arquivo\NFs"
```

| Flag | Descrição |
|---|---|
| `--base PATH` | Pasta base com `Entrada/Organizadas/Relatorios` (padrão: ao lado do script). |
| `--input PATH` | Sobrescreve a pasta de entrada. |
| `--output PATH` | Sobrescreve a pasta de saída. |
| `--report PATH` | Sobrescreve a pasta de relatórios. |
| `--verbose` `-v` | Mostra mensagem detalhada por arquivo. |

### Saídas geradas

| Caminho | Conteúdo |
|---|---|
| `Notas/Organizadas/AAAA/AAAA-MM/` | PDFs copiados e renomeados |
| `Notas/Organizadas/SEM_DATA/` | PDFs sem data identificada |
| `Notas/Relatorios/relatorio_notas.xlsx` | Relatório consolidado |
| `Notas/Relatorios/log_erros.txt` | Log com append entre execuções |

### Empacotamento

```powershell
pyinstaller OrganizeInvoices.spec
```

O executável é gerado em `dist/`.

### Testes

```powershell
pip install pytest
pytest
```

Os testes em `tests/test_parsers.py` cobrem os parsers de data, CNPJ, número e a sanitização de nomes.

### Observações técnicas

- PDFs escaneados como imagem podem não ter texto extraível (status `ERRO` no log).
- A extração usa regex e heurísticas; layouts de NF brasileiras variam bastante.
- O script **copia** os PDFs; nunca remove os originais.
- PDFs de entrada, relatórios e arquivos organizados são ignorados pelo Git para evitar publicação acidental de documentos fiscais.
- Contato: `anselmo.janio@gmail.com`.

### Licença

Distribuído sob **MIT**. Veja [`LICENSE`](./LICENSE).

---

## 🇺🇸 English

### Overview

The script reads invoice PDFs, extracts text with `pdfplumber`, identifies the issue date, invoice number, issuer and CNPJ, validates the CNPJ check digits, and copies files into a chronological structure under `Notas/Organizadas/`. It also produces an Excel report and an error log.

### Features

- Batch PDF reading.
- Heuristic extraction of date, number, issuer and CNPJ.
- **CNPJ check digit validation**.
- **MD5-based deduplication**: running twice on the same input doesn't duplicate files.
- Year/month folder organization.
- Chronological renaming.
- Consolidated `.xlsx` report; **automatic fallback** if the file is open in Excel.
- **Appends** to the error log (preserves history across runs).
- `argparse` CLI (`--base`, `--input`, `--output`, `--report`, `--verbose`).
- PyInstaller packaging.

### Default structure

```text
Notas/
├── Entrada/        # drop your PDFs here
├── Organizadas/    # PDFs organized by year/month (generated)
└── Relatorios/     # relatorio_notas.xlsx + log_erros.txt (generated)
```

The `Notas/` folder is created next to the script (`OrganizeInvoices.py`) or the packaged executable. Use `--base PATH` to point elsewhere.

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

Options:

```powershell
python OrganizeInvoices.py --base "D:\Notas" --verbose
python OrganizeInvoices.py --input "C:\Downloads\NFs" --output "D:\Arquivo\NFs"
```

| Flag | Description |
|---|---|
| `--base PATH` | Base folder containing `Entrada/Organizadas/Relatorios` (default: next to the script). |
| `--input PATH` | Override the input folder. |
| `--output PATH` | Override the output folder. |
| `--report PATH` | Override the report folder. |
| `--verbose` `-v` | Print per-file detail. |

### Generated outputs

| Path | Content |
|---|---|
| `Notas/Organizadas/YYYY/YYYY-MM/` | Copied and renamed PDFs |
| `Notas/Organizadas/SEM_DATA/` | PDFs without an identified date |
| `Notas/Relatorios/relatorio_notas.xlsx` | Consolidated report |
| `Notas/Relatorios/log_erros.txt` | Log appended across runs |

### Packaging

```powershell
pyinstaller OrganizeInvoices.spec
```

The executable is produced in `dist/`.

### Tests

```powershell
pip install pytest
pytest
```

Tests in `tests/test_parsers.py` cover the date, CNPJ and invoice-number parsers plus filename sanitization.

### Technical notes

- Scanned image-only PDFs may not have extractable text (status `ERRO` in the log).
- Extraction uses regex and heuristics; Brazilian invoice layouts vary significantly.
- The script **copies** PDFs; it never deletes originals.
- Input PDFs, reports and organized files are ignored by Git to avoid accidental publication of fiscal documents.
- Contact: `anselmo.janio@gmail.com`.

### License

Distributed under **MIT**. See [`LICENSE`](./LICENSE).

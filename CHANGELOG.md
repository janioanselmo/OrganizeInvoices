# Changelog

Histórico de versões do OrganizeInvoices.
OrganizeInvoices release history.

Formato: [Keep a Changelog](https://keepachangelog.com/).

---

## [v1.1.0] — 2026-05-17 — Auditoria + endurecimento dos parsers / Audit + parser hardening

### 🇧🇷 PT-BR

#### Confiabilidade do caminho base
- `BASE_DIR` resolvido via `script_dir()` (`Path(__file__).resolve().parent`, ou `sys.executable` quando empacotado). Antes o caminho era `Path("notas")` relativo ao cwd — o `.exe` PyInstaller criava pastas em local inesperado.
- Capitalização alinhada: o código usa `Notas/Entrada/Organizadas/Relatorios` (igual aos diretórios reais no Windows do projeto).
- Novo CLI com `argparse`: `--base`, `--input`, `--output`, `--report`, `--verbose`.

#### Parsers mais robustos
- `parse_date` agora aceita `dd/mm/aaaa`, `dd-mm-aaaa`, `dd.mm.aaaa`, `yyyy-mm-dd`, `yyyy/mm/dd`, e rejeita anos absurdos (<1990 ou >ano atual + 1).
- `extract_date` rejeita datas precedidas de "vencimento", "pagamento", "validade", "impressão", "recebimento" no fallback genérico — evita confundir vencimento com emissão.
- `extract_cnpj` valida **dígito verificador** antes de aceitar. Antes pegava qualquer sequência de 14 dígitos (incluindo trechos de código de barras de boleto). Saída sempre formatada com pontuação.
- `extract_invoice_number` prioriza padrões de NF-e/NFC-e/"nota fiscal eletrônica" antes de cair em "número" genérico.
- `extract_emitente` recebe o CNPJ extraído e, se não achar label explícito, busca o nome **nas 3 linhas acima do CNPJ no texto** — heurística mais alinhada com o layout real de DANFEs.

#### Robustez de I/O
- `generate_excel_report` trata `PermissionError` (arquivo `.xlsx` aberto no Excel) gravando em fallback `_YYYYMMDD_HHMMSS.xlsx` em vez de explodir.
- `append_error_log` agora faz **append** com cabeçalho de timestamp, preservando histórico entre execuções (antes sobrescrevia).
- `copy_with_dedupe` evita duplicar PDFs idênticos (compara MD5 com arquivos do destino). Rodar 2× com a mesma entrada não infla mais a `Organizadas/`.

#### Arquitetura
- Novo módulo `parsers.py` com toda a lógica pura (data, CNPJ, número, emitente, sanitização, hash). Sem dependência de `pandas`/`pdfplumber` — torna os parsers testáveis isoladamente.
- `OrganizeInvoices.py` agora só cuida de I/O, CLI, Excel/log. ~300 linhas.

#### Governança
- Adicionado `LICENSE` (MIT, Copyright Jânio Anselmo).
- README enxugado, seção "Auditoria / Audit" com checklist genérico removida.
- Docstring do script corrigida (nome do arquivo certo, requirements alinhado com `requirements.txt`).
- Smoke tests em `tests/test_parsers.py` cobrindo data, CNPJ, número, sanitização (27 testes, todos passando).
- **Bug encontrado pelos testes**: `_DATE_RE` usava `(?:\d{2}|\d{4})` no ano, mas a alternância da regex prefere o ramo da esquerda — então `02/02/2024` capturava `02/02/20` e `parse_date` falhava. Corrigido para `\d{4}` direto. Bug existia em produção antes desta rodada.

### 🇺🇸 English

#### Base path reliability
- `BASE_DIR` resolved via `script_dir()` (`Path(__file__).resolve().parent`, or `sys.executable` when frozen). Previously the path was `Path("notas")` relative to the cwd — the PyInstaller `.exe` created folders in unexpected locations.
- Capitalization aligned: code now uses `Notas/Entrada/Organizadas/Relatorios` (matching the actual project directories on Windows).
- New CLI via `argparse`: `--base`, `--input`, `--output`, `--report`, `--verbose`.

#### More robust parsers
- `parse_date` now accepts `dd/mm/yyyy`, `dd-mm-yyyy`, `dd.mm.yyyy`, `yyyy-mm-dd`, `yyyy/mm/dd`, and rejects implausible years (<1990 or >current year + 1).
- `extract_date` rejects dates preceded by "vencimento", "pagamento", "validade", "impressão", "recebimento" in the generic fallback — avoids treating the due date as the issue date.
- `extract_cnpj` validates the **check digits** before accepting. Previously matched any 14-digit run (including parts of payment-slip barcodes). Output is always formatted with punctuation.
- `extract_invoice_number` prioritizes NF-e/NFC-e/"nota fiscal eletrônica" patterns before the generic "número".
- `extract_emitente` takes the extracted CNPJ and, if no explicit label is found, looks for the name **in the 3 lines above the CNPJ in the text** — closer to real DANFE layout.

#### I/O robustness
- `generate_excel_report` handles `PermissionError` (the `.xlsx` open in Excel) by writing a `_YYYYMMDD_HHMMSS.xlsx` fallback instead of crashing.
- `append_error_log` now **appends** with a timestamped header, preserving history between runs (previously overwrote).
- `copy_with_dedupe` avoids duplicating identical PDFs (MD5 comparison against destination files). Running twice on the same input no longer inflates `Organizadas/`.

#### Architecture
- New `parsers.py` module holding all pure logic (date, CNPJ, invoice number, issuer, sanitization, hash). No `pandas`/`pdfplumber` dependency — parsers are testable in isolation.
- `OrganizeInvoices.py` now handles only I/O, CLI, Excel/log. ~300 lines.

#### Governance
- Added `LICENSE` (MIT, Copyright Jânio Anselmo).
- Slimmed the README; removed the generic-checklist "Auditoria / Audit" section.
- Fixed the script docstring (right filename, requirements aligned with `requirements.txt`).
- Smoke tests in `tests/test_parsers.py` covering date, CNPJ, invoice number, sanitization (27 tests, all passing).
- **Bug caught by the tests**: `_DATE_RE` used `(?:\d{2}|\d{4})` for the year, but regex alternation prefers the left branch — so `02/02/2024` captured `02/02/20` and `parse_date` failed. Fixed to `\d{4}` directly. The bug existed in production before this round.

---

## [v1.0.0] — Base inicial / Initial baseline

- Leitura em lote de PDFs em `notas/entrada/` com `pdfplumber`.
- Extração heurística de data, número, emitente e CNPJ.
- Organização em `notas/organizadas/AAAA/AAAA-MM/`.
- Relatório consolidado em `relatorio_notas.xlsx`.
- Log de erros em `log_erros.txt`.
- Empacotamento PyInstaller (`OrganizeInvoices.spec`).

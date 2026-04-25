# OrganizeInvoices

🧾 Utilitário em Python para organizar notas fiscais em PDF por data de emissão, gerar relatório em Excel e registrar arquivos que precisam de revisão.

## 📌 Visão geral

O projeto lê PDFs de notas fiscais, extrai texto com `pdfplumber`, identifica informações como data de emissão, número da nota, emitente e CNPJ, e copia os arquivos para uma estrutura cronológica.

O resultado do processamento é um conjunto de PDFs organizados por ano/mês, um relatório `.xlsx` e um log de erros/revisões.

## ✨ Funcionalidades

- Leitura de arquivos PDF em lote.
- Extração automática de texto.
- Identificação heurística de:
  - data de emissão;
  - número da nota;
  - emitente;
  - CNPJ.
- Renomeação cronológica dos arquivos.
- Organização em pastas por ano e mês.
- Geração de relatório Excel.
- Geração de log para arquivos com erro ou sem data identificada.

## 📁 Estrutura esperada

Ao executar o script, a estrutura abaixo é criada automaticamente:

```text
notas/
├── entrada/
├── organizadas/
└── relatorios/
```

Coloque os PDFs que deseja processar em:

```text
notas/entrada/
```

## 📁 Arquivos do projeto

| Arquivo | Descrição |
| --- | --- |
| `OrganizeInvoices.py` | Script principal de organização |
| `requirements.txt` | Dependências Python |
| `OrganizeInvoices.spec` | Configuração PyInstaller para empacotamento |

## 🛠️ Instalação

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## ▶️ Execução

```powershell
python OrganizeInvoices.py
```

## 📊 Saídas geradas

| Caminho | Conteúdo |
| --- | --- |
| `notas/organizadas/AAAA/AAAA-MM/` | PDFs copiados e renomeados |
| `notas/organizadas/SEM_DATA/` | PDFs sem data identificada |
| `notas/relatorios/relatorio_notas.xlsx` | Relatório consolidado |
| `notas/relatorios/log_erros.txt` | Log de erros e revisões |

## 📦 Empacotamento

Quando o PyInstaller estiver instalado:

```powershell
pyinstaller OrganizeInvoices.spec
```

O executável será gerado em `dist/`.

## 🧪 Observações técnicas

- PDFs escaneados como imagem podem não ter texto extraível; nesses casos, o arquivo será marcado como erro.
- A extração de dados usa expressões regulares e heurísticas, pois layouts de notas fiscais variam bastante.
- O script copia os PDFs para a pasta organizada; ele não remove os arquivos originais.
- Os PDFs de entrada, relatórios e arquivos organizados são ignorados pelo Git para evitar publicação acidental de documentos fiscais.

## 📄 Licença

Projeto público pessoal de Jânio Anselmo.

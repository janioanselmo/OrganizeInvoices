#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Organizador cronológico de notas fiscais em PDF.

Funcionalidades:
- Lê PDFs de uma pasta de entrada.
- Extrai texto do PDF.
- Tenta identificar data de emissão e número da nota.
- Gera relatório em Excel.
- Copia os arquivos para pastas por ano/mês.
- Renomeia os arquivos em padrão cronológico.
- Gera log dos arquivos com erro.

Requisitos:
    pip install pdfplumber pandas openpyxl

Uso:
    python organizar_notas.py

Estrutura esperada:
    notas/
    ├── entrada/
    ├── organizadas/
    └── relatorios/
"""

from __future__ import annotations

import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import pdfplumber


# =========================
# CONFIGURAÇÕES
# =========================
BASE_DIR = Path("notas")
INPUT_DIR = BASE_DIR / "entrada"
OUTPUT_DIR = BASE_DIR / "organizadas"
REPORT_DIR = BASE_DIR / "relatorios"

EXCEL_REPORT = REPORT_DIR / "relatorio_notas.xlsx"
ERROR_LOG = REPORT_DIR / "log_erros.txt"


# =========================
# MODELOS
# =========================
@dataclass
class NotaFiscalInfo:
    """Estrutura com os dados extraídos de um PDF de nota fiscal."""

    arquivo_original: str
    caminho_original: str
    data_emissao: Optional[datetime]
    numero_nota: Optional[str]
    emitente: Optional[str]
    cnpj: Optional[str]
    status: str
    mensagem: str
    nome_novo_arquivo: Optional[str]
    pasta_destino: Optional[str]


# =========================
# UTILITÁRIOS
# =========================
def ensure_directories() -> None:
    """Cria as pastas necessárias, se não existirem."""
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)


def sanitize_filename(text: str) -> str:
    """
    Remove caracteres problemáticos para nome de arquivo.

    Args:
        text: Texto original.

    Returns:
        Texto seguro para nome de arquivo.
    """
    text = re.sub(r"[^\w\s\-\.]", "", text, flags=re.UNICODE)
    text = re.sub(r"\s+", "_", text.strip())
    return text[:80] if text else "SEM_NOME"


def extract_text_from_pdf(pdf_path: Path) -> str:
    """
    Extrai texto de todas as páginas do PDF.

    Args:
        pdf_path: Caminho do PDF.

    Returns:
        Texto concatenado do PDF.
    """
    pages_text = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text.strip():
                pages_text.append(text)

    return "\n".join(pages_text)


def extract_date(text: str) -> Optional[datetime]:
    """
    Tenta extrair data de emissão do texto do PDF.

    Estratégia:
    1. Busca padrões próximos a 'data de emissão'.
    2. Busca datas genéricas dd/mm/aaaa.
    3. Retorna a primeira data válida encontrada.

    Args:
        text: Texto do PDF.

    Returns:
        Data encontrada ou None.
    """
    normalized = " ".join(text.split())

    priority_patterns = [
        r"data\s+de\s+emiss[aã]o[:\s]*([0-3]?\d/[0-1]?\d/\d{4})",
        r"emiss[aã]o[:\s]*([0-3]?\d/[0-1]?\d/\d{4})",
        r"dt\.\s*emiss[aã]o[:\s]*([0-3]?\d/[0-1]?\d/\d{4})",
    ]

    generic_patterns = [
        r"\b([0-3]?\d/[0-1]?\d/\d{4})\b",
        r"\b([0-3]?\d-[0-1]?\d-\d{4})\b",
    ]

    for pattern in priority_patterns:
        match = re.search(pattern, normalized, flags=re.IGNORECASE)
        if match:
            return parse_date(match.group(1))

    for pattern in generic_patterns:
        matches = re.findall(pattern, normalized, flags=re.IGNORECASE)
        for date_str in matches:
            parsed = parse_date(date_str)
            if parsed:
                return parsed

    return None


def parse_date(date_str: str) -> Optional[datetime]:
    """
    Converte string de data em datetime.

    Args:
        date_str: String de data.

    Returns:
        datetime ou None.
    """
    date_str = date_str.strip()

    for fmt in ("%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue

    return None


def extract_invoice_number(text: str) -> Optional[str]:
    """
    Tenta extrair número da nota fiscal do texto.

    Args:
        text: Texto do PDF.

    Returns:
        Número da nota ou None.
    """
    normalized = " ".join(text.split())

    patterns = [
        r"n[uú]mero\s+da\s+nota[:\s]*([A-Z0-9\-\/\.]+)",
        r"n[ºo°]\s*[:\-]?\s*([A-Z0-9\-\/\.]+)",
        r"nota\s+fiscal\s+eletr[oô]nica\s*[:\-]?\s*([A-Z0-9\-\/\.]+)",
        r"n[uú]mero[:\s]*([A-Z0-9\-\/\.]{1,20})",
    ]

    for pattern in patterns:
        match = re.search(pattern, normalized, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip()

    return None


def extract_cnpj(text: str) -> Optional[str]:
    """
    Extrai o primeiro CNPJ encontrado no texto.

    Args:
        text: Texto do PDF.

    Returns:
        CNPJ ou None.
    """
    match = re.search(r"\b(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\b", text)
    if match:
        return match.group(1)

    match = re.search(r"\b(\d{14})\b", text)
    if match:
        return match.group(1)

    return None


def extract_emitente(text: str) -> Optional[str]:
    """
    Tenta extrair o nome do emitente.

    Observação:
    Em PDFs de NF o padrão varia muito. Aqui usamos heurística simples.

    Args:
        text: Texto do PDF.

    Returns:
        Nome do emitente ou None.
    """
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    candidate_patterns = [
        r"emitente[:\s]*(.+)",
        r"fornecedor[:\s]*(.+)",
        r"raz[aã]o\s+social[:\s]*(.+)",
    ]

    for line in lines[:50]:
        for pattern in candidate_patterns:
            match = re.search(pattern, line, flags=re.IGNORECASE)
            if match:
                value = match.group(1).strip()
                if len(value) > 2:
                    return value

    return None


def build_new_filename(
    data_emissao: Optional[datetime],
    numero_nota: Optional[str],
    emitente: Optional[str],
    original_suffix: str,
    sequence: int,
) -> str:
    """
    Monta nome novo do arquivo.

    Args:
        data_emissao: Data extraída.
        numero_nota: Número da nota.
        emitente: Nome do emitente.
        original_suffix: Extensão do arquivo.
        sequence: Índice sequencial para evitar colisão.

    Returns:
        Nome de arquivo padronizado.
    """
    date_part = data_emissao.strftime("%Y-%m-%d") if data_emissao else "SEM_DATA"
    invoice_part = sanitize_filename(numero_nota or "SEM_NUMERO")
    emitente_part = sanitize_filename(emitente or "SEM_EMITENTE")

    return (
        f"{date_part}_NF_{invoice_part}_{emitente_part}_{sequence:03d}"
        f"{original_suffix.lower()}"
    )


def copy_file_safely(src: Path, dst: Path) -> None:
    """
    Copia o arquivo sem alterar o original.

    Args:
        src: Origem.
        dst: Destino.
    """
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)


# =========================
# PROCESSAMENTO PRINCIPAL
# =========================
def process_pdf(pdf_path: Path, sequence: int) -> NotaFiscalInfo:
    """
    Processa um único PDF.

    Args:
        pdf_path: Caminho do PDF.
        sequence: Índice sequencial.

    Returns:
        Dados extraídos do arquivo.
    """
    try:
        text = extract_text_from_pdf(pdf_path)

        if not text.strip():
            return NotaFiscalInfo(
                arquivo_original=pdf_path.name,
                caminho_original=str(pdf_path.resolve()),
                data_emissao=None,
                numero_nota=None,
                emitente=None,
                cnpj=None,
                status="ERRO",
                mensagem="PDF sem texto extraível.",
                nome_novo_arquivo=None,
                pasta_destino=None,
            )

        data_emissao = extract_date(text)
        numero_nota = extract_invoice_number(text)
        emitente = extract_emitente(text)
        cnpj = extract_cnpj(text)

        if data_emissao:
            year_month_folder = (
                OUTPUT_DIR
                / str(data_emissao.year)
                / f"{data_emissao.year}-{data_emissao.month:02d}"
            )
        else:
            year_month_folder = OUTPUT_DIR / "SEM_DATA"

        new_filename = build_new_filename(
            data_emissao=data_emissao,
            numero_nota=numero_nota,
            emitente=emitente,
            original_suffix=pdf_path.suffix,
            sequence=sequence,
        )

        destination = year_month_folder / new_filename
        copy_file_safely(pdf_path, destination)

        status = "OK" if data_emissao else "REVISAR"
        mensagem = (
            "Arquivo processado com sucesso."
            if data_emissao
            else "Data não identificada automaticamente."
        )

        return NotaFiscalInfo(
            arquivo_original=pdf_path.name,
            caminho_original=str(pdf_path.resolve()),
            data_emissao=data_emissao,
            numero_nota=numero_nota,
            emitente=emitente,
            cnpj=cnpj,
            status=status,
            mensagem=mensagem,
            nome_novo_arquivo=new_filename,
            pasta_destino=str(destination.parent.resolve()),
        )

    except Exception as exc:  # pylint: disable=broad-except
        return NotaFiscalInfo(
            arquivo_original=pdf_path.name,
            caminho_original=str(pdf_path.resolve()),
            data_emissao=None,
            numero_nota=None,
            emitente=None,
            cnpj=None,
            status="ERRO",
            mensagem=f"Falha no processamento: {exc}",
            nome_novo_arquivo=None,
            pasta_destino=None,
        )


def generate_excel_report(items: list[NotaFiscalInfo]) -> None:
    """
    Gera o relatório em Excel.

    Args:
        items: Lista de resultados processados.
    """
    records = []

    for item in items:
        records.append(
            {
                "arquivo_original": item.arquivo_original,
                "caminho_original": item.caminho_original,
                "data_emissao": (
                    item.data_emissao.strftime("%Y-%m-%d")
                    if item.data_emissao
                    else None
                ),
                "numero_nota": item.numero_nota,
                "emitente": item.emitente,
                "cnpj": item.cnpj,
                "status": item.status,
                "mensagem": item.mensagem,
                "nome_novo_arquivo": item.nome_novo_arquivo,
                "pasta_destino": item.pasta_destino,
            }
        )

    df = pd.DataFrame(records)

    if not df.empty:
        df = df.sort_values(
            by=["data_emissao", "arquivo_original"],
            ascending=[True, True],
            na_position="last",
        )

    with pd.ExcelWriter(EXCEL_REPORT, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="NotasOrganizadas")


def generate_error_log(items: list[NotaFiscalInfo]) -> None:
    """
    Gera um log de erros e arquivos para revisão.

    Args:
        items: Lista de resultados processados.
    """
    lines = []
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    lines.append(f"Log gerado em: {now_str}")
    lines.append("=" * 80)

    for item in items:
        if item.status in {"ERRO", "REVISAR"}:
            lines.append(f"Arquivo: {item.arquivo_original}")
            lines.append(f"Status: {item.status}")
            lines.append(f"Mensagem: {item.mensagem}")
            lines.append(f"Caminho: {item.caminho_original}")
            lines.append("-" * 80)

    ERROR_LOG.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    """Função principal do script."""
    ensure_directories()

    pdf_files = sorted(INPUT_DIR.glob("*.pdf"))

    if not pdf_files:
        print(f"Nenhum PDF encontrado em: {INPUT_DIR.resolve()}")
        print("Coloque os arquivos PDF dentro da pasta 'entrada' e execute novamente.")
        return

    results = []

    print(f"Encontrados {len(pdf_files)} arquivos PDF.")
    print("Processando...\n")

    for idx, pdf_file in enumerate(pdf_files, start=1):
        result = process_pdf(pdf_file, sequence=idx)
        results.append(result)
        print(f"[{idx:03d}/{len(pdf_files):03d}] {pdf_file.name} -> {result.status}")

    generate_excel_report(results)
    generate_error_log(results)

    total_ok = sum(1 for item in results if item.status == "OK")
    total_revisar = sum(1 for item in results if item.status == "REVISAR")
    total_error = sum(1 for item in results if item.status == "ERRO")

    print("\nProcessamento concluído.")
    print(f"OK: {total_ok}")
    print(f"REVISAR: {total_revisar}")
    print(f"ERRO: {total_error}")
    print(f"Relatório Excel: {EXCEL_REPORT.resolve()}")
    print(f"Log: {ERROR_LOG.resolve()}")


if __name__ == "__main__":
    main()
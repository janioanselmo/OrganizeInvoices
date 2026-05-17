#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Organizador cronológico de notas fiscais em PDF.

Funcionalidades:
- Lê PDFs de uma pasta de entrada (padrão: ``Notas/Entrada`` ao lado do script).
- Extrai texto do PDF.
- Tenta identificar data de emissão, número da nota, emitente e CNPJ.
- Valida o dígito verificador do CNPJ extraído.
- Deduplica por hash MD5 do conteúdo do PDF.
- Gera relatório em Excel.
- Copia os arquivos para pastas por ano/mês.
- Renomeia os arquivos em padrão cronológico.
- Faz append no log de erros (preserva histórico entre execuções).

Uso:
    python OrganizeInvoices.py [--base PATH] [--input PATH] [--output PATH]
                               [--report PATH] [--verbose]

Estrutura padrão (relativa ao script ou a ``--base``):
    Notas/
    ├── Entrada/
    ├── Organizadas/
    └── Relatorios/

As dependências oficiais ficam em ``requirements.txt``.
"""

from __future__ import annotations

import argparse
import shutil
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import pdfplumber

from parsers import (
    build_new_filename,
    extract_cnpj,
    extract_date,
    extract_emitente,
    extract_invoice_number,
    file_md5,
)


# =========================
# CAMINHOS PADRÃO
# =========================
def script_dir() -> Path:
    """Pasta onde o script (ou o ``.exe`` empacotado) está.

    PyInstaller define ``sys._MEIPASS`` quando o programa está rodando como
    onefile/onefolder. Mesmo aí, queremos a pasta do executável real, não a
    temporária descompactada — daí o ``sys.executable``.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


DEFAULT_BASE = script_dir() / "Notas"


@dataclass(frozen=True)
class Paths:
    """Conjunto de pastas usado em uma execução."""

    base: Path
    input_dir: Path
    output_dir: Path
    report_dir: Path
    excel_report: Path
    error_log: Path

    @classmethod
    def from_base(cls, base: Path) -> "Paths":
        base = base.resolve()
        report = base / "Relatorios"
        return cls(
            base=base,
            input_dir=base / "Entrada",
            output_dir=base / "Organizadas",
            report_dir=report,
            excel_report=report / "relatorio_notas.xlsx",
            error_log=report / "log_erros.txt",
        )


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
def ensure_directories(paths: Paths) -> None:
    """Cria as pastas necessárias, se não existirem."""
    paths.input_dir.mkdir(parents=True, exist_ok=True)
    paths.output_dir.mkdir(parents=True, exist_ok=True)
    paths.report_dir.mkdir(parents=True, exist_ok=True)


# =========================
# EXTRAÇÃO DE TEXTO
# =========================
def extract_text_from_pdf(pdf_path: Path) -> str:
    """Concatena o texto de todas as páginas do PDF."""
    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text.strip():
                pages_text.append(text)
    return "\n".join(pages_text)


# =========================
# CÓPIA E DEDUPE
# =========================
def copy_with_dedupe(src: Path, dst: Path) -> tuple[Path, bool]:
    """Copia ``src`` para ``dst``, evitando duplicar conteúdo já presente.

    Retorna ``(caminho_final, novo)``. Se um arquivo com o mesmo MD5 já
    existir no diretório de destino, devolve o caminho dele e ``novo=False``
    em vez de criar uma cópia extra. Isso evita inchaço quando o operador
    roda o script duas vezes com os mesmos PDFs na ``Entrada``.
    """
    dst.parent.mkdir(parents=True, exist_ok=True)
    src_hash = file_md5(src)
    if dst.parent.exists():
        for existing in dst.parent.iterdir():
            if existing.is_file() and file_md5(existing) == src_hash:
                return existing, False
    shutil.copy2(src, dst)
    return dst, True


# =========================
# PROCESSAMENTO DE UM PDF
# =========================
def process_pdf(pdf_path: Path, sequence: int, paths: Paths) -> NotaFiscalInfo:
    """Processa um único PDF — extrai dados, copia e registra resultado."""
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
                mensagem="PDF sem texto extraível (provavelmente imagem escaneada).",
                nome_novo_arquivo=None,
                pasta_destino=None,
            )

        data_emissao = extract_date(text)
        numero_nota = extract_invoice_number(text)
        cnpj = extract_cnpj(text)
        emitente = extract_emitente(text, cnpj=cnpj)

        if data_emissao:
            year_month_folder = (
                paths.output_dir
                / str(data_emissao.year)
                / f"{data_emissao.year}-{data_emissao.month:02d}"
            )
        else:
            year_month_folder = paths.output_dir / "SEM_DATA"

        new_filename = build_new_filename(
            data_emissao=data_emissao,
            numero_nota=numero_nota,
            emitente=emitente,
            original_suffix=pdf_path.suffix,
            sequence=sequence,
        )

        final_path, is_new = copy_with_dedupe(pdf_path, year_month_folder / new_filename)

        if data_emissao:
            status = "OK"
            mensagem = (
                "Arquivo processado com sucesso."
                if is_new
                else f"Duplicado de {final_path.name}; cópia não criada."
            )
        else:
            status = "REVISAR"
            mensagem = "Data não identificada automaticamente."

        return NotaFiscalInfo(
            arquivo_original=pdf_path.name,
            caminho_original=str(pdf_path.resolve()),
            data_emissao=data_emissao,
            numero_nota=numero_nota,
            emitente=emitente,
            cnpj=cnpj,
            status=status,
            mensagem=mensagem,
            nome_novo_arquivo=final_path.name,
            pasta_destino=str(final_path.parent.resolve()),
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


# =========================
# RELATÓRIOS
# =========================
def _records_from_items(items: list[NotaFiscalInfo]) -> list[dict]:
    return [
        {
            "arquivo_original": item.arquivo_original,
            "caminho_original": item.caminho_original,
            "data_emissao": (
                item.data_emissao.strftime("%Y-%m-%d") if item.data_emissao else None
            ),
            "numero_nota": item.numero_nota,
            "emitente": item.emitente,
            "cnpj": item.cnpj,
            "status": item.status,
            "mensagem": item.mensagem,
            "nome_novo_arquivo": item.nome_novo_arquivo,
            "pasta_destino": item.pasta_destino,
        }
        for item in items
    ]


def generate_excel_report(items: list[NotaFiscalInfo], paths: Paths) -> Path:
    """Grava o relatório em Excel. Se o arquivo estiver aberto, salva uma
    cópia com timestamp e mostra o caminho efetivo.
    """
    df = pd.DataFrame(_records_from_items(items))
    if not df.empty:
        df = df.sort_values(
            by=["data_emissao", "arquivo_original"],
            ascending=[True, True],
            na_position="last",
        )

    target = paths.excel_report
    try:
        with pd.ExcelWriter(target, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="NotasOrganizadas")
        return target
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback = target.with_name(f"{target.stem}_{timestamp}{target.suffix}")
        with pd.ExcelWriter(fallback, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="NotasOrganizadas")
        print(
            f"\nAtenção: '{target.name}' está aberto em outro programa. "
            f"Salvei o relatório em '{fallback.name}'.",
            file=sys.stderr,
        )
        return fallback


def append_error_log(items: list[NotaFiscalInfo], paths: Paths) -> None:
    """Append no log de erros (preserva histórico entre execuções)."""
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        "",
        "=" * 80,
        f"Execução em: {now_str}",
        "=" * 80,
    ]

    flagged = [item for item in items if item.status in {"ERRO", "REVISAR"}]
    if not flagged:
        lines.append("Nenhum arquivo precisou de revisão nesta execução.")
    else:
        for item in flagged:
            lines.append(f"Arquivo: {item.arquivo_original}")
            lines.append(f"Status: {item.status}")
            lines.append(f"Mensagem: {item.mensagem}")
            lines.append(f"Caminho: {item.caminho_original}")
            lines.append("-" * 80)

    with paths.error_log.open("a", encoding="utf-8") as handle:
        handle.write("\n".join(lines) + "\n")


# =========================
# CLI
# =========================
def _parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="OrganizeInvoices",
        description="Organiza PDFs de notas fiscais por data de emissão.",
    )
    parser.add_argument(
        "--base",
        type=Path,
        default=DEFAULT_BASE,
        help=f"Pasta base que contém Entrada/Organizadas/Relatorios (padrão: {DEFAULT_BASE}).",
    )
    parser.add_argument("--input", type=Path, default=None, help="Sobrescreve a pasta de entrada.")
    parser.add_argument("--output", type=Path, default=None, help="Sobrescreve a pasta de saída.")
    parser.add_argument("--report", type=Path, default=None, help="Sobrescreve a pasta de relatórios.")
    parser.add_argument("--verbose", "-v", action="store_true", help="Mostra detalhes por arquivo.")
    return parser.parse_args(argv)


def _paths_from_args(args: argparse.Namespace) -> Paths:
    paths = Paths.from_base(args.base)
    if args.input or args.output or args.report:
        report = (args.report or paths.report_dir).resolve()
        return Paths(
            base=paths.base,
            input_dir=(args.input or paths.input_dir).resolve(),
            output_dir=(args.output or paths.output_dir).resolve(),
            report_dir=report,
            excel_report=report / paths.excel_report.name,
            error_log=report / paths.error_log.name,
        )
    return paths


def main(argv: Optional[list[str]] = None) -> int:
    args = _parse_args(argv)
    paths = _paths_from_args(args)
    ensure_directories(paths)

    pdf_files = sorted(paths.input_dir.glob("*.pdf"))

    if not pdf_files:
        print(f"Nenhum PDF encontrado em: {paths.input_dir}")
        print("Coloque os arquivos PDF dentro da pasta de entrada e execute novamente.")
        return 0

    print(f"Encontrados {len(pdf_files)} arquivos PDF.")
    print(f"Pasta base: {paths.base}")
    print("Processando...\n")

    results: list[NotaFiscalInfo] = []
    for idx, pdf_file in enumerate(pdf_files, start=1):
        result = process_pdf(pdf_file, sequence=idx, paths=paths)
        results.append(result)
        if args.verbose:
            print(f"[{idx:03d}/{len(pdf_files):03d}] {pdf_file.name} -> {result.status} | {result.mensagem}")
        else:
            print(f"[{idx:03d}/{len(pdf_files):03d}] {pdf_file.name} -> {result.status}")

    excel_path = generate_excel_report(results, paths)
    append_error_log(results, paths)

    total_ok = sum(1 for item in results if item.status == "OK")
    total_revisar = sum(1 for item in results if item.status == "REVISAR")
    total_error = sum(1 for item in results if item.status == "ERRO")

    print("\nProcessamento concluído.")
    print(f"OK:      {total_ok}")
    print(f"REVISAR: {total_revisar}")
    print(f"ERRO:    {total_error}")
    print(f"Relatório Excel: {excel_path}")
    print(f"Log:             {paths.error_log}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

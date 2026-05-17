"""Pure parsers used by OrganizeInvoices.

This module has zero runtime dependencies (no ``pandas``, no ``pdfplumber``)
so the parser logic can be smoke-tested in isolation. Anything that reads a
PDF, writes Excel or touches the filesystem with side effects stays in
``OrganizeInvoices.py``.
"""

from __future__ import annotations

import hashlib
import re
from datetime import datetime
from pathlib import Path
from typing import Optional


# =========================
# FILENAME
# =========================
def sanitize_filename(text: str) -> str:
    """Remove caracteres problemáticos e limita o tamanho do nome do arquivo."""
    text = re.sub(r"[^\w\s\-\.]", "", text, flags=re.UNICODE)
    text = re.sub(r"\s+", "_", text.strip())
    return text[:80] if text else "SEM_NOME"


def build_new_filename(
    data_emissao: Optional[datetime],
    numero_nota: Optional[str],
    emitente: Optional[str],
    original_suffix: str,
    sequence: int,
) -> str:
    """Monta o nome padronizado do arquivo de destino."""
    date_part = data_emissao.strftime("%Y-%m-%d") if data_emissao else "SEM_DATA"
    invoice_part = sanitize_filename(numero_nota or "SEM_NUMERO")
    emitente_part = sanitize_filename(emitente or "SEM_EMITENTE")
    return (
        f"{date_part}_NF_{invoice_part}_{emitente_part}_{sequence:03d}"
        f"{original_suffix.lower()}"
    )


# =========================
# HASH
# =========================
def file_md5(path: Path, chunk: int = 1024 * 1024) -> str:
    """Hash MD5 do conteúdo (usado para dedupe)."""
    digest = hashlib.md5()  # noqa: S324 — dedupe only, not security
    with path.open("rb") as handle:
        while True:
            block = handle.read(chunk)
            if not block:
                break
            digest.update(block)
    return digest.hexdigest()


# =========================
# DATAS
# =========================
_DATE_FORMATS = ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%Y/%m/%d")
_MIN_PLAUSIBLE_YEAR = 1990
# Year must be 4 digits to avoid greedy-alternation trap: `(?:\d{2}|\d{4})`
# stops at 2 digits because regex `|` prefers the leftmost branch, so
# "02/02/2024" used to capture "02/02/20" and then parse_date failed.
_DATE_RE = (
    r"([0-3]?\d[\/\-\.][0-1]?\d[\/\-\.]\d{4}|\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})"
)


def parse_date(date_str: str) -> Optional[datetime]:
    """Converte string em ``datetime`` aceitando os formatos comuns.

    Datas claramente implausíveis (ano < 1990 ou mais de 1 ano no futuro) são
    rejeitadas para reduzir falso-positivo de números aleatórios no PDF.
    """
    date_str = date_str.strip()
    for fmt in _DATE_FORMATS:
        try:
            parsed = datetime.strptime(date_str, fmt)
        except ValueError:
            continue
        if parsed.year < _MIN_PLAUSIBLE_YEAR:
            return None
        if parsed.year > datetime.now().year + 1:
            return None
        return parsed
    return None


def extract_date(text: str) -> Optional[datetime]:
    """Extrai a data de emissão do texto do PDF.

    1. Padrões explícitos de "data de emissão" / "emissão" / "dt. emissão"
       ganham prioridade absoluta.
    2. No fallback genérico, **rejeita** datas próximas a labels como
       "vencimento", "pagamento", "validade", "impressão" ou "recebimento" —
       evita tratar vencimento como emissão.
    """
    normalized = " ".join(text.split())

    priority_patterns = [
        rf"data\s+de\s+emiss[aã]o[:\s]*{_DATE_RE}",
        rf"emiss[aã]o[:\s]*{_DATE_RE}",
        rf"dt\.?\s*emiss[aã]o[:\s]*{_DATE_RE}",
        rf"data\s+da\s+emiss[aã]o[:\s]*{_DATE_RE}",
    ]
    for pattern in priority_patterns:
        match = re.search(pattern, normalized, flags=re.IGNORECASE)
        if match:
            parsed = parse_date(match.group(1))
            if parsed:
                return parsed

    bad_label = re.compile(
        r"(vencimento|pagamento|validade|impress[aã]o|recebimento)[:\s]*$",
        re.IGNORECASE,
    )
    for match in re.finditer(_DATE_RE, normalized):
        prefix = normalized[max(0, match.start() - 20): match.start()]
        if bad_label.search(prefix):
            continue
        parsed = parse_date(match.group(1))
        if parsed:
            return parsed
    return None


# =========================
# NÚMERO DA NOTA
# =========================
def extract_invoice_number(text: str) -> Optional[str]:
    """Extrai o número da nota fiscal.

    Padrões priorizam termos típicos de DANFE/NF-e antes de cair em "número"
    genérico (que casaria com telefone, CEP etc.).
    """
    normalized = " ".join(text.split())

    patterns = [
        r"nf[\-\s]?e?\s+n[ºo°]?\s*[:\-]?\s*([0-9][0-9\.\-\/]{0,19})",
        r"nfc[\-\s]?e?\s+n[ºo°]?\s*[:\-]?\s*([0-9][0-9\.\-\/]{0,19})",
        r"nota\s+fiscal\s+eletr[oô]nica\s*n[ºo°]?\s*[:\-]?\s*([0-9][0-9\.\-\/]{0,19})",
        r"n[uú]mero\s+da\s+nota[:\s]*([A-Z0-9\-\/\.]{1,20})",
        r"n[ºo°]\s*[:\-]?\s*([0-9][0-9\.\-\/]{1,19})",
    ]
    for pattern in patterns:
        match = re.search(pattern, normalized, flags=re.IGNORECASE)
        if match:
            value = match.group(1).strip(" .-/")
            if value:
                return value
    return None


# =========================
# CNPJ
# =========================
def _validate_cnpj_digits(digits: str) -> bool:
    """Valida CNPJ pelo dígito verificador (padrão Receita Federal)."""
    if len(digits) != 14 or len(set(digits)) == 1:
        return False

    def calc(slice_: str, weights: list[int]) -> int:
        total = sum(int(d) * w for d, w in zip(slice_, weights))
        rest = total % 11
        return 0 if rest < 2 else 11 - rest

    first = calc(digits[:12], [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2])
    second = calc(digits[:13], [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2])
    return digits[-2:] == f"{first}{second}"


def extract_cnpj(text: str) -> Optional[str]:
    """Retorna o primeiro CNPJ **válido** encontrado no texto.

    Antes a função aceitava qualquer sequência de 14 dígitos, gerando falso
    positivo em códigos de barras de boleto. Agora o dígito verificador é
    sempre conferido. O formato com pontuação é mantido na saída.
    """
    for match in re.finditer(
        r"\b(\d{2})\.(\d{3})\.(\d{3})/(\d{4})-(\d{2})\b", text
    ):
        digits = "".join(match.groups())
        if _validate_cnpj_digits(digits):
            return match.group(0)

    for match in re.finditer(r"\b(\d{14})\b", text):
        digits = match.group(1)
        if _validate_cnpj_digits(digits):
            return (
                f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/"
                f"{digits[8:12]}-{digits[12:]}"
            )
    return None


# =========================
# EMITENTE
# =========================
def extract_emitente(text: str, cnpj: Optional[str] = None) -> Optional[str]:
    """Tenta identificar o nome do emitente.

    1. Procura labels explícitos ("razão social", "emitente", "fornecedor").
    2. Se não achar, e um ``cnpj`` for fornecido, procura nas 3 linhas
       acima da linha onde o CNPJ aparece — em DANFEs o nome do emitente
       costuma ficar imediatamente acima do CNPJ.
    """
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    candidate_patterns = [
        r"raz[aã]o\s+social[:\s]*(.+)",
        r"emitente[:\s]*(.+)",
        r"fornecedor[:\s]*(.+)",
    ]
    for line in lines[:80]:
        for pattern in candidate_patterns:
            match = re.search(pattern, line, flags=re.IGNORECASE)
            if match:
                value = match.group(1).strip(" :-")
                if len(value) > 2:
                    return value

    if cnpj:
        for index, line in enumerate(lines):
            if cnpj in line:
                for back in range(1, 4):
                    if index - back < 0:
                        break
                    candidate = lines[index - back]
                    if (
                        len(candidate) > 3
                        and not re.match(r"^[0-9\W]+$", candidate)
                        and not re.search(
                            r"endere[cç]o|cep|inscri[cç][aã]o|telefone",
                            candidate,
                            re.IGNORECASE,
                        )
                    ):
                        return candidate
                break
    return None

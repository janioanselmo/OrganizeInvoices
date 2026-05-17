"""Smoke tests for the OrganizeInvoices parsers.

The script handles real fiscal documents, so a regression in a regex would
silently start filing invoices under the wrong date or CNPJ. These tests
lock down the current behavior of each pure parser. They do NOT require a
real PDF — we feed the parsers raw text.
"""

from datetime import datetime

from parsers import (
    _validate_cnpj_digits,
    extract_cnpj,
    extract_date,
    extract_invoice_number,
    parse_date,
    sanitize_filename,
)


# ----------------------------------------------------------------------
# parse_date
# ----------------------------------------------------------------------
class TestParseDate:
    def test_brazilian_slash(self):
        assert parse_date("15/03/2024") == datetime(2024, 3, 15)

    def test_brazilian_dash(self):
        assert parse_date("15-03-2024") == datetime(2024, 3, 15)

    def test_brazilian_dot(self):
        assert parse_date("15.03.2024") == datetime(2024, 3, 15)

    def test_iso(self):
        assert parse_date("2024-03-15") == datetime(2024, 3, 15)

    def test_iso_slash(self):
        assert parse_date("2024/03/15") == datetime(2024, 3, 15)

    def test_invalid_text(self):
        assert parse_date("nada-aqui") is None

    def test_implausible_old_year(self):
        # Year < 1990 is rejected to avoid matching random sequences.
        assert parse_date("15/03/1899") is None

    def test_implausible_future_year(self):
        far_future = datetime.now().year + 5
        assert parse_date(f"15/03/{far_future}") is None


# ----------------------------------------------------------------------
# extract_date
# ----------------------------------------------------------------------
class TestExtractDate:
    def test_explicit_label_wins(self):
        text = (
            "Vencimento: 30/04/2024\n"
            "Data de emissão: 15/03/2024\n"
            "Outro campo qualquer 10/01/2023"
        )
        assert extract_date(text) == datetime(2024, 3, 15)

    def test_fallback_skips_due_date(self):
        # No explicit "emissão" label; fallback must skip dates preceded by
        # "vencimento".
        text = "Vencimento: 30/04/2024\nReferente a 15/03/2024"
        assert extract_date(text) == datetime(2024, 3, 15)

    def test_returns_none_when_no_plausible_date(self):
        assert extract_date("Documento sem nenhuma data dentro") is None

    def test_emissao_label_with_dot(self):
        text = "Dt. Emissão: 02/02/2024"
        assert extract_date(text) == datetime(2024, 2, 2)


# ----------------------------------------------------------------------
# CNPJ
# ----------------------------------------------------------------------
class TestCnpj:
    # 11.222.333/0001-81 is a well-known valid CNPJ check-digit example.
    VALID_PUNCT = "11.222.333/0001-81"
    VALID_DIGITS = "11222333000181"

    def test_validate_punctuated(self):
        assert _validate_cnpj_digits(self.VALID_DIGITS) is True

    def test_validate_invalid_dv(self):
        # Same numbers, wrong final check digit.
        assert _validate_cnpj_digits("11222333000182") is False

    def test_validate_repeated_digits_rejected(self):
        # All zeros / all ones must NOT be considered valid even though the
        # arithmetic check digit happens to match.
        assert _validate_cnpj_digits("00000000000000") is False
        assert _validate_cnpj_digits("11111111111111") is False

    def test_extract_with_punctuation(self):
        text = f"Emitente Ltda CNPJ {self.VALID_PUNCT} Rua A 123"
        assert extract_cnpj(text) == self.VALID_PUNCT

    def test_extract_without_punctuation_returns_formatted(self):
        text = f"CNPJ {self.VALID_DIGITS} acabou."
        assert extract_cnpj(text) == self.VALID_PUNCT

    def test_extract_rejects_invalid_14_digit_run(self):
        # Looks like a 14-digit number (e.g. a slice of a payment-slip barcode)
        # but has invalid check digits — must be rejected.
        text = "Boleto 12345678901234 vencimento"
        assert extract_cnpj(text) is None


# ----------------------------------------------------------------------
# invoice number
# ----------------------------------------------------------------------
class TestExtractInvoiceNumber:
    def test_nfe_pattern(self):
        text = "NF-e Nº: 000.123.456 emitida em..."
        assert extract_invoice_number(text) == "000.123.456"

    def test_nota_fiscal_eletronica(self):
        text = "Nota fiscal eletrônica nº 987654"
        assert extract_invoice_number(text) == "987654"

    def test_generic_numero_da_nota(self):
        text = "Número da nota: AB-123/4"
        assert extract_invoice_number(text) == "AB-123/4"

    def test_no_match_returns_none(self):
        assert extract_invoice_number("Texto sem indicação de número") is None


# ----------------------------------------------------------------------
# sanitize_filename
# ----------------------------------------------------------------------
class TestSanitizeFilename:
    def test_keeps_simple_text(self):
        assert sanitize_filename("EmpresaTeste") == "EmpresaTeste"

    def test_replaces_spaces_with_underscore(self):
        assert sanitize_filename("Empresa Teste LTDA") == "Empresa_Teste_LTDA"

    def test_strips_forbidden_chars(self):
        assert sanitize_filename("a/b\\c:d*e?f") == "abcdef"

    def test_empty_falls_back(self):
        assert sanitize_filename("") == "SEM_NOME"
        assert sanitize_filename("   ") == "SEM_NOME"

    def test_caps_length(self):
        assert len(sanitize_filename("a" * 200)) == 80

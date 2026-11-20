"""Utilidades de texto, fechas y nombres de archivo."""

from __future__ import annotations

import re
import sys
import unicodedata
from datetime import datetime

import pandas as pd

INVALID_FILENAME_CHARS = '<>:"/\\|?*'

RESERVED_WINDOWS_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
}


def configure_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def normalize_text(value) -> str:
    text = str(value or "").strip().upper()
    text = "".join(c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn")
    return " ".join(text.split())


def safe_filename(text: str, max_stem: int = 120) -> str:
    text = str(text or "").strip()
    text = "".join(c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn")
    cleaned = "".join(
        "_" if (c in INVALID_FILENAME_CHARS or ord(c) < 32) else c for c in text
    )
    cleaned = " ".join(cleaned.split()).strip(" .")

    if max_stem > 0 and len(cleaned) > max_stem:
        cleaned = cleaned[:max_stem].rstrip(" .")

    if cleaned.upper() in RESERVED_WINDOWS_NAMES:
        cleaned = f"_{cleaned}"

    return cleaned or "sin_nombre"


def classify_profile(raw_profile) -> str | None:
    profile = normalize_text(raw_profile)
    if "AUXILIAR" in profile:
        return "AUXILIAR"
    if "ENFERM" in profile:
        return "ENFERMERO"
    if "MEDIC" in profile:
        return "MEDICO"
    return None


def clean_document(doc_value) -> str:
    if doc_value is None:
        return ""
    if not isinstance(doc_value, str) and pd.isna(doc_value):
        return ""

    if isinstance(doc_value, (int, float)):
        try:
            if float(doc_value).is_integer():
                return str(int(doc_value))
        except Exception:
            pass
    elif hasattr(doc_value, "item"):
        try:
            return str(int(doc_value))
        except Exception:
            pass

    text = str(doc_value).strip().replace(" ", "")
    if text.endswith(".0"):
        text = text[:-2]
    return "".join(ch for ch in text if ch.isdigit())


def format_date_ddmmyyyy(value) -> str:
    """Convierte la fecha del Excel al formato DD/MM/AAAA del sitio."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""

    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")

    if hasattr(value, "to_pydatetime"):
        return value.to_pydatetime().strftime("%d/%m/%Y")

    text = str(value).strip()
    if not text:
        return ""

    if re.match(r"^\d{2}/\d{2}/\d{4}$", text):
        return text

    if re.match(r"^\d{4}-\d{2}-\d{2}", text):
        try:
            return datetime.strptime(text[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            pass

    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    if pd.notna(parsed):
        return parsed.strftime("%d/%m/%Y")

    return text


def parse_date_for_form(value) -> str:
    """Compatibilidad: devuelve AAAA-MM-DD."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if hasattr(value, "to_pydatetime"):
        return value.to_pydatetime().strftime("%Y-%m-%d")
    text = str(value).strip()
    if not text:
        return ""
    if len(text) >= 10 and text[0:4].isdigit() and text[4] in ("/", "-"):
        return text.replace("/", "-")[:10]
    parsed = pd.to_datetime(text, dayfirst=False, errors="coerce")
    if pd.notna(parsed):
        return parsed.strftime("%Y-%m-%d")
    return text.replace("/", "-")


def map_document_type_ruaf(raw_value, default: str = "5|CC") -> str:
    """Mapea el tipo de documento del Excel al valor del sitio RUAF.

    El sitio usa valores compuestos ``codigo|sigla``:
    5|CC, 4|CE, 6|PA, 3|TI, 2|RC, 1|MS, 7|AS, 10|CD, 12|CN,
    13|SC, 14|PE, 15|PT, 16|PC.
    """
    text = normalize_text(raw_value).replace(".", "").replace(" ", "")
    if not text:
        return default

    mapping = {
        "CC": "5|CC",
        "CCV": "5|CC",
        "CEDULACIUDADANIA": "5|CC",
        "CEDULADECIUDADANIA": "5|CC",
        "CE": "4|CE",
        "CX": "4|CE",
        "CEDULAEXTRANJERIA": "4|CE",
        "CEDULADEEXTRANJERIA": "4|CE",
        "PA": "6|PA",
        "PAS": "6|PA",
        "PASAPORTE": "6|PA",
        "TI": "3|TI",
        "TARJETAIDENTIDAD": "3|TI",
        "TARJETADEIDENTIDAD": "3|TI",
        "RC": "2|RC",
        "REGISTROCIVIL": "2|RC",
        "MS": "1|MS",
        "MENORSINIDENTIFICACION": "1|MS",
        "AS": "7|AS",
        "ADULTOSINIDENTIFICACION": "7|AS",
        "CD": "10|CD",
        "CARNETDIPLOMATICO": "10|CD",
        "CN": "12|CN",
        "CERTIFICADODENACIDOVIVO": "12|CN",
        "SC": "13|SC",
        "SALVACONDUCTODEPERMANENCIA": "13|SC",
        "PE": "14|PE",
        "PERMISOESPECIALDEPERMANENCIA": "14|PE",
        "PT": "15|PT",
        "PERMISOPORPROTECCIONTEMPORAL": "15|PT",
        "PC": "16|PC",
        "PERMISOESPECIALDEPERMANENCIATUTOR": "16|PC",
    }
    return mapping.get(text, default)

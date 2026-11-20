"""Lectura del Excel de caracterización para la consulta RUAF."""

from __future__ import annotations

import logging
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl.utils import get_column_letter

from perfiles import CAMPOS
from utils import (
    classify_profile,
    clean_document,
    format_date_ddmmyyyy,
    map_document_type_ruaf,
    normalize_text,
)


class ExcelError(Exception):
    """Error controlado al leer el Excel."""


class ColumnasFaltantes(ExcelError):
    def __init__(self, campos: list[str]):
        self.campos = campos
        super().__init__("Faltan columnas obligatorias: " + ", ".join(campos))


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip() == ""


def list_sheets(path: str | Path) -> list[str]:
    path = Path(path)
    if not path.exists():
        raise ExcelError(f"No existe el archivo: {path}")
    try:
        with pd.ExcelFile(path) as excel:
            return list(excel.sheet_names)
    except Exception as exc:
        raise ExcelError(f"No se pudo abrir el Excel: {exc}") from exc


def resolve_sheet(path: str | Path, sheet_name: str | int | None) -> str:
    sheets = list_sheets(path)
    if not sheets:
        raise ExcelError("El Excel no tiene hojas.")
    if sheet_name in (None, ""):
        return sheets[0]
    if isinstance(sheet_name, int):
        try:
            return sheets[sheet_name]
        except IndexError as exc:
            raise ExcelError(f"La hoja #{sheet_name} no existe.") from exc
    name = str(sheet_name)
    if name in sheets:
        return name
    normalized = {normalize_text(s): s for s in sheets}
    key = normalize_text(name)
    if key in normalized:
        return normalized[key]
    raise ExcelError(f"La hoja '{name}' no existe. Hojas disponibles: {', '.join(sheets)}")


def read_table(
    path: str | Path,
    sheet_name: str | int | None = None,
    header_row: int = 1,
) -> tuple[pd.DataFrame, str]:
    path = Path(path)
    if not path.exists():
        raise ExcelError(f"No existe el archivo: {path}")
    sheet = resolve_sheet(path, sheet_name)
    header_row = max(1, int(header_row or 1))
    try:
        df = pd.read_excel(path, sheet_name=sheet, header=header_row - 1, dtype=object)
    except Exception as exc:
        raise ExcelError(f"No se pudo leer la hoja '{sheet}': {exc}") from exc
    if not df.empty:
        df = df[df.notna().any(axis=1)]
    return df, sheet


def column_options(df: pd.DataFrame) -> list[dict[str, str]]:
    options: list[dict[str, str]] = []
    used: dict[str, int] = {}
    for index, original in enumerate(df.columns):
        name = str(original).strip()
        if not name or name.startswith("Unnamed:") or name.lower() == "nan":
            name = "(sin título)"
        count = used.get(name, 0) + 1
        used[name] = count
        suffix = f" ({count})" if count > 1 else ""
        letter = get_column_letter(index + 1)
        options.append(
            {
                "original": str(original),
                "nombre": name,
                "letra": letter,
                "display": f"{letter} · {name}{suffix}",
            }
        )
    return options


def detect_mapping(headers: list[str]) -> dict[str, str]:
    normalized: dict[str, str] = {}
    for header in headers:
        key = normalize_text(header)
        if key and key not in normalized:
            normalized[key] = header

    mapping: dict[str, str] = {}
    for campo in CAMPOS:
        chosen = None
        for candidate in campo.candidatos:
            key = normalize_text(candidate)
            if key in normalized:
                chosen = normalized[key]
                break
        if chosen is None:
            for candidate in campo.candidatos:
                key = normalize_text(candidate)
                for norm, original in normalized.items():
                    if key and key in norm:
                        chosen = original
                        break
                if chosen:
                    break
        if chosen:
            mapping[campo.clave] = chosen
    return mapping


def preview_rows(df: pd.DataFrame, limit: int = 5) -> list[list[str]]:
    rows: list[list[str]] = []
    for _, row in df.head(limit).iterrows():
        values = []
        for value in row.tolist():
            if _is_blank(value):
                values.append("")
            elif isinstance(value, (datetime, date)):
                values.append(value.strftime("%d/%m/%Y"))
            elif isinstance(value, float) and value.is_integer():
                values.append(str(int(value)))
            else:
                values.append(str(value))
        rows.append(values)
    return rows


def inspect(
    path: str | Path,
    sheet_name: str | int | None = None,
    header_row: int = 1,
) -> dict[str, Any]:
    df, sheet = read_table(path, sheet_name=sheet_name, header_row=header_row)
    options = column_options(df)
    headers = [option["original"] for option in options]
    return {
        "hoja": sheet,
        "hojas": list_sheets(path),
        "columnas": options,
        "mapeo_detectado": detect_mapping(headers),
        "vista_previa": preview_rows(df, 5),
        "total_filas": int(len(df)),
        "fila_encabezado": int(header_row),
    }


def load_users(
    path: str | Path,
    sheet_name: str | int | None = None,
    header_row: int = 1,
    column_map: dict[str, str] | None = None,
    selected_profiles: list[str] | None = None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, Any]]:
    """Lee el Excel y devuelve los usuarios en el ORDEN de las filas."""
    df, sheet = read_table(path, sheet_name=sheet_name, header_row=header_row)
    options = column_options(df)
    headers = [option["original"] for option in options]

    detected = detect_mapping(headers)
    mapping: dict[str, str] = {}
    for campo in CAMPOS:
        override = (column_map or {}).get(campo.clave)
        if override and override in headers:
            mapping[campo.clave] = override
        elif campo.clave in detected:
            mapping[campo.clave] = detected[campo.clave]

    missing = [campo.etiqueta for campo in CAMPOS if campo.obligatorio and campo.clave not in mapping]
    if missing:
        raise ColumnasFaltantes(missing)

    selected = {p.upper() for p in selected_profiles} if selected_profiles else None

    users: list[dict[str, Any]] = []
    incomplete: list[dict[str, Any]] = []
    duplicated: list[dict[str, Any]] = []
    seen: dict[str, int] = {}
    total = 0
    filtered = 0

    for idx, row in df.iterrows():
        values = [row[mapping[campo.clave]] for campo in CAMPOS if campo.clave in mapping]
        if all(_is_blank(value) for value in values):
            continue

        total += 1
        fila_excel = int(idx) + int(header_row) + 1

        doc = clean_document(row[mapping["documento"]])
        fecha = format_date_ddmmyyyy(row[mapping["fecha"]])
        nombre = "" if _is_blank(row[mapping["nombre"]]) else str(row[mapping["nombre"]]).strip()
        apellido = "" if _is_blank(row[mapping["apellido"]]) else str(row[mapping["apellido"]]).strip()
        perfil_raw = row[mapping["perfil"]]
        perfil_grupo = classify_profile(perfil_raw)

        tipo_raw = ""
        if "tipo_documento" in mapping:
            raw = row[mapping["tipo_documento"]]
            tipo_raw = "" if _is_blank(raw) else str(raw).strip()
        tipo_documento = map_document_type_ruaf(tipo_raw, default="5|CC")

        faltantes = []
        if not doc:
            faltantes.append("documento")
        if not fecha:
            faltantes.append("fecha de expedición")
        if not nombre:
            faltantes.append("nombres")
        if not apellido:
            faltantes.append("apellidos")
        if not perfil_grupo:
            faltantes.append(f"perfil ({perfil_raw!r})")

        if faltantes:
            incomplete.append(
                {
                    "fila": fila_excel,
                    "documento": doc or "",
                    "nombre": f"{nombre} {apellido}".strip(),
                    "motivo": "dato_incompleto",
                    "detalle": "Faltan: " + ", ".join(faltantes),
                }
            )
            continue

        if doc in seen:
            duplicated.append(
                {
                    "fila": fila_excel,
                    "documento": doc,
                    "nombre": f"{nombre} {apellido}".strip(),
                    "primera_fila": seen[doc],
                }
            )
            continue
        seen[doc] = fila_excel

        if selected and perfil_grupo not in selected:
            filtered += 1
            continue

        users.append(
            {
                "documento": doc,
                "fecha_expedicion": fecha,
                "nombre": nombre,
                "apellido": apellido,
                "perfil_raw": str(perfil_raw).strip() if not _is_blank(perfil_raw) else "",
                "perfil_grupo": perfil_grupo,
                "tipo_documento": tipo_documento,
                "tipo_documento_raw": tipo_raw,
                "fila_excel": fila_excel,
            }
        )

    stats = {
        "hoja": sheet,
        "fila_encabezado": int(header_row),
        "total_filas": total,
        "validos": len(users),
        "incompletos": len(incomplete),
        "duplicados": len(duplicated),
        "filtrados_por_perfil": filtered,
        "columnas": {campo.clave: mapping.get(campo.clave) for campo in CAMPOS},
        "duplicados_detalle": duplicated[:20],
    }

    logging.info(
        "Excel '%s' hoja '%s': %s filas, %s válidos, %s incompletos, %s duplicados",
        Path(path).name, sheet, total, len(users), len(incomplete), len(duplicated),
    )
    return users, incomplete, stats

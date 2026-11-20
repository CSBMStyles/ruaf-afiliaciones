"""Extrae los datos del PDF oficial del RUAF y genera un Excel consolidado.

Lee los PDFs descargados por ``obtener-ruaf-sispro`` (Reporte Detallado del
RUAF), extrae las secciones de información básica, salud, pensiones, ARL,
compensación familiar, cesantías, pensionados y programas de asistencia social,
y arma un libro de Excel con:

- ``Empleados``: una fila por empleado con los datos clave.
- ``Detalle``: una fila por sección y afiliación.
- ``Resumen``: conteos y cobertura.

Uso:
    python extraer_datos.py --carpeta "../RUAF/AFILIACIONES"
    python extraer_datos.py --carpeta "../RUAF/AFILIACIONES" --pptx
"""

from __future__ import annotations

import argparse
import re
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    import pymupdf as fitz  # PyMuPDF
except Exception:  # pragma: no cover
    try:
        import fitz
    except Exception as exc:
        raise SystemExit("Falta PyMuPDF. Instala con: pip install PyMuPDF") from exc

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from config import load_settings
from sources import load_users

SECTION_LABELS = {
    "basica": "INFORMACION BASICA",
    "salud": "AFILIACION A SALUD",
    "pension": "AFILIACION A PENSIONES",
    "arl": "AFILIACION A RIESGOS LABORALES",
    "compensacion": "AFILIACION A COMPENSACION FAMILIAR",
    "cesantias": "AFILIACION A CESANTIAS",
    "pensionados": "PENSIONADOS",
    "programas": "VINCULACION A PROGRAMAS DE ASISTENCIA SOCIAL",
}

SECTION_TITLES = {
    "salud": "Salud",
    "pension": "Pensión",
    "arl": "ARL",
    "compensacion": "Compensación familiar",
    "cesantias": "Cesantías",
    "pensionados": "Pensionados",
    "programas": "Programas de asistencia social",
}

SECTION_ORDER = ["salud", "pension", "arl", "compensacion", "cesantias", "pensionados", "programas"]

HEADER_ALIASES = {
    "administradora": ("ADMINISTRADORA", "ENTIDAD", "EPS", "AFP", "ARL", "CAJA"),
    "regimen": ("REGIMEN",),
    "fecha": (
        "FECHA AFILIACION",
        "FECHA DE AFILIACION",
        "FECHA VINCULACION",
        "FECHA DE VINCULACION",
        "FECHA",
    ),
    "estado": ("ESTADO DE AFILIACION", "ESTADO"),
    "tipo": ("TIPO DE AFILIADO", "TIPO AFILIADO"),
    "municipio": ("DEPARTAMENTO -> MUNICIPIO", "MUNICIPIO LABORA", "MUNICIPIO"),
    "actividad": ("ACTIVIDAD ECONOMICA",),
    "programa": ("PROGRAMA",),
}

EMPRESA_COLUMNS = [
    "Fila Excel",
    "Documento",
    "Tipo documento",
    "Apellidos",
    "Nombres",
    "Perfil",
    "Fecha expedición",
    "Fecha de corte",
    "Fecha de consulta",
    "Sexo",
    "Estado general",
    "Salud - Administradora",
    "Salud - Régimen",
    "Salud - Fecha afiliación",
    "Salud - Estado",
    "Salud - Tipo afiliado",
    "Salud - Departamento/Municipio",
    "Pensión - Régimen",
    "Pensión - Administradora",
    "Pensión - Fecha afiliación",
    "Pensión - Estado",
    "ARL - Administradora",
    "ARL - Fecha afiliación",
    "ARL - Estado",
    "ARL - Actividad económica",
    "ARL - Municipio",
    "Compensación - Administradora",
    "Compensación - Estado",
    "Cesantías - Administradora",
    "Cesantías - Estado",
    "Pensionado - Estado",
    "Programas - Detalle",
    "Observaciones",
    "PDF",
]

DETAIL_COLUMNS = [
    "Documento",
    "Nombre",
    "Perfil",
    "Sección",
    "Administradora",
    "Régimen",
    "Fecha afiliación",
    "Estado",
    "Tipo afiliado",
    "Departamento/Municipio",
    "Actividad económica",
    "Municipio labora",
    "Observación",
    "PDF",
]


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def _norm(value) -> str:
    text = str(value or "").strip().upper()
    text = "".join(c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn")
    return " ".join(text.split())


def _clean(value) -> str:
    return " ".join(str(value or "").split())


def _digits(value) -> str:
    return "".join(ch for ch in str(value or "") if ch.isdigit())


def _norm_date(value) -> str:
    text = _clean(value)
    if not text:
        return ""
    match = re.match(r"^(\d{4})-(\d{2})-(\d{2})", text)
    if match:
        return f"{match.group(3)}/{match.group(2)}/{match.group(1)}"
    match = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", text)
    if match:
        return f"{int(match.group(1)):02d}/{int(match.group(2)):02d}/{match.group(3)}"
    return text


def _norm_datetime(value) -> str:
    text = _clean(value)
    if not text:
        return ""
    for fmt in ("%m/%d/%Y %I:%M:%S %p", "%d/%m/%Y %I:%M:%S %p"):
        try:
            return datetime.strptime(text, fmt).strftime("%d/%m/%Y %I:%M %p")
        except ValueError:
            pass
    return text


def _section_from_text(value) -> str | None:
    text = _norm(value)
    if not text:
        return None
    for key, label in SECTION_LABELS.items():
        if text.startswith(label) or label in text:
            return key
    return None


def _header_key(value) -> str | None:
    text = _norm(value)
    if not text or "FECHA DE CORTE" in text:
        return None
    for key, aliases in HEADER_ALIASES.items():
        if any(alias in text for alias in aliases):
            return key
    return None


def _map_row(headers, data) -> dict[str, str]:
    row: dict[str, str] = {}
    for header, value in zip(headers, data):
        key = _header_key(header)
        text = _clean(value)
        if not key or not text:
            continue
        row[key] = _norm_date(text) if key == "fecha" else text
    if not row:
        values = [_clean(v) for v in data if _clean(v)]
        if values:
            row["detalle"] = " | ".join(values)
    return row


def _is_active(row: dict[str, str]) -> bool:
    return "ACTIV" in _norm(row.get("estado", ""))


def _pick(rows: list[dict[str, str]]) -> dict[str, str]:
    if not rows:
        return {}
    for row in rows:
        if _is_active(row):
            return row
    return rows[0]


def _value(row: dict[str, str], key: str, default: str = "No reportado") -> str:
    return _clean(row.get(key, "")) or default


# ---------------------------------------------------------------------------
# Lectura del PDF
# ---------------------------------------------------------------------------

def parse_pdf(path: str | Path) -> dict[str, Any]:
    path = Path(path)
    doc = fitz.open(path)
    sections: dict[str, list[dict[str, str]]] = {key: [] for key in SECTION_ORDER}
    seen_sections: set[str] = set()
    empty_sections: set[str] = set()
    events: list[dict[str, Any]] = []
    basic_rows: list[list[Any]] = []

    for page_no, page in enumerate(doc):
        blocks = page.get_text("blocks")

        # Primero los encabezados y mensajes de texto de la página.
        for block in blocks:
            text = _clean(block[4])
            if not text:
                continue
            if text.lower().startswith("no se han reportado"):
                events.append({"page": page_no, "y": block[1], "type": "empty", "key": None})
                continue
            key = _section_from_text(text)
            if key and key != "basica":
                events.append({"page": page_no, "y": block[1], "type": "heading", "key": key})

        # Luego las tablas.
        for table in page.find_tables().tables:
            rows = table.extract()
            if not rows:
                continue
            first = _clean(rows[0][0]) if rows[0] else ""
            key = _section_from_text(first)

            if key == "basica":
                basic_rows = rows
                continue

            if not key:
                key = _heading_above(events, page_no, table.bbox[1])

            if not key:
                continue

            seen_sections.add(key)
            events.append({"page": page_no, "y": table.bbox[1], "type": "heading", "key": key})

            headers = rows[1] if len(rows) > 1 else []
            data_rows = rows[2:] if len(rows) > 2 else []
            for data in data_rows:
                if not any(_clean(cell) for cell in data):
                    continue
                mapped = _map_row(headers, data)
                if mapped:
                    sections[key].append(mapped)

    # Asignar los "No se han reportado" al último encabezado anterior.
    current = None
    for event in sorted(events, key=lambda item: (item["page"], item["y"])):
        if event["type"] == "heading":
            current = event["key"]
        elif event["type"] == "empty" and current:
            empty_sections.add(current)

    for key, rows in sections.items():
        if rows:
            empty_sections.discard(key)

    text_all = "\n".join(page.get_text() for page in doc)

    fecha_corte = ""
    dates = re.findall(r"\b(20\d{2}-\d{2}-\d{2})\b", text_all)
    if dates:
        fecha_corte = Counter(dates).most_common(1)[0][0]

    fecha_consulta = ""
    match = re.search(
        r"Fecha:\s*(\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s*[AP]M)",
        text_all,
        re.IGNORECASE,
    )
    if match:
        fecha_consulta = _norm_datetime(match.group(1))

    documento = ""
    nombres = ""
    apellidos = ""
    sexo = ""

    if len(basic_rows) > 2:
        data = basic_rows[2]
        if data:
            documento = _digits(data[0])
        if len(data) > 2:
            nombres = " ".join(_clean(data[i]) for i in (1, 2) if i < len(data) and _clean(data[i]))
        if len(data) > 5:
            apellidos = " ".join(_clean(data[i]) for i in (4, 5) if i < len(data) and _clean(data[i]))
        if len(data) > 7:
            sexo = _clean(data[7])

    if not documento:
        match = re.search(r"(\d{6,})", path.stem)
        if match:
            documento = match.group(1)

    if not apellidos or not nombres:
        # Respaldo desde el nombre del archivo: APELLIDOS NOMBRES - CEDULA.pdf
        stem = path.stem
        if " - " in stem:
            left = stem.split(" - ")[0]
            if not apellidos:
                apellidos = left
            if not nombres:
                nombres = ""

    doc.close()

    return {
        "pdf": str(path),
        "documento": documento,
        "nombres": nombres,
        "apellidos": apellidos,
        "sexo": sexo,
        "fecha_corte": _norm_date(fecha_corte),
        "fecha_consulta": fecha_consulta,
        "secciones": sections,
        "no_reportado": sorted(empty_sections),
    }


def _heading_above(events: list[dict[str, Any]], page_no: int, y: float) -> str | None:
    best = None
    for event in sorted(events, key=lambda item: (item["page"], item["y"])):
        if event["type"] != "heading":
            continue
        if event["page"] < page_no or (event["page"] == page_no and event["y"] <= y):
            best = event["key"]
    return best


# ---------------------------------------------------------------------------
# Construcción de registros
# ---------------------------------------------------------------------------

def _load_metadata(excel_path, sheet_name=None, header_row=1, column_map=None) -> dict[str, dict]:
    if not excel_path:
        return {}
    path = Path(excel_path)
    if not path.exists():
        return {}
    try:
        users, _, _ = load_users(
            path,
            sheet_name=sheet_name,
            header_row=header_row,
            column_map=column_map or {},
        )
    except Exception as exc:
        print(f"[WARN] No se pudo leer el Excel de caracterización: {exc}")
        return {}
    return {user["documento"]: user for user in users}


def _observaciones(secciones: dict[str, list[dict[str, str]]]) -> str:
    """Novedades críticas: salud, pensión y ARL.

    Las secciones complementarias (compensación, cesantías, pensionados y
    programas) quedan en sus propias columnas y en la hoja Detalle.
    """
    notes = []
    for key in ("salud", "pension", "arl"):
        rows = secciones.get(key, [])
        title = SECTION_TITLES[key]
        if not rows:
            notes.append(f"Sin {title.lower()}")
            continue
        if len(rows) > 1:
            notes.append(f"{title}: {len(rows)} registros (ver Detalle)")
        if not any(_is_active(row) for row in rows):
            estado = _value(rows[0], "estado")
            notes.append(f"{title}: {estado}")
    return "; ".join(notes)


def _estado_general(secciones: dict[str, list[dict[str, str]]]) -> str:
    salud = secciones.get("salud", [])
    pension = secciones.get("pension", [])
    arl = secciones.get("arl", [])
    if not salud and not pension and not arl:
        return "Sin afiliaciones"
    if salud and pension and arl and all(
        any(_is_active(row) for row in rows) for rows in (salud, pension, arl)
    ):
        return "Completo"
    return "Parcial"


def build_flat_record(parsed: dict[str, Any], meta: dict | None = None) -> dict[str, Any]:
    meta = meta or {}
    secciones = parsed.get("secciones", {})
    salud = _pick(secciones.get("salud", []))
    pension = _pick(secciones.get("pension", []))
    arl = _pick(secciones.get("arl", []))
    compensacion = _pick(secciones.get("compensacion", []))
    cesantias = _pick(secciones.get("cesantias", []))
    pensionado = _pick(secciones.get("pensionados", []))
    programas = secciones.get("programas", [])

    programas_detalle = "; ".join(
        " | ".join(v for v in row.values() if v) for row in programas
    ) or "No reportado"

    return {
        "fila_excel": meta.get("fila_excel", ""),
        "documento": parsed.get("documento", "") or meta.get("documento", ""),
        "tipo_documento": meta.get("tipo_documento_raw", "") or meta.get("tipo_documento", ""),
        "apellidos": meta.get("apellido", "") or parsed.get("apellidos", ""),
        "nombres": meta.get("nombre", "") or parsed.get("nombres", ""),
        "perfil": meta.get("perfil_grupo", "") or meta.get("perfil_raw", ""),
        "fecha_expedicion": meta.get("fecha_expedicion", ""),
        "fecha_corte": parsed.get("fecha_corte", ""),
        "fecha_consulta": parsed.get("fecha_consulta", ""),
        "sexo": parsed.get("sexo", ""),
        "estado_general": _estado_general(secciones),
        "salud_administradora": _value(salud, "administradora"),
        "salud_regimen": _value(salud, "regimen"),
        "salud_fecha": _value(salud, "fecha"),
        "salud_estado": _value(salud, "estado"),
        "salud_tipo": _value(salud, "tipo"),
        "salud_municipio": _value(salud, "municipio"),
        "pension_regimen": _value(pension, "regimen"),
        "pension_administradora": _value(pension, "administradora"),
        "pension_fecha": _value(pension, "fecha"),
        "pension_estado": _value(pension, "estado"),
        "arl_administradora": _value(arl, "administradora"),
        "arl_fecha": _value(arl, "fecha"),
        "arl_estado": _value(arl, "estado"),
        "arl_actividad": _value(arl, "actividad"),
        "arl_municipio": _value(arl, "municipio"),
        "compensacion_administradora": _value(compensacion, "administradora"),
        "compensacion_estado": _value(compensacion, "estado"),
        "cesantias_administradora": _value(cesantias, "administradora"),
        "cesantias_estado": _value(cesantias, "estado"),
        "pensionado_estado": _value(pensionado, "estado"),
        "programas_detalle": programas_detalle,
        "observaciones": _observaciones(secciones),
        "pdf": parsed.get("pdf", ""),
        "secciones": secciones,
        "no_reportado": parsed.get("no_reportado", []),
    }


def construir_registros(
    output_dir,
    excel_path=None,
    sheet_name=None,
    header_row=1,
    column_map=None,
) -> list[dict[str, Any]]:
    output_dir = Path(output_dir)
    if not output_dir.exists():
        raise FileNotFoundError(f"No existe la carpeta de PDFs: {output_dir}")

    metadata = _load_metadata(excel_path, sheet_name, header_row, column_map)
    pdfs = sorted(
        p for p in output_dir.rglob("*.pdf")
        if p.is_file() and not p.name.upper().startswith("REPORTE_")
    )

    registros: dict[str, dict[str, Any]] = {}
    for pdf in pdfs:
        try:
            parsed = parse_pdf(pdf)
        except Exception as exc:
            print(f"[WARN] No se pudo leer {pdf.name}: {exc}")
            continue
        documento = parsed.get("documento", "")
        if not documento:
            print(f"[WARN] PDF sin documento identificable: {pdf.name}")
            continue
        meta = metadata.get(documento, {})
        flat = build_flat_record(parsed, meta)
        previous = registros.get(documento)
        if previous is None or pdf.stat().st_mtime >= Path(previous["pdf"]).stat().st_mtime:
            if previous is not None:
                flat["observaciones"] = (
                    flat["observaciones"] + "; PDF duplicado, se usó el más reciente"
                ).strip("; ")
            registros[documento] = flat

    def sort_key(record):
        fila = record.get("fila_excel", "")
        try:
            return (0, int(fila))
        except (TypeError, ValueError):
            return (1, record.get("documento", ""))

    return sorted(registros.values(), key=sort_key)


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="142749")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(vertical="top", wrap_text=True)


def _style_header(ws, columns):
    ws.append(columns)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def _style_body(ws):
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = WRAP
            cell.border = BORDER


def _employee_values(record: dict[str, Any]) -> list:
    return [
        record.get("fila_excel", ""),
        record.get("documento", ""),
        record.get("tipo_documento", ""),
        record.get("apellidos", ""),
        record.get("nombres", ""),
        record.get("perfil", ""),
        record.get("fecha_expedicion", ""),
        record.get("fecha_corte", ""),
        record.get("fecha_consulta", ""),
        record.get("sexo", ""),
        record.get("estado_general", ""),
        record.get("salud_administradora", ""),
        record.get("salud_regimen", ""),
        record.get("salud_fecha", ""),
        record.get("salud_estado", ""),
        record.get("salud_tipo", ""),
        record.get("salud_municipio", ""),
        record.get("pension_regimen", ""),
        record.get("pension_administradora", ""),
        record.get("pension_fecha", ""),
        record.get("pension_estado", ""),
        record.get("arl_administradora", ""),
        record.get("arl_fecha", ""),
        record.get("arl_estado", ""),
        record.get("arl_actividad", ""),
        record.get("arl_municipio", ""),
        record.get("compensacion_administradora", ""),
        record.get("compensacion_estado", ""),
        record.get("cesantias_administradora", ""),
        record.get("cesantias_estado", ""),
        record.get("pensionado_estado", ""),
        record.get("programas_detalle", ""),
        record.get("observaciones", ""),
        record.get("pdf", ""),
    ]


def generar_excel(registros: list[dict[str, Any]], output_path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Empleados"
    _style_header(ws, EMPRESA_COLUMNS)
    for record in registros:
        ws.append(_employee_values(record))
    _style_body(ws)

    widths = [9, 14, 9, 24, 24, 12, 12, 12, 16, 6, 13,
              34, 13, 12, 13, 13, 22,
              20, 38, 12, 15,
              34, 12, 13, 24, 18,
              28, 15, 28, 15, 15, 30, 45, 55]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for row in ws.iter_rows(min_row=2, min_col=34, max_col=34):
        for cell in row:
            if cell.value:
                cell.hyperlink = "file://" + str(Path(cell.value).resolve())
                cell.style = "Hyperlink"

    detalle = wb.create_sheet("Detalle")
    _style_header(detalle, DETAIL_COLUMNS)
    for record in registros:
        nombre = f"{record.get('apellidos', '')} {record.get('nombres', '')}".strip()
        for key in SECTION_ORDER:
            rows = record.get("secciones", {}).get(key, [])
            if not rows:
                detalle.append([
                    record.get("documento", ""), nombre, record.get("perfil", ""),
                    SECTION_TITLES[key], "No reportado", "", "", "No reportado", "", "", "", "", "", record.get("pdf", ""),
                ])
                continue
            for row in rows:
                detalle.append([
                    record.get("documento", ""), nombre, record.get("perfil", ""),
                    SECTION_TITLES[key],
                    _value(row, "administradora", ""),
                    _value(row, "regimen", ""),
                    _value(row, "fecha", ""),
                    _value(row, "estado", ""),
                    _value(row, "tipo", ""),
                    _value(row, "municipio", ""),
                    _value(row, "actividad", ""),
                    _value(row, "municipio", ""),
                    _value(row, "detalle", ""),
                    record.get("pdf", ""),
                ])
    _style_body(detalle)
    for index, width in enumerate([14, 32, 12, 26, 38, 14, 12, 16, 14, 24, 24, 20, 30, 50], start=1):
        detalle.column_dimensions[get_column_letter(index)].width = width

    resumen = wb.create_sheet("Resumen")
    _style_header(resumen, ["Indicador", "Valor"])
    total = len(registros)
    counts = {
        "Total de empleados con PDF": total,
        "Con salud reportada": sum(1 for r in registros if r.get("salud_administradora") != "No reportado"),
        "Sin salud reportada": sum(1 for r in registros if r.get("salud_administradora") == "No reportado"),
        "Con pensión reportada": sum(1 for r in registros if r.get("pension_administradora") != "No reportado"),
        "Sin pensión reportada": sum(1 for r in registros if r.get("pension_administradora") == "No reportado"),
        "Con ARL reportada": sum(1 for r in registros if r.get("arl_administradora") != "No reportado"),
        "Sin ARL reportada": sum(1 for r in registros if r.get("arl_administradora") == "No reportado"),
        "Con compensación familiar": sum(1 for r in registros if r.get("compensacion_administradora") != "No reportado"),
        "Con cesantías": sum(1 for r in registros if r.get("cesantias_administradora") != "No reportado"),
        "Con pensionado": sum(1 for r in registros if r.get("pensionado_estado") != "No reportado"),
        "Con programas de asistencia": sum(1 for r in registros if r.get("programas_detalle") != "No reportado"),
        "Estado general Completo": sum(1 for r in registros if r.get("estado_general") == "Completo"),
        "Estado general Parcial": sum(1 for r in registros if r.get("estado_general") == "Parcial"),
        "Estado general Sin afiliaciones": sum(1 for r in registros if r.get("estado_general") == "Sin afiliaciones"),
        "Empleados con observaciones": sum(1 for r in registros if r.get("observaciones")),
    }
    for key, value in counts.items():
        resumen.append([key, value])
    _style_body(resumen)
    resumen.column_dimensions["A"].width = 36
    resumen.column_dimensions["B"].width = 12

    resumen.append([])
    resumen.append(["Sección", "Con reporte", "Sin reporte", "Activos"])
    header_row = resumen.max_row
    for cell in resumen[header_row]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    for key in SECTION_ORDER:
        con = sum(1 for r in registros if r.get("secciones", {}).get(key))
        sin = total - con
        activos = sum(
            1 for r in registros
            if any(_is_active(row) for row in r.get("secciones", {}).get(key, []))
        )
        resumen.append([SECTION_TITLES[key], con, sin, activos])
    _style_body(resumen)

    wb.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv=None) -> int:
    settings = load_settings()
    parser = argparse.ArgumentParser(description="Extrae datos de los PDFs del RUAF a Excel.")
    parser.add_argument("--carpeta", default=str(settings.output_dir), help="Carpeta con los PDFs")
    parser.add_argument("--excel", default=str(settings.excel_path), help="Excel de caracterización")
    parser.add_argument("--hoja", default=settings.sheet_name, help="Hoja del Excel")
    parser.add_argument("--fila-encabezado", type=int, default=settings.header_row, help="Fila de encabezados")
    parser.add_argument("--salida", default="", help="Ruta del Excel de salida")
    parser.add_argument("--pptx", action="store_true", help="Generar también el informe PPTX")
    parser.add_argument("--pptx-salida", default="", help="Ruta del PPTX de salida")
    args = parser.parse_args(argv)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    reports_dir = Path(settings.reports_dir)
    salida = Path(args.salida) if args.salida else reports_dir / f"DATOS_RUAF_{timestamp}.xlsx"

    registros = construir_registros(
        args.carpeta,
        excel_path=args.excel,
        sheet_name=args.hoja,
        header_row=args.fila_encabezado,
        column_map=settings.column_map,
    )
    generar_excel(registros, salida)
    print(f"[OK] Empleados: {len(registros)}")
    print(f"[OK] Excel: {salida}")

    if args.pptx:
        from generar_pptx import generar_pptx

        pptx_path = Path(args.pptx_salida) if args.pptx_salida else reports_dir / f"INFORME_RUAF_{timestamp}.pptx"
        generar_pptx(registros, pptx_path)
        print(f"[OK] PPTX: {pptx_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Genera un informe ejecutivo en PPTX a partir de los datos del RUAF.

Toma la lista de registros que produce ``extraer_datos.py`` y arma una
presentación con portada, resumen ejecutivo, cobertura, estados, perfiles,
principales EPS/AFP/ARL, empleados con novedades y conclusiones.

Uso:
    python generar_pptx.py --carpeta "../RUAF/AFILIACIONES" --salida "../RUAF/INFORME_RUAF.pptx"
"""

from __future__ import annotations

import argparse
import os
from collections import Counter
from datetime import datetime
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

EMPRESA_INFORME = os.getenv("EMPRESA_INFORME", "Mi empresa")

AZUL = RGBColor(0x0B, 0x5F, 0xA5)
AZUL_OSCURO = RGBColor(0x08, 0x3C, 0x6B)
VERDE = RGBColor(0x2E, 0x7D, 0x32)
AMARILLO = RGBColor(0xF2, 0xA9, 0x00)
ROJO = RGBColor(0xC5, 0x22, 0x1F)
GRIS = RGBColor(0x5F, 0x63, 0x68)
GRIS_CLARO = RGBColor(0xF2, 0xF5, 0xF9)
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)
NEGRO = RGBColor(0x20, 0x21, 0x24)

ANCHO = 13.333
ALTO = 7.5


# ---------------------------------------------------------------------------
# Utilidades de diseño
# ---------------------------------------------------------------------------

def _textbox(slide, left, top, width, height, text, size=14, bold=False,
             color=NEGRO, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    frame = box.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    paragraph = frame.paragraphs[0]
    paragraph.text = text
    paragraph.alignment = align
    run = paragraph.runs[0]
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def _rect(slide, left, top, width, height, color, line=None):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
    shape.shadow.inherit = False
    return shape


def _base_slide(prs, title, subtitle=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, Inches(ANCHO), Inches(0.18)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = AZUL
    bar.line.fill.background()

    _textbox(slide, 0.6, 0.35, 12.1, 0.55, title, size=25, bold=True, color=AZUL_OSCURO)
    if subtitle:
        _textbox(slide, 0.6, 0.92, 12.1, 0.3, subtitle, size=11, color=GRIS)
    _textbox(
        slide, 0.6, 7.12, 12.1, 0.25,
        f"{EMPRESA_INFORME}  ·  RUAF - Afiliaciones  ·  Informe generado automáticamente",
        size=8, color=GRIS,
    )
    return slide


def _footer_date(prs, fecha):
    # Se usa en la portada; las demás diapositivas llevan el pie fijo.
    pass


def _kpi(slide, left, top, width, height, value, label, color=AZUL):
    _rect(slide, left, top, width, height, GRIS_CLARO, line=RGBColor(0xDD, 0xE3, 0xEA))
    _textbox(slide, left + 0.12, top + 0.12, width - 0.24, 0.55, str(value), size=28, bold=True, color=color)
    _textbox(slide, left + 0.12, top + 0.72, width - 0.24, 0.42, label, size=9, color=GRIS)


def _style_chart(chart, title, legend=True):
    chart.has_title = True
    chart.chart_title.text_frame.text = title
    for paragraph in chart.chart_title.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(13)
            run.font.bold = True
            run.font.color.rgb = AZUL_OSCURO
    chart.font.size = Pt(9)
    chart.font.color.rgb = NEGRO
    chart.has_legend = legend
    if legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(9)


def _color_series(series, color):
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = color


def _slide_table(slide, headers, rows, top=1.35, width=12.1, height=5.4, left=0.6):
    row_count = len(rows) + 1
    shape = slide.shapes.add_table(
        row_count, len(headers), Inches(left), Inches(top), Inches(width), Inches(height)
    )
    table = shape.table
    table.rows[0].height = Inches(0.38)
    for index in range(1, row_count):
        table.rows[index].height = Inches(0.32)
    for index, header in enumerate(headers):
        cell = table.cell(0, index)
        cell.text = header
        cell.fill.solid()
        cell.fill.fore_color.rgb = AZUL
        for paragraph in cell.text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            for run in paragraph.runs:
                run.font.size = Pt(9)
                run.font.bold = True
                run.font.color.rgb = BLANCO
    for row_index, row in enumerate(rows, start=1):
        for col_index, value in enumerate(row):
            cell = table.cell(row_index, col_index)
            cell.text = str(value)
            cell.fill.solid()
            cell.fill.fore_color.rgb = BLANCO if row_index % 2 else GRIS_CLARO
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.LEFT
                for run in paragraph.runs:
                    run.font.size = Pt(8)
                    run.font.color.rgb = NEGRO
    return table


# ---------------------------------------------------------------------------
# Estadísticas
# ---------------------------------------------------------------------------

def _has(value) -> bool:
    return bool(value) and str(value).strip().lower() != "no reportado"


def _active(value) -> bool:
    return "activ" in str(value or "").lower()


def _stats(registros):
    total = len(registros)
    con_salud = sum(1 for r in registros if _has(r.get("salud_administradora")))
    con_pension = sum(1 for r in registros if _has(r.get("pension_administradora")))
    con_arl = sum(1 for r in registros if _has(r.get("arl_administradora")))
    con_comp = sum(1 for r in registros if _has(r.get("compensacion_administradora")))
    con_ces = sum(1 for r in registros if _has(r.get("cesantias_administradora")))
    con_prog = sum(1 for r in registros if _has(r.get("programas_detalle")))
    novedades = [r for r in registros if r.get("observaciones")]
    return {
        "total": total,
        "con_salud": con_salud,
        "con_pension": con_pension,
        "con_arl": con_arl,
        "con_comp": con_comp,
        "con_ces": con_ces,
        "con_prog": con_prog,
        "novedades": novedades,
        "perfiles": Counter(r.get("perfil") or "Sin perfil" for r in registros),
        "estado_general": Counter(r.get("estado_general") or "Sin dato" for r in registros),
        "eps": Counter(r.get("salud_administradora") for r in registros if _has(r.get("salud_administradora"))),
        "afp": Counter(r.get("pension_administradora") for r in registros if _has(r.get("pension_administradora"))),
        "arl": Counter(r.get("arl_administradora") for r in registros if _has(r.get("arl_administradora"))),
    }


def _pct(part, total):
    return f"{(100 * part / total):.1f}%" if total else "0.0%"


# ---------------------------------------------------------------------------
# Diapositivas
# ---------------------------------------------------------------------------

def _portada(prs, total, fecha):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fondo = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(ANCHO), Inches(ALTO))
    fondo.fill.solid()
    fondo.fill.fore_color.rgb = AZUL_OSCURO
    fondo.line.fill.background()

    barra = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(2.05), Inches(1.2), Inches(0.08)
    )
    barra.fill.solid()
    barra.fill.fore_color.rgb = VERDE
    barra.line.fill.background()

    _textbox(slide, 0.8, 1.35, 11.8, 0.6, "INFORME RUAF", size=20, bold=True, color=RGBColor(0xBF, 0xD7, 0xEA))
    _textbox(slide, 0.8, 2.3, 11.8, 1.1, "Afiliaciones de los empleados", size=38, bold=True, color=BLANCO)
    _textbox(
        slide, 0.8, 3.55, 11.8, 0.5,
        "Salud, pensiones, riesgos laborales y demás afiliaciones reportadas en el RUAF",
        size=15, color=RGBColor(0xD9, 0xE6, 0xF2),
    )
    _textbox(slide, 0.8, 4.55, 11.8, 0.4, f"Fecha del informe: {fecha}", size=13, color=BLANCO)
    _textbox(slide, 0.8, 4.95, 11.8, 0.4, f"Empleados analizados: {total}", size=13, color=BLANCO)
    _textbox(slide, 0.8, 6.55, 11.8, 0.35, f"{EMPRESA_INFORME}  ·  Consulta autorizada del RUAF", size=10, color=RGBColor(0xBF, 0xD7, 0xEA))


def _resumen(prs, stats):
    slide = _base_slide(prs, "Resumen ejecutivo", "Indicadores generales de la consulta")
    total = stats["total"]
    _kpi(slide, 0.6, 1.35, 2.35, 1.25, total, "Empleados con PDF", AZUL)
    _kpi(slide, 3.12, 1.35, 2.35, 1.25, stats["con_salud"], f"Con salud ({_pct(stats['con_salud'], total)})", VERDE)
    _kpi(slide, 5.64, 1.35, 2.35, 1.25, stats["con_pension"], f"Con pensión ({_pct(stats['con_pension'], total)})", AZUL)
    _kpi(slide, 8.16, 1.35, 2.35, 1.25, stats["con_arl"], f"Con ARL ({_pct(stats['con_arl'], total)})", AZUL)
    _kpi(slide, 10.68, 1.35, 2.05, 1.25, len(stats["novedades"]), "Con novedades", ROJO if stats["novedades"] else VERDE)

    bullets = [
        f"Se analizaron {total} empleados con reporte oficial del RUAF.",
        f"Cobertura: salud {_pct(stats['con_salud'], total)}, pensión {_pct(stats['con_pension'], total)}, ARL {_pct(stats['con_arl'], total)}.",
        f"Compensación familiar: {_pct(stats['con_comp'], total)}; cesantías: {_pct(stats['con_ces'], total)}; programas: {_pct(stats['con_prog'], total)}.",
        f"{len(stats['novedades'])} empleados presentan secciones sin reportar o estados no activos.",
    ]
    box = slide.shapes.add_textbox(Inches(0.7), Inches(2.95), Inches(7.3), Inches(2.4))
    frame = box.text_frame
    frame.word_wrap = True
    for index, text in enumerate(bullets):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = "•  " + text
        paragraph.space_after = Pt(8)
        for run in paragraph.runs:
            run.font.size = Pt(12)
            run.font.color.rgb = NEGRO

    # Tabla de estado general
    rows = [[key, value] for key, value in stats["estado_general"].most_common()]
    _slide_table(
        slide, ["Estado general", "Empleados"], rows,
        top=2.95, width=4.2, height=2.0, left=8.45,
    )
    table = slide.shapes[-1].table
    table.columns[0].width = Inches(2.7)
    table.columns[1].width = Inches(1.5)


def _cobertura(prs, stats):
    slide = _base_slide(prs, "Cobertura por sección", "Empleados con y sin reporte en cada componente")
    total = stats["total"]
    categories = ["Salud", "Pensión", "ARL", "Compensación", "Cesantías", "Programas"]
    con = [stats["con_salud"], stats["con_pension"], stats["con_arl"], stats["con_comp"], stats["con_ces"], stats["con_prog"]]
    sin = [total - value for value in con]

    data = CategoryChartData()
    data.categories = categories
    data.add_series("Con reporte", con)
    data.add_series("Sin reporte", sin)

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_STACKED, Inches(0.7), Inches(1.45), Inches(11.9), Inches(5.2), data
    ).chart
    _style_chart(chart, "Cobertura por sección (número de empleados)")
    _color_series(chart.series[0], VERDE)
    _color_series(chart.series[1], RGBColor(0xD9, 0xE2, 0xEC))
    plot = chart.plots[0]
    plot.gap_width = 70
    plot.has_data_labels = True
    plot.data_labels.font.size = Pt(8)
    plot.data_labels.position = XL_LABEL_POSITION.CENTER


def _estados(prs, stats, registros):
    slide = _base_slide(prs, "Estados de afiliación", "Activos, inactivos y sin reporte por sección")
    total = stats["total"]

    def counts(key, admin_key):
        con = sum(1 for r in registros if _has(r.get(admin_key)))
        activos = sum(1 for r in registros if _active(r.get(key)))
        return activos, max(0, con - activos), total - con

    salud = counts("salud_estado", "salud_administradora")
    pension = counts("pension_estado", "pension_administradora")
    arl = counts("arl_estado", "arl_administradora")

    data = CategoryChartData()
    data.categories = ["Salud", "Pensión", "ARL"]
    data.add_series("Activo", [salud[0], pension[0], arl[0]])
    data.add_series("Inactivo / otro", [salud[1], pension[1], arl[1]])
    data.add_series("No reportado", [salud[2], pension[2], arl[2]])

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_STACKED, Inches(0.7), Inches(1.45), Inches(11.9), Inches(5.2), data
    ).chart
    _style_chart(chart, "Estado de afiliación por sección")
    _color_series(chart.series[0], VERDE)
    _color_series(chart.series[1], AMARILLO)
    _color_series(chart.series[2], RGBColor(0xD9, 0xE2, 0xEC))
    plot = chart.plots[0]
    plot.gap_width = 70
    plot.has_data_labels = True
    plot.data_labels.font.size = Pt(8)
    plot.data_labels.position = XL_LABEL_POSITION.CENTER


def _perfiles(prs, stats):
    slide = _base_slide(prs, "Empleados por perfil", "Distribución de la base analizada")
    data = CategoryChartData()
    categories = [key for key, _ in stats["perfiles"].most_common()]
    values = [value for _, value in stats["perfiles"].most_common()]
    data.categories = categories
    data.add_series("Empleados", values)

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE, Inches(1.2), Inches(1.55), Inches(10.9), Inches(4.9), data
    ).chart
    _style_chart(chart, "Distribución por perfil")
    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.data_labels.font.size = Pt(10)
    plot.data_labels.show_percentage = True
    plot.data_labels.show_value = False
    colors = [AZUL, VERDE, AMARILLO, ROJO, RGBColor(0x7B, 0x1F, 0xA2)]
    for index, point in enumerate(plot.series[0].points):
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = colors[index % len(colors)]


def _top_bar(prs, title, subtitle, counter, color=AZUL, limit=8):
    slide = _base_slide(prs, title, subtitle)
    items = [(name, value) for name, value in counter.most_common(limit) if name]
    if not items:
        _textbox(slide, 0.8, 3.0, 11.7, 0.6, "Sin datos para mostrar.", size=16, color=GRIS, align=PP_ALIGN.CENTER)
        return
    data = CategoryChartData()
    data.categories = [name for name, _ in items]
    data.add_series("Empleados", [value for _, value in items])
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.7), Inches(1.45), Inches(11.9), Inches(5.2), data
    ).chart
    _style_chart(chart, f"Top {len(items)}", legend=False)
    _color_series(chart.series[0], color)
    plot = chart.plots[0]
    plot.gap_width = 45
    plot.has_data_labels = True
    plot.data_labels.font.size = Pt(9)
    plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    chart.category_axis.tick_labels.font.size = Pt(8)
    try:
        chart.category_axis.reverse_order = True
    except Exception:
        pass


def _novedades(prs, stats):
    slide = _base_slide(prs, "Empleados con novedades", "Salud, pensión o ARL sin reportar o con estado no activo")
    novedades = stats["novedades"]
    if not novedades:
        _textbox(slide, 0.8, 2.8, 11.7, 0.8, "No se encontraron novedades: todos los empleados tienen las secciones principales reportadas.", size=16, color=VERDE, align=PP_ALIGN.CENTER)
        return

    headers = ["Documento", "Nombre", "Perfil", "Salud", "Pensión", "ARL", "Observaciones"]
    rows = []
    for record in novedades[:12]:
        nombre = f"{record.get('apellidos', '')} {record.get('nombres', '')}".strip()
        rows.append([
            record.get("documento", ""),
            nombre[:32],
            record.get("perfil", ""),
            record.get("salud_estado", "No reportado")[:14],
            record.get("pension_estado", "No reportado")[:14],
            record.get("arl_estado", "No reportado")[:14],
            record.get("observaciones", "")[:55],
        ])
    _slide_table(slide, headers, rows, top=1.35, height=4.9)
    table = slide.shapes[-1].table
    widths = [1.0, 2.1, 1.1, 1.0, 1.0, 0.9, 4.9]
    for index, width in enumerate(widths):
        table.columns[index].width = Inches(width)

    if len(novedades) > 12:
        _textbox(
            slide, 0.6, 6.55, 12.1, 0.35,
            f"Se muestran 12 de {len(novedades)} casos. El detalle completo está en el Excel de datos.",
            size=9, color=GRIS,
        )


def _conclusiones(prs, stats):
    slide = _base_slide(prs, "Conclusiones y recomendaciones", "Lectura rápida para la gestión")
    total = stats["total"]
    bullets = [
        f"La base analizada tiene {total} empleados con reporte del RUAF.",
        f"Salud: {stats['con_salud']} ({_pct(stats['con_salud'], total)}); pensión: {stats['con_pension']} ({_pct(stats['con_pension'], total)}); ARL: {stats['con_arl']} ({_pct(stats['con_arl'], total)}).",
        f"Estado general: {stats['estado_general'].get('Completo', 0)} completos, {stats['estado_general'].get('Parcial', 0)} parciales y {stats['estado_general'].get('Sin afiliaciones', 0)} sin afiliaciones.",
        f"{len(stats['novedades'])} empleados tienen alguna novedad para revisar con la administradora respectiva.",
        "El detalle completo, con una fila por empleado y por sección, está en el Excel generado junto con este informe.",
    ]
    box = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.7), Inches(4.6))
    frame = box.text_frame
    frame.word_wrap = True
    for index, text in enumerate(bullets):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = "•  " + text
        paragraph.space_after = Pt(12)
        for run in paragraph.runs:
            run.font.size = Pt(14)
            run.font.color.rgb = NEGRO

    _rect(slide, 0.8, 6.05, 11.7, 0.65, GRIS_CLARO, line=RGBColor(0xDD, 0xE3, 0xEA))
    _textbox(
        slide, 0.95, 6.17, 11.4, 0.4,
        "Recomendación: priorizar los casos sin ARL o sin pensión y verificar los estados no activos con cada administradora.",
        size=10, color=AZUL_OSCURO,
    )


# ---------------------------------------------------------------------------
# Generación
# ---------------------------------------------------------------------------

def generar_pptx(registros, output_path, fecha=None):
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fecha = fecha or datetime.now().strftime("%d/%m/%Y %H:%M")

    prs = Presentation()
    prs.slide_width = Inches(ANCHO)
    prs.slide_height = Inches(ALTO)

    stats = _stats(registros)
    _portada(prs, stats["total"], fecha)
    if stats["total"] == 0:
        slide = _base_slide(prs, "Sin datos", "No se encontraron PDFs para analizar")
        _textbox(
            slide, 0.8, 3.0, 11.7, 0.8,
            "No hay PDFs del RUAF en la carpeta seleccionada. Ejecuta primero la consulta.",
            size=16, color=GRIS, align=PP_ALIGN.CENTER,
        )
        prs.save(output_path)
        return output_path

    _resumen(prs, stats)
    _cobertura(prs, stats)
    _estados(prs, stats, registros)
    _perfiles(prs, stats)
    _top_bar(prs, "Principales EPS", "Salud: administradoras con más empleados", stats["eps"], AZUL)
    _top_bar(prs, "Principales AFP", "Pensiones: administradoras con más empleados", stats["afp"], VERDE)
    _top_bar(prs, "Principales ARL", "Riesgos laborales: administradoras con más empleados", stats["arl"], AMARILLO)
    _novedades(prs, stats)
    _conclusiones(prs, stats)

    prs.save(output_path)
    return output_path


def main(argv=None):
    parser = argparse.ArgumentParser(description="Genera el informe PPTX del RUAF.")
    parser.add_argument("--carpeta", required=True, help="Carpeta con los PDFs")
    parser.add_argument("--excel", default="", help="Excel de caracterización")
    parser.add_argument("--salida", required=True, help="Ruta del PPTX")
    args = parser.parse_args(argv)

    from extraer_datos import construir_registros

    registros = construir_registros(args.carpeta, excel_path=args.excel or None)
    generar_pptx(registros, args.salida)
    print(f"[OK] PPTX: {args.salida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

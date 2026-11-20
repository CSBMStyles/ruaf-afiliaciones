"""Orquesta la generación del Excel de datos y el informe PPTX del RUAF."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import extraer_datos
from generar_pptx import generar_pptx


def generar_todo(
    output_dir,
    reports_dir,
    excel_path=None,
    timestamp=None,
    sheet_name=None,
    header_row=1,
    column_map=None,
):
    """Genera el Excel de datos y el PPTX a partir de los PDFs descargados."""
    timestamp = timestamp or datetime.now().strftime("%Y%m%d_%H%M%S")
    reports_dir = Path(reports_dir)
    reports_dir.mkdir(parents=True, exist_ok=True)

    registros = extraer_datos.construir_registros(
        output_dir,
        excel_path=excel_path,
        sheet_name=sheet_name,
        header_row=header_row,
        column_map=column_map,
    )

    excel_out = reports_dir / f"DATOS_RUAF_{timestamp}.xlsx"
    extraer_datos.generar_excel(registros, excel_out)

    pptx_out = reports_dir / f"INFORME_RUAF_{timestamp}.pptx"
    generar_pptx(registros, pptx_out)

    return {
        "excel": str(excel_out),
        "pptx": str(pptx_out),
        "registros": len(registros),
    }


def main(argv=None):
    import argparse

    from config import load_settings

    settings = load_settings()
    parser = argparse.ArgumentParser(description="Genera Excel y PPTX del RUAF.")
    parser.add_argument("--carpeta", default=str(settings.output_dir))
    parser.add_argument("--excel", default=str(settings.excel_path))
    parser.add_argument("--salida", default=str(settings.reports_dir))
    args = parser.parse_args(argv)

    result = generar_todo(args.carpeta, args.salida, excel_path=args.excel)
    print(f"[OK] Registros: {result['registros']}")
    print(f"[OK] Excel: {result['excel']}")
    print(f"[OK] PPTX:  {result['pptx']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Declaración de campos del proyecto RUAF - Afiliaciones."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Campo:
    clave: str
    etiqueta: str
    candidatos: tuple[str, ...]
    obligatorio: bool = True


CAMPOS: tuple[Campo, ...] = (
    Campo(
        clave="documento",
        etiqueta="Documento / Cédula",
        candidatos=(
            "Identificacion",
            "Identificación",
            "Numero Documento",
            "Número de documento",
            "Numero de documento",
            "Documento",
            "Cedula",
            "Cédula",
        ),
    ),
    Campo(
        clave="fecha",
        etiqueta="Fecha de expedición",
        candidatos=(
            "Expedicion",
            "Expedición",
            "Fecha Expedicion",
            "Fecha de expedicion",
            "Fecha Expedición",
            "Fecha de Expedición",
        ),
    ),
    Campo(
        clave="nombre",
        etiqueta="Nombres",
        candidatos=("Nombres", "Nombre"),
    ),
    Campo(
        clave="apellido",
        etiqueta="Apellidos",
        candidatos=("Apellidos", "Apellido"),
    ),
    Campo(
        clave="perfil",
        etiqueta="Perfil / Cargo",
        candidatos=("Perfil", "Cargo", "Rol"),
    ),
    Campo(
        clave="tipo_documento",
        etiqueta="Tipo de documento (opcional)",
        candidatos=(
            "TipoDocumento",
            "Tipo Documento",
            "Tipo de documento",
            "Tipo de Documento",
            "Tipo doc",
        ),
        obligatorio=False,
    ),
)

PROYECTO = {
    "id": "ruaf",
    "nombre": "RUAF - Afiliaciones",
    "entidad": "SISPRO - Ministerio de Salud",
    "tipo_captcha": "imagen_ia_manual",
    "salida_por_perfil": True,
    "perfiles": ("AUXILIAR", "ENFERMERO", "MEDICO"),
    "filename_pattern": "apellido_primero",
}

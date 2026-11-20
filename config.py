"""Configuración persistente del proyecto RUAF - Afiliaciones.

- ``config.json``: rutas, hoja, mapeo y comportamiento (sin secretos).
- ``.env``: solo ajustes opcionales como MANUAL_TIMEOUT.
"""

from __future__ import annotations

import json
import os
import platform
import re
import shutil
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path
from typing import Any

try:
    from dotenv import load_dotenv
except Exception:  # pragma: no cover
    load_dotenv = None


PROJECT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = PROJECT_DIR / "config.json"
ENV_FILE = PROJECT_DIR / ".env"

_RUNTIME_FIELDS = {"interactive", "reset_state"}
WORKSPACE_DIR = PROJECT_DIR.parents[2] if len(PROJECT_DIR.parents) > 2 else PROJECT_DIR

_PROJECT_FIELDS = {"state_file", "temp_dir", "error_dir"}
_PORTABLE_FIELDS = {"excel_path", "output_dir", "reports_dir"} | _PROJECT_FIELDS


def _safe_parent(index: int) -> Path | None:
    parents = PROJECT_DIR.parents
    return parents[index] if len(parents) > index else None


def default_excel_path() -> Path:
    candidates: list[Path] = []
    vacunaciones = _safe_parent(2)
    if vacunaciones:
        candidates.extend(
            [
                vacunaciones / "RUAF" / "CARACTERIZACION I.xlsx",
                vacunaciones / "RUAF" / "CARACTERIZACION.xlsx",
                vacunaciones / "CARACTERIZACION.xlsx",
            ]
        )
    codigo = _safe_parent(1)
    if codigo:
        candidates.append(codigo / "CARACTERIZACION.xlsx")
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0] if candidates else PROJECT_DIR / "CARACTERIZACION.xlsx"


def default_output_dir() -> Path:
    return WORKSPACE_DIR / "RUAF" / "AFILIACIONES"


def default_reports_dir() -> Path:
    return WORKSPACE_DIR / "RUAF"


def _as_path(value: Any, fallback: Path) -> Path:
    if value in (None, ""):
        return fallback
    return Path(os.path.expandvars(os.path.expanduser(str(value))))


@dataclass
class Settings:
    excel_path: Path = field(default_factory=default_excel_path)
    output_dir: Path = field(default_factory=default_output_dir)
    reports_dir: Path = field(default_factory=default_reports_dir)
    state_file: Path = field(default_factory=lambda: PROJECT_DIR / "estado_ruaf.json")
    temp_dir: Path = field(default_factory=lambda: PROJECT_DIR / "temp_pdf")
    error_dir: Path = field(default_factory=lambda: PROJECT_DIR / "errors")

    sheet_name: str | int | None = None
    header_row: int = 1
    column_map: dict[str, str] = field(default_factory=dict)

    profiles: list[str] | None = None
    batch_size: int | None = None
    headless: bool | None = False
    resume: bool = True
    reset_state: bool = False

    captcha_max_attempts: int = 5
    captcha_ai_attempts: int = 3
    captcha_mode: str = "auto"
    manual_timeout: int = 240
    download_timeout: int = 30
    driver_backend: str = "selenium"
    filename_pattern: str = "apellido_primero"
    generate_reports: bool = True

    interactive: bool = True

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        for key in _RUNTIME_FIELDS:
            data.pop(key, None)
        for key, value in list(data.items()):
            if isinstance(value, Path):
                if key in _PORTABLE_FIELDS:
                    base = PROJECT_DIR if key in _PROJECT_FIELDS else WORKSPACE_DIR
                    try:
                        data[key] = str(value.relative_to(base))
                        continue
                    except ValueError:
                        pass
                data[key] = str(value)
        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "Settings":
        settings = cls()
        valid = {f.name for f in fields(cls)}
        for key, value in (data or {}).items():
            if key not in valid:
                continue

            if key in _PORTABLE_FIELDS:
                fallback = getattr(settings, key)
                raw = str(value)

                if raw and not Path(raw).is_absolute():
                    base = PROJECT_DIR if key in _PROJECT_FIELDS else WORKSPACE_DIR
                    candidate = base / raw
                else:
                    candidate = _as_path(value, fallback)

                if key == "excel_path" and not candidate.exists():
                    candidate = fallback

                if key in _PROJECT_FIELDS and not candidate.exists():
                    try:
                        inside = candidate.resolve().is_relative_to(PROJECT_DIR.resolve())
                    except Exception:
                        inside = str(candidate).startswith(str(PROJECT_DIR))
                    if not inside:
                        candidate = fallback

                setattr(settings, key, candidate)
            elif key == "header_row":
                try:
                    settings.header_row = max(1, int(value))
                except (TypeError, ValueError):
                    pass
            elif key in ("captcha_max_attempts", "captcha_ai_attempts", "manual_timeout", "download_timeout"):
                try:
                    setattr(settings, key, max(1, int(value)))
                except (TypeError, ValueError):
                    pass
            elif key == "generate_reports":
                settings.generate_reports = bool(value)
            elif key == "captcha_mode":
                mode = str(value or "").strip().lower()
                if mode in ("auto", "manual"):
                    settings.captcha_mode = mode
            elif key == "batch_size":
                try:
                    settings.batch_size = None if value is None else max(0, int(value))
                except (TypeError, ValueError):
                    pass
            elif key == "profiles":
                if value in (None, "", []):
                    settings.profiles = None
                elif isinstance(value, str):
                    settings.profiles = [value]
                else:
                    settings.profiles = [str(v) for v in value]
            elif key == "column_map" and isinstance(value, dict):
                settings.column_map = {str(k): str(v) for k, v in value.items() if v}
            elif key == "headless":
                settings.headless = None if value is None else bool(value)
            else:
                setattr(settings, key, value)
        return settings


def _env_int(name: str, default: int) -> int:
    raw = os.getenv(name, "").strip()
    if not raw:
        return default
    try:
        return max(1, int(raw))
    except ValueError:
        return default


def load_settings(path: Path | None = None) -> Settings:
    path = Path(path) if path else CONFIG_FILE
    if not path.exists():
        settings = Settings()
        settings.captcha_max_attempts = _env_int("CAPTCHA_MAX_ATTEMPTS", settings.captcha_max_attempts)
        settings.captcha_ai_attempts = _env_int("CAPTCHA_AI_ATTEMPTS", settings.captcha_ai_attempts)
        settings.manual_timeout = _env_int("MANUAL_TIMEOUT", settings.manual_timeout)
        mode = os.getenv("CAPTCHA_MODE", "").strip().lower()
        if mode in ("auto", "manual"):
            settings.captcha_mode = mode
        settings.download_timeout = _env_int("DOWNLOAD_TIMEOUT", settings.download_timeout)
        backend = os.getenv("DRIVER_BACKEND", "").strip()
        if backend:
            settings.driver_backend = backend
        generate = os.getenv("GENERATE_REPORTS", "").strip().lower()
        if generate:
            settings.generate_reports = generate not in ("0", "false", "no", "n")
        return settings

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return Settings()
    return Settings.from_dict(data)


def save_settings(settings: Settings, path: Path | None = None) -> Path:
    path = Path(path) if path else CONFIG_FILE
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(settings.to_dict(), ensure_ascii=False, indent=2)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(payload + "\n", encoding="utf-8")
    os.replace(tmp, path)
    return path


# ----------------------------------------------------------------------
# .env
# ----------------------------------------------------------------------

def load_environment() -> None:
    if load_dotenv is None:
        return
    if ENV_FILE.exists():
        load_dotenv(ENV_FILE, override=False)


_ENV_LINE = re.compile(r"^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$")


def read_env_file(path: Path | None = None) -> dict[str, str]:
    path = Path(path) if path else ENV_FILE
    values: dict[str, str] = {}
    if not path.exists():
        return values
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        match = _ENV_LINE.match(line)
        if not match:
            continue
        key, raw = match.group(1), match.group(2).strip()
        if len(raw) >= 2 and raw[0] == raw[-1] and raw[0] in ("'", '"'):
            raw = raw[1:-1]
        values[key] = raw
    return values


def _env_value(value: str) -> str:
    value = str(value or "")
    if not value:
        return ""
    if any(ch in value for ch in (" ", "#", "\t")) or value != value.strip():
        return '"' + value.replace('"', '\\"') + '"'
    return value


def update_env_file(values: dict[str, str], path: Path | None = None) -> Path:
    path = Path(path) if path else ENV_FILE
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = path.read_text(encoding="utf-8").splitlines() if path.exists() else []
    pending = {k: str(v) for k, v in values.items()}
    output: list[str] = []
    for line in lines:
        match = _ENV_LINE.match(line)
        if match and match.group(1) in pending:
            key = match.group(1)
            output.append(f"{key}={_env_value(pending.pop(key))}")
        else:
            output.append(line)
    for key, value in pending.items():
        output.append(f"{key}={_env_value(value)}")
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text("\n".join(output).rstrip() + "\n", encoding="utf-8")
    os.replace(tmp, path)
    return path


def chrome_path() -> Path | None:
    system = platform.system()
    candidates: list[Path] = []
    if system == "Windows":
        for env in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
            base = os.environ.get(env)
            if base:
                candidates.append(Path(base) / "Google" / "Chrome" / "Application" / "chrome.exe")
    elif system == "Darwin":
        candidates.append(Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"))
        candidates.append(Path.home() / "Applications/Google Chrome.app/Contents/MacOS/Google Chrome")
    else:
        for name in ("google-chrome", "google-chrome-stable", "chromium", "chromium-browser"):
            found = shutil.which(name)
            if found:
                candidates.append(Path(found))
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


load_environment()

"""Consulta masiva de afiliaciones en el RUAF - SISPRO.

El sitio usa un captcha de imagen de 5 caracteres. La aplicación llena el
formulario y lo resuelve con IA (Gemini/OpenAI); si falla, muestra la imagen
ampliada para que el usuario escriba el código. Luego consulta, imprime el
reporte como PDF y guarda el Excel y el PPTX.
"""

from __future__ import annotations

# En macOS, el Python de python.org no trae los certificados CA y
# undetected-chromedriver falla al descargar el chromedriver.
import os as _os

try:
    import certifi as _certifi

    _os.environ.setdefault("SSL_CERT_FILE", _certifi.where())
    _os.environ.setdefault("REQUESTS_CA_BUNDLE", _certifi.where())
except Exception:
    pass

import argparse
import base64
import csv
import json
import logging
import platform
import shutil
import time
from datetime import datetime
from pathlib import Path

from selenium.common.exceptions import (
    NoAlertPresentException,
    TimeoutException,
    UnexpectedAlertPresentException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

import captcha_solver
from config import Settings, load_settings
from sources import load_users
from utils import configure_console, safe_filename

configure_console()

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

FILTRO_URL = "https://ruaf.sispro.gov.co/Filtro.aspx"
RESULT_URL_MARKER = "AfiliacionPersona.aspx"

ID_TERMS_ACEPTO = "MainContent_RadioButtonList1_0"
ID_TERMS_ENVIAR = "MainContent_btnEnviar"
ID_TIPO = "ddlTiposDocumentos"
ID_DOC = "MainContent_txbNumeroIdentificacion"
ID_FECHA = "MainContent_datepicker"
ID_CAPTCHA = "MainContent_txtCaptcha"
ID_BTN_VERIFY = "MainContent_btnVerify"
ID_BTN_CONSULTAR = "MainContent_btnConsultar"
ID_EXPORT_BTN = "ctl00_MainContent_rvConsulta_ctl09_ctl04_ctl00_ButtonLink"
ID_EXPORT_MENU = "ctl00_MainContent_rvConsulta_ctl09_ctl04_ctl00_Menu"

CHROME_UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36"
)
if platform.system() == "Windows":
    CHROME_UA = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36"
    )

PROJECT_DIR = Path(__file__).resolve().parent

PROFILE_FOLDER_MAP = {
    "AUXILIAR": "AUXILIAR",
    "ENFERMERO": "ENFERMERO",
    "MEDICO": "MEDICO",
}

SETTINGS: Settings = load_settings()

EXCEL_PATH = Path(SETTINGS.excel_path)
BASE_OUTPUT_DIR = Path(SETTINGS.output_dir)
REPORTS_DIR = Path(SETTINGS.reports_dir)
STATE_FILE = Path(SETTINGS.state_file)
TEMP_DIR = Path(SETTINGS.temp_dir)
ERROR_DIR = Path(SETTINGS.error_dir)
DOWNLOAD_DIR = TEMP_DIR / "downloads"

LAST_STATS: dict = {}
_STOP_CHECKER = lambda: False
_CAPTCHA_CALLBACK = None


class AbortExecution(Exception):
    pass


class SkipUser(Exception):
    pass


def set_stop_checker(checker=None) -> None:
    global _STOP_CHECKER
    _STOP_CHECKER = checker or (lambda: False)


def set_captcha_callback(callback=None) -> None:
    global _CAPTCHA_CALLBACK
    _CAPTCHA_CALLBACK = callback


def stop_requested() -> bool:
    try:
        return bool(_STOP_CHECKER())
    except Exception:
        return False


def apply_settings(settings: Settings) -> Settings:
    global SETTINGS, EXCEL_PATH, BASE_OUTPUT_DIR, REPORTS_DIR, STATE_FILE, TEMP_DIR, ERROR_DIR, DOWNLOAD_DIR

    SETTINGS = settings
    EXCEL_PATH = Path(settings.excel_path)
    BASE_OUTPUT_DIR = Path(settings.output_dir)
    REPORTS_DIR = Path(settings.reports_dir)
    STATE_FILE = Path(settings.state_file)
    TEMP_DIR = Path(settings.temp_dir)
    ERROR_DIR = Path(settings.error_dir)
    DOWNLOAD_DIR = TEMP_DIR / "downloads"
    return SETTINGS


# ---------------------------------------------------------------------------
# Estado
# ---------------------------------------------------------------------------

def load_state():
    if not STATE_FILE.exists():
        return {"procesados": [], "fallidos": [], "ultima_actualizacion": None}
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)
            data.setdefault("procesados", [])
            data.setdefault("fallidos", [])
            return data
    except Exception:
        return {"procesados": [], "fallidos": [], "ultima_actualizacion": None}


def save_state(state):
    state["ultima_actualizacion"] = datetime.now().isoformat(timespec="seconds")
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(STATE_FILE, "w", encoding="utf-8") as file:
        json.dump(state, file, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Carpetas y navegador
# ---------------------------------------------------------------------------

def configure_folders():
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    ERROR_DIR.mkdir(parents=True, exist_ok=True)
    BASE_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    for folder in PROFILE_FOLDER_MAP.values():
        (BASE_OUTPUT_DIR / folder).mkdir(parents=True, exist_ok=True)


def _download_prefs():
    return {
        "download.default_directory": str(DOWNLOAD_DIR),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
    }


def create_driver(headless: bool = False, backend: str | None = None):
    """Crea el driver. Para el captcha manual debe ser visible."""
    backend = (backend or SETTINGS.driver_backend or "auto").strip().lower()
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    args = [
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu",
        "--window-size=1400,1000",
        "--lang=es-CO",
    ]
    if headless:
        args.append("--headless=new")

    driver = None
    if backend in ("auto", "uc"):
        try:
            import undetected_chromedriver as uc

            options = uc.ChromeOptions()
            for arg in args:
                options.add_argument(arg)
            options.add_argument(f"--user-agent={CHROME_UA}")
            try:
                options.add_experimental_option("prefs", _download_prefs())
            except Exception:
                pass
            driver = uc.Chrome(options=options)
            logging.info("Driver: undetected-chromedriver")
        except Exception as exc:
            if backend == "uc":
                raise
            logging.warning("undetected-chromedriver no disponible (%s). Usando Selenium...", str(exc)[:160])

    if driver is None:
        from selenium import webdriver

        options = webdriver.ChromeOptions()
        for arg in args:
            options.add_argument(arg)
        options.add_argument(f"--user-agent={CHROME_UA}")
        options.add_experimental_option("prefs", _download_prefs())
        driver = webdriver.Chrome(options=options)
        logging.info("Driver: Selenium estándar")

    for command, payload in (
        ("Browser.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": str(DOWNLOAD_DIR),
            "eventsEnabled": True,
        }),
        ("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": str(DOWNLOAD_DIR),
        }),
    ):
        try:
            driver.execute_cdp_cmd(command, payload)
        except Exception:
            pass
    try:
        driver.set_script_timeout(30)
    except Exception:
        pass
    return driver


def focus_chrome():
    """Intenta traer Chrome al frente (macOS). En Windows el usuario hace clic."""
    if platform.system() == "Darwin":
        try:
            import subprocess

            subprocess.Popen(["osascript", "-e", 'tell application "Google Chrome" to activate'])
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Formulario y captcha
# ---------------------------------------------------------------------------

def _set_value(driver, element, value: str):
    driver.execute_script("arguments[0].value=''; arguments[0].focus();", element)
    try:
        element.send_keys(value)
    except Exception:
        pass
    if element.get_attribute("value") != value:
        driver.execute_script(
            "arguments[0].value = arguments[1];"
            "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
            element, value,
        )


def ensure_form(driver, wait):
    """Abre el formulario y acepta los términos si la sesión es nueva o expiró."""
    driver.get(FILTRO_URL)
    deadline = time.time() + 60
    last_click = 0.0
    clicks = 0
    while time.time() < deadline:
        if driver.find_elements(By.ID, ID_DOC):
            return True
        radios = driver.find_elements(By.ID, ID_TERMS_ACEPTO)
        if radios and clicks < 3 and (clicks == 0 or time.time() - last_click > 8):
            try:
                driver.execute_script("arguments[0].click();", radios[0])
                driver.execute_script(
                    "arguments[0].click();",
                    driver.find_element(By.ID, ID_TERMS_ENVIAR),
                )
                clicks += 1
                last_click = time.time()
                time.sleep(2)
            except Exception:
                time.sleep(0.5)
        else:
            time.sleep(0.5)
    raise TimeoutException("No se pudo entrar al formulario del RUAF.")


def fill_form(driver, wait, user, settings: Settings):
    select = Select(wait.until(EC.presence_of_element_located((By.ID, ID_TIPO))))
    try:
        select.select_by_value(user.get("tipo_documento") or "5|CC")
    except Exception:
        select.select_by_value("5|CC")

    _set_value(driver, wait.until(EC.visibility_of_element_located((By.ID, ID_DOC))), user["documento"])
    _set_value(driver, wait.until(EC.visibility_of_element_located((By.ID, ID_FECHA))), user["fecha_expedicion"])
    _set_value(driver, driver.find_element(By.ID, ID_CAPTCHA), "")


def fetch_captcha(driver) -> bytes:
    """Obtiene la imagen del captcha con fetch; la etiqueta <img> suele dar 404."""
    script = """
    const cb = arguments[arguments.length - 1];
    const img = document.querySelector('img[alt="Captcha"]');
    if (!img) { cb(null); return; }
    fetch(img.src, {credentials: 'include'})
      .then(r => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.arrayBuffer(); })
      .then(b => {
        let s = '';
        const a = new Uint8Array(b);
        for (let i = 0; i < a.length; i++) s += String.fromCharCode(a[i]);
        cb(btoa(s));
      })
      .catch(e => cb(null));
    """
    raw = driver.execute_async_script(script)
    if not raw:
        raise RuntimeError("No se pudo obtener la imagen del captcha.")
    data = base64.b64decode(raw)
    if len(data) < 1000:
        raise RuntimeError("La imagen del captcha llegó vacía.")
    return data


def click_verify(driver, wait) -> bool:
    try:
        button = wait.until(EC.presence_of_element_located((By.ID, ID_BTN_VERIFY)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", button)
        driver.execute_script("arguments[0].click();", button)
        return True
    except Exception:
        try:
            driver.find_element(By.ID, ID_BTN_VERIFY).click()
            return True
        except Exception:
            return False


def wait_validation(driver, timeout: int = 30):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            body = driver.find_element(By.TAG_NAME, "body").text
        except UnexpectedAlertPresentException:
            try:
                alert = driver.switch_to.alert
                text = alert.text
                alert.accept()
                low = text.lower()
                if "válid" in low or "valid" in low:
                    return True
                return False
            except Exception:
                body = ""
        except Exception:
            body = ""
        if "Texto Válido" in body:
            return True
        if "Texto Inválido" in body:
            return False
        time.sleep(0.4)
    return None


def click_consultar(driver, wait) -> bool:
    try:
        button = wait.until(EC.presence_of_element_located((By.ID, ID_BTN_CONSULTAR)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", button)
        driver.execute_script("arguments[0].click();", button)
        return True
    except Exception:
        try:
            driver.find_element(By.ID, ID_BTN_CONSULTAR).click()
            return True
        except Exception:
            return False


def _clean_site_error(body: str) -> str:
    text = " ".join(str(body or "").split())
    for marker in ("Error", "La fecha", "No se encontr", "no corresponde", "no coincide"):
        if marker.lower() in text.lower():
            start = max(0, text.lower().find(marker.lower()) - 20)
            return text[start:start + 300].strip(" :.-")
    return text[:300]


def wait_result(driver, documento: str, timeout: int = 60):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if stop_requested():
            return "detenido", "El usuario detuvo el proceso."

        try:
            alert = driver.switch_to.alert
            text = alert.text
            alert.accept()
            return "datos_no_coinciden", _clean_site_error(text)
        except NoAlertPresentException:
            pass
        except Exception:
            pass

        try:
            body = driver.find_element(By.TAG_NAME, "body").text
        except Exception:
            body = ""
        url = driver.current_url

        if RESULT_URL_MARKER in url:
            if documento in body and (
                "Afiliaciones de una Persona" in body
                or "INFORMACIÓN BASICA" in body
                or "INFORMACION BASICA" in body
            ):
                return "ok", body[:300]
        elif "Filtro.aspx" in url:
            low = body.lower()
            if "no corresponde" in low or "no coincide" in low or "no es correcta" in low:
                return "datos_no_coinciden", _clean_site_error(body)
            if "no se encontr" in low and documento in body:
                return "sin_resultado", _clean_site_error(body)
            if "error" in low and "aceptar" in low:
                return "error_sitio", _clean_site_error(body)

        time.sleep(0.8)

    return "sin_resultado", "No apareció el reporte en el tiempo esperado."


def build_filename(user, pattern: str) -> str:
    if pattern == "documento_primero":
        stem = f"{user['documento']} {user['nombre']} {user['apellido']}"
    else:
        stem = f"{user['apellido']} {user['nombre']} - {user['documento']}"
    return f"{safe_filename(stem, max_stem=120)}.pdf"


def _is_valid_existing_pdf(path: Path, documento: str) -> bool:
    """True si el PDF existente contiene el reporte y no una página de error."""
    try:
        import pymupdf as fitz

        doc = fitz.open(path)
        text = "\n".join(page.get_text() for page in doc)
        doc.close()
    except Exception:
        return False
    low = text.lower()
    if "sesión asp.net ha expirado" in low or "session expired" in low:
        return False
    if "información basica" not in low and "informacion basica" not in low:
        return False
    return bool(documento) and documento in text


def _find_valid_pdf(documento: str) -> Path | None:
    for folder in PROFILE_FOLDER_MAP.values():
        base = BASE_OUTPUT_DIR / folder
        if not base.exists():
            continue
        for path in base.glob("*.pdf"):
            if documento in path.stem and _is_valid_existing_pdf(path, documento):
                return path
    return None


def build_output_path(user, pattern: str) -> Path:
    folder = BASE_OUTPUT_DIR / PROFILE_FOLDER_MAP[user["perfil_grupo"]]
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / build_filename(user, pattern)

    if not target.exists():
        return target

    if not _is_valid_existing_pdf(target, user["documento"]):
        # Si el PDF existente es la página de error de ASP.NET, se sobrescribe.
        return target

    stem, suffix = target.stem, target.suffix
    i = 2
    while True:
        candidate = folder / f"{stem} ({i}){suffix}"
        if not candidate.exists():
            return candidate
        i += 1


def _clear_downloads():
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    for file in DOWNLOAD_DIR.iterdir():
        if file.is_file():
            try:
                file.unlink()
            except Exception:
                pass


def wait_for_download(timeout: int):
    deadline = time.time() + timeout
    last_size = -1
    stable = 0
    while time.time() < deadline:
        files = [f for f in DOWNLOAD_DIR.glob("*.pdf") if f.is_file()]
        if files:
            newest = max(files, key=lambda f: f.stat().st_mtime)
            size = newest.stat().st_size
            if size > 0 and size == last_size:
                stable += 1
                if stable >= 2 and not list(DOWNLOAD_DIR.glob("*.crdownload")):
                    return newest
            else:
                stable = 0
            last_size = size
        time.sleep(0.5)
    return None


REPORT_CONTENT_ID = "VisibleReportContentctl00_MainContent_rvConsulta_ctl13"

_PRINT_SCRIPT = """
const content = document.getElementById('VisibleReportContentctl00_MainContent_rvConsulta_ctl13');
if (!content) return false;
const clone = content.cloneNode(true);
clone.querySelectorAll('img').forEach(img => {
  if (!img.complete || img.naturalWidth === 0) img.remove();
});
document.body.innerHTML = '';
document.body.appendChild(clone);
const style = document.createElement('style');
style.id = 'ruaf-print-style';
style.innerHTML = `
  @page { size: A4 landscape; margin: 0.2in; }
  body { margin: 0 !important; background: #fff !important; }
  #VisibleReportContentctl00_MainContent_rvConsulta_ctl13 { width: 100% !important; }
  #VisibleReportContentctl00_MainContent_rvConsulta_ctl13 * {
    page-break-before: auto !important;
    page-break-after: auto !important;
    page-break-inside: auto !important;
  }
  img[alt="Imagen de informes"] { display: none !important; }
`;
document.head.appendChild(style);
return true;
"""


def _valid_report_pdf(path: Path, documento: str) -> tuple[bool, str]:
    try:
        import pymupdf as fitz
    except Exception:
        return path.stat().st_size > 10000, ""
    try:
        doc = fitz.open(path)
        text = "\n".join(page.get_text() for page in doc)
        doc.close()
    except Exception as exc:
        return False, f"No se pudo leer el PDF: {exc}"
    low = text.lower()
    if "sesión asp.net ha expirado" in low or "session expired" in low:
        return False, "El PDF contiene el error de sesión de ASP.NET."
    if "información basica" not in low and "informacion basica" not in low:
        return False, "El PDF no contiene la información básica del reporte."
    if documento and documento not in text:
        return False, "El PDF no contiene el documento consultado."
    return True, ""


def print_report_pdf(driver, user, pattern: str, settings: Settings) -> Path:
    """Imprime el reporte con el panel de impresión (Page.printToPDF).

    No usa el botón Exportar del visor, que puede devolver una página de
    sesión ASP.NET expirada.
    """
    target = build_output_path(user, pattern)
    WebDriverWait(driver, 45).until(
        EC.presence_of_element_located((By.ID, REPORT_CONTENT_ID))
    )
    ok = driver.execute_script(_PRINT_SCRIPT)
    if not ok:
        raise RuntimeError("No se encontró el contenido del reporte para imprimir.")
    time.sleep(0.8)

    result = driver.execute_cdp_cmd(
        "Page.printToPDF",
        {
            "printBackground": True,
            "landscape": True,
            "scale": 1.0,
            "displayHeaderFooter": False,
            "marginTop": 0.2,
            "marginBottom": 0.2,
            "marginLeft": 0.2,
            "marginRight": 0.2,
        },
    )
    data = base64.b64decode(result["data"])
    if len(data) < 10000:
        raise RuntimeError("El PDF de impresión salió demasiado pequeño.")
    target.write_bytes(data)

    valid, detail = _valid_report_pdf(target, user["documento"])
    if not valid:
        try:
            target.unlink()
        except Exception:
            pass
        raise RuntimeError(detail or "El PDF no corresponde al reporte.")

    return target


def save_error_screenshot(driver, user, label="error"):
    try:
        name = f"{label}_{user.get('documento', 'sin_doc')}_{int(time.time())}.png"
        path = ERROR_DIR / name
        driver.save_screenshot(str(path))
        return str(path)
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Flujo por persona
# ---------------------------------------------------------------------------

def _default_captcha_dialog(image_bytes, attempt=1):
    """Diálogo Tkinter para usar la CLI sin la interfaz principal."""
    try:
        import io
        import tkinter as tk
        from tkinter import ttk

        from PIL import Image, ImageTk

        root = tk.Tk()
        root.title("Captcha RUAF")
        root.resizable(False, False)
        try:
            root.attributes("-topmost", True)
        except Exception:
            pass

        image = Image.open(io.BytesIO(image_bytes))
        image = image.resize((image.width * 3, image.height * 3), Image.LANCZOS)
        photo = ImageTk.PhotoImage(image)
        root._photo = photo

        frame = ttk.Frame(root, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=f"Intento {attempt} · Escribe los caracteres que ves:").pack(anchor="w")
        ttk.Label(frame, image=photo).pack(pady=10)
        entry = ttk.Entry(frame, width=12, font=("Helvetica", 20), justify="center")
        entry.pack(pady=(0, 6))
        entry.focus_set()
        result = {"code": None}

        def accept():
            code = entry.get().strip().upper().replace(" ", "")
            if len(code) >= 3:
                result["code"] = code
                root.destroy()

        def omit():
            root.destroy()

        ttk.Button(frame, text="Aceptar", command=accept).pack(side="right", pady=10)
        ttk.Button(frame, text="Omitir", command=omit).pack(side="right", padx=8, pady=10)
        entry.bind("<Return>", lambda _e: accept())
        root.bind("<Escape>", lambda _e: omit())
        root.mainloop()
        return result["code"]
    except Exception as exc:
        logging.error("No se pudo mostrar el diálogo del captcha: %s", exc)
        return None


def process_user(driver, wait, user, settings: Settings, on_status=None):
    attempts = []
    last_reason, last_detail = "sin_resultado", ""

    def status(message):
        if on_status:
            try:
                on_status(message)
            except Exception:
                pass

    try:
        ensure_form(driver, wait)
        fill_form(driver, wait, user, settings)

        captcha_ok = False
        for attempt in range(1, max(1, settings.captcha_max_attempts) + 1):
            if stop_requested():
                return False, attempts, "detenido", "El usuario detuvo el proceso.", ""

            try:
                image = fetch_captcha(driver)
            except Exception as exc:
                logging.warning("Captcha no disponible para %s: %s", user["documento"], exc)
                ensure_form(driver, wait)
                fill_form(driver, wait, user, settings)
                try:
                    image = fetch_captcha(driver)
                except Exception as exc2:
                    last_reason = "captcha_no_resuelto"
                    last_detail = f"No se pudo cargar el captcha: {exc2}"
                    attempts.append({
                        "intento": attempt, "metodo": "manual",
                        "resultado": last_reason, "detalle": last_detail,
                    })
                    break

            metodo = "manual"
            code = None
            ai_attempts = max(1, int(settings.captcha_ai_attempts or 3))
            if settings.captcha_mode != "manual" and attempt <= ai_attempts:
                status(f"La IA está leyendo el captcha de {user['apellido']} {user['nombre']}…")
                try:
                    code, provider = captcha_solver.solve_captcha_auto(
                        image, allow_manual=False
                    )
                    metodo = f"ia:{provider}"
                    logging.info(
                        "IA leyó el captcha de %s: %s (%s)",
                        user["documento"], code, provider,
                    )
                except Exception as exc:
                    logging.warning(
                        "La IA no pudo leer el captcha de %s: %s",
                        user["documento"], str(exc)[:200],
                    )
                    code = None

                if not code and attempt < ai_attempts:
                    attempts.append({
                        "intento": attempt, "metodo": "ia",
                        "resultado": "error_ia", "detalle": "La IA no devolvió un código.",
                    })
                    ensure_form(driver, wait)
                    fill_form(driver, wait, user, settings)
                    continue

            if not code:
                metodo = "manual"
                status(f"Escribe el captcha de {user['apellido']} {user['nombre']} en el diálogo")
                if _CAPTCHA_CALLBACK:
                    code = _CAPTCHA_CALLBACK(image, attempt)
                else:
                    code = _default_captcha_dialog(image, attempt)
            if not code:
                last_reason = "captcha_no_resuelto"
                last_detail = "No se ingresó el captcha."
                attempts.append({
                    "intento": attempt, "metodo": metodo,
                    "resultado": last_reason, "detalle": last_detail,
                })
                break

            _set_value(driver, driver.find_element(By.ID, ID_CAPTCHA), str(code).strip().upper())
            if not click_verify(driver, wait):
                last_reason = "error_navegador"
                last_detail = "No se encontró el botón Verificar."
                attempts.append({
                    "intento": attempt, "metodo": metodo,
                    "resultado": last_reason, "detalle": last_detail,
                })
                break

            valid = wait_validation(driver)
            attempts.append({
                "intento": attempt,
                "metodo": metodo,
                "resultado": "captcha_ok" if valid else "captcha_invalido",
                "detalle": "Captcha aceptado." if valid else "El sitio rechazó el captcha.",
            })
            if valid:
                captcha_ok = True
                break

            logging.warning("Captcha rechazado para %s (intento %s)", user["documento"], attempt)
            ensure_form(driver, wait)
            fill_form(driver, wait, user, settings)

        if not captcha_ok:
            if last_reason == "captcha_no_resuelto":
                screenshot = save_error_screenshot(driver, user, label=last_reason)
                return False, attempts, last_reason, last_detail, screenshot
            last_reason = "captcha_invalido"
            last_detail = "El sitio rechazó el captcha después de varios intentos."
            screenshot = save_error_screenshot(driver, user, label=last_reason)
            return False, attempts, last_reason, last_detail, screenshot

        status(f"Consultando {user['apellido']} {user['nombre']}…")
        if not click_consultar(driver, wait):
            last_reason = "error_navegador"
            last_detail = "No se encontró el botón Consultar."
            screenshot = save_error_screenshot(driver, user, label=last_reason)
            return False, attempts, last_reason, last_detail, screenshot

        result, detail = wait_result(driver, user["documento"])
        attempts.append({
            "intento": len(attempts) + 1,
            "metodo": "manual",
            "resultado": result,
            "detalle": detail,
        })
        if result != "ok":
            screenshot = save_error_screenshot(driver, user, label=result)
            return False, attempts, result, detail, screenshot

        try:
            path = print_report_pdf(driver, user, settings.filename_pattern, settings)
        except Exception as exc:
            last_reason = "error_pdf"
            last_detail = f"{type(exc).__name__}: {str(exc)[:200]}"
            screenshot = save_error_screenshot(driver, user, label=last_reason)
            return False, attempts, last_reason, last_detail, screenshot

        logging.info(
            "[OK] Guardado: %s | Perfil: %s | Doc: %s",
            path.name, user["perfil_grupo"], user["documento"],
        )
        return True, attempts, "", "", str(path)

    except (TimeoutException, WebDriverException) as exc:
        last_reason = "error_navegador"
        last_detail = f"{type(exc).__name__}: {str(exc)[:200]}"
        logging.error("Error de navegador con %s: %s", user["documento"], last_detail)
    except Exception as exc:
        last_reason = "error_inesperado"
        last_detail = f"{type(exc).__name__}: {str(exc)[:200]}"
        logging.error("Error inesperado con %s: %s", user["documento"], last_detail)

    screenshot = save_error_screenshot(driver, user, label=last_reason)
    return False, attempts, last_reason, last_detail, screenshot


# ---------------------------------------------------------------------------
# Informe
# ---------------------------------------------------------------------------

MOTIVOS_REPORTE = {
    "dato_incompleto": (
        "Faltan datos en el Excel.",
        "Completa la fila y vuelve a intentar.",
    ),
    "captcha_no_resuelto": (
        "No se ingresó el captcha.",
        "Vuelve a intentar y escribe el código en el diálogo.",
    ),
    "captcha_invalido": (
        "El sitio rechazó el captcha varias veces.",
        "Reintenta; si persiste, revisa la imagen o cambia de red.",
    ),
    "sin_resultado": (
        "No apareció el reporte de afiliaciones.",
        "Reintenta; revisa la captura.",
    ),
    "error_navegador": (
        "Falló el navegador.",
        "Reintenta; puede ser un timeout o un bloqueo temporal.",
    ),
    "error_inesperado": (
        "Error inesperado.",
        "Revisa la captura y el detalle.",
    ),
    "formulario_vacio": (
        "El sitio reportó campos obligatorios vacíos.",
        "Revisa los datos de esa fila en el Excel.",
    ),
    "datos_no_coinciden": (
        "El documento o los datos no coinciden.",
        "Revisa la cédula y los nombres en el Excel.",
    ),
    "fecha_no_corresponde": (
        "La fecha de expedición no corresponde.",
        "Revisa la columna Expedicion del Excel.",
    ),
    "sin_descarga": (
        "El sitio no entregó el PDF.",
        "Reintenta; si persiste, revisa la captura.",
    ),
    "error_pdf": (
        "No se pudo generar el PDF del reporte.",
        "Reintenta; revisa la captura y el visor del reporte.",
    ),
    "error_sitio": (
        "El sitio mostró un error.",
        "Revisa el mensaje y los datos del Excel.",
    ),
    "sesion_expirada": (
        "La sesión del RUAF expiró.",
        "La aplicación la renueva sola; reintenta si persiste.",
    ),
    "detenido": (
        "Proceso detenido.",
        "Vuelve a iniciar para continuar.",
    ),
}


def _motivo_info(motivo: str):
    return MOTIVOS_REPORTE.get(motivo, ("Error no clasificado.", "Revisa el detalle y la captura."))


def _clean_detail(text) -> str:
    return " ".join(str(text or "").split())[:300]


def write_failure_report(fallos, incomplete_rows, timestamp):
    if not fallos and not incomplete_rows:
        return None, None, None

    rows = []
    for fila in incomplete_rows:
        rows.append({
            "fila_excel": fila.get("fila", ""),
            "documento": fila.get("documento", ""),
            "nombre": fila.get("nombre", ""),
            "perfil": "",
            "motivo": fila.get("motivo", "dato_incompleto"),
            "detalle": _clean_detail(fila.get("detalle", "")),
            "intentos": 0,
            "metodo": "",
            "captura": "",
        })

    for fallo in fallos:
        intentos = fallo.get("intentos", [])
        metodo = ""
        if intentos:
            metodo = str(intentos[-1].get("metodo", "") or "")
        rows.append({
            "fila_excel": fallo.get("fila_excel", ""),
            "documento": fallo.get("documento", ""),
            "nombre": fallo.get("nombre", ""),
            "perfil": fallo.get("perfil", ""),
            "motivo": fallo.get("motivo", "sin_resultado"),
            "detalle": _clean_detail(fallo.get("detalle", "")),
            "intentos": len(intentos),
            "metodo": metodo,
            "captura": fallo.get("captura", ""),
        })

    rows.sort(key=lambda r: (int(r["fila_excel"]) if str(r["fila_excel"]).isdigit() else 10**9))

    csv_path = BASE_OUTPUT_DIR / f"REPORTE_FALLOS_{timestamp}.csv"
    txt_path = BASE_OUTPUT_DIR / f"REPORTE_FALLOS_{timestamp}.txt"
    xlsx_path = BASE_OUTPUT_DIR / f"REPORTE_FALLOS_{timestamp}.xlsx"

    headers = [
        "Fila Excel", "Documento", "Nombre", "Perfil", "Motivo",
        "Qué pasó", "Qué hacer", "Mensaje del sitio", "Intentos", "Método", "Captura",
    ]

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        for row in rows:
            explicacion, que_hacer = _motivo_info(row["motivo"])
            writer.writerow([
                row["fila_excel"], row["documento"], row["nombre"], row["perfil"],
                row["motivo"], explicacion, que_hacer, row["detalle"],
                row["intentos"], row["metodo"], row["captura"],
            ])

    fecha_txt = timestamp
    try:
        fecha_txt = datetime.strptime(timestamp, "%Y%m%d_%H%M%S").strftime("%d/%m/%Y %H:%M")
    except Exception:
        pass

    from collections import Counter

    contador = Counter(row["motivo"] for row in rows)

    lineas = [
        "REPORTE DE FALLOS - RUAF - AFILIACIONES",
        f"Generado: {fecha_txt}",
        f"Excel: {EXCEL_PATH}",
        f"Carpeta de salida: {BASE_OUTPUT_DIR}",
        "",
        f"Total de fallos: {len(rows)}",
        "",
        "RESUMEN",
    ]
    for motivo, cantidad in contador.most_common():
        explicacion, _ = _motivo_info(motivo)
        lineas.append(f"  {motivo}: {cantidad}  ({explicacion})")

    lineas += ["", "DETALLE (en el orden del Excel)", ""]
    for index, row in enumerate(rows, start=1):
        explicacion, que_hacer = _motivo_info(row["motivo"])
        lineas.append(
            f"{index}. Fila {row['fila_excel']} | {row['documento']} | "
            f"{row['nombre']} | {row['perfil'] or '(sin perfil)'}"
        )
        lineas.append(f"   Motivo: {row['motivo']}")
        lineas.append(f"   Qué pasó: {explicacion}")
        lineas.append(f"   Qué hacer: {que_hacer}")
        if row["detalle"]:
            lineas.append(f"   Mensaje del sitio: {row['detalle']}")
        metodo = f" | Método: {row['metodo']}" if row["metodo"] else ""
        lineas.append(f"   Intentos: {row['intentos']}{metodo}")
        if row["captura"]:
            lineas.append(f"   Captura: {row['captura']}")
        lineas.append("")

    txt_path.write_text("\n".join(lineas) + "\n", encoding="utf-8")

    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.utils import get_column_letter

        wb = Workbook()
        ws = wb.active
        ws.title = "Fallos"

        header_fill = PatternFill("solid", fgColor="142749")
        header_font = Font(color="FFFFFF", bold=True)
        thin = Side(style="thin", color="D9D9D9")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        ws.append(headers)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        for row in rows:
            explicacion, que_hacer = _motivo_info(row["motivo"])
            captura_nombre = Path(row["captura"]).name if row["captura"] else ""
            ws.append([
                row["fila_excel"], row["documento"], row["nombre"], row["perfil"],
                row["motivo"], explicacion, que_hacer, row["detalle"],
                row["intentos"], row["metodo"], captura_nombre,
            ])
            row_index = ws.max_row
            for cell in ws[row_index]:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = border
            if row["captura"]:
                cell = ws.cell(row=row_index, column=11)
                cell.hyperlink = "file://" + str(Path(row["captura"]).resolve())
                cell.style = "Hyperlink"

        for index, width in enumerate([9, 14, 34, 12, 20, 45, 48, 55, 9, 12, 45], start=1):
            ws.column_dimensions[get_column_letter(index)].width = width
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        resumen = wb.create_sheet("Resumen")
        resumen.append(["Motivo", "Cantidad", "Qué significa", "Qué hacer"])
        for cell in resumen[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        for motivo, cantidad in contador.most_common():
            explicacion, que_hacer = _motivo_info(motivo)
            resumen.append([motivo, cantidad, explicacion, que_hacer])
        for row in resumen.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = border
        for index, width in enumerate([24, 10, 55, 60], start=1):
            resumen.column_dimensions[get_column_letter(index)].width = width
        resumen.freeze_panes = "A2"

        wb.save(xlsx_path)
    except Exception as exc:
        logging.warning("No se pudo generar el reporte Excel: %s", exc)
        xlsx_path = None

    return csv_path, txt_path, xlsx_path


# ---------------------------------------------------------------------------
# Lectura del Excel
# ---------------------------------------------------------------------------

def load_users_from_excel(excel_path, selected_profiles=None, sheet_name=None,
                          header_row=None, column_map=None):
    sheet = SETTINGS.sheet_name if sheet_name is None else sheet_name
    header = SETTINGS.header_row if header_row is None else header_row
    mapping = SETTINGS.column_map if column_map is None else column_map

    users, incomplete, stats = load_users(
        excel_path,
        sheet_name=sheet,
        header_row=header,
        column_map=mapping,
        selected_profiles=selected_profiles,
    )
    globals()["LAST_STATS"] = stats
    logging.info(
        "Columnas detectadas -> Documento: %r | Nombres: %r | Apellidos: %r | Perfil: %r | Tipo doc: %r",
        stats["columnas"].get("documento"), stats["columnas"].get("nombre"),
        stats["columnas"].get("apellido"), stats["columnas"].get("perfil"),
        stats["columnas"].get("tipo_documento"),
    )
    return users, incomplete


# ---------------------------------------------------------------------------
# CLI / ejecución
# ---------------------------------------------------------------------------

def choose_profiles():
    print("\nPerfiles para procesar:")
    print("  1. AUXILIAR")
    print("  2. ENFERMERO")
    print("  3. MEDICO")
    print("  4. TODOS")
    option = input("Selecciona opción (1-4) [4]: ").strip() or "4"
    return {
        "1": ["AUXILIAR"],
        "2": ["ENFERMERO"],
        "3": ["MEDICO"],
        "4": None,
    }.get(option, None)


def choose_batch_size(total_pending):
    value = input(f"Cantidad por tanda (0 = todos, máximo {total_pending}) [0]: ").strip() or "0"
    try:
        size = int(value)
    except ValueError:
        size = 0
    return total_pending if size <= 0 else min(size, total_pending)


def build_arg_parser():
    parser = argparse.ArgumentParser(
        description="Consulta masiva de afiliaciones en el RUAF - SISPRO."
    )
    parser.add_argument("--config", help="Archivo config.json alternativo")
    parser.add_argument("--excel", help="Ruta del Excel de caracterización")
    parser.add_argument("--salida", help="Carpeta base de salida")
    parser.add_argument("--hoja", help="Nombre o índice de la hoja")
    parser.add_argument("--fila-encabezado", type=int, help="Fila (1-based) de encabezados")
    parser.add_argument("--perfil", choices=["AUXILIAR", "ENFERMERO", "MEDICO", "TODOS"])
    parser.add_argument("--tanda", type=int, help="Cantidad de registros en esta corrida")
    parser.add_argument("--timeout-manual", type=int, help="Segundos de espera del captcha manual")
    parser.add_argument("--si", dest="no_prompts", action="store_true", help="No preguntar nada")
    parser.add_argument("--reiniciar-estado", action="store_true", help="Ignorar el estado previo")
    parser.add_argument("--dry-run", action="store_true", help="Solo validar el Excel")
    parser.add_argument("--headless", action="store_true", help="No recomendado: el captcha es manual")
    parser.add_argument("--informes", action="store_true", help="Generar Excel de datos y PPTX al terminar")
    parser.add_argument("--sin-informes", dest="sin_informes", action="store_true", help="No generar Excel/PPTX")
    return parser


def _future_dates(users):
    hoy = datetime.now().date()
    futuras = []
    for user in users:
        try:
            fecha = datetime.strptime(user["fecha_expedicion"], "%d/%m/%Y").date()
        except Exception:
            continue
        if fecha > hoy:
            futuras.append(user)
    return futuras


def dry_run(settings: Settings):
    apply_settings(settings)
    print("\n" + "=" * 70)
    print("  VALIDACIÓN DE DATOS RUAF (sin navegador)")
    print("=" * 70)
    print(f"Excel:       {EXCEL_PATH}")
    print(f"Hoja:        {settings.sheet_name or '(primera)'}")
    print(f"Encabezados: fila {settings.header_row}")
    print("=" * 70)

    try:
        users, incomplete, stats = load_users(
            EXCEL_PATH,
            sheet_name=settings.sheet_name,
            header_row=settings.header_row,
            column_map=settings.column_map,
            selected_profiles=settings.profiles,
        )
    except Exception as exc:
        print(f"\n[ERROR] {exc}")
        return 2

    print(f"Hoja leída:      {stats['hoja']}")
    print(f"Filas con datos: {stats['total_filas']}")
    print(f"Válidos:         {stats['validos']}")
    print(f"Incompletos:     {stats['incompletos']}")
    print(f"Duplicados:      {stats['duplicados']}")
    print(f"Filtrados:       {stats['filtrados_por_perfil']}")

    futuras = _future_dates(users)
    if futuras:
        print(f"\n[WARN] {len(futuras)} fechas de expedición son futuras (posible error del Excel):")
        for user in futuras[:10]:
            print(f"  fila {user['fila_excel']}: {user['documento']} -> {user['fecha_expedicion']}")

    print("\nColumnas usadas:")
    for clave, columna in stats["columnas"].items():
        print(f"  - {clave:14s} -> {columna}")

    if users:
        print(f"\nPrimero: {users[0]['apellido']} {users[0]['nombre']} ({users[0]['documento']})")
        print(f"Último:  {users[-1]['apellido']} {users[-1]['nombre']} ({users[-1]['documento']})")

    if incomplete:
        print("\nPrimeros incompletos:")
        for fila in incomplete[:10]:
            print(f"  fila {fila['fila']}: {fila['detalle']}")

    print("\n[OK] Validación terminada.")
    return 0


def run(settings: Settings, on_progress=None, on_status=None, should_stop=None, on_captcha=None):
    apply_settings(settings)
    set_stop_checker(should_stop)
    set_captcha_callback(on_captcha)

    interactive = bool(settings.interactive)
    detenido = False

    print("\n" + "=" * 70)
    print("  RUAF - AFILIACIONES (SISPRO - MINISTERIO DE SALUD)")
    print("=" * 70)
    print(f"Sistema:      {platform.system()} {platform.machine()}")
    print(f"Excel fuente: {EXCEL_PATH}")
    print(f"Salida base:  {BASE_OUTPUT_DIR}")
    print("=" * 70)

    configure_folders()

    if not Path(EXCEL_PATH).exists():
        print(f"\n[ERROR] No existe el Excel: {EXCEL_PATH}")
        return {"codigo": 2, "success": 0, "failed": 0, "total": 0, "detenido": False}

    if settings.profiles is not None:
        selected_profiles = None if not settings.profiles else list(settings.profiles)
    elif not interactive:
        selected_profiles = None
    else:
        selected_profiles = choose_profiles()

    try:
        users, incomplete_rows = load_users_from_excel(
            EXCEL_PATH,
            selected_profiles=selected_profiles,
            sheet_name=settings.sheet_name,
            header_row=settings.header_row,
            column_map=settings.column_map,
        )
    except Exception as exc:
        print(f"\n[ERROR] Error leyendo el Excel: {exc}")
        return {"codigo": 2, "success": 0, "failed": 0, "total": 0, "detenido": False}

    if not users:
        print("\n[ERROR] No se encontraron usuarios válidos para ese filtro.")
        return {"codigo": 2, "success": 0, "failed": 0, "total": 0, "detenido": False}

    state = load_state()
    processed_set = set(state.get("procesados", []))
    if processed_set and not settings.reset_state:
        valid_processed = {doc for doc in processed_set if _find_valid_pdf(doc)}
        removed = len(processed_set) - len(valid_processed)
        if removed:
            print(f"[WARN] {removed} procesados previos no tienen un PDF válido; se vuelven a consultar.")
        processed_set = valid_processed
        state["procesados"] = sorted(processed_set)

    print(f"\nUsuarios válidos: {len(users)}")
    print(f"Filas con datos incompletos: {len(incomplete_rows)}")
    print(f"Procesados previos en estado: {len(processed_set)}")

    futuras = _future_dates(users)
    if futuras:
        print(f"[WARN] {len(futuras)} fechas de expedición son futuras; posible error del Excel.")

    if not settings.reset_state:
        if interactive:
            if settings.resume:
                answer = input("¿Reanudar omitiendo procesados previos? [S/n]: ").strip().lower()
                resume = answer != "n"
            else:
                answer = input("¿Reanudar omitiendo procesados previos? [s/N]: ").strip().lower()
                resume = answer == "s"
        else:
            resume = bool(settings.resume)

        if resume:
            users = [u for u in users if u["documento"] not in processed_set]

    if not users:
        print("\n[OK] No hay pendientes para los perfiles seleccionados.")
        return {"codigo": 0, "success": 0, "failed": 0, "total": 0, "detenido": False}

    if settings.batch_size is not None:
        batch_size = settings.batch_size if settings.batch_size > 0 else len(users)
    elif not interactive:
        batch_size = len(users)
    else:
        batch_size = choose_batch_size(len(users))
    users = users[:batch_size]

    total = len(users)
    headless = bool(settings.headless)
    if headless:
        print("\n[WARN] El captcha es manual: se abrirá Chrome visible de todos modos.")
        headless = False

    print(f"\n[INFO] Esta tanda procesará {total} usuarios.")
    print("[INFO] La IA intentará leer el captcha; si falla, te pedirá escribirlo.")
    print("[INFO] Después del captcha, la aplicación consulta e imprime el PDF sola.")

    driver = None
    success = 0
    failed = 0
    fallos = []
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if on_progress:
        try:
            on_progress(0, total, 0, 0)
        except Exception:
            pass

    try:
        driver = create_driver(headless=headless)
        try:
            driver.maximize_window()
            focus_chrome()
        except Exception:
            pass
        wait = WebDriverWait(driver, 30)

        for idx, user in enumerate(users, start=1):
            if stop_requested():
                detenido = True
                print("\n[INFO] Detenido por el usuario.")
                break

            print(
                f"\n[{idx}/{total}] {user['apellido']} {user['nombre']} "
                f"({user['documento']}) - {user['perfil_grupo']}"
            )
            if on_status:
                try:
                    on_status(f"[{idx}/{total}] {user['apellido']} {user['nombre']}")
                except Exception:
                    pass

            ok, intentos, motivo, detalle, captura = process_user(
                driver, wait, user, settings, on_status=on_status
            )

            if motivo == "detenido":
                detenido = True
                print("\n[INFO] Detenido por el usuario.")
                break

            if ok:
                success += 1
                processed_set.add(user["documento"])
                state.setdefault("fallidos", [])
                if user["documento"] in state["fallidos"]:
                    state["fallidos"].remove(user["documento"])
            else:
                failed += 1
                fallos.append({
                    "documento": user["documento"],
                    "nombre": f"{user['nombre']} {user['apellido']}",
                    "perfil": user["perfil_grupo"],
                    "fila_excel": user.get("fila_excel", ""),
                    "motivo": motivo or "sin_resultado",
                    "detalle": detalle,
                    "intentos": intentos,
                    "captura": captura,
                })
                state.setdefault("fallidos", [])
                if user["documento"] not in state["fallidos"]:
                    state["fallidos"].append(user["documento"])

            state["procesados"] = sorted(processed_set)
            save_state(state)

            if on_progress:
                try:
                    on_progress(idx, total, success, failed)
                except Exception:
                    pass

            time.sleep(0.8)

    except KeyboardInterrupt:
        detenido = True
        print("\n\n[INFO] Proceso interrumpido por el usuario.")
    except Exception as exc:
        print(f"\n[ERROR] Error general: {exc}")
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    csv_path, txt_path, xlsx_path = write_failure_report(fallos, incomplete_rows, timestamp)

    datos_xlsx = ""
    pptx_path = ""
    if settings.generate_reports:
        try:
            import generar_informes

            resultado = generar_informes.generar_todo(
                output_dir=BASE_OUTPUT_DIR,
                reports_dir=REPORTS_DIR,
                excel_path=EXCEL_PATH,
                timestamp=timestamp,
                sheet_name=settings.sheet_name,
                header_row=settings.header_row,
                column_map=settings.column_map,
            )
            datos_xlsx = resultado.get("excel", "")
            pptx_path = resultado.get("pptx", "")
            if datos_xlsx:
                print(f"[OK] Excel de datos: {datos_xlsx}")
            if pptx_path:
                print(f"[OK] Informe PPTX:   {pptx_path}")
        except Exception as exc:
            print(f"[WARN] No se pudieron generar los informes: {exc}")

    print("\n" + "=" * 70)
    print("RESUMEN")
    print(f"[OK] Exitosos: {success}")
    print(f"[ERROR] Fallidos: {failed}")
    print(f"Carpeta base: {BASE_OUTPUT_DIR}")
    print(f"Estado: {STATE_FILE}")
    if txt_path:
        print(f"Informe de fallos: {txt_path}")
        print(f"Informe (CSV):     {csv_path}")
        if xlsx_path:
            print(f"Informe (Excel):   {xlsx_path}")
    else:
        print("[OK] No hubo fallos que reportar.")
    print("=" * 70)

    return {
        "codigo": 0,
        "success": success,
        "failed": failed,
        "total": total,
        "detenido": detenido,
        "csv": str(csv_path) if csv_path else "",
        "txt": str(txt_path) if txt_path else "",
        "xlsx": str(xlsx_path) if xlsx_path else "",
        "datos_xlsx": datos_xlsx,
        "pptx": pptx_path,
        "carpeta": str(BASE_OUTPUT_DIR),
    }


def main(argv=None):
    args = build_arg_parser().parse_args(argv)
    settings = load_settings(Path(args.config) if args.config else None)

    if args.excel:
        settings.excel_path = Path(args.excel)
    if args.salida:
        settings.output_dir = Path(args.salida)
    if args.hoja is not None:
        settings.sheet_name = args.hoja
    if args.fila_encabezado is not None:
        settings.header_row = max(1, args.fila_encabezado)
    if args.perfil:
        settings.profiles = None if args.perfil == "TODOS" else [args.perfil]
    if args.tanda is not None:
        settings.batch_size = args.tanda
    if args.timeout_manual is not None:
        settings.manual_timeout = max(30, args.timeout_manual)
    if args.reiniciar_estado:
        settings.reset_state = True
    if args.no_prompts:
        settings.interactive = False
    if args.headless:
        settings.headless = True
    if args.informes:
        settings.generate_reports = True
    if args.sin_informes:
        settings.generate_reports = False

    if args.dry_run:
        return dry_run(settings)

    return run(settings)["codigo"]


if __name__ == "__main__":
    raise SystemExit(main())

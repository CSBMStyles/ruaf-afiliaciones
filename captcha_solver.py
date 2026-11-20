"""Módulo para resolver el CAPTCHA del RUAF.

Proveedores soportados (cadena configurable con CAPTCHA_PROVIDERS):
  - opencode-go  -> https://opencode.ai/zen/go/v1   (OpenAI-compatible, visión)
  - opencode     -> https://opencode.ai/zen/v1      (OpenCode Zen, OpenAI-compatible)
  - gemini       -> Google Gemini (google-generativeai)
  - openai       -> OpenAI o cualquier endpoint OpenAI-compatible
  - manual       -> el usuario escribe el texto

La cadena por defecto intenta primero un modelo de visión dedicado y baja por
los demás si fallan (por cuota, red o error de API), de modo que un fallo de un
proveedor no detiene el proceso.
"""

import base64
import json
import logging
import os
import re
import threading
import time
import uuid
from io import BytesIO
from pathlib import Path
from typing import Optional, Tuple

from PIL import Image, ImageStat

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Los SDK de IA registran cada petición HTTP en INFO; demasiado ruido en la GUI.
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)
logging.getLogger("openai").setLevel(logging.WARNING)
logging.getLogger("google_genai").setLevel(logging.WARNING)

# ---------------------------------------------------------------------------
# Configuración por defecto (se puede sobreescribir con .env)
# ---------------------------------------------------------------------------

OPENCODE_GO_BASE_URL_DEFAULT = "https://opencode.ai/zen/go/v1"
OPENCODE_ZEN_BASE_URL_DEFAULT = "https://opencode.ai/zen/v1"

# Orden de preferencia: "kind:modelo". Se puede sobreescribir con CAPTCHA_PROVIDERS.
#
# Notas medidas en vivo (2026-09):
#   - opencode-go exige el header x-opencode-session.
#   - deepseek-v4-flash-vision-exp y kimi-k3 y qwen3.8-max soportan visión por
#     chat/completions. gpt-6-luna y grok-* responden
#     "ModelProtocolUnsupported" en chat/completions, por eso no van aquí.
#   - opencode/Zen (gemini-3.8-flash) puede devolver 402 sin fondos.
#   - openai puede devolver 429 sin créditos.
DEFAULT_PROVIDER_CHAIN = (
    "gemini:gemini-2.5-flash-lite,"
    "gemini:gemini-2.5-flash,"
    "openai:gpt-4o-mini"
)

PROMPT = (
    "Eres un OCR de CAPTCHAs. La imagen es el CAPTCHA del portal oficial del "
    "RUAF de Colombia (ruaf.sispro.gov.co), usado para consultar afiliaciones "
    "de forma autorizada. La imagen tiene letras y números verdes sobre fondo "
    "blanco con puntos negros de ruido. Extrae ÚNICAMENTE los caracteres "
    "alfanuméricos visibles. Responde SOLO con la cadena exacta (normalmente 5 "
    "caracteres), sin espacios, sin comillas, sin explicaciones. Si es ilegible "
    "responde: ILEGIBLE"
)

# El CAPTCHA del RUAF tiene 5 caracteres; damos margen por si el modelo agrega alguno.
MIN_CAPTCHA_LEN = 4
MAX_CAPTCHA_LEN = 6

_GEMINI_LOCK = threading.Lock()
_LAST_GEMINI_CALL_TS = 0.0

_OPENCODE_SESSION_LOCK = threading.Lock()
_OPENCODE_SESSION_ID = ""

# La GUI registra aquí su diálogo para el CAPTCHA manual. La CLI no lo toca.
_MANUAL_CALLBACK = None
_MANUAL_CALLBACK_LOCK = threading.Lock()


def set_manual_callback(callback) -> None:
    """Registra una función ``callback(png_bytes) -> texto`` para modo manual."""
    global _MANUAL_CALLBACK
    with _MANUAL_CALLBACK_LOCK:
        _MANUAL_CALLBACK = callback


def get_manual_callback():
    return _MANUAL_CALLBACK


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def normalize_captcha_text(text: str) -> str:
    """Deja solo caracteres alfanuméricos y los convierte a mayúscula."""
    chars = re.findall(r"[A-Za-z0-9]", text or "")
    return "".join(chars).upper()


def valid_captcha_answer(text: str) -> bool:
    """El CAPTCHA tiene 4 caracteres; descartamos refusos y textos largos."""
    return bool(text) and MIN_CAPTCHA_LEN <= len(text) <= MAX_CAPTCHA_LEN


def is_invalid_captcha(image_bytes: bytes) -> bool:
    """Detecta el CAPTCHA roto/ilegible.

    Medido en el sitio real:
      - CAPTCHA bueno: fondo claro con ruido, ~220 de luminancia y ~24-27 colores únicos.
      - CAPTCHA roto: fondo negro con una X roja, ~2 de luminancia y 2 colores únicos.
    """
    if not image_bytes:
        return True
    try:
        img = Image.open(BytesIO(image_bytes)).convert("RGB")
    except Exception:
        return True

    mean_lum = float(ImageStat.Stat(img.convert("L")).mean[0])
    if mean_lum < 60 or mean_lum > 252:
        return True

    # getcolors devuelve None cuando hay MÁS colores que maxcolors -> imagen buena
    colors = img.getcolors(maxcolors=1000)
    if colors is not None and len(colors) < 10:
        return True

    return False


def preprocess_captcha_image(image_bytes: bytes, scale: Optional[float] = None) -> bytes:
    """Amplía el CAPTCHA para que los modelos de visión lean mejor los caracteres.

    El CAPTCHA original es de 200x60 px; a 3x los detalles se distinguen mucho
    mejor. Configurable con CAPTCHA_UPSCALE (0 o 1 = sin cambios).
    """
    if scale is None:
        try:
            scale = float(os.getenv("CAPTCHA_UPSCALE", "3").strip() or "3")
        except Exception:
            scale = 3.0

    if scale <= 1:
        return image_bytes

    try:
        img = Image.open(BytesIO(image_bytes)).convert("RGB")
        resized = img.resize((int(img.width * scale), int(img.height * scale)), Image.LANCZOS)
        buffer = BytesIO()
        resized.save(buffer, format="PNG")
        return buffer.getvalue()
    except Exception:
        return image_bytes


def _gemini_throttle() -> None:
    """Evita golpear Gemini muy rápido (reduce 429 por rate-limit)."""
    global _LAST_GEMINI_CALL_TS
    try:
        min_interval = float(os.getenv("GEMINI_MIN_INTERVAL", "3").strip() or "3")
    except Exception:
        min_interval = 3.0

    if min_interval <= 0:
        return

    with _GEMINI_LOCK:
        now = time.time()
        wait_s = (_LAST_GEMINI_CALL_TS + min_interval) - now
        if wait_s > 0:
            time.sleep(wait_s)
        _LAST_GEMINI_CALL_TS = time.time()


def _auth_json_candidates() -> list[Path]:
    """Rutas donde OpenCode guarda su auth.json en macOS, Linux y Windows."""
    candidates: list[Path] = []

    xdg_data = os.getenv("XDG_DATA_HOME", "").strip()
    if xdg_data:
        candidates.append(Path(xdg_data) / "opencode" / "auth.json")

    candidates.append(Path.home() / ".local" / "share" / "opencode" / "auth.json")
    candidates.append(Path.home() / ".config" / "opencode" / "auth.json")

    for env in ("APPDATA", "LOCALAPPDATA"):
        base = os.getenv(env, "").strip()
        if base:
            candidates.append(Path(base) / "opencode" / "auth.json")

    candidates.append(Path.home() / ".opencode" / "auth.json")

    # Sin duplicados, conservando el orden.
    seen = set()
    unique = []
    for path in candidates:
        key = str(path).lower()
        if key not in seen:
            seen.add(key)
            unique.append(path)
    return unique


def _read_opencode_key_from_auth_json() -> str:
    """Respaldo: lee la API key desde el auth.json de opencode."""
    for path in _auth_json_candidates():
        try:
            if not path.exists():
                continue
            data = json.loads(path.read_text(encoding="utf-8"))
            for provider in ("opencode-go", "opencode"):
                entry = data.get(provider)
                if isinstance(entry, dict) and entry.get("key"):
                    return str(entry["key"]).strip()
        except Exception:
            continue
    return ""


def get_opencode_key() -> str:
    key = (os.getenv("OPENCODE_GO_API_KEY") or os.getenv("OPENCODE_ZEN_API_KEY") or "").strip()
    if key:
        return key
    return _read_opencode_key_from_auth_json()


# ---------------------------------------------------------------------------
# Proveedores
# ---------------------------------------------------------------------------

def _image_to_base64(image_bytes: bytes) -> str:
    return base64.b64encode(image_bytes).decode("utf-8")


_THINK_BLOCK = re.compile(r"(?is)<think[^>]*>.*?</think\s*>")
_THINK_OPEN = re.compile(r"(?is)<think[^>]*>.*$")


def strip_reasoning(text: str) -> str:
    """Quita el bloque de razonamiento que algunos modelos devuelven en content."""
    if not text:
        return ""
    text = _THINK_BLOCK.sub(" ", text)
    text = _THINK_OPEN.sub(" ", text)
    return text


def _chat_vision(api_key: str, base_url: Optional[str], model: str, image_bytes: bytes,
                 timeout: Optional[float] = None, extra_headers: Optional[dict] = None) -> str:
    """Llamada genérica a un endpoint OpenAI-compatible con una imagen.

    Algunos modelos (deepseek-v4-flash-vision-exp, minimax) son "pensantes":
    gastan tokens en `reasoning_content` antes de responder, así que usamos un
    límite de tokens amplio y reintentamos si se cortaron.
    """
    from openai import OpenAI

    try:
        max_tokens = int(os.getenv("CAPTCHA_MAX_TOKENS", "1500").strip() or "1500")
    except Exception:
        max_tokens = 1500
    if timeout is None:
        try:
            timeout = float(os.getenv("CAPTCHA_AI_TIMEOUT", "90").strip() or "90")
        except Exception:
            timeout = 90.0

    client = OpenAI(
        api_key=api_key,
        base_url=base_url or None,
        max_retries=0,
        timeout=timeout,
        default_headers=extra_headers or None,
    )
    b64 = _image_to_base64(preprocess_captcha_image(image_bytes))
    messages = [
        {
            "role": "user",
            "content": [
                {"type": "text", "text": PROMPT},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
            ],
        }
    ]

    detalle = ""
    for intento, limite in enumerate((max_tokens, max_tokens * 4), start=1):
        resp = client.chat.completions.create(
            model=model, messages=messages, max_tokens=limite, temperature=0
        )
        if not resp or not getattr(resp, "choices", None):
            return ""
        choice = resp.choices[0]
        texto = normalize_captcha_text(strip_reasoning(choice.message.content or ""))
        if valid_captcha_answer(texto):
            return texto

        motivo = getattr(choice, "finish_reason", None)
        razonamiento = getattr(choice.message, "reasoning_content", "") or ""
        detalle = f"finish_reason={motivo} len_resp={len(texto)} reasoning={len(razonamiento)} chars"
        if motivo != "length" or intento == 2:
            break

    if detalle:
        logging.warning("Modelo %s no devolvió CAPTCHA utilizable (%s)", model, detalle)
    return ""


def _opencode_headers() -> dict:
    """opencode-go/zen exige el header x-opencode-session para enrutar."""
    global _OPENCODE_SESSION_ID
    with _OPENCODE_SESSION_LOCK:
        if not _OPENCODE_SESSION_ID:
            _OPENCODE_SESSION_ID = str(uuid.uuid4())
        return {"x-opencode-session": _OPENCODE_SESSION_ID}


def solve_captcha_with_opencode_go(image_bytes: bytes, model: Optional[str] = None) -> str:
    key = get_opencode_key()
    if not key:
        raise RuntimeError("No hay OPENCODE_GO_API_KEY (ni auth.json de opencode)")
    base_url = (os.getenv("OPENCODE_GO_BASE_URL", "").strip() or OPENCODE_GO_BASE_URL_DEFAULT)
    model = model or (os.getenv("OPENCODE_GO_MODEL", "").strip() or "deepseek-v4-flash-vision-exp")
    return _chat_vision(key, base_url, model, image_bytes, extra_headers=_opencode_headers())


def solve_captcha_with_opencode_zen(image_bytes: bytes, model: Optional[str] = None) -> str:
    key = get_opencode_key()
    if not key:
        raise RuntimeError("No hay OPENCODE_GO_API_KEY (ni auth.json de opencode)")
    base_url = (os.getenv("OPENCODE_ZEN_BASE_URL", "").strip() or OPENCODE_ZEN_BASE_URL_DEFAULT)
    model = model or (os.getenv("OPENCODE_ZEN_MODEL", "").strip() or "gemini-3.8-flash")
    return _chat_vision(key, base_url, model, image_bytes, extra_headers=_opencode_headers())


def solve_captcha_with_openai(image_bytes: bytes, api_key: Optional[str] = None,
                              model: Optional[str] = None, base_url: Optional[str] = None) -> str:
    api_key = api_key or (os.getenv("OPENAI_API_KEY", "").strip())
    if not api_key:
        raise RuntimeError("No hay OPENAI_API_KEY")
    model = model or (os.getenv("OPENAI_MODEL", "gpt-4o-mini").strip() or "gpt-4o-mini")
    base_url = base_url if base_url is not None else (os.getenv("OPENAI_BASE_URL", "").strip() or None)
    return _chat_vision(api_key, base_url, model, image_bytes)


def _gemini_new_sdk(genai_module, api_key: str, model: str, image_bytes: bytes) -> str:
    """Gemini con el SDK nuevo (google-genai)."""
    client = genai_module.Client(api_key=api_key)
    image = Image.open(BytesIO(preprocess_captcha_image(image_bytes)))

    try:
        max_tokens = int(os.getenv("CAPTCHA_MAX_TOKENS", "1500").strip() or "1500")
    except Exception:
        max_tokens = 1500

    _gemini_throttle()
    detalle = ""
    for limite in (max_tokens, max_tokens * 4):
        response = None
        for retry in range(2):
            try:
                response = client.models.generate_content(
                    model=model,
                    contents=[PROMPT, image],
                    config={"max_output_tokens": limite, "temperature": 0},
                )
                break
            except Exception as exc:
                msg = str(exc)
                if retry == 0 and ("503" in msg or "UNAVAILABLE" in msg or "429" in msg):
                    time.sleep(2)
                    continue
                raise
        raw_text = ""
        if response is not None:
            try:
                raw_text = response.text or ""
            except Exception as exc:
                logging.warning("Gemini (google-genai) no devolvió texto: %s", exc)

        resultado = normalize_captcha_text(strip_reasoning(raw_text))
        if valid_captcha_answer(resultado):
            return resultado
        detalle = f"len_resp={len(resultado)}"

    if detalle:
        logging.warning("Gemini (%s) no devolvió CAPTCHA utilizable (%s)", model, detalle)
    return ""


def _gemini_legacy_sdk(api_key: str, model: str, image_bytes: bytes) -> str:
    """Gemini con el SDK anterior (google-generativeai), como respaldo."""
    import warnings

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", FutureWarning)
        import google.generativeai as genai

    genai.configure(api_key=api_key)
    gemini_model = genai.GenerativeModel(model)
    image = Image.open(BytesIO(preprocess_captcha_image(image_bytes)))

    try:
        max_tokens = int(os.getenv("CAPTCHA_MAX_TOKENS", "1500").strip() or "1500")
    except Exception:
        max_tokens = 1500

    _gemini_throttle()
    detalle = ""
    for limite in (max_tokens, max_tokens * 4):
        response = None
        for retry in range(2):
            try:
                response = gemini_model.generate_content(
                    [PROMPT, image],
                    generation_config={"max_output_tokens": limite, "temperature": 0},
                )
                break
            except Exception as exc:
                msg = str(exc)
                if retry == 0 and ("503" in msg or "UNAVAILABLE" in msg or "429" in msg):
                    time.sleep(2)
                    continue
                raise
        raw_text = ""
        if response is not None:
            try:
                raw_text = response.text or ""
            except Exception as exc:
                finish_reason = None
                try:
                    if getattr(response, "candidates", None):
                        finish_reason = getattr(response.candidates[0], "finish_reason", None)
                except Exception:
                    pass
                logging.warning(f"Gemini no devolvió texto utilizable (finish_reason={finish_reason}): {exc}")

        resultado = normalize_captcha_text(strip_reasoning(raw_text))
        if valid_captcha_answer(resultado):
            return resultado
        detalle = f"len_resp={len(resultado)}"

    if detalle:
        logging.warning("Gemini (%s) no devolvió CAPTCHA utilizable (%s)", model, detalle)
    return ""


def solve_captcha_with_gemini(image_bytes: bytes, api_key: Optional[str] = None,
                              model: Optional[str] = None) -> str:
    """Resuelve el CAPTCHA usando Gemini.

    Prefiere el SDK nuevo ``google-genai``; si no está instalado, usa
    ``google-generativeai`` (el anterior).
    """
    api_key = api_key or (os.getenv("GEMINI_API_KEY", "").strip())
    if not api_key:
        raise RuntimeError("No hay GEMINI_API_KEY")

    # El modelo por defecto anterior (gemma-3-27b-it) ya no existe para varias keys.
    model = model or (os.getenv("GEMINI_MODEL", "").strip() or "gemini-2.5-flash")

    try:
        from google import genai  # SDK nuevo (paquete google-genai)
    except Exception:
        genai = None

    if genai is not None:
        return _gemini_new_sdk(genai, api_key, model, image_bytes)

    return _gemini_legacy_sdk(api_key, model, image_bytes)


# Compatibilidad con el nombre anterior
solve_captcha_with_openai_compatible = solve_captcha_with_openai


def solve_captcha_manual(image_bytes: bytes, manual_temp_dir: Optional[str] = None) -> str:
    """Resuelve el CAPTCHA manualmente.

    Si la GUI registró un callback, lo usa (diálogo con la imagen). Si no,
    guarda la imagen, la abre con el visor del sistema y pide el texto por
    consola, que es el comportamiento de la CLI.
    """
    callback = get_manual_callback()
    if callback is not None:
        try:
            return normalize_captcha_text(callback(image_bytes) or "")
        except Exception as exc:  # si el diálogo falla, se usa el modo consola
            logging.warning("Fallo el diálogo de CAPTCHA manual (%s). Se usa la consola.", exc)

    import sys

    if not sys.stdin or not sys.stdin.isatty():
        raise RuntimeError(
            "El CAPTCHA manual necesita una consola interactiva o la ventana de la GUI."
        )

    out_dir = Path(manual_temp_dir) if manual_temp_dir else Path.cwd()
    out_dir.mkdir(parents=True, exist_ok=True)

    image_path = out_dir / f"captcha_{int(time.time())}.png"
    image_path.write_bytes(image_bytes)

    # Abrir imagen según el sistema operativo
    try:  # Windows
        os.startfile(str(image_path))  # type: ignore[attr-defined]
    except Exception:
        try:
            import subprocess
            import sys

            if sys.platform == "darwin":
                subprocess.Popen(["open", str(image_path)])
            else:
                subprocess.Popen(["xdg-open", str(image_path)])
        except Exception:
            pass

    text = input(f"CAPTCHA manual (ver: {image_path}): ").strip()
    return normalize_captcha_text(text)


# ---------------------------------------------------------------------------
# Cadena de proveedores
# ---------------------------------------------------------------------------

def _get_chain():
    raw = (os.getenv("CAPTCHA_PROVIDERS", "").strip() or DEFAULT_PROVIDER_CHAIN)
    chain = []
    for item in raw.split(","):
        item = item.strip()
        if not item:
            continue
        if ":" in item:
            kind, model = item.split(":", 1)
        else:
            kind, model = item, ""
        chain.append((kind.strip().lower(), model.strip()))
    return chain


def solve_with_provider(kind: str, model: str, image_bytes: bytes,
                        manual_temp_dir: Optional[str] = None) -> str:
    """Resuelve con un proveedor concreto (sin cadena de respaldo)."""
    kind = (kind or "").strip().lower()
    model = (model or "").strip()
    if kind == "opencode-go":
        return solve_captcha_with_opencode_go(image_bytes, model or None)
    if kind in ("opencode", "opencode-zen", "zen"):
        return solve_captcha_with_opencode_zen(image_bytes, model or None)
    if kind == "openai":
        return solve_captcha_with_openai(image_bytes, model=model or None)
    if kind == "gemini":
        return solve_captcha_with_gemini(image_bytes, model=model or None)
    if kind == "manual":
        return solve_captcha_manual(image_bytes, manual_temp_dir)
    raise RuntimeError(f"Proveedor no soportado: {kind}")


def solve_captcha_auto(image_bytes: bytes, manual_temp_dir: Optional[str] = None,
                       allow_manual: bool = False, rotate: int = 0) -> Tuple[str, str]:
    """Recorre la cadena de proveedores y devuelve (texto, proveedor_usado).

    `rotate` permite empezar la cadena por otro modelo: en los reintentos conviene
    preguntarle primero a un modelo distinto, porque los modelos suelen fallar en
    caracteres diferentes y así aumenta la probabilidad de acertar.

    Lanza la última excepción si ninguno funcionó.
    """
    errors = []
    chain = _get_chain()
    if chain and rotate:
        n = rotate % len(chain)
        chain = chain[n:] + chain[:n]

    for kind, model in chain:
        try:
            if kind == "opencode-go":
                text = solve_captcha_with_opencode_go(image_bytes, model or None)
            elif kind in ("opencode", "opencode-zen", "zen"):
                text = solve_captcha_with_opencode_zen(image_bytes, model or None)
            elif kind == "openai":
                key = os.getenv("OPENAI_API_KEY", "").strip()
                if not key:
                    raise RuntimeError("OPENAI_API_KEY vacío")
                text = solve_captcha_with_openai(image_bytes, model=model or None)
            elif kind == "gemini":
                key = os.getenv("GEMINI_API_KEY", "").strip()
                if not key or "tu_api_key" in key.lower():
                    raise RuntimeError("GEMINI_API_KEY vacía")
                text = solve_captcha_with_gemini(image_bytes, model=model or None)
            elif kind == "manual":
                text = solve_captcha_manual(image_bytes, manual_temp_dir)
            else:
                raise RuntimeError(f"Proveedor no soportado: {kind}")

            if valid_captcha_answer(text):
                logging.info(f"CAPTCHA resuelto por {kind}:{model or '(default)'} -> {text}")
                return text, f"{kind}:{model or 'default'}"

            errors.append(f"{kind}:{model or 'default'} -> respuesta inválida {text!r}")
            logging.warning("Proveedor %s:%s devolvió una respuesta no válida: %r", kind, model or "default", text)
        except Exception as exc:
            msg = str(exc)
            errors.append(f"{kind}:{model or 'default'} -> {type(exc).__name__}: {msg[:180]}")
            logging.warning(
                "Proveedor %s:%s falló (%s): %s",
                kind, model or "default", type(exc).__name__, msg[:180],
            )
            if "429" in msg or "Too Many" in msg or "too many" in msg or "quota" in msg.lower():
                time.sleep(min(8, 2 + len(errors)))
            continue

    if allow_manual:
        text = solve_captcha_manual(image_bytes, manual_temp_dir)
        if valid_captcha_answer(text):
            return text, "manual"

    logging.warning("Ningún proveedor de IA pudo resolver el CAPTCHA:\n  - " + "\n  - ".join(errors))
    raise RuntimeError("; ".join(errors) if errors else "Sin proveedores configurados")


# ---------------------------------------------------------------------------
# API antigua (compatibilidad con versiones anteriores)
# ---------------------------------------------------------------------------

def solve_captcha(image_bytes: bytes, mode: str = "ia",
                  manual_temp_dir: Optional[str] = None) -> str:
    selected = (mode or "ia").strip().lower()
    if selected == "manual":
        return solve_captcha_manual(image_bytes, manual_temp_dir=manual_temp_dir)

    if selected in ("ia", "auto"):
        try:
            text, _ = solve_captcha_auto(image_bytes, manual_temp_dir=manual_temp_dir,
                                         allow_manual=(selected == "auto"))
            return text
        except Exception:
            if selected == "auto":
                return solve_captcha_manual(image_bytes, manual_temp_dir=manual_temp_dir)
            raise

    raise Exception(f"Modo no soportado: {selected}")


def capture_captcha_image(captcha_element) -> bytes:
    """Captura la imagen del CAPTCHA desde el elemento web (respaldo)."""
    return captcha_element.screenshot_as_png

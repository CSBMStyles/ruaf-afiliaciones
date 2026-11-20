# Consulta RUAF - Afiliaciones (SISPRO)

Automatiza la consulta de **afiliaciones de una persona en el Sistema** en el
Registro Único de Afiliados (RUAF) del Ministerio de Salud:

`https://ruaf.sispro.gov.co/Filtro.aspx`

La aplicación llena el formulario, resuelve el **captcha con IA** (y manual si
falla), consulta, **imprime el reporte como PDF** y genera un **Excel de datos**
y un **informe PPTX** con el análisis.

---

## Qué hace

- Lee los empleados desde `RUAF/CARACTERIZACION I.xlsx`, hoja `CONVENIO EJECUCION`.
- Respeta **exactamente el orden de las filas del Excel**.
- Permite elegir hoja, fila de encabezados y qué columna tiene cada dato:
  documento, fecha de expedición, nombres, apellidos, perfil y tipo de documento.
- Llena el formulario del RUAF:
  - Tipo de documento (`CC`, `CE`, `PA`, etc.).
  - Número de documento.
  - Fecha de expedición (`DD/MM/AAAA`).
  - Acepta los términos de uso.
- Obtiene el captcha de imagen y lo resuelve con **IA** (Gemini/OpenAI).
- Si la IA falla o el sitio rechaza el código, muestra la imagen **3× más grande**
  en un diálogo para que escribas el captcha.
- Pulsa **Verificar** y, si el sitio lo rechaza, pide un captcha nuevo y reintenta.
- Cuando el captcha es válido, pulsa **Consultar** y espera el reporte.
- Genera el PDF con el **panel de impresión** del navegador (A4 horizontal),
  sin usar el botón Exportar del visor.
- Guarda el PDF como `APELLIDO NOMBRE - CEDULA.pdf` en la carpeta del perfil.
- Guarda el avance en `estado_ruaf.json` para reanudar.
- Genera `REPORTE_FALLOS_<fecha>.xlsx`, `.csv` y `.txt` si hay fallos.
- Genera `DATOS_RUAF_<fecha>.xlsx` con una fila por empleado y una hoja `Detalle`.
- Genera `INFORME_RUAF_<fecha>.pptx` con portada, resumen, cobertura, estados,
  perfiles, principales EPS/AFP/ARL, novedades y conclusiones.

---

## Estructura de salida

```
RUAF/
├── AFILIACIONES/
│   ├── AUXILIAR/
│   │   ├── PEREZ GOMEZ ANA MARIA - 1000000001.pdf
│   │   └── ...
│   ├── ENFERMERO/
│   └── MEDICO/
├── DATOS_RUAF_20261103_153000.xlsx
└── INFORME_RUAF_20261103_153000.pptx
```

El PDF se genera con el **panel de impresión** del navegador (A4 horizontal),
no con el botón Exportar. Sale limpio, sin la página de “sesión ASP.NET
expirada”, y normalmente tiene **2 páginas**.

---

## El CAPTCHA, paso a paso

1. La aplicación abre Chrome y entra al formulario (acepta los términos si es
   una sesión nueva).
2. Llena tipo de documento, número y fecha de expedición.
3. Obtiene la imagen del captcha y la envía a la IA:
   - Primero `gemini-2.5-flash-lite`.
   - Luego `gemini-2.5-flash`.
   - Después OpenAI (`gpt-4o-mini`), si hay créditos.
4. Si la IA acierta, llena el campo y pulsa **Verificar**.
5. Si el sitio responde **“Texto Válido”**, pulsa **Consultar**.
6. Si responde **“Texto Inválido”**, pide un captcha nuevo y vuelve a intentar
   con IA.
7. Después de los intentos configurados con IA, muestra la imagen ampliada para
   que escribas el código manualmente.
8. Cuando el reporte carga, lo imprime como PDF y sigue con la siguiente persona.

Configura `GEMINI_API_KEY` en `.env`. Si no hay clave, la IA falla o se agota la
cuota, el modo manual sigue disponible. No necesitas pagar 2captcha.

> La sesión del RUAF dura **5 minutos de inactividad**. Si se vence, la
> aplicación vuelve a aceptar los términos automáticamente.

## Instalación

### macOS

```bash
cd Codigo/ANTECEDENTES/obtener-ruaf-sispro
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Abre la interfaz con doble clic en **`iniciar.command`** o:

```bash
python app.py
```

Si macOS bloquea el chromedriver:

```bash
xattr -d com.apple.quarantine "$HOME/Library/Application Support/undetected_chromedriver/undetected_chromedriver"
```

### Windows

1. Instala **Python 3.11 o superior** y marca **“Add python.exe to PATH”**.
2. Instala **Google Chrome**.
3. En PowerShell:

```powershell
cd Codigo\ANTECEDENTES\obtener-ruaf-sispro
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Si PowerShell bloquea la activación:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

4. Abre la interfaz con doble clic en **`iniciar.bat`** o:

```powershell
python app.py
```

---

## La interfaz

### 1 · Datos

- Excel de caracterización y hoja.
- Fila de encabezados.
- Mapeo de columnas: documento, fecha de expedición, nombres, apellidos, perfil
  y tipo de documento (opcional).
- Vista previa con las columnas mapeadas primero.
- **Validar datos**: cuenta válidos, incompletos y duplicados; avisa de fechas
  futuras.

### 2 · Salida y navegador

- Carpeta de PDFs (`RUAF/AFILIACIONES` por defecto).
- Carpeta de informes (`RUAF` por defecto).
- Nombre del PDF:
  - `APELLIDO NOMBRE - CEDULA` (recomendado).
  - `CEDULA NOMBRE APELLIDO`.
- Driver: `selenium` (recomendado), `auto` o `uc`.
- Espera máxima para descargar el PDF.

### 3 · CAPTCHA (IA + manual)

- Instrucciones y tiempo de espera.
- Modo: **IA + manual** (recomendado) o **Solo manual**.
- Intentos con IA antes de pedir el captcha manual.
- **Probar CAPTCHA**: abre Chrome, intenta leer el captcha con IA y, si falla,
  muestra el diálogo. No consulta ninguna cédula.

### 4 · Ejecución

- Perfil (`TODOS`, `AUXILIAR`, `ENFERMERO`, `MEDICO`).
- Cantidad por tanda (`0` = todos).
- Reanudar o reiniciar el estado.
- Intentos totales por persona si el captcha es rechazado.
- Barra de progreso y botón **Detener**.

### 5 · Informes

- **Generar Excel de datos**.
- **Generar informe PPTX**.
- **Generar ambos**.
- **Generar automáticamente al terminar la consulta** (activado por defecto).
- Abrir la carpeta de informes.

---

## Uso por línea de comandos

```bash
python main.py
python main.py --perfil AUXILIAR --tanda 10 --si
python main.py --perfil TODOS --si
python main.py --reiniciar-estado --si
python main.py --excel "/ruta/CARACTERIZACION I.xlsx" --salida "/ruta/AFILIACIONES"
python main.py --dry-run --si
python main.py --sin-informes --si
```

`--dry-run` valida el Excel sin abrir Chrome.

Para generar solo los informes a partir de PDFs existentes:

```bash
python extraer_datos.py --carpeta "../RUAF/AFILIACIONES" --pptx
python generar_informes.py --carpeta "../RUAF/AFILIACIONES" --salida "../RUAF"
```

---

## El Excel de datos

### Hoja `Empleados`

Una fila por empleado:

- **Identificación:** fila del Excel, documento, tipo, apellidos, nombres,
  perfil y fecha de expedición.
- **Reporte:** fecha de corte, fecha de consulta, sexo, estado general y PDF.
- **Salud:** administradora (EPS), régimen, fecha de afiliación, estado, tipo de
  afiliado y departamento/municipio.
- **Pensión:** régimen, administradora (AFP), fecha de afiliación y estado.
- **ARL:** administradora, fecha de afiliación, estado, actividad económica y
  municipio donde labora.
- **Otras secciones:** compensación familiar, cesantías, pensionados y
  programas de asistencia social.
- **Observaciones:** novedades de salud, pensión o ARL.

### Hoja `Detalle`

Una fila por sección y afiliación, con todos los registros que traiga el
reporte. Si una persona tiene varias filas en una sección, aquí aparecen todas.

### Hoja `Resumen`

Conteos de cobertura, estados y secciones reportadas.

---

## El informe PPTX

El nombre de la empresa se puede cambiar con `EMPRESA_INFORME` en `.env`.


Incluye:

1. Portada.
2. Resumen ejecutivo con indicadores.
3. Cobertura por sección.
4. Estados de afiliación (activo, inactivo, no reportado).
5. Distribución por perfil.
6. Principales EPS.
7. Principales AFP.
8. Principales ARL.
9. Empleados con novedades críticas.
10. Conclusiones y recomendaciones.

El detalle completo queda en el Excel; el PPTX es el resumen para presentar.

---

## Reporte de fallos

El `.xlsx` tiene dos hojas:

- **Fallos:** una fila por persona con `Fila Excel`, `Documento`, `Nombre`,
  `Perfil`, `Motivo`, `Qué pasó`, `Qué hacer`, `Mensaje del sitio`, `Intentos`,
  `Método` y `Captura`.
- **Resumen:** cuántos fallos hay por motivo y qué significa cada uno.

El `.txt` es la versión corta y el `.csv` sirve para filtrar en Excel.

---

## Problemas frecuentes

### El captcha no aparece o sale en blanco

La etiqueta `<img>` del sitio a veces responde 404 en automatización. La
aplicación no usa esa imagen directamente: la obtiene con `fetch` desde
JavaScript y la muestra en el diálogo. Si aun así falla, la app recarga el
formulario y reintenta.

### El captcha se rechaza

- La app reintenta con un captcha nuevo; la IA suele acertar en el siguiente.
- Si Gemini devuelve 503 o se agota la cuota, la app pasa a Flash, a OpenAI o al
  modo manual.
- Revisa que `GEMINI_API_KEY` esté configurada en `.env`.
- Aumenta los intentos con IA en la pestaña 3.
- Si el problema persiste, cambia de red o de navegador.

### La sesión se vence

El RUAF cierra la sesión tras **5 minutos de inactividad**. La app vuelve a
aceptar los términos y continúa; si una persona queda a medias, aparece en el
reporte y se puede reintentar.

### El PDF no se guarda

- El PDF se genera con el panel de impresión, no con el botón Exportar.
- Verifica que la carpeta de PDFs tenga permisos de escritura.
- Si el antivirus bloquea la escritura, usa otra carpeta local.
- Revisa la captura de error en `errors/` y el reporte de fallos.

### “La fecha de expedición no corresponde”

Revisa la columna `Expedicion` del Excel. Debe ser `DD/MM/AAAA` y coincidir con
la cédula. La validación avisa si hay fechas futuras.

### “This version of ChromeDriver only supports Chrome version …”

`undetected-chromedriver` puede tener un driver viejo. El programa cae
automáticamente a **Selenium estándar**, que descarga el driver correcto. Si
quieres forzarlo:

```bash
DRIVER_BACKEND=selenium
```

o elige `selenium` en la pestaña 2.

### Windows

- Si el antivirus bloquea `undetected_chromedriver.exe`, usa `selenium`.
- Si el proyecto está en OneDrive, mueve la carpeta de salida a una ruta local.
- Los nombres se recortan a 120 caracteres para no superar el límite de Windows.

---

## Archivos

- `app.py`: lanzador de la interfaz.
- `gui.py`: ventana, pestañas y diálogo del captcha.
- `main.py`: flujo de consulta, captcha, PDF, estado y reportes.
- `captcha_solver.py`: resuelve el captcha con Gemini/OpenAI y respaldo manual.
- `config.py`: `config.json` y `.env`.
- `sources.py`: lectura del Excel, hojas, mapeo y orden.
- `utils.py`: fechas `DD/MM/AAAA`, tipos de documento y nombres seguros.
- `perfiles.py`: campos del proyecto.
- `extraer_datos.py`: lee los PDFs y genera el Excel de datos.
- `generar_pptx.py`: genera el informe PPTX.
- `generar_informes.py`: orquesta Excel + PPTX.
- `estado_ruaf.json`: control de avance.
- `errors/`: capturas cuando falla una consulta.
- `REPORTE_FALLOS_*.xlsx` / `.csv` / `.txt`: informe de fallos.
- `DATOS_RUAF_*.xlsx`: datos consolidados.
- `INFORME_RUAF_*.pptx`: informe ejecutivo.

---

## Uso autorizado

Este proyecto está pensado para que una entidad autorizada consulte las
afiliaciones de sus colaboradores. Usa la información solo para los fines
institucionales previstos y trata los datos personales conforme a la ley.

---

## Datos y privacidad

Este repositorio contiene únicamente código y documentación. **No incluye**:

- API keys ni archivos `.env`.
- Bases de datos de empleados, cédulas, nombres, correos o teléfonos.
- PDFs, Excel, PPTX, CSV, capturas o reportes generados por las consultas.
- Configuraciones locales (`config.json`) ni estados de proceso.

Para usarlo, proporciona tu propia base de datos autorizada y configura tus
propias credenciales. No subas datos personales al repositorio.

El archivo `config.example.json` muestra una configuración mínima; cópialo a
`config.json` o deja que la interfaz lo cree.

## Licencia

MIT. Consulta el archivo `LICENSE`.


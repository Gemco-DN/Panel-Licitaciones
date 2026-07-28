# Panel de Licitaciones GEMCO

Panel visual tipo TV corporativa que muestra las licitaciones próximas a cerrar (ventana de 20 días), con modo presentador rotativo. Se actualiza automáticamente todos los días vía GitHub Actions: el workflow lee el Excel de SharePoint por Microsoft Graph API, genera un HTML, y lo publica en GitHub Pages. También se puede correr manualmente si hace falta.

## Arquitectura general

```
Excel (SharePoint)
        ↓  [Microsoft Graph API, MSAL Client Credentials Flow]
   python actualizar_panel_v2.py (corre en GitHub Actions, cron diario 11:00 UTC,
   o localmente / vía .bat si se necesita forzar una actualización)
        ↓
   Lee PPublicas + Checklist, cruza por ID Licitación
        ↓
   Genera index.html (con datos embebidos + logos en base64)
        ↓  [git add / commit / push, automático dentro del script]
   GitHub Pages publica la URL pública
        ↓
   TV con navegador en modo kiosko, apuntando a esa URL
```

**Punto clave:** el proceso SÍ está 100% automatizado. GitHub Actions corre el script a diario sin intervención humana, usando un App Registration de Azure AD (secrets `SP_TENANT_ID`, `SP_CLIENT_ID`, `SP_CLIENT_SECRET`, `SP_FILE_URL` configurados en el repo) para leer el Excel de SharePoint vía Graph API. Correr el script manualmente (`.bat` o `python actualizar_panel_v2.py`) sigue siendo posible para forzar una actualización fuera de horario, pero ya no es necesario para el funcionamiento normal — ver sección "Automatización" (antes "Por qué no está 100% automatizado").

## Estructura de archivos del repo

```
panel-licitaciones/
├── actualizar_panel_v2.py     ← script principal (usar este, no v1)
├── index.html                  ← generado automáticamente, NO editar a mano
├── README.md
└── PNG Empresas/                ← logos fuente, NO se suben a GitHub
    ├── Gemco-logo-blanco-v2.png       (usado, incrustado en base64)
    ├── Isotipos_-_copia.png           (usado, incrustado en base64)
    ├── INCARDIA-logo-blanco-v2.png    (sin uso todavía)
    └── MMQ-logo-blanco-v2.png         (sin uso todavía)
```

URL pública del panel: `https://mguajardoe-cyber.github.io/panel-licitaciones/`

## Origen y estructura de los datos

El Excel fuente vive en SharePoint. El script ya no depende de que esté sincronizado localmente en ningún PC: lo descarga directo desde la nube vía Microsoft Graph API (`descargar_excel_desde_sharepoint()` en `actualizar_panel_v2.py`), usando el link de "Compartir" del archivo (`SP_FILE_URL`) y un App Registration de Azure AD (Client Credentials Flow con `msal`).

Otra persona (no Martín) mantiene y actualiza este archivo a diario. **Riesgo conocido:** si el script corre mientras esa persona tiene el Excel abierto en edición, puede levantar datos parciales o desactualizados. No hay mitigación automática — se maneja coordinando horarios (el cron corre a las 07:00-08:00 Chile, antes del horario laboral típico).

### Hoja `PPublicas` (fuente principal)
Encabezados en fila 2, datos desde fila 3. Columnas relevantes usadas:
- `ID`, `ESTADO`, `LINEA DE NEGOCIO`, `VENDEDOR`, `CLIENTE LICITACIÓN`, `EQUIPAMIENTO`, `FECHA DE CIERRE`

`ESTADO` tiene 4 valores posibles: `Confirmada`, `Sin confirmar`, `Descartada`, `Revocada`. El panel solo muestra las dos primeras.

### Hoja `Checklist` (datos de garantía, cruzados por ID)
Encabezados en fila 2. Columna de cruce: `ID Licitación` (columna B).
Columnas usadas: `Requiere garantía de seriedad (SI/NO)`, `Tipo de garantía`, `Monto garantía`, `Moneda garantía`, `Banco/empresa emisora`, `Fecha inicio garantía`, `Fecha fin garantía`.

**Cuidado:** la columna "Requiere garantía" tiene inconsistencias de formato en el Excel real (mezcla `SI`/`si`/`NO`/`N/A`/vacío). El script normaliza a mayúsculas antes de comparar — cualquier cambio futuro a esta lógica debe mantener esa normalización.

El cruce entre hojas replica un BUSCARV de Excel: se hace con `pandas.merge()` por `ID_norm` (ID limpiado de espacios). Si hay IDs duplicados en `Checklist`, se usa la primera coincidencia (`drop_duplicates(keep="first")`).

## Reglas de negocio ya decididas

- **Ventana temporal:** próximos 20 días desde la fecha de ejecución del script.
- **Estados incluidos:** Confirmada y Sin confirmar (mostradas en secciones separadas, "Sin confirmar" se presenta como "Por confirmar" en la UI).
- **Sección de garantía:** solo se muestra (como fila extra debajo de la licitación) si `Requiere garantía` = SI. Si es NO, N/A, o vacío, no se muestra nada extra.
- **Urgencia visual:** barra de color por fila — rojo si cierra en ≤2 días, naranja si 3-5 días, verde/gris si más de 5 días.
- **Modo presentador:** 8 filas por pantalla, 12 segundos por pantalla, transición con fade + desplazamiento.
- **Slide de marca:** al final de cada ciclo completo (después del último slide de "Por confirmar"), aparece el isotipo de GEMCO a pantalla completa por 5 segundos, luego reinicia el ciclo.
- **Interactividad:** clic en cualquier parte de la pantalla, o barra espaciadora, pausa/reanuda la rotación automática. Hay un indicador visual ("⏸ Panel en pausa...") cuando está pausado.
- **Auto-refresh:** la página se recarga sola cada 30 minutos (para levantar una actualización nueva sin que alguien toque la TV), salvo que esté pausada.

## Decisiones de diseño y por qué

- **Paleta clara (no oscura):** inspirada en una referencia de dashboard SaaS que el usuario compartió. El header sí usa una franja oscura con degradado — es la excepción, necesaria porque los logos GEMCO son versión "blanco" (pensados para fondo oscuro) y sobre fondo claro quedarían invisibles.
- **Imágenes incrustadas en base64 dentro del HTML** (no como archivos `<img src="ruta">` separados): evita problemas de rutas relativas al servir en GitHub Pages. La carpeta `PNG Empresas` SÍ está subida al repo (no está en `.gitignore`) porque el runner de GitHub Actions parte de un checkout limpio y necesita los PNG ahí para poder incrustarlos — si no estuvieran versionados, el workflow en la nube no tendría de dónde leerlos.
- **GitHub Pages en vez de Netlify:** Netlify pasó a un modelo de créditos (300/mes, ~15 créditos por deploy) que se agotó rápido con actualizaciones diarias en un proyecto anterior del usuario. GitHub Pages no tiene ese modelo de facturación — los límites son "blandos" (soft limits: 100GB bandwidth/mes, 10 builds/hora) y no aplican de forma realista a este caso de uso interno de bajo tráfico.

## Automatización (Azure AD + GitHub Actions)

El permiso de Azure AD / Graph API ya se consiguió y está implementado: hay un App Registration en Entra ID (Client Credentials Flow) cuyas credenciales viven como secrets del repo (`SP_TENANT_ID`, `SP_CLIENT_ID`, `SP_CLIENT_SECRET`, `SP_FILE_URL`), consumidos por `.github/workflows/actualizar-panel.yml`. El workflow corre a diario (`cron: 0 11 * * *`, ~07:00-08:00 Chile) y también se puede disparar a mano desde GitHub (`workflow_dispatch`). Nada de esto depende del PC de Martín estando prendido.

**Verificación pendiente (no confirmable desde acá):** confirmar en GitHub → Settings → Secrets del repo `Gemco-DN/Panel-Licitaciones` que los 4 secrets estén cargados y vigentes (el Client Secret de Azure AD tiene fecha de expiración — revisar cuándo vence).

## Pendientes / próximos pasos

- [ ] Confirmar visualmente en la TV real que el header oscuro + logo se ven bien a la distancia de visualización real.
- [ ] Configurar el navegador de la PC de la TV en modo kiosko apuntando a la URL de GitHub Pages.
- [ ] Evaluar si el número de deploys diarios (un push por día) se mantiene lejos de cualquier límite práctico de GitHub Pages (hoy no hay indicio de que vaya a ser un problema).
- [ ] Revisar fecha de expiración del Client Secret de Azure AD y calendarizar su renovación antes de que venza (si vence, el workflow diario empieza a fallar en silencio hasta que alguien lo note).
- [ ] Los logos de Incardia y MMQ están disponibles pero sin uso — evaluar si en algún momento se quiere mostrar más de una marca en el panel (decisión explícita: por ahora NO, solo GEMCO).
- [ ] `actualizar_panel.py` (v1) sigue en la carpeta sin trackear en git — evaluar si conviene borrarlo para no confundir con v2, que es la versión vigente.

## Comandos útiles

```bash
# Instalar dependencias (una sola vez)
pip install pandas openpyxl

# Correr la actualización manual
python actualizar_panel_v2.py

# Ver el estado de git del repo
git status

# Ver historial de actualizaciones del panel
git log --oneline
```

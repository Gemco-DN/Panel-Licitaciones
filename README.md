# Panel de Licitaciones GEMCO

Panel TV de licitaciones próximas a cerrar, con datos leídos desde SharePoint (Excel) vía Microsoft Graph API y publicado en GitHub Pages. Se actualiza solo, todos los días, vía GitHub Actions.

URL pública: https://mguajardoe-cyber.github.io/panel-licitaciones/

Detalle completo de arquitectura, datos y decisiones de diseño en [CONTEXTO_PROYECTO.md](CONTEXTO_PROYECTO.md).

## Uso rápido

```bash
pip install -r requirements.txt

# Variables de entorno requeridas (ver CONTEXTO_PROYECTO.md):
# SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET, SP_FILE_URL

python actualizar_panel_v2.py
```

En Windows, también se puede correr con doble clic a `Actualizar_Panel.bat` (requiere `python` en el PATH y las variables de entorno configuradas).
# Preparación para publicación pública en GitHub

La publicación no se ha realizado. Antes de crear el repositorio público:

1. Usar `ConstructionAssistant/opencode.public.example.json` como configuración genérica; no incluir `ConstructionAssistant/opencode.json`, que contiene una ruta de usuario. En una instalación local, copiar la plantilla a `opencode.json` y ajustar el directorio si hace falta; ese archivo quedó agregado a `.gitignore`.
2. Mantener fuera del repositorio `.env`, claves, configuraciones locales, rutas personales, archivos temporales, `node_modules/` y `dist/` (estos dos últimos ya están excluidos en `.gitignore`).
3. Revisar el contenido de los libros Excel, capturas y video. Excluirlos de la publicación si contienen datos personales, rutas, información de una obra real o resultados no destinados a difusión. No se ha hecho aquí una revisión visual completa del video.
4. Revisar que `README.md`, `PROMPT.md`, declaración de IA, licencia y ejemplos describan la versión que se publicará; ejecutar `npm ci` y `npm test` desde Windows con Excel instalado para revalidar el paquete final.
5. Inspeccionar la lista exacta de archivos del commit y el diff antes de crear o publicar un repositorio.

La configuración de ejemplo es intencionalmente mínima y requiere que OpenCode ejecute el MCP desde la carpeta `ConstructionAssistant`, después de compilar con `npm run build`.

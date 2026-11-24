# Investigación de integración con Microsoft Excel

## Hallazgos

- Sistema detectado: Windows 10/11 (PowerShell 5.1 disponible).
- Excel aparece registrado en `Get-StartApps` como `Microsoft.Office.EXCEL.EXE.15`.
- El ProgID COM `Excel.Application` está registrado y apunta al CLSID `{00024500-0000-0000-C000-000000000046}`.
- La ruta registrada para automatización es `C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE /automation`.
- Node.js 26 está instalado; el proyecto ya contiene el SDK MCP oficial y Zod.

## Vía elegida y motivo

Se utiliza la automatización COM de Excel a través de PowerShell, invocado desde Node con `execFile` (mediante `src/lib/exec.ts`). Excel ofrece un modelo de objetos COM oficial y permite abrir libros existentes, leer hojas y guardar resultados directamente en el libro indicado por el usuario. No se construyen archivos Excel sustitutos ni se automatiza la interfaz gráfica.

Los valores proporcionados al modelo se envían al proceso como variables de entorno JSON; el script PowerShell es constante y no concatena entradas del usuario. Se valida en Node la ruta, existencia y extensión del libro. Para registrar cálculos se exige `confirmar: true` y se anexan filas a las hojas administradas `Cubicaciones`, `Presupuesto` o `Avance`, conservando las demás hojas. La lectura de hojas y la auditoría son de solo lectura. `construction_crear_libro_desde_plantilla` abre una plantilla `.xlsx` en modo de solo lectura mediante COM y guarda una copia nueva; no reemplaza un destino existente.

## Alternativas revisadas

1. **API HTTP/WebSocket o plugin:** no se detectó una API local de Excel disponible para este caso. Microsoft Graph es una alternativa en la nube, pero requiere autenticación y no es una conexión directa al Excel local solicitado.
2. **CLI:** no se detectó una interfaz de línea de comandos de Excel que permita operar libros con las funciones necesarias.
3. **Edición de archivos de datos:** se descarta como vía principal porque el requerimiento pide operar sobre libros reales mediante Excel y no reemplazar Excel con la generación de archivos.
4. **COM (elegida):** está instalado y registrado en este computador y proporciona acceso al modelo de objetos de Excel.
5. **Automatización de interfaz:** descartada por ser más frágil y menos verificable que COM.

## Limitaciones conocidas

- Requiere Windows, Microsoft Excel instalado y permisos para iniciar automatización COM.
- El libro debe existir y ser accesible, estar en formato `.xlsx`, `.xlsm` o `.xls` y no estar bloqueado de forma que impida la edición.
- El MCP crea una instancia COM independiente, abre el libro solicitado y la cierra al terminar. No intenta apropiarse de una instancia interactiva ya abierta.
- La automatización real del libro depende del Excel del computador; las pruebas automatizadas cubren cálculos, validaciones y el protocolo, no guardan cambios en libros del usuario.

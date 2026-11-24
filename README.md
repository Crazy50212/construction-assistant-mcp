# Construction Assistant MCP

Servidor MCP en TypeScript para cálculos y control de Ingeniería en Obras Civiles mediante Microsoft Excel de escritorio en Windows. Usa el SDK oficial `@modelcontextprotocol/sdk`, validación Zod y automatización real de Excel COM mediante PowerShell.

## Funciones

El libro de trabajo conserva cuatro hojas: **Cubicaciones**, **Presupuesto**, **Avance** y **Resumen**. Las tablas `tblCubicaciones`, `tblPresupuesto` y `tblAvance` mantienen fórmulas y referencias estructuradas; los vínculos se actualizan al agregar, editar o eliminar registros en las tablas.

| Herramienta MCP | Uso | Acceso al libro |
|---|---|---|
| `excel_listar_hojas` | Inspecciona las hojas de un libro existente. | Solo lectura |
| `construction_auditar_integridad` | Audita vínculos, unidades, datos inválidos, fórmulas, sobre-ejecuciones, partidas sin avance y gráficos. | Solo lectura |
| `construction_crear_libro_desde_plantilla` | Crea una copia `.xlsx` desde una plantilla `.xlsx`; falla si el destino ya existe. | Crea un archivo nuevo mediante Excel COM |
| `construction_cubicaciones` | Calcula superficie, volumen geométrico y volumen ajustado, y registra la cubicación. | Agrega datos y fórmulas |
| `construction_presupuesto` | Agrega partidas, calcula costos, subtotal, IVA y total. | Modifica Presupuesto y actualiza Resumen/gráficos |
| `construction_control_avance` | Registra cantidades ejecutadas, avance, saldo, costo y exceso. | Modifica Avance y actualiza vínculos/Resumen |

Las tres herramientas que guardan filas requieren `confirmar: true`. El auditor es de solo lectura. La copia desde plantilla nunca sobrescribe un destino existente; actualmente acepta plantillas `.xlsx` para preservar sus elementos sin conversión silenciosa.

## Fórmulas, vínculos y clasificación

- **Cubicaciones:** superficie = largo × ancho; volumen = superficie × profundidad; volumen ajustado = volumen × factor. Se aceptan dimensiones grandes dentro del rango numérico de Excel. Las entradas vacías, no numéricas, negativas o con productos fuera del rango producen advertencias y fórmulas vacías, no errores de hoja.
- **Presupuesto:** costo de partida = cantidad × precio unitario; subtotal = suma de costos; IVA usa la tasa configurada (19% por defecto); total = subtotal + IVA.
- **Avance:** avance físico = ejecutado / presupuestado; saldo = `MAX(presupuestado − ejecutado, 0)`; costo ejecutado = ejecutado × precio unitario; el exceso se registra por separado.
- Cantidad y precio se vinculan por nombre y unidad compatible. Se admiten variantes usuales de m²/m2 y m³/m3. Los duplicados compatibles agregan cantidades y calculan un precio unitario ponderado por costo/cantidad.
- `Resumen` separa unidades y no mezcla cantidades de m³, kg, unidades u otras magnitudes. Contiene indicadores de datos por revisar, sobre-ejecución, avances sin vínculo presupuestario y partidas de presupuesto sin avance por nombre.
- El auditor distingue `VINCULADO`, `SIN VÍNCULO`, `UNIDAD INCOMPATIBLE`, `REGISTRO MANUAL`, `SIN AVANCE`, `DATOS INVÁLIDOS`, `FORMULA SOBRESCRITA` y `SOBREEJECUCIÓN`. Las clasificaciones pueden coexistir, por ejemplo una partida vinculada y sobre-ejecutada.
- `Resumen` conserva un gráfico lineal y una alternativa logarítmica, ambos referidos a las mismas partidas y costos de Presupuesto.

## Arquitectura

- `src/index.ts`: servidor MCP por `stdio`; carga automáticamente las herramientas de `src/tools/`.
- `src/tools/construction-assistant.ts`: esquemas Zod y registro de las seis herramientas.
- `src/lib/calculos.ts`: cálculos independientes de Excel.
- `src/lib/excel-com.ts`: PowerShell/Excel COM, tablas, validaciones, fórmulas, auditoría y gráficos. Los datos viajan como JSON por variables de entorno; el script PowerShell se ejecuta desde un archivo temporal UTF-8.
- `src/lib/excel-plan.ts`, `excel-sync-templates.ts`, `excel-formatos.ts` y `excel-hojas.ts`: planes de hoja, vínculos, formatos y detección de hojas predeterminadas.
- `tests/`: pruebas Vitest.
- `scripts/smoke-test.mjs`: handshake, descubrimiento y descripciones MCP.

## Requisitos

- Windows y Microsoft Excel de escritorio con `Excel.Application` COM disponible.
- Node.js 18 o posterior y npm.
- Permiso del usuario para iniciar Excel mediante COM y escribir en el libro de destino.
- Una plantilla/libro accesible en formato `.xlsx`, `.xlsm` o `.xls` para las herramientas existentes. La herramienta de copia desde plantilla requiere `.xlsx`.

La automatización Excel COM no es compatible con macOS ni Linux. Las validaciones de celda pueden ser eludidas por importaciones o escrituras externas vía COM; las fórmulas protegidas, el Resumen y `construction_auditar_integridad` detectan los datos resultantes, pero no bloquean la aplicación externa que los escribe.

## Instalación y ejecución

Desde el directorio del proyecto:

```powershell
npm ci
npm run build
npm test
npm start
```

El transporte MCP es `stdio`. `opencode.json` registra `construction-assistant` con timeout de 90 segundos, acorde con las pruebas de tablas grandes. En el checkout original, `command` apunta a la ruta absoluta de `dist/index.js`; en `ENTREGA_FINAL/ConstructionAssistant`, la configuración usa `cwd` para la carpeta del código y un comando relativo. Si se mueve cualquiera de esas carpetas, actualiza la ruta o `cwd`. Después de editar `opencode.json`, cierra y reinicia opencode para que recargue la configuración.

Para desarrollo se puede usar `npm run dev`. Excel debe estar instalado, aunque no es necesario abrirlo manualmente antes de cada llamada MCP.

## Ejemplos de uso

1. **Auditoría:** “Audita `C:\Obras\Puente.xlsx` y dime qué vínculos no coinciden, qué partidas no tienen avance y si hay errores de fórmulas.”
2. **Nuevo libro:** “Crea `C:\Obras\Puente_Fase1.xlsx` desde `C:\Plantillas\Construction_Assistant.xlsx`; no reemplaces ningún archivo existente.”
3. **Cubicación y presupuesto:** “En `C:\Obras\Puente_Fase1.xlsx`, registra la excavación de 10 × 2 × 1,5 m con factor 1,2 y agrega su cantidad ajustada al presupuesto a $15.000 por m³.”
4. **Control de avance:** “Registra 190 m³ ejecutados de 180 presupuestados, muestra saldo y exceso por separado y actualiza el Resumen.”

La tasa de IVA se pasa como decimal (`0.19` = 19%). Las dimensiones se expresan en metros y los precios en la moneda del presupuesto (el formato de demostración es CLP).

## Pruebas y validación de Fase 4

`npm test` ejecuta build, pruebas Vitest y smoke test MCP. En el cierre final, las cinco suites y las 24 pruebas pasaron; también pasaron el handshake y el listado de herramientas. La validación COM de `Demo_Construction_Assistant_Phase4.xlsx` incluyó guardado/reapertura, vínculos, entradas inválidas, borrado/restauración, fórmulas sobrescritas, duplicados, gráficos y tablas de hasta 2.000 filas.

Tiempos medidos de anexado (Presupuesto / Avance):

| Filas en cada tabla | Presupuesto | Avance |
|---:|---:|---:|
| 100 | 4,4 s | 5,7 s |
| 250 | 8,4 s | 7,6 s |
| 500 | 9,5 s | 12,6 s |
| 1.000 | 14,9 s | 21,5 s |
| 1.500 | 23,2 s | 31,2 s |
| 2.000 | 30,6 s | 42,3 s |

Los tiempos dependen de Excel, CPU y del número de filas. En llamadas adicionales sobre tablas de 2.000 filas, el pico observado fue de aproximadamente 226–228 MB para Excel y 66 MB para Node; una operación de Avance llegó a 52,3 s. El timeout de PowerShell y el de OpenCode están configurados en 90 s. No se midieron tamaños mayores de 2.000 partidas.

En el libro Phase4 validado, el Resumen muestra 1.200.756 m³ de volumen acumulado, subtotal de $264.284.920, IVA de $50.214.134,80 y total de $314.499.054,80. El costo ejecutado es $396.172.755; no es comparable directamente con el volumen porque incorpora partidas de diversas unidades. Los indicadores de auditoría registraron 7 avances sin vínculo, 5 partidas presupuestadas sin avance por nombre, 7 sobre-ejecuciones y cero datos inválidos o problemas de fórmula.

## Archivos de demostración

- `Demo_Construction_Assistant_Integrated_Final.xlsx`: original conservado.
- `Demo_Construction_Assistant_Phase2.xlsx` y `Demo_Construction_Assistant_Phase3.xlsx`: versiones previas conservadas.
- `Demo_Construction_Assistant_Phase4.xlsx`: copia nueva generada desde la plantilla Phase3 mediante MCP; contiene las cuatro hojas, tablas, fórmulas, validaciones y los dos gráficos.

## Limitaciones conocidas

- La hoja `Estado de avance` conserva los estados numéricos y de sobre-ejecución; la clasificación del vínculo se informa en el resultado del auditor y sus contadores del Resumen.
- Los indicadores de partidas de Presupuesto sin avance se cuentan por nombre, y el auditor también evalúa nombre y unidad normalizada para identificar incompatibilidades.
- Los valores CLP se calculan con precisión decimal; el formato de celda los muestra sin fracciones de peso.
- No se midieron tablas mayores de 2.000 partidas; los tiempos pueden aproximarse o superar el timeout de 90 segundos al escalar más.

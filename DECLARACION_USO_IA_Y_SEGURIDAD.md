# Declaración de uso de IA y seguridad

## Uso de inteligencia artificial

Se utilizó un asistente de IA como apoyo durante el análisis, implementación, revisión y documentación del servidor MCP. La IA ayudó a explorar alternativas de integración, proponer y revisar código TypeScript/PowerShell, detectar inconsistencias entre la respuesta del MCP y los valores calculados por Excel, y preparar documentación. El estudiante conserva la responsabilidad de comprender, revisar y defender el diseño, los cálculos, las pruebas y las decisiones de integración presentadas. Las pruebas de ejecución con Excel se realizaron sobre copias de trabajo, no sobre libros fuente.

## Seguridad y manejo de datos

- La integración opera localmente en Windows mediante Excel COM; no requiere enviar el libro a un servicio en la nube.
- Las rutas y datos se validan antes de su uso. Los procesos externos se invocan con argumentos separados y los datos se transportan como JSON, no concatenados en comandos de shell.
- Las herramientas de escritura solicitan `confirmar: true`. La creación desde plantilla exige un destino nuevo y no sobrescribe un archivo ya existente.
- Las herramientas de lectura y auditoría no guardan cambios. La auditoría contrasta fórmulas, vínculos, unidades, errores y resultados.
- No se deben publicar libros, capturas ni grabaciones que contengan nombres, rutas, precios, información de clientes o datos de obras sin autorización. Revisar los artefactos de demostración antes de hacerlos públicos.
- No se encontraron credenciales incrustadas en el código durante la revisión realizada. La configuración local de OpenCode contiene una ruta de usuario y debe sustituirse por la plantilla genérica antes de publicar el código.

**Nota para la entrega:** ajustar esta declaración si el proceso real de elaboración difiere de lo indicado y confirmar que las políticas del curso aceptan el uso descrito.

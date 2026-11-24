# Demostración del MCP: relleno granular de subbase

## Preparación

La demostración se ejecutó mediante el cliente MCP y Excel COM de escritorio sobre una copia nueva de `Demo_Construction_Assistant_Phase4.xlsx`. El libro original de Fase 4 no se modificó. El resultado se entrega en `Demostracion_MCP_Caso_Obra_Civil.xlsx`.

## Secuencia y resultados

1. **Cubicación:** relleno de 8 m × 3 m × 0,9 m, con factor de ajuste 1,2. Superficie: **24 m²**; volumen geométrico: **21,6 m³**; volumen ajustado: **25,92 m³**.
2. **Presupuesto:** partida “Relleno granular subbase demostracion”, unidad m³ ajustado, cantidad 25,92 y precio unitario CLP 150.000. Costo: **CLP 3.888.000**; IVA 19%: **CLP 738.720**; total: **CLP 4.626.720**.
3. **Avance:** se registraron 30 m³ ajustados ejecutados. El avance físico fue **115,74%**, el saldo pendiente **0**, la sobre-ejecución **4,08 m³** y el costo ejecutado **CLP 4.500.000**. La clasificación fue `VINCULADO` y `SOBREEJECUCIÓN`.
4. **Vínculo presupuestario:** se pasó CLP 90.000 como precio de respaldo al registrar el avance. Como la partida y la unidad coincidían con Presupuesto, Excel usó el precio vinculado de **CLP 150.000**. La respuesta MCP se contrastó con esos valores reales de Excel.
5. **Auditoría:** no se detectaron problemas de fórmulas, inconsistencias de resumen, errores de celda ni referencias rotas en esta ejecución. Los gráficos de costos mantuvieron categorías y valores alineados con Presupuesto.

## Cómo revisar

Abrir `Demostracion_MCP_Caso_Obra_Civil.xlsx` en Excel y revisar las hojas **Cubicaciones**, **Presupuesto**, **Avance** y **Resumen**. La hoja Avance conserva las fórmulas y el vínculo al precio presupuestado; la alerta de sobre-ejecución queda visible en la fila agregada. El libro contiene los registros escritos por el MCP; esta ficha resume los resultados verificados en la ejecución.

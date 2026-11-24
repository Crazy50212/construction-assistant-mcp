import { describe, expect, it } from "vitest";
import { inferirFormatosColumnas } from "../src/lib/excel-formatos.js";

describe("inferencia de formatos de columnas Excel", () => {
  it("clasifica cantidades de cubicación sin moneda", () => {
    expect(inferirFormatosColumnas([
      "Concepto", "Tipo", "Largo (m)", "Ancho (m)", "Profundidad (m)",
      "Superficie (m²)", "Volumen (m³)", "Factor ajuste", "Volumen ajustado (m³)",
    ])).toEqual([
      { columna: 3, tipo: "cantidad" },
      { columna: 4, tipo: "cantidad" },
      { columna: 5, tipo: "cantidad" },
      { columna: 6, tipo: "cantidad" },
      { columna: 7, tipo: "cantidad" },
      { columna: 8, tipo: "cantidad" },
      { columna: 9, tipo: "cantidad" },
    ]);
  });

  it("mantiene cantidades separadas de CLP y trata avance físico como porcentaje", () => {
    expect(inferirFormatosColumnas([
      "Partida", "Unidad", "Cantidad presupuestada", "Cantidad ejecutada",
      "Avance físico (%)", "Cantidad pendiente", "Precio unitario", "Costo ejecutado",
    ])).toEqual([
      { columna: 3, tipo: "cantidad" },
      { columna: 4, tipo: "cantidad" },
      { columna: 5, tipo: "porcentaje" },
      { columna: 6, tipo: "cantidad" },
      { columna: 7, tipo: "clp" },
      { columna: 8, tipo: "clp" },
    ]);
  });

  it("identifica subtotal, IVA y totales como moneda incluso con tildes", () => {
    expect(inferirFormatosColumnas(["Subtotal", "IVA", "Total", "Costo total"])).toEqual([
      { columna: 1, tipo: "clp" },
      { columna: 2, tipo: "clp" },
      { columna: 3, tipo: "clp" },
      { columna: 4, tipo: "clp" },
    ]);
  });
});

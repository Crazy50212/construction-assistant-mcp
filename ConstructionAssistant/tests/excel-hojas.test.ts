import { describe, expect, it } from "vitest";
import { esNombreHojaPredeterminada } from "../src/lib/excel-hojas.js";

describe("detección de hojas predeterminadas de Excel", () => {
  it.each(["Hoja1", "Hoja2", "Sheet1", "sheet3"])('identifica "%s" como nombre predeterminado', (name) => {
    expect(esNombreHojaPredeterminada(name)).toBe(true);
  });

  it.each(["Presupuesto", "Resumen", "Hoja de cálculo", "Sheet_Data1"])('no identifica "%s" como nombre predeterminado', (name) => {
    expect(esNombreHojaPredeterminada(name)).toBe(false);
  });
});

import { describe, expect, it } from "vitest";
import {
  FORMULA_CANTIDAD_DESDE_CUBICACION,
  FORMULA_CANTIDAD_DESDE_PRESUPUESTO,
  FORMULA_PRECIO_DESDE_PRESUPUESTO,
  materializarFormulaEnlace,
} from "../src/lib/excel-sync-templates.js";

describe("fórmulas de sincronización entre hojas", () => {
  it("vincula cubicación al concepto exacto y respeta superficie, volumen o ajuste por unidad", () => {
    const formula = materializarFormulaEnlace(FORMULA_CANTIDAD_DESDE_CUBICACION, 7, 12.5);
    expect(formula).toContain("COUNTIF(tblCubicaciones[Concepto],$A7)");
    expect(formula).toContain("SUMIFS(tblCubicaciones[Superficie (m²)],tblCubicaciones[Concepto],$A7)");
    expect(formula).toContain("tblCubicaciones[Superficie (m²)]");
    expect(formula).toContain("tblCubicaciones[Volumen (m³)]");
    expect(formula).toContain("tblCubicaciones[Volumen ajustado (m³)]");
    expect(formula).toContain(",12.5,");
    expect(formula).not.toContain("SUMPRODUCT");
  });

  it("vincula Avance por nombre y unidad normalizada con fórmulas SUMIFS", () => {
    const quantityFormula = materializarFormulaEnlace(FORMULA_CANTIDAD_DESDE_PRESUPUESTO, 9, 40);
    const priceFormula = materializarFormulaEnlace(FORMULA_PRECIO_DESDE_PRESUPUESTO, 9, 2500);
    for (const formula of [quantityFormula, priceFormula]) {
      expect(formula).toContain("tblPresupuesto[Partida],$A9");
      expect(formula).toContain("tblPresupuesto[Unidad]");
      expect(formula).toContain('LOWER(TRIM($B9))="m3"');
      expect(formula).toContain('tblPresupuesto[Unidad],"m³"');
      expect(formula).not.toContain("INDEX((tblPresupuesto");
      expect(formula).not.toContain("{fila}");
      expect(formula).not.toContain("{fallback}");
    }
    expect(quantityFormula).toContain("SUMIFS(tblPresupuesto[Cantidad]");
    expect(quantityFormula).toContain(",40,");
    expect(priceFormula).toContain("SUMIFS(tblPresupuesto[Costo total]");
    expect(priceFormula).toContain("SUMIFS(tblPresupuesto[Cantidad]");
    expect(priceFormula).toContain(",2500,");
  });
});

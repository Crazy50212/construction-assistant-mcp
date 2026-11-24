import { describe, expect, it } from "vitest";
import { crearPlanHojaExcel, materializarFormula } from "../src/lib/excel-plan.js";

describe("planes de fórmulas y actualización de hojas Excel", () => {
  it("genera fórmulas de costo y subtotales, e inserta presupuesto antes de sus totales", () => {
    const plan = crearPlanHojaExcel("Presupuesto", 0.19);
    expect(plan.estrategia).toBe("insertar-antes-del-subtotal");
    expect(materializarFormula(plan.formulasFila.E, 12)).toBe('=IF(AND(ISNUMBER(C12),C12>0,ISNUMBER(D12),D12>=0),IFERROR(C12*D12,""),"")');
    expect(plan.formulasTotales).toEqual({
      subtotal: "=SUM(E{firstDataRow}:E{lastDataRow})",
      iva: "=E{subtotalRow}*0.19",
      total: "=E{subtotalRow}+E{ivaRow}",
    });
  });

  it("planifica cubicaciones como fórmulas geométricas sobre longitudes y no moneda", () => {
    const plan = crearPlanHojaExcel("Cubicaciones");
    expect(plan.formulasFila).toEqual({
      F: '=IF(OR(NOT(ISNUMBER(C{fila})),C{fila}<=0,NOT(ISNUMBER(D{fila})),D{fila}<=0),"",IFERROR(C{fila}*D{fila},""))',
      G: '=IF(OR(NOT(ISNUMBER(F{fila})),NOT(ISNUMBER(E{fila})),E{fila}<=0),"",IFERROR(F{fila}*E{fila},""))',
      I: '=IF(OR(NOT(ISNUMBER(G{fila})),NOT(ISNUMBER(H{fila})),H{fila}<=0),"",IFERROR(G{fila}*H{fila},""))',
    });
    expect(plan.estrategia).toBe("agregar-al-final");
  });

  it("calcula avance, pendiente y costo con fórmulas de fila", () => {
    const plan = crearPlanHojaExcel("Avance");
    expect(materializarFormula(plan.formulasFila.E, 8)).toBe('=IFERROR(IF(AND(ISNUMBER(C8),C8>0,ISNUMBER(D8),D8>=0),D8/C8,""),"")');
    expect(materializarFormula(plan.formulasFila.F, 8)).toBe('=IF(AND(ISNUMBER(C8),C8>0,ISNUMBER(D8),D8>=0),MAX(C8-D8,0),"")');
    expect(materializarFormula(plan.formulasFila.H, 8)).toBe('=IF(AND(ISNUMBER(D8),D8>=0,ISNUMBER(G8),G8>=0),IFERROR(D8*G8,""),"")');
    expect(materializarFormula(plan.formulasFila.I, 8)).toBe('=IF(OR(NOT(ISNUMBER(C8)),C8<=0,NOT(ISNUMBER(D8)),D8<0,NOT(ISNUMBER(G8)),G8<0,NOT(ISNUMBER(H8))),"REVISAR DATOS",IF(D8>C8,"SOBREEJECUCIÓN","OK"))');
    expect(materializarFormula(plan.formulasFila.J, 8)).toBe('=IF(AND(ISNUMBER(C8),C8>0,ISNUMBER(D8),D8>=0),MAX(D8-C8,0),"")');
  });
});

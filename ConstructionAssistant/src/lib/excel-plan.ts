export type HojaConstruction = "Cubicaciones" | "Presupuesto" | "Avance";
export type EstrategiaActualizacion = "agregar-al-final" | "insertar-antes-del-subtotal";

export interface PlanHojaExcel {
  estrategia: EstrategiaActualizacion;
  formulasFila: Record<string, string>;
  formulasTotales?: { subtotal: string; iva: string; total: string };
}

export function crearPlanHojaExcel(hoja: HojaConstruction, tasaIva = 0.19): PlanHojaExcel {
  if (!Number.isFinite(tasaIva) || tasaIva < 0 || tasaIva > 1) {
    throw new Error("La tasa de IVA debe estar entre 0 y 1.");
  }

  if (hoja === "Cubicaciones") {
    return {
      estrategia: "agregar-al-final",
      formulasFila: {
        F: '=IF(OR(NOT(ISNUMBER(C{fila})),C{fila}<=0,NOT(ISNUMBER(D{fila})),D{fila}<=0),"",IFERROR(C{fila}*D{fila},""))',
        G: '=IF(OR(NOT(ISNUMBER(F{fila})),NOT(ISNUMBER(E{fila})),E{fila}<=0),"",IFERROR(F{fila}*E{fila},""))',
        I: '=IF(OR(NOT(ISNUMBER(G{fila})),NOT(ISNUMBER(H{fila})),H{fila}<=0),"",IFERROR(G{fila}*H{fila},""))',
      },
    };
  }

  if (hoja === "Avance") {
    return {
      estrategia: "agregar-al-final",
      formulasFila: {
        E: '=IFERROR(IF(AND(ISNUMBER(C{fila}),C{fila}>0,ISNUMBER(D{fila}),D{fila}>=0),D{fila}/C{fila},""),"")',
        F: '=IF(AND(ISNUMBER(C{fila}),C{fila}>0,ISNUMBER(D{fila}),D{fila}>=0),MAX(C{fila}-D{fila},0),"")',
        H: '=IF(AND(ISNUMBER(D{fila}),D{fila}>=0,ISNUMBER(G{fila}),G{fila}>=0),IFERROR(D{fila}*G{fila},""),"")',
        I: '=IF(OR(NOT(ISNUMBER(C{fila})),C{fila}<=0,NOT(ISNUMBER(D{fila})),D{fila}<0,NOT(ISNUMBER(G{fila})),G{fila}<0,NOT(ISNUMBER(H{fila}))),"REVISAR DATOS",IF(D{fila}>C{fila},"SOBREEJECUCIÓN","OK"))',
        J: '=IF(AND(ISNUMBER(C{fila}),C{fila}>0,ISNUMBER(D{fila}),D{fila}>=0),MAX(D{fila}-C{fila},0),"")',
      },
    };
  }

  return {
    estrategia: "insertar-antes-del-subtotal",
    formulasFila: { E: '=IF(AND(ISNUMBER(C{fila}),C{fila}>0,ISNUMBER(D{fila}),D{fila}>=0),IFERROR(C{fila}*D{fila},""),"")' },
    formulasTotales: {
      subtotal: "=SUM(E{firstDataRow}:E{lastDataRow})",
      iva: `=E{subtotalRow}*${tasaIva.toString()}`,
      total: "=E{subtotalRow}+E{ivaRow}",
    },
  };
}

export function materializarFormula(formula: string, fila: number): string {
  return formula.replaceAll("{fila}", String(fila));
}

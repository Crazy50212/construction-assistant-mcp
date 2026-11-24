export type TipoFormatoExcel = "clp" | "cantidad" | "porcentaje";

export interface FormatoColumnaExcel {
  columna: number;
  tipo: TipoFormatoExcel;
}

function normalizarEncabezado(encabezado: string): string {
  return encabezado
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

export function inferirFormatosColumnas(encabezados: unknown[]): FormatoColumnaExcel[] {
  const formatos: FormatoColumnaExcel[] = [];

  encabezados.forEach((valor, indice) => {
    if (typeof valor !== "string") return;
    const encabezado = normalizarEncabezado(valor);

    // La regla de porcentaje tiene prioridad: nunca debe heredar CLP.
    if (encabezado.includes("avance fisico") || encabezado.includes("porcentaje") || /%/.test(encabezado)) {
      formatos.push({ columna: indice + 1, tipo: "porcentaje" });
    } else if (/precio unitario|costo|subtotal|\biva\b|\btotal\b/.test(encabezado)) {
      formatos.push({ columna: indice + 1, tipo: "clp" });
    } else if (/cantidad|largo|ancho|profundidad|superficie|volumen|pendiente|exceso|\bsaldo\b|factor/.test(encabezado)) {
      formatos.push({ columna: indice + 1, tipo: "cantidad" });
    }
  });

  return formatos;
}

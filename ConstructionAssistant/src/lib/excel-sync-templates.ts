export const FORMULA_CANTIDAD_DESDE_CUBICACION =
  '=IFERROR(IF(COUNTIF(tblCubicaciones[Concepto],$A{fila})=0,{fallback},IF(OR(LOWER(TRIM($B{fila}))="m2",LOWER(TRIM($B{fila}))="m²",LOWER(TRIM($B{fila}))="m^2"),SUMIFS(tblCubicaciones[Superficie (m²)],tblCubicaciones[Concepto],$A{fila}),IF(OR(LOWER(TRIM($B{fila}))="m3",LOWER(TRIM($B{fila}))="m³",LOWER(TRIM($B{fila}))="m^3"),SUMIFS(tblCubicaciones[Volumen (m³)],tblCubicaciones[Concepto],$A{fila}),IF(OR(LOWER(TRIM($B{fila}))="m3 ajustado",LOWER(TRIM($B{fila}))="m³ ajustado"),SUMIFS(tblCubicaciones[Volumen ajustado (m³)],tblCubicaciones[Concepto],$A{fila}),{fallback})))) ,"")';

const GRUPOS_UNIDAD = [
  { aliases: ["m2", "m²", "m^2"], dimension: "superficie" },
  { aliases: ["m3", "m³", "m^3"], dimension: "volumen" },
  { aliases: ["m3 ajustado", "m³ ajustado"], dimension: "volumenAjustado" },
] as const;

function criterioUnidad(unidad: string): string {
  return `tblPresupuesto[Unidad],"${unidad}"`;
}

function conteoPresupuesto(fila: string, unidades?: readonly string[]): string {
  if (!unidades) return `COUNTIFS(tblPresupuesto[Partida],$A${fila},tblPresupuesto[Unidad],$B${fila})`;
  return unidades
    .map((unidad) => `COUNTIFS(tblPresupuesto[Partida],$A${fila},${criterioUnidad(unidad)})`)
    .join("+");
}

function sumaPresupuesto(columna: string, fila: string, unidades?: readonly string[]): string {
  if (!unidades) return `SUMIFS(tblPresupuesto[${columna}],tblPresupuesto[Partida],$A${fila},tblPresupuesto[Unidad],$B${fila})`;
  return unidades
    .map((unidad) => `SUMIFS(tblPresupuesto[${columna}],tblPresupuesto[Partida],$A${fila},${criterioUnidad(unidad)})`)
    .join("+");
}

function formulaPorGrupo(fila: string, group: (typeof GRUPOS_UNIDAD)[number], valueColumn: string, weightedPrice: boolean): string {
  const condition = `OR(${group.aliases.map((unit) => `LOWER(TRIM($B${fila}))="${unit}"`).join(",")})`;
  const count = conteoPresupuesto(fila, group.aliases);
  if (!weightedPrice) {
    return `IF(${condition},IF(${count}=0,{fallback},${sumaPresupuesto(valueColumn, fila, group.aliases)}),`;
  }
  const validCostCount = group.aliases
    .map((unit) => `COUNTIFS(tblPresupuesto[Partida],$A${fila},${criterioUnidad(unit)},tblPresupuesto[Costo total],">=0")`)
    .join("+");
  const totalCost = sumaPresupuesto("Costo total", fila, group.aliases);
  const totalQuantity = sumaPresupuesto("Cantidad", fila, group.aliases);
  return `IF(${condition},IF(${count}=0,{fallback},IF(${validCostCount}=${count},IFERROR((${totalCost})/(${totalQuantity}),""),"")),`;
}

function crearFormulaPresupuesto(valueColumn: string, weightedPrice: boolean): string {
  const fila = "{fila}";
  const groupBranches = GRUPOS_UNIDAD.map((group) => formulaPorGrupo(fila, group, valueColumn, weightedPrice)).join("");
  const exactCount = conteoPresupuesto(fila);
  const exactValue = weightedPrice
    ? `IF(${exactCount}=0,{fallback},IF(COUNTIFS(tblPresupuesto[Partida],$A${fila},tblPresupuesto[Unidad],$B${fila},tblPresupuesto[Costo total],">=0")=${exactCount},IFERROR(SUMIFS(tblPresupuesto[Costo total],tblPresupuesto[Partida],$A${fila},tblPresupuesto[Unidad],$B${fila})/SUMIFS(tblPresupuesto[Cantidad],tblPresupuesto[Partida],$A${fila},tblPresupuesto[Unidad],$B${fila}),""),""))`
    : `IF(${exactCount}=0,{fallback},${sumaPresupuesto(valueColumn, fila)})`;
  return `=IFERROR(${groupBranches}${exactValue}${")".repeat(GRUPOS_UNIDAD.length)},"")`;
}

export const FORMULA_CANTIDAD_DESDE_PRESUPUESTO = crearFormulaPresupuesto("Cantidad", false);

export const FORMULA_PRECIO_DESDE_PRESUPUESTO = crearFormulaPresupuesto("Precio unitario", true);

export function materializarFormulaEnlace(formula: string, fila: number, fallback: number): string {
  return formula
    .replaceAll("{fila}", String(fila))
    .replaceAll("{fallback}", String(fallback));
}

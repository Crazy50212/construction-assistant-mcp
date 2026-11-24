export interface Cubicacion {
  superficieM2: number;
  volumenM3: number;
  volumenCorregidoM3: number;
}

const VALOR_MAXIMO_EXCEL = 9.99999999999999e307;

export function calcularCubicacion(
  largoM: number,
  anchoM: number,
  profundidadM: number,
  factor: number,
): Cubicacion {
  const valores = [largoM, anchoM, profundidadM, factor];
  if (valores.some((valor) => !Number.isFinite(valor) || valor <= 0)) {
    throw new Error("Las dimensiones y el factor deben ser números finitos mayores que cero.");
  }
  const superficieM2 = largoM * anchoM;
  const volumenM3 = superficieM2 * profundidadM;
  const volumenCorregidoM3 = volumenM3 * factor;
  if ([superficieM2, volumenM3, volumenCorregidoM3].some((valor) =>
    !Number.isFinite(valor) || valor <= 0 || valor > VALOR_MAXIMO_EXCEL)) {
    throw new Error("Las dimensiones o el factor exceden el rango numérico compatible con Excel; verifica las unidades y los valores.");
  }
  return { superficieM2, volumenM3, volumenCorregidoM3 };
}

export interface PartidaPresupuesto {
  nombre: string;
  unidad: string;
  cantidad: number;
  precioUnitario: number;
}

export function calcularPresupuesto(partidas: PartidaPresupuesto[], tasaIva: number) {
  if (partidas.length === 0) throw new Error("Debe incluir al menos una partida.");
  if (!Number.isFinite(tasaIva) || tasaIva < 0 || tasaIva > 1) {
    throw new Error("La tasa de IVA debe estar entre 0 y 1.");
  }
  const lineas = partidas.map((partida) => {
    if (!Number.isFinite(partida.cantidad) || partida.cantidad <= 0 ||
      !Number.isFinite(partida.precioUnitario) || partida.precioUnitario < 0) {
      throw new Error(`Cantidad o precio no válido en la partida «${partida.nombre}».`);
    }
    return { ...partida, costoTotal: partida.cantidad * partida.precioUnitario };
  });
  const subtotal = lineas.reduce((suma, linea) => suma + linea.costoTotal, 0);
  const iva = subtotal * tasaIva;
  return { lineas, subtotal, tasaIva, iva, total: subtotal + iva };
}

export interface PartidaAvance {
  nombre: string;
  unidad: string;
  cantidadPresupuestada: number;
  cantidadEjecutada: number;
  precioUnitario: number;
}

export function calcularAvance(partidas: PartidaAvance[]) {
  if (partidas.length === 0) throw new Error("Debe incluir al menos una partida.");
  const lineas = partidas.map((partida) => {
    if (!Number.isFinite(partida.cantidadPresupuestada) || partida.cantidadPresupuestada <= 0 ||
      !Number.isFinite(partida.cantidadEjecutada) || partida.cantidadEjecutada < 0 ||
      !Number.isFinite(partida.precioUnitario) || partida.precioUnitario < 0) {
      throw new Error(`Cantidades o precio no válidos en la partida «${partida.nombre}».`);
    }
    const porcentajeAvance = (partida.cantidadEjecutada / partida.cantidadPresupuestada) * 100;
    const sobreEjecucion = Math.max(partida.cantidadEjecutada - partida.cantidadPresupuestada, 0);
    return {
      ...partida,
      porcentajeAvance,
      costoEjecutado: partida.cantidadEjecutada * partida.precioUnitario,
      cantidadPendiente: Math.max(partida.cantidadPresupuestada - partida.cantidadEjecutada, 0),
      sobreEjecucion,
      estadoAvance: sobreEjecucion > 0 ? "SOBREEJECUCIÓN" : "OK",
    };
  });
  const totalPresupuestado = lineas.reduce((suma, linea) => suma + linea.cantidadPresupuestada * linea.precioUnitario, 0);
  const costoEjecutado = lineas.reduce((suma, linea) => suma + linea.costoEjecutado, 0);
  const cantidadPresupuestadaTotal = lineas.reduce((suma, linea) => suma + linea.cantidadPresupuestada, 0);
  const cantidadEjecutadaTotal = lineas.reduce((suma, linea) => suma + linea.cantidadEjecutada, 0);
  const totalesPorUnidad = new Map<string, { presupuestada: number; ejecutada: number }>();
  for (const linea of lineas) {
    const acumulado = totalesPorUnidad.get(linea.unidad) ?? { presupuestada: 0, ejecutada: 0 };
    acumulado.presupuestada += linea.cantidadPresupuestada;
    acumulado.ejecutada += linea.cantidadEjecutada;
    totalesPorUnidad.set(linea.unidad, acumulado);
  }
  const avancePonderadoPorUnidad = Object.fromEntries(
    [...totalesPorUnidad].map(([unidad, cantidades]) => [
      unidad,
      cantidades.presupuestada === 0 ? 0 : (cantidades.ejecutada / cantidades.presupuestada) * 100,
    ]),
  );
  return {
    lineas,
    costoEjecutado,
    totalPresupuestado,
    porcentajeAvancePonderado: totalesPorUnidad.size === 1 && cantidadPresupuestadaTotal > 0
      ? (cantidadEjecutadaTotal / cantidadPresupuestadaTotal) * 100
      : null,
    avancePonderadoPorUnidad,
  };
}

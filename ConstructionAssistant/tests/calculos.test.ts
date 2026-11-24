import { describe, expect, it } from "vitest";
import { calcularAvance, calcularCubicacion, calcularPresupuesto } from "../src/lib/calculos.js";

describe("cálculos de obras civiles", () => {
  it("calcula superficie, volumen excavado y volumen ajustado", () => {
    expect(calcularCubicacion(10, 2, 1.5, 1.2)).toEqual({
      superficieM2: 20,
      volumenM3: 30,
      volumenCorregidoM3: 36,
    });
  });

  it("rechaza dimensiones no positivas", () => {
    expect(() => calcularCubicacion(0, 2, 1, 1)).toThrow(/mayores que cero/);
  });

  it("acepta dimensiones grandes sin un límite fijo arbitrario", () => {
    expect(calcularCubicacion(150, 400, 20, 1)).toEqual({
      superficieM2: 60000,
      volumenM3: 1200000,
      volumenCorregidoM3: 1200000,
    });
  });

  it("rechaza resultados que exceden el rango numérico de Excel", () => {
    expect(() => calcularCubicacion(1e200, 1e200, 1, 1)).toThrow(/rango numérico compatible con Excel/);
  });

  it("calcula subtotal, IVA del 19% y total del presupuesto", () => {
    const resultado = calcularPresupuesto([
      { nombre: "Excavación", unidad: "m3", cantidad: 10, precioUnitario: 5000 },
      { nombre: "Hormigón", unidad: "m3", cantidad: 2, precioUnitario: 80000 },
    ], 0.19);
    expect(resultado.subtotal).toBe(210000);
    expect(resultado.iva).toBe(39900);
    expect(resultado.total).toBe(249900);
  });

  it("calcula avance físico y costos, permitiendo identificar sobre ejecución", () => {
    const resultado = calcularAvance([
      { nombre: "Excavación", unidad: "m3", cantidadPresupuestada: 100, cantidadEjecutada: 60, precioUnitario: 1000 },
      { nombre: "Relleno", unidad: "m3", cantidadPresupuestada: 20, cantidadEjecutada: 25, precioUnitario: 2000 },
    ]);
    expect(resultado.lineas.map((linea) => linea.porcentajeAvance)).toEqual([60, 125]);
    expect(resultado.lineas[1].cantidadPendiente).toBe(0);
    expect(resultado.lineas[1].sobreEjecucion).toBe(5);
    expect(resultado.lineas[1].estadoAvance).toBe("SOBREEJECUCIÓN");
    expect(resultado.costoEjecutado).toBe(110000);
    expect(resultado.porcentajeAvancePonderado).toBeCloseTo(70.8333333);
    expect(resultado.avancePonderadoPorUnidad.m3).toBeCloseTo(70.8333333);
  });

  it("rechaza cantidades ejecutadas negativas", () => {
    expect(() => calcularAvance([
      { nombre: "Hormigón", unidad: "m3", cantidadPresupuestada: 10, cantidadEjecutada: -1, precioUnitario: 1 },
    ])).toThrow(/no válidos/);
  });

  it("pondera avance por cantidad dentro de cada unidad sin mezclar magnitudes", () => {
    const resultado = calcularAvance([
      { nombre: "Excavación", unidad: "m3", cantidadPresupuestada: 100, cantidadEjecutada: 50, precioUnitario: 1000 },
      { nombre: "Enfierradura", unidad: "kg", cantidadPresupuestada: 1000, cantidadEjecutada: 250, precioUnitario: 1 },
    ]);
    expect(resultado.porcentajeAvancePonderado).toBeNull();
    expect(resultado.avancePonderadoPorUnidad).toEqual({ m3: 50, kg: 25 });
  });
});

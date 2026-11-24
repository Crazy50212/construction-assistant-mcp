import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { calcularAvance, calcularCubicacion, calcularPresupuesto } from "../lib/calculos.js";
import { crearLibroDesdePlantilla, ejecutarExcel, type ExcelCell } from "../lib/excel-com.js";
import { error, json } from "../lib/respuestas.js";

const rutaSchema = z.string().min(1).max(500).describe("Ruta absoluta a un libro Excel existente (.xlsx, .xlsm o .xls)");
const rutaPlantillaSchema = z.string().trim().min(1).max(500).describe("Ruta absoluta a una plantilla Excel .xlsx existente");
const rutaDestinoSchema = z.string().trim().min(1).max(500).describe("Ruta absoluta nueva para guardar el libro de salida .xlsx; no se sobrescriben archivos existentes");
const nombreSchema = z.string().trim().min(1).max(120);

export function register(server: McpServer): void {
  server.registerTool(
    "excel_listar_hojas",
    {
      title: "Listar hojas del libro Excel",
      description: "Lee los nombres de las hojas de un libro Excel existente usando COM. Úsala para revisar el contenido estructural del libro antes de registrar cálculos; no modifica el archivo.",
      inputSchema: { ruta_archivo: rutaSchema },
      annotations: { readOnlyHint: true },
    },
    async ({ ruta_archivo }) => {
      try {
        return json(await ejecutarExcel(ruta_archivo, "sheets"));
      } catch (e) {
        return error(`No se pudieron leer las hojas de Excel: ${e instanceof Error ? e.message : String(e)}`);
      }
    },
  );

  server.registerTool(
    "construction_auditar_integridad",
    {
      title: "Auditar integridad del libro de construcción",
      description: "Lee mediante Excel COM las cuatro hojas y las tablas de Cubicaciones, Presupuesto y Avance. Clasifica vínculos, unidades incompatibles, registros manuales, partidas sin avance, datos inválidos y sobre-ejecuciones; compara fórmulas y errores sin modificar el libro.",
      inputSchema: { ruta_archivo: rutaSchema },
      annotations: { readOnlyHint: true },
    },
    async ({ ruta_archivo }) => {
      try {
        return json(await ejecutarExcel(ruta_archivo, "audit"));
      } catch (e) {
        return error(`No se pudo auditar el libro Excel: ${e instanceof Error ? e.message : String(e)}`);
      }
    },
  );

  server.registerTool(
    "construction_crear_libro_desde_plantilla",
    {
      title: "Crear libro de construcción desde una plantilla",
      description: "Copia un libro Excel plantilla mediante Excel COM a una ruta .xlsx nueva y conserva las hojas, tablas, fórmulas, formatos y gráficos. Úsala para iniciar una demostración o proyecto sin modificar la plantilla; falla si el destino ya existe.",
      inputSchema: {
        ruta_plantilla: rutaPlantillaSchema,
        ruta_destino: rutaDestinoSchema,
        confirmar: z.literal(true).describe("Confirma que se cree un nuevo archivo en la ruta de destino"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ ruta_plantilla, ruta_destino }) => {
      try {
        return json(await crearLibroDesdePlantilla(ruta_plantilla, ruta_destino));
      } catch (e) {
        return error(`No se pudo crear el libro desde la plantilla: ${e instanceof Error ? e.message : String(e)}`);
      }
    },
  );

  server.registerTool(
    "construction_cubicaciones",
    {
      title: "Cubicaciones de obras civiles",
      description: "Calcula superficie y volumen de una sección rectangular (excavación, relleno u obra), con factor opcional de ajuste, y registra dimensiones y fórmulas en la hoja Cubicaciones del libro indicado. El concepto debe usar el mismo nombre de la partida de Presupuesto para vincular su cantidad automáticamente.",
      inputSchema: {
        ruta_archivo: rutaSchema,
        nombre: nombreSchema.describe("Identificación de la cubicación, por ejemplo Excavación de fundaciones"),
        tipo: z.enum(["Excavación", "Relleno", "Obra"]).describe("Tipo de cubicación que se registrará"),
        largo_m: z.number().finite().positive().describe("Largo en metros; debe ser un número finito mayor que cero"),
        ancho_m: z.number().finite().positive().describe("Ancho en metros; debe ser un número finito mayor que cero"),
        profundidad_m: z.number().finite().positive().describe("Profundidad o espesor en metros; debe ser un número finito mayor que cero"),
        factor_ajuste: z.number().finite().positive().default(1).describe("Factor positivo de esponjamiento/ajuste del volumen; use 1 si no aplica"),
        confirmar: z.literal(true).describe("Confirma que se agregue un registro a la hoja Cubicaciones; es obligatorio para guardar"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ ruta_archivo, nombre, tipo, largo_m, ancho_m, profundidad_m, factor_ajuste }) => {
      try {
        const resultado = calcularCubicacion(largo_m, ancho_m, profundidad_m, factor_ajuste);
        const rows: ExcelCell[][] = [
          ["Cubicación de obras civiles"],
          ["Concepto", "Tipo", "Largo (m)", "Ancho (m)", "Profundidad (m)", "Superficie (m²)", "Volumen (m³)", "Factor ajuste", "Volumen ajustado (m³)"],
          [nombre, tipo, largo_m, ancho_m, profundidad_m, resultado.superficieM2, resultado.volumenM3, factor_ajuste, resultado.volumenCorregidoM3],
        ];
        const excel = await ejecutarExcel(ruta_archivo, "append", { sheetBase: "Cubicaciones", rows });
        return json({ ...resultado, excel });
      } catch (e) {
        return error(`No se pudo calcular o registrar la cubicación: ${e instanceof Error ? e.message : String(e)}`);
      }
    },
  );

  server.registerTool(
    "construction_presupuesto",
    {
      title: "Presupuesto de obras civiles",
      description: "Calcula un presupuesto por partidas, con cantidad, unidad, precio unitario, subtotal, IVA configurable (19% por defecto) y total. Agrega nuevas partidas a la misma hoja Presupuesto administrada, conserva las existentes y mantiene fórmulas de cantidades vinculadas, costos y totales actualizables.",
      inputSchema: {
        ruta_archivo: rutaSchema,
        partidas: z.array(z.object({
          nombre: nombreSchema.describe("Nombre de la partida, por ejemplo Hormigón H25 en fundaciones"),
          unidad: z.string().trim().min(1).max(20).describe("Unidad de medida, por ejemplo m3, m2, kg o unidad"),
          cantidad: z.number().finite().positive().max(1_000_000_000).describe("Cantidad presupuestada"),
          precio_unitario: z.number().finite().nonnegative().max(1_000_000_000_000).describe("Precio unitario en la moneda del presupuesto"),
        }).describe("Partida del presupuesto")).min(1).max(500).describe("Partidas que formarán el presupuesto"),
        tasa_iva: z.number().finite().min(0).max(1).default(0.19).describe("Tasa de IVA entre 0 y 1; 0.19 equivale a 19%"),
        confirmar: z.literal(true).describe("Confirma que se agreguen partidas a la hoja Presupuesto; es obligatorio para guardar"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ ruta_archivo, partidas, tasa_iva }) => {
      try {
        const partidasCalculo = partidas.map(({ nombre, unidad, cantidad, precio_unitario }) => ({ nombre, unidad, cantidad, precioUnitario: precio_unitario }));
        const resultado = calcularPresupuesto(partidasCalculo, tasa_iva);
        const rows: ExcelCell[][] = [
          ["Presupuesto de obras civiles"],
          ["Partida", "Unidad", "Cantidad", "Precio unitario", "Costo total"],
          ...resultado.lineas.map((linea) => [linea.nombre, linea.unidad, linea.cantidad, linea.precioUnitario, linea.costoTotal]),
        ];
        const excel = await ejecutarExcel(ruta_archivo, "append", { sheetBase: "Presupuesto", rows, tasaIva: tasa_iva });
        return json({ ...resultado, excel });
      } catch (e) {
        return error(`No se pudo calcular o registrar el presupuesto: ${e instanceof Error ? e.message : String(e)}`);
      }
    },
  );

  server.registerTool(
    "construction_control_avance",
    {
      title: "Control de avance de obra",
      description: "Compara cantidades presupuestadas y ejecutadas por partida, vincula cantidades y precios a Presupuesto cuando coinciden el nombre y la unidad (incluidas variantes m²/m2 y m³/m3), calcula porcentaje físico, saldo no negativo, costo ejecutado y alerta de sobre-ejecución con fórmulas de Excel.",
      inputSchema: {
        ruta_archivo: rutaSchema,
        partidas: z.array(z.object({
          nombre: nombreSchema.describe("Nombre de la partida"),
          unidad: z.string().trim().min(1).max(20).describe("Unidad de medida de la partida"),
          cantidad_presupuestada: z.number().finite().positive().max(1_000_000_000).describe("Cantidad total presupuestada"),
          cantidad_ejecutada: z.number().finite().nonnegative().max(1_000_000_000).describe("Cantidad acumulada ejecutada a la fecha"),
          precio_unitario: z.number().finite().nonnegative().max(1_000_000_000_000).describe("Precio unitario en la moneda del presupuesto"),
        }).describe("Partida cuyo avance se desea controlar")).min(1).max(500).describe("Partidas presupuestadas con sus cantidades ejecutadas"),
        confirmar: z.literal(true).describe("Confirma que se agreguen registros a la hoja Avance; es obligatorio para guardar"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ ruta_archivo, partidas }) => {
      try {
        const resultado = calcularAvance(partidas.map((partida) => ({
          nombre: partida.nombre,
          unidad: partida.unidad,
          cantidadPresupuestada: partida.cantidad_presupuestada,
          cantidadEjecutada: partida.cantidad_ejecutada,
          precioUnitario: partida.precio_unitario,
        })));
        const rows: ExcelCell[][] = [
          ["Control de avance de obra"],
          ["Partida", "Unidad", "Cantidad presupuestada", "Cantidad ejecutada", "Avance físico (%)", "Cantidad pendiente", "Precio unitario", "Costo ejecutado", "Estado de avance", "Exceso ejecutado"],
          ...resultado.lineas.map((linea) => [linea.nombre, linea.unidad, linea.cantidadPresupuestada, linea.cantidadEjecutada, linea.porcentajeAvance, linea.cantidadPendiente, linea.precioUnitario, linea.costoEjecutado, linea.estadoAvance, linea.sobreEjecucion]),
        ];
        const excel = await ejecutarExcel(ruta_archivo, "append", { sheetBase: "Avance", rows }) as {
          appendedItems?: Array<{
            nombre: string;
            unidad: string;
            cantidadPresupuestada: number;
            cantidadEjecutada: number;
            precioUnitario: number;
          }>;
        };
        const partidasExcel = excel.appendedItems?.map((partida, index) => ({
          ...partida,
          nombre: partidas[index]?.nombre ?? partida.nombre,
          unidad: partidas[index]?.unidad ?? partida.unidad,
        }));
        const partidasCalculadas = partidasExcel?.map((partida) => ({
          nombre: partida.nombre,
          unidad: partida.unidad,
          cantidadPresupuestada: partida.cantidadPresupuestada,
          cantidadEjecutada: partida.cantidadEjecutada,
          precioUnitario: partida.precioUnitario,
        }));
        const resultadoExcel = partidasCalculadas ? calcularAvance(partidasCalculadas) : resultado;
        return json({ ...resultadoExcel, excel: { ...excel, ...(partidasExcel ? { appendedItems: partidasExcel } : {}) } });
      } catch (e) {
        return error(`No se pudo calcular o registrar el control de avance: ${e instanceof Error ? e.message : String(e)}`);
      }
    },
  );
}

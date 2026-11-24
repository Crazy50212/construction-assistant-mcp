export const PATRON_HOJA_PREDETERMINADA = /^(?:Hoja|Sheet)\d+$/i;

export function esNombreHojaPredeterminada(nombre: string): boolean {
  return PATRON_HOJA_PREDETERMINADA.test(nombre.trim());
}

// src/config.js - CONFIG FINAL 6 GRUPOS + SHEETS

export const SPREADSHEET_ID = process.env.SPREADSHEET_ID || '1v0a4vHLE5GdX3l6Gxb3Xv0a4vHLE5GdX' // <-- pon tu ID real si no está en env

export const GRUPO_REPORTES_TRINIDAD_ID = '120363412984528459@g.us'
export const GRUPO_PRUEBAS_ID = '120363430350253017@g.us'
export const GRUPO_COYOACAN_ID = '120363430474175735@g.us'
export const GRUPO_BUCARELI_ID = '120363410610062461@g.us'
export const GRUPO_JUAREZ_ID = '120363410610062461@g.us'
export const GRUPO_CHECADOR_COYOACAN_ID = '120363403668034900@g.us'
export const GRUPO_CHECADOR_BUCARELI_ID = '120363411739089744@g.us'

export const SUCURSALES = [
  { nombre: 'Coyoacan', lat: 19.346, lng: -99.16, rEnt: 150, rSal: 300 },
  { nombre: 'Bucareli', lat: 19.43, lng: -99.15, rEnt: 150, rSal: 300 },
  { nombre: 'Juarez', lat: 19.43, lng: -99.15, rEnt: 150, rSal: 300 },
  { nombre: 'Hotel', lat: 19.34, lng: -99.16, rEnt: 150, rSal: 300 },
]

export const GRUPOS = {
  REPORTES: [GRUPO_REPORTES_TRINIDAD_ID, GRUPO_PRUEBAS_ID],
  GERENTES: [GRUPO_COYOACAN_ID, GRUPO_BUCARELI_ID],
  CHECADORES: [GRUPO_CHECADOR_COYOACAN_ID, GRUPO_CHECADOR_BUCARELI_ID]
}

export function getTipoGrupo(jid){
  if(GRUPOS.REPORTES.includes(jid)) return 'REPORTES'
  if(GRUPOS.CHECADORES.includes(jid)) return 'CHECADORES'
  if(GRUPOS.GERENTES.includes(jid)) return 'GERENTES'
  return null
}

export const PAQUETES = {
  REPORTES_PARA_GERENTES: /(info fer|asistencia hoy|resumen|reporte|compras de hoy)/i
}

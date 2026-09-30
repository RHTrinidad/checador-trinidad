export const SPREADSHEET_ID = process.env.SPREADSHEET_ID
export const SPREADSHEET_COMPRAS_ID = process.env.SPREADSHEET_COMPRAS_ID || "1eidbKAX5QAaDvDj_oe5BBQis3TA0WUohVZE0MH4cug4"
export const SHEET_RESUMEN = 'RESUMEN'
export const SHEET_INSUMOS = 'INSUMOS'

const GRUPO_REPORTES_RAW = process.env.GRUPO_REPORTES_ID || "120363412984528459@g.us,120363430350253017@g.us"
let GRUPO_REPORTES_IDS = GRUPO_REPORTES_RAW.split(/[, \n]+/).map(s => s.trim()).filter(s => s.includes('@g.us'))
export const GRUPO_COYOACAN_ID = (process.env.GRUPO_COYOACAN_ID || "120363430474175735@g.us").trim()
export const GRUPO_BUCARELI_ID = (process.env.GRUPO_BUCARELI_ID || "120363410610062461@g.us").trim()
export const GRUPO_JUAREZ_ID = (process.env.GRUPO_JUAREZ_ID || GRUPO_BUCARELI_ID).trim()
export const GRUPO_CHECADOR_COYOACAN_ID = (process.env.GRUPO_CHECADOR_COYOACAN_ID || "120363403668034900@g.us").trim()
export const GRUPO_CHECADOR_BUCARELI_ID = (process.env.GRUPO_CHECADOR_BUCARELI_ID || "120363411739089744@g.us").trim()

const TODOS = [GRUPO_COYOACAN_ID, GRUPO_BUCARELI_ID, GRUPO_JUAREZ_ID, GRUPO_CHECADOR_COYOACAN_ID, GRUPO_CHECADOR_BUCARELI_ID]
GRUPO_REPORTES_IDS = GRUPO_REPORTES_IDS.filter(id =>!TODOS.includes(id))
export { GRUPO_REPORTES_IDS }
export const GRUPOS = { REPORTES: GRUPO_REPORTES_IDS, CHECADORES: [GRUPO_CHECADOR_COYOACAN_ID, GRUPO_CHECADOR_BUCARELI_ID], GERENTES: [GRUPO_COYOACAN_ID, GRUPO_BUCARELI_ID, GRUPO_JUAREZ_ID] }
export const PAQUETES = { REPORTES_PARA_GERENTES: /^(numero|número|num|tel|telefono|teléfono|info|ficha|datos|dato|asistencia hoy|resumen|reporte|checador|reporte x unidad|compras)/i }
export const SUCURSALES = [
  { id:"COYOACAN", nombre:"Trinidad Coyoacan", lat:19.352525, lng:-99.161817, rEnt:150, rSal:150 },
  { id:"JUAREZ", nombre:"Trinidad Juarez", lat:19.4314119, lng:-99.1512074, rEnt:150, rSal:150 },
  { id:"HOTEL", nombre:"Servicio Hotel", lat:19.351770, lng:-99.165458, rEnt:150, rSal:150 }
]
export function getTipoGrupo(jid){
  if(GRUPOS.REPORTES.includes(jid)) return 'REPORTES';
  if(GRUPOS.CHECADORES.includes(jid)) return 'CHECADORES';
  if(GRUPOS.GERENTES.includes(jid)) return 'GERENTES';
  return null
}

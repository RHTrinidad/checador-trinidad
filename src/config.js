// src/config.js - CORREGIDO 150/150 TODO

export const SPREADSHEET_ID=process.env.SPREADSHEET_ID||''
export const SPREADSHEET_COMPRAS_ID=process.env.SPREADSHEET_COMPRAS_ID||SPREADSHEET_ID
export const SPREADSHEET_RECETAS_ID=process.env.SPREADSHEET_RECETAS_ID||''
export const SHEET_RESUMEN='Resumen'
export const SHEET_INSUMOS='Insumos'

// --- SECCION GRUPOS WHATSAPP ---
export const GRUPO_REPORTES_TRINIDAD_ID='120363412984528459@g.us'
export const GRUPO_PRUEBAS_ID='120363430350253017@g.us'
export const GRUPO_COYOACAN_ID='120363430350253017@g.us'
export const GRUPO_BUCARELI_ID='120363410610062461@g.us'
export const GRUPO_CHECADOR_COYOACAN_ID='120363403668034900@g.us'
export const GRUPO_CHECADOR_BUCARELI_ID='120363411739089744@g.us'
export const GRUPO_COMPRAS_ID='120363404272185958@g.us'

// --- SECCION SUCURSALES Y COORDENADAS ---
export const SUCURSALES=[
  {nombre:'Coyoacan',lat:19.352525,lng:-99.161817,rEnt:150,rSal:150},
  {nombre:'Hotel',lat:19.351783,lng:-99.165471,rEnt:150,rSal:150},
  {nombre:'Bucareli',lat:19.431443,lng:-99.151142,rEnt:150,rSal:150},
]

// --- SECCION AGRUPACION DE GRUPOS ---
export const GRUPOS={
  REPORTES:[
    GRUPO_REPORTES_TRINIDAD_ID,
    GRUPO_PRUEBAS_ID
  ],

  COMPRAS:[
    GRUPO_COMPRAS_ID
  ],

  GERENTES:[
    GRUPO_COYOACAN_ID,
    GRUPO_BUCARELI_ID
  ],

  CHECADORES:[
    GRUPO_CHECADOR_COYOACAN_ID,
    GRUPO_CHECADOR_BUCARELI_ID
  ]
}

export function getTipoGrupo(jid){
  if(GRUPOS.REPORTES.includes(jid))
    return'REPORTES'

  if(GRUPOS.COMPRAS.includes(jid))
    return'COMPRAS'

  if(GRUPOS.CHECADORES.includes(jid))
    return'CHECADORES'

  if(GRUPOS.GERENTES.includes(jid))
    return'GERENTES'

  return null
}

export const PAQUETES={
  REPORTES_PARA_GERENTES:
    /(info fer|asistencia hoy|resumen|reporte|compras de hoy)/i
}

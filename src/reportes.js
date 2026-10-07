import ExcelJS from 'exceljs'
import fs from 'fs'
import path from 'path'
import os from 'os'
import opentype from 'opentype.js'
import {SPREADSHEET_ID} from './config.js'
import {getRows,sheetsClient,getHorarioBaseMap} from './sheets.js'
import {fechaLaboral,horaMX,minutos,parseFechaMX,normaliza,parseHorarioRango,getRangoSemana,sucursalCoincideConFiltro,scoreEmpleado,calcularExtra} from './utils.js'
import {iniciarLista}from'./flujo_numerico.js'
const DIAS_RETARDOS=30
const DIAS_CRITICOS=45
const CUTOFF_RETARDOS='2026-10-02'

const CATALOGO_REPORTES=[
{comando:'info [nombre]',descripcion:'Datos generales del empleado',grupos:['GERENTES_COYOACAN','GERENTES_BUCARELI','REPORTES']},
{comando:'datos [nombre]',descripcion:'Datos generales del empleado',grupos:['GERENTES_COYOACAN','GERENTES_BUCARELI','REPORTES']},
{comando:'banco [nombre]',descripcion:'Banco y dato bancario disponible',grupos:['GERENTES_COYOACAN','GERENTES_BUCARELI','REPORTES']},
{comando:'cuenta [nombre]',descripcion:'Cuenta bancaria',grupos:['GERENTES_COYOACAN','GERENTES_BUCARELI','REPORTES']},
{comando:'clabe [nombre]',descripcion:'CLABE bancaria',grupos:['GERENTES_COYOACAN','GERENTES_BUCARELI','REPORTES']},
{comando:'clave [nombre]',descripcion:'CLABE bancaria',grupos:['GERENTES_COYOACAN','GERENTES_BUCARELI','REPORTES']},
{comando:'asistencia hoy',descripcion:'Asistencia del día',grupos:['GERENTES_COYOACAN','GERENTES_BUCARELI','REPORTES']},
{comando:'resumen [nombre]',descripcion:'Resumen semanal del empleado',grupos:['GERENTES_COYOACAN','GERENTES_BUCARELI','REPORTES']},
{comando:'reporte [nombre]',descripcion:'Reporte semanal por sucursal',grupos:['GERENTES_COYOACAN','GERENTES_BUCARELI','REPORTES']},
{comando:'faltas Coyoacán',descripcion:'Faltas de Coyoacán de los últimos 45 días',grupos:['GERENTES_COYOACAN','REPORTES']},
{comando:'faltas Bucareli',descripcion:'Faltas de Bucareli de los últimos 45 días',grupos:['GERENTES_BUCARELI','REPORTES']},
{comando:'retardos Coyoacán',descripcion:'Retardos de Coyoacán de los últimos 45 días',grupos:['GERENTES_COYOACAN','REPORTES']},
{comando:'retardos Bucareli',descripcion:'Retardos de Bucareli de los últimos 45 días',grupos:['GERENTES_BUCARELI','REPORTES']},
{comando:'críticos',descripcion:'Empleados con 3 o más faltas equivalentes',grupos:['REPORTES']},
{comando:'graves',descripcion:'Empleados con incidencias graves',grupos:['REPORTES']},
{comando:'reporte semanal',descripcion:'Reporte semanal',grupos:['REPORTES']},
{comando:'reporte mensual',descripcion:'Reporte mensual',grupos:['REPORTES']},
{comando:'jornada semanal',descripcion:'Jornada semanal por sucursal, área o empleado en imagen',grupos:['GERENTES_COYOACAN','GERENTES_BUCARELI','REPORTES']},
{comando:'excel',descripcion:'Generar reporte en Excel',grupos:['REPORTES']},
{comando:'compras de hoy',descripcion:'Compras registradas del día',grupos:['GERENTES_COYOACAN','GERENTES_BUCARELI','REPORTES']}
]

function grupoActual(filtroGrupo){
  const f=normaliza(filtroGrupo)
  if(f==='coyoacan'||f==='coyoacan hotel'||f==='hotel')return 'GERENTES_COYOACAN'
  if(f==='bucareli')return 'GERENTES_BUCARELI'
  return 'REPORTES'
}

function obtenerReportesDisponibles(filtroGrupo){
  const grupo=grupoActual(filtroGrupo)
  return CATALOGO_REPORTES.filter(r=>r.grupos.includes(grupo))
}

function textoListaReportes(filtroGrupo){
  const grupo=grupoActual(filtroGrupo)
  const lista=obtenerReportesDisponibles(filtroGrupo)
  const personal=[],bancarios=[],asistencia=[],incidencias=[],generales=[],compras=[]

  for(const r of lista){
    if(r.comando.startsWith('info')||r.comando.startsWith('datos')){personal.push(r);continue}
    if(r.comando.startsWith('banco')||r.comando.startsWith('cuenta')||r.comando.startsWith('clabe')||r.comando.startsWith('clave')){bancarios.push(r);continue}
    if(r.comando.startsWith('asistencia')||r.comando.startsWith('resumen')||r.comando.startsWith('reporte [')){asistencia.push(r);continue}
    if(r.comando.startsWith('faltas')||r.comando.startsWith('retardos')||r.comando.startsWith('críticos')||r.comando.startsWith('graves')){incidencias.push(r);continue}
    if(r.comando.startsWith('reporte semanal')||r.comando.startsWith('reporte mensual')||r.comando.startsWith('jornada semanal')||r.comando.startsWith('excel')){generales.push(r);continue}
    if(r.comando.startsWith('compras'))compras.push(r)
  }

  let txt='📋 *LISTA DE REPORTES*'
  if(grupo==='GERENTES_COYOACAN')txt+='\n📍 Coyoacán / Hotel'
  else if(grupo==='GERENTES_BUCARELI')txt+='\n📍 Bucareli'
  else txt+='\n📊 Reportes / Pruebas'

  if(personal.length){
    txt+='\n\n👤 *PERSONAL*'
    for(const r of personal)txt+=`\n• ${r.comando}`
  }
  if(bancarios.length){
    txt+='\n\n🏦 *BANCARIOS*'
    for(const r of bancarios)txt+=`\n• ${r.comando}`
  }
  if(asistencia.length){
    txt+='\n\n🕐 *ASISTENCIA*'
    for(const r of asistencia)txt+=`\n• ${r.comando}`
  }
  if(incidencias.length){
    txt+='\n\n📊 *INCIDENCIAS*'
    for(const r of incidencias)txt+=`\n• ${r.comando}`
  }
  if(generales.length){
    txt+='\n\n📈 *REPORTES GENERALES*'
    for(const r of generales)txt+=`\n• ${r.comando}`
  }
  if(compras.length){
    txt+='\n\n🛒 *COMPRAS*'
    for(const r of compras)txt+=`\n• ${r.comando}`
  }

  txt+='\n\nEscribe cualquiera de los comandos de la lista.'
  return txt
}

function tel10(v){
  return (v||'').toString().replace(/\D/g,'').slice(-10)
}

function nombreEmpleado(r){
  return (r?.[3]||r?.[1]||'').toString().trim()
}

function sucursalEmpleado(r){
  return (r?.[2]||'').toString().trim()
}

function esDescansoRegistro(f){
  const d=(f?.[3]||'').toString().trim().toUpperCase()
  const e=(f?.[4]||'').toString().trim().toUpperCase()
  const m=(f?.[12]||'').toString().trim().toUpperCase()
  return d==='DESCANSO'||e==='DESCANSO'||m==='DESCANSO'
}

function esTrabajoDescanso(f){
  return (f?.[4]||'').toString().trim().toUpperCase()==='TRABAJO EN DESCANSO'
}

function esLibreHorario(v){
  const p=parseHorarioRango(v)
  return p?.entrada==='LIBRE'
}

function getHorarioDia(row,dia){
  const mapa={1:row?.[3],2:row?.[4],3:row?.[5],4:row?.[6],5:row?.[7],6:row?.[8],0:row?.[9]}
  return (mapa[dia]||'').toString().trim()
}

function fechaClave(f){
  const x=parseFechaMX(f)
  if(!x||isNaN(x.getTime()))return null
  return `${x.getFullYear()}-${String(x.getMonth()+1).padStart(2,'0')}-${String(x.getDate()).padStart(2,'0')}`
}

function fechaDesdeHoy(dias){
  const hoy=new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'}))
  hoy.setHours(0,0,0,0)
  hoy.setDate(hoy.getDate()-dias)
  return hoy
}

function fechaEnRango(f,desde,hasta){
  const x=parseFechaMX(f)
  if(!x||isNaN(x.getTime()))return false
  x.setHours(0,0,0,0)
  return x>=desde&&x<=hasta
}

function minutosTrabajados(v){
  if(v===undefined||v===null||v==='')return 0
  const s=String(v).trim()
  if(s==='8')return 480
  const m=s.match(/^(\d+):(\d{2})/)
  if(m)return Number(m[1])*60+Number(m[2])
  const n=Number(s)
  return Number.isFinite(n)?Math.round(n*60):0
}

function formatoHoras(min){
  min=Math.max(0,Math.round(min||0))
  return `${Math.floor(min/60)}:${String(min%60).padStart(2,'0')}h`
}

function sucursalCoincideAsignada(suc,filtro){
  const x=normaliza(suc)
  const f=normaliza(filtro)
  if(!f)return true
  if(f==='coyoacan'||f==='coyoacan hotel'||f==='hotel')return x.includes('coyo')||x.includes('hotel')||x.includes('trinidad')
  if(f==='bucareli')return x.includes('bucareli')
  return x.includes(f)
}

async function cargarDatos(){
  const sClient=await sheetsClient()

  const [asis,emp,base]=await Promise.all([
    sClient.spreadsheets.values.get({spreadsheetId:SPREADSHEET_ID,range:'Asistencia!A2:M'}),
    getRows('Empleados!A:T'),
    getRows('Horario_Base!A2:K')
  ])

  return {sClient,asis:asis.data.values||[],emp,base}
}

function mapaEmpleados(empRows){
  const map={}
  for(const r of empRows){
    const tel=tel10(r[0])
    if(tel)map[tel]=r
  }
  return map
}

function mapaBase(baseRows){
  const map={}
  for(const r of baseRows){
    const tel=tel10(r[0])
    if(tel)map[tel]=r
  }
  return map
}

// =====================================================
// ASISTENCIA DE HOY
// =====================================================

export async function asistenciaHoy(filtroSucursal,jid,sock,areaFiltro=''){
  const {asis,base}=await cargarDatos()
  const fLab=fechaLaboral()
  const ahoraMin=minutos(horaMX())
  const mx=new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'}))
  const diaNum=mx.getDay()
  const grupos={}
  for(const r of base){
    const sucBase=(r[2]||'').toString()
   if(filtroSucursal!=='todas'&&!sucursalCoincideAsignada(sucBase,filtroSucursal))continue 
    const area=normalizaArea(r[10])
    if(areaFiltro&&!jornadaCoincideArea(area,areaFiltro))continue
    const nombre=r[1]||''
    const tel=tel10(r[0])
    if(!tel)continue
    const horario=getHorarioDia(r,diaNum)
    if(!horario)continue
    const claveArea=area||'sinarea'
    const clave=`${normaliza(sucBase)}|${claveArea}`
    if(!grupos[clave]){
      grupos[clave]={
        sucursal:sucBase,
        area:claveArea,
        llego:[],
        retardo:[],
        falta:[],
        futuro:[],
        descanso:[]
      }
    }
    const g=grupos[clave]
    if(horario.toLowerCase().includes('descanso')){
      g.descanso.push(`• ${nombre} - 💤 DESCANSO`)
      continue
    }
    const parsed=parseHorarioRango(horario)
    if(!parsed)continue
    if(parsed.entrada==='LIBRE'){
      const registro=asis.find(a=>tel10(a[0])===tel&&a[2]===fLab&&a[3])
      if(registro)g.llego.push(`• ${nombre} - Entró ${registro[3]} en ${registro[6]||'-'} ✅`)
      else g.futuro.push(`• ${nombre} - LIBRE`)
      continue
    }
    const registro=asis.find(a=>tel10(a[0])===tel&&a[2]===fLab&&a[3]&&!esDescansoRegistro(a))
    if(registro){
      const entrada=registro[3]
      const dif=minutos(entrada)-minutos(parsed.entrada)
      if(esTrabajoDescanso(registro)){
        g.llego.push(`• ${nombre} - TRABAJO EN DESCANSO - ${entrada} ✅`)
      }else if(dif>15){
        g.retardo.push(`• ${nombre} - Prog ${parsed.entrada} - Entró ${entrada} - ⏰ ${dif}m tarde - ${registro[6]||''}`)
      }else{
        g.llego.push(`• ${nombre} - Prog ${parsed.entrada} - Entró ${entrada} ✅ - ${registro[6]||''}`)
      }
    }else{
      const dif=ahoraMin-minutos(parsed.entrada)
      if(dif<0)g.futuro.push(`• ${nombre} - Prog ${parsed.entrada}`)
      else g.falta.push(`• ${nombre} - Prog ${parsed.entrada} - ❌ ${dif}m sin llegar`)
    }
  }
  const gruposOrdenados=Object.values(grupos).sort((a,b)=>{
    const sa=normaliza(a.sucursal)
    const sb=normaliza(b.sucursal)
    const aa=a.area
    const ab=b.area
    const ordenSuc=x=>x.includes('coyo')||x.includes('hotel')||x.includes('trinidad')?0:x.includes('bucareli')?1:9
    const ordenArea=x=>x==='cocina'?0:x==='barra'?1:x==='salon'?2:9
    return ordenSuc(sa)-ordenSuc(sb)||ordenArea(aa)-ordenArea(ab)
  })
  let txt=`📍 *ASISTENCIA HOY ${fLab}* ${horaMX()}\n`
  for(const g of gruposOrdenados){
    const tituloSuc=nombreSucursalJornada(g.sucursal)
    const tituloArea=g.area==='sinarea'?'SIN AREA':nombreAreaJornada(g.area)
    txt+=`\n━━━━━━━━━━━━━━━━━━\n📍 *${tituloSuc} - ${tituloArea}*\n━━━━━━━━━━━━━━━━━━\n`
    txt+=`\n✅ *A TIEMPO (${g.llego.length}):*\n${g.llego.join('\n')||'-'}\n`
    txt+=`\n⏰ *RETARDOS (${g.retardo.length}):*\n${g.retardo.join('\n')||'-'}\n`
    txt+=`\n❌ *NO HAN LLEGADO (${g.falta.length}):*\n${g.falta.join('\n')||'Todos llegaron'}\n`
    txt+=`\n⏳ *PRÓXIMOS (${g.futuro.length}):*\n${g.futuro.join('\n')||'-'}\n`
    txt+=`\n💤 *DESCANSOS (${g.descanso.length}):*\n${g.descanso.join('\n')||'-'}\n`
  }
  if(!gruposOrdenados.length){
    txt+='\nNo se encontraron empleados para el filtro solicitado.'
  }
  await sock.sendMessage(jid,{text:txt})
}

// =====================================================
// JORNADA SEMANAL - IMAGEN
// =====================================================

function escaparSvg(v){
  return String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;')
}

function abreviarHorario(v){
  const s=String(v||'').trim()
  if(!s)return '-'
  if(normaliza(s)==='descanso')return 'DESCANSO'
  if(normaliza(s)==='libre')return 'LIBRE'
  return s
}

function filtroJornada(filtroGrupo,sucursalFiltro=''){
  const f=normaliza(sucursalFiltro||filtroGrupo)
  if(f.includes('bucareli'))return ['bucareli']
  if(f.includes('coyoacan')||f.includes('hotel')||f.includes('trinidad'))return ['coyoacan','hotel']
  return ['coyoacan','hotel','bucareli']
}

function jornadaCoincideSucursal(sucursal,sucursales){
  const x=normaliza(sucursal)
  if(!x)return false

  return sucursales.some(s=>{
    const f=normaliza(s)
    if(f==='coyoacan')return x.includes('coyo')||x.includes('trinidad')
    if(f==='hotel')return x.includes('hotel')
    if(f==='bucareli')return x.includes('bucareli')
    return x.includes(f)
  })
}

function normalizaArea(area){
  const a=normaliza(area)
  if(a==='cocina')return 'cocina'
  if(a==='barra')return 'barra'
  if(a==='salon')return 'salon'
  return ''
}

function jornadaCoincideArea(area,areaFiltro){
  if(!areaFiltro)return true
  return normalizaArea(area)===normalizaArea(areaFiltro)
}

function nombreAreaJornada(area){
  const a=normalizaArea(area)
  if(a==='cocina')return 'COCINA'
  if(a==='barra')return 'BARRA'
  if(a==='salon')return 'SALON'
  return ''
}

function jornadaCoincideEmpleado(nombre,empleadoFiltro){
  if(!empleadoFiltro)return true
  const n=normaliza(nombre)
  const f=normaliza(empleadoFiltro)
  if(!n||!f)return false
  return n.includes(f)
}

function obtenerFiltrosJornada(texto,filtroGrupo){
  const low=normaliza(texto)
  let sucursal=''
  let area=''
  let empleado=''

  if(low.includes('bucareli'))sucursal='bucareli'
  else if(low.includes('coyoacan')||low.includes('hotel')||low.includes('trinidad'))sucursal='coyoacan'

  if(low.includes('cocina'))area='cocina'
  else if(low.includes('barra'))area='barra'
  else if(low.includes('salon'))area='salon'

  let resto=low

  resto=resto
    .replace(/\bcoyoacan\b/g,'')
    .replace(/\bhotel\b/g,'')
    .replace(/\btrinidad\b/g,'')
    .replace(/\bbucareli\b/g,'')
    .replace(/\bcocina\b/g,'')
    .replace(/\bbarra\b/g,'')
    .replace(/\bsalon\b/g,'')
    .replace(/\s+/g,' ')
    .trim()

  if(resto)empleado=resto

  if(!sucursal){
    const grupo=normaliza(filtroGrupo)
    if(grupo==='bucareli')sucursal='bucareli'
    else if(grupo==='coyoacan'||grupo==='coyoacan hotel'||grupo==='hotel')sucursal='coyoacan'
  }

  return {sucursal,area,empleado}
}

function textoTituloJornada(sucursal,area,empleado=''){
  let titulo=sucursal==='bucareli'?'BUCARELI':sucursal==='coyoacan'?'COYOACAN':'TODAS LAS SUCURSALES'
  const a=nombreAreaJornada(area)
  if(a)titulo+=` - ${a}`
  if(empleado)titulo+=` - ${empleado.toUpperCase()}`
  return titulo
}

function anchoTextoSvg(texto,max){
  const s=String(texto||'')
  if(s.length<=max)return s
  return `${s.slice(0,Math.max(1,max-1))}...`
}

function textoComoPath(font,texto,x,y,fontSize,fill='#111111',anchor='start'){
  const s=String(texto??'')
  if(!s)return ''

  let xReal=x

  if(anchor==='middle'){
    const ancho=font.getAdvanceWidth(s,fontSize,{kerning:true})
    xReal=x-(ancho/2)
  }else if(anchor==='end'){
    const ancho=font.getAdvanceWidth(s,fontSize,{kerning:true})
    xReal=x-ancho
  }

  const p=font.getPath(s,xReal,y,fontSize,{kerning:true})
  const d=p.toPathData(2)

  if(!d)return ''
  return `<path d="${d}" fill="${fill}"/>`
}
function rangoSemanaJornada(){
  const mx=new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'}))
  mx.setHours(0,0,0,0)
  const lunes=new Date(mx)
  lunes.setDate(mx.getDate()-(mx.getDay()||7)+1)
  const domingo=new Date(lunes)
  domingo.setDate(lunes.getDate()+6)
  const fmt=d=>`${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`
  return `Semana: ${fmt(lunes)} - ${fmt(domingo)}`
}

function nombreSucursalJornada(suc){
  const x=normaliza(suc)
  if(x.includes('bucareli'))return 'BUCARELI'
  if(x.includes('coyo')||x.includes('hotel')||x.includes('trinidad'))return 'COYOACAN'
  return suc||''
}
// =====================================================
// GENERAR IMAGEN
// =====================================================

async function generarImagenJornada(filas,filtroGrupo,jid,sock,sucursalFiltro='',areaFiltro='',empleadoFiltro=''){
  console.log('🔥 ENTRE A generarImagenJornada - VERSION NUEVA')
  const {default:sharp}=await import('sharp')
  const sucursales=filtroJornada(filtroGrupo,sucursalFiltro)
  const datos=filas
    .filter(r=>jornadaCoincideSucursal(r[2],sucursales))
    .filter(r=>jornadaCoincideArea(r[10],areaFiltro))
    .filter(r=>jornadaCoincideEmpleado(r[1],empleadoFiltro))
    .map(r=>({
      nombre:(r[1]||'').toString().trim(),
      sucursal:(r[2]||'').toString().trim(),
      area:normalizaArea(r[10]),
      lun:abreviarHorario(r[3]),
      mar:abreviarHorario(r[4]),
      mie:abreviarHorario(r[5]),
      jue:abreviarHorario(r[6]),
      vie:abreviarHorario(r[7]),
      sab:abreviarHorario(r[8]),
      dom:abreviarHorario(r[9])
    }))
    .filter(r=>r.nombre)
    .sort((a,b)=>normaliza(a.nombre).localeCompare(normaliza(b.nombre)))

  if(!datos.length){
    await sock.sendMessage(jid,{text:'No se encontraron empleados para la jornada solicitada.'})
    return
  }

  function buscarFuente(dir){
    if(!fs.existsSync(dir))return null
    const encontrados=[]
    function recorrer(actual){
      let items=[]
      try{items=fs.readdirSync(actual,{withFileTypes:true})}catch{return}
      for(const item of items){
        const full=path.join(actual,item.name)
        if(item.isDirectory()){recorrer(full);continue}
        const nombre=item.name.toLowerCase()
        if(nombre.endsWith('.ttf')||nombre.endsWith('.otf')||nombre.endsWith('.woff'))encontrados.push(full)
      }
    }
    recorrer(dir)
    if(!encontrados.length)return null
    const prioridad=[/roboto-latin-400-normal\.woff$/i,/roboto-400-normal\.woff$/i,/roboto-regular\.woff$/i,/roboto-regular\.ttf$/i,/roboto-regular\.otf$/i]
    for(const patron of prioridad){
      const encontrada=encontrados.find(x=>patron.test(x))
      if(encontrada)return encontrada
    }
    const regular=encontrados.find(x=>{
      const n=path.basename(x).toLowerCase()
      return (n.includes('latin')&&n.includes('400')&&!n.includes('italic'))||n.includes('regular')||n.includes('400-normal')
    })
    if(regular)return regular
    const ttf=encontrados.find(x=>x.toLowerCase().endsWith('.ttf'))
    if(ttf)return ttf
    const otf=encontrados.find(x=>x.toLowerCase().endsWith('.otf'))
    if(otf)return otf
    return encontrados.find(x=>x.toLowerCase().endsWith('.woff'))||null
  }

  const fontDir=path.join(process.cwd(),'node_modules','@fontsource','roboto')
  const fontPath=buscarFuente(fontDir)
  if(!fontPath)throw new Error('No se encontró una fuente compatible en @fontsource/roboto')

  console.log('FUENTE JORNADA:',fontPath)

  let font
  try{font=opentype.loadSync(fontPath)}catch(error){
    console.error('ERROR CARGANDO FUENTE:',error)
    throw new Error(`No se pudo cargar la fuente: ${fontPath}`)
  }

  function pathTexto(texto,x,y,size,fill='#111111',anchor='start'){
    const s=String(texto??'').replace(/\r/g,'').replace(/\n/g,' ')
    if(!s)return ''
    try{
      const ancho=font.getAdvanceWidth(s,size,{kerning:true})
      const xReal=anchor==='middle'?x-ancho/2:anchor==='end'?x-ancho:x
      const d=font.getPath(s,xReal,y,size,{kerning:true}).toPathData(3)
      return d?`<path d="${d}" fill="${fill}"/>`:''
    }catch(error){
      console.error('ERROR TEXTO:',s,error)
      return ''
    }
  }

  const grupos={}
  for(const r of datos){
    const suc=nombreSucursalJornada(r.sucursal)
    const area=r.area||'sinarea'
    const key=`${suc}|${area}`
    if(!grupos[key])grupos[key]={sucursal:suc,area,datos:[]}
    grupos[key].datos.push(r)
  }

  const ordenSuc=['COYOACAN','BUCARELI']
  const ordenArea=['cocina','barra','salon','sinarea']
  const bloques=Object.values(grupos).sort((a,b)=>{
    const sa=ordenSuc.indexOf(a.sucursal),sb=ordenSuc.indexOf(b.sucursal)
    if(sa!==sb)return (sa<0?99:sa)-(sb<0?99:sb)
    const aa=ordenArea.indexOf(a.area),ab=ordenArea.indexOf(b.area)
    return (aa<0?99:aa)-(ab<0?99:ab)
  })

  const anchoEmpleado=320,anchoDia=170,margen=30,altoCabecera=150,altoFila=55
  const ancho=margen*2+anchoEmpleado+anchoDia*7
  let alto=altoCabecera+30
  for(const b of bloques)alto+=55+altoFila+(b.datos.length*altoFila)+25

  const rects=[`<rect x="0" y="0" width="${ancho}" height="${alto}" fill="#ffffff"/>`]
  const textos=[]

  textos.push(pathTexto('JORNADA SEMANAL',margen,55,30,'#111111'))
  textos.push(pathTexto(rangoSemanaJornada(),margen,92,20,'#555555'))
  textos.push(pathTexto(fechaLaboral(),ancho-margen,55,18,'#777777','end'))

  let y=altoCabecera
  const dias=[['LUN',170],['MAR',170],['MIE',170],['JUE',170],['VIE',170],['SAB',170],['DOM',170]]

  for(const bloque of bloques){
    const tituloArea=bloque.area==='sinarea'?'SIN AREA':nombreAreaJornada(bloque.area)
    const titulo=bloque.sucursal+(tituloArea?` - ${tituloArea}`:'')

    rects.push(`<rect x="${margen}" y="${y}" width="${ancho-2*margen}" height="55" fill="#dddddd"/>`)
    textos.push(pathTexto(titulo,margen+15,y+36,21,'#111111'))
    y+=55

    const columnas=[['EMPLEADO',anchoEmpleado],...dias]
    let x=margen

    for(const [tituloCol,w] of columnas){
      rects.push(`<rect x="${x}" y="${y}" width="${w}" height="${altoFila}" fill="#eeeeee" stroke="#cccccc" stroke-width="1"/>`)
      textos.push(pathTexto(tituloCol,x+w/2,y+35,17,'#111111','middle'))
      x+=w
    }

    y+=altoFila

    for(const r of bloque.datos){
      const valores=[anchoTextoSvg(r.nombre,38),r.lun,r.mar,r.mie,r.jue,r.vie,r.sab,r.dom]
      x=margen
      for(let i=0;i<columnas.length;i++){
        const w=columnas[i][1]
        rects.push(`<rect x="${x}" y="${y}" width="${w}" height="${altoFila}" fill="#ffffff" stroke="#dddddd" stroke-width="1"/>`)
        textos.push(pathTexto(valores[i],x+w/2,y+35,i===0?15:14,'#111111','middle'))
        x+=w
      }
      y+=altoFila
    }
    y+=25
  }

  const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${ancho}" height="${alto}" viewBox="0 0 ${ancho} ${alto}">${rects.join('\n')}${textos.join('\n')}</svg>`
  const buffer=await sharp(Buffer.from(svg)).png().toBuffer()

  const sufijoSucursal=sucursalFiltro?String(sucursalFiltro).replace(/\s+/g,'_'):'Todas'
  const sufijoArea=areaFiltro?`_${normalizaArea(areaFiltro)}`:''
  const sufijoEmpleado=empleadoFiltro?`_${normaliza(empleadoFiltro).replace(/\s+/g,'_')}`:''
  const fileName=`Jornada_Semanal_${sufijoSucursal}${sufijoArea}${sufijoEmpleado}_${fechaLaboral().replace(/\//g,'-')}.png`

  await sock.sendMessage(jid,{
    image:buffer,
    mimetype:'image/png',
    fileName,
    caption:`JORNADA SEMANAL\n${rangoSemanaJornada()}`
  })
}
// =====================================================
// JORNADA SEMANAL - CARGAR HORARIO BASE
// =====================================================

async function jornadaSemanal(
  filtroGrupo,
  jid,
  sock,
  sucursalFiltro='',
  areaFiltro='',
  empleadoFiltro=''
){

  console.log('================================')
  console.log('🚀 INICIANDO JORNADA SEMANAL')
  console.log('GRUPO:',filtroGrupo)
  console.log('SUCURSAL:',sucursalFiltro||'TODAS')
  console.log('AREA:',areaFiltro||'TODAS')
  console.log('EMPLEADO:',empleadoFiltro||'TODOS')
  console.log('================================')

  try{

    const filas=await getRows('Horario_Base!A2:K')

    console.log(
      'FILAS HORARIO_BASE:',
      filas.length
    )

    if(!filas.length){
      await sock.sendMessage(jid,{
        text:'No hay empleados registrados en Horario_Base.'
      })
      return
    }

    await generarImagenJornada(
      filas,
      filtroGrupo,
      jid,
      sock,
      sucursalFiltro,
      areaFiltro,
      empleadoFiltro
    )

  }catch(error){

    console.error(
      'ERROR EN jornadaSemanal:',
      error
    )

    await sock.sendMessage(jid,{
      text:'No se pudo generar la jornada semanal. Revisa los logs de Railway.'
    })
  }
}

// =====================================================
// BUSCAR EMPLEADO
// =====================================================

async function buscarEmpleado(nombreBuscar,filtroGrupo=null){

  const empRows=await getRows('Empleados!A:T')
  const buscar=normaliza(nombreBuscar)
  const candidatos=[]

  for(const r of empRows.slice(1)){

    const corto=r[1]||''
    const completo=r[3]||''
    const suc=r[2]||''

    if(!sucursalCoincideConFiltro(suc,filtroGrupo))continue

    const score=scoreEmpleado(
      corto,
      completo,
      buscar
    )

    if(score>=0)candidatos.push({r,score})
  }

  candidatos.sort((a,b)=>b.score-a.score)

  return candidatos
}

// =====================================================
// DATOS GENERALES DEL EMPLEADO
// =====================================================

async function enviarDatosEmpleado(
  nombreBuscar,
  jid,
  sock,
  filtroGrupo
){

  const candidatos=await buscarEmpleado(
    nombreBuscar,
    filtroGrupo
  )

  if(!candidatos.length){
    await sock.sendMessage(jid,{
      text:'No se encontró empleado en esta sucursal.'
    })
    return
  }

  if(
    candidatos.length>1&&
    candidatos[0].score===candidatos[1].score
  ){

    const opciones=candidatos
      .slice(0,5)
      .map(x=>`• ${nombreEmpleado(x.r)} - ${sucursalEmpleado(x.r)}`)
      .join('\n')

    await sock.sendMessage(jid,{
      text:`Encontré más de un empleado:
${opciones}

Especifica un poco más el nombre.`
    })

    return
  }

  const r=candidatos[0].r

  const status=(r[16]||'')
    .toString()
    .trim()
    .toUpperCase()

  let txt=`📋 *${nombreEmpleado(r)}*`

  if(status&&status!=='ACTIVO'){

    txt+=
`\n📅 Ingreso: ${r[5]||'-'}
\n📤 Baja: ${r[17]||'-'}
\n📝 Motivo de baja: ${r[18]||'-'}
\n🔄 Reingreso: ${r[19]||'-'}`

  }else{

    txt+=
`\n📍 ${r[2]||'-'}
\n📱 ${r[0]||'-'}
\n💼 Puesto: ${r[4]||'-'}
\n🚨 Contacto de emergencia: ${r[6]||'-'}
\n📞 Tel. emergencia: ${r[7]||'-'}`
  }

  await sock.sendMessage(jid,{text:txt})
}

// =====================================================
// DATOS BANCARIOS
// =====================================================

async function enviarDatosBancarios(
  nombreBuscar,
  jid,
  sock,
  filtroGrupo
){

  const candidatos=await buscarEmpleado(
    nombreBuscar,
    filtroGrupo
  )

  if(!candidatos.length){
    await sock.sendMessage(jid,{
      text:'No se encontró empleado en esta sucursal.'
    })
    return
  }

  if(
    candidatos.length>1&&
    candidatos[0].score===candidatos[1].score
  ){

    const opciones=candidatos
      .slice(0,5)
      .map(x=>`• ${nombreEmpleado(x.r)} - ${sucursalEmpleado(x.r)}`)
      .join('\n')

    await sock.sendMessage(jid,{
      text:`Encontré más de un empleado:
${opciones}

Especifica un poco más el nombre.`
    })

    return
  }

  const r=candidatos[0].r

  const banco=(r[12]||'').toString().trim()
  const clabe=(r[13]||'').toString().trim()
  const cuenta=(r[14]||'').toString().trim()
  const tarjeta=(r[15]||'').toString().trim()

  let tipo=''
  let dato=''

  if(clabe){
    tipo='CLABE'
    dato=clabe
  }else if(cuenta){
    tipo='Cuenta'
    dato=cuenta
  }else if(tarjeta){
    tipo='Tarjeta'
    dato=tarjeta
  }

  if(!banco&&!dato){
    await sock.sendMessage(jid,{
      text:`🏦 ${nombreEmpleado(r)}
Sin datos bancarios registrados.`
    })
    return
  }

  let txt=`🏦 *${nombreEmpleado(r)}*
Banco: ${banco||'-'}`

  if(tipo)txt+=`\n${tipo}: ${dato}`
  else txt+='\nSin CLABE, cuenta o tarjeta registrada.'

  await sock.sendMessage(jid,{text:txt})
}

// =====================================================
// RESUMEN EMPLEADO
// =====================================================

export async function resumenEmpleado(
  nombreBuscar,
  jidRespuesta,
  sock,
  tipo='actual'
){

  const sClient=await sheetsClient()

  const [asisRes,baseMap]=await Promise.all([
    sClient.spreadsheets.values.get({
      spreadsheetId:SPREADSHEET_ID,
      range:'Asistencia!A2:M'
    }),
    getHorarioBaseMap()
  ])

  const filas=asisRes.data.values||[]

  const buscar=normaliza(
    nombreBuscar.replace(
      /actual|pasada|pasado|esta semana|hoy/g,
      ''
    )
  )

  const {lunes,domingo,rangoTxt}=getRangoSemana(tipo)

  const info=
    baseMap[buscar]||
    Object.values(baseMap).find(v=>
      normaliza(v.nombreOriginal).includes(buscar)
    )||
    {
      descansos:new Set(),
      horas:{}
    }

  let diasSem=0
  let retSem=0
  let minSem=0
  let horasMin=0
  let extraTotal=0

  const detalle=[]
  const sin=[]

  for(const f of filas){

    const n=normaliza(f[1]||'')
    if(!n.includes(buscar))continue

    const fe=parseFechaMX(f[2])
    if(!fe)continue

    fe.setHours(0,0,0,0)

    if(fe<lunes||fe>domingo)continue

    if(esDescansoRegistro(f)){
      detalle.push(`• ${f[2]}: DESCANSO`)
      continue
    }

    if(!f[3])continue

    diasSem++

    if(esTrabajoDescanso(f)){

      if(f[5]){
        const mt=minutosTrabajados(
          calcularExtra(
            f[3],
            f[5],
            null,
            null
          ).trabajadas
        )

        horasMin+=mt
        extraTotal+=mt
      }

    }else{

      const ret=(f[4]||'').match(/(\d+)\s*min/)

      if(ret){
        retSem++
        minSem+=parseInt(ret[1])
      }

      if(!f[5])sin.push(f[2])

      const progDia=info.horas?.[fe.getDay()]

      if(f[3]&&f[5]&&progDia){

        const calc=calcularExtra(
          f[3],
          f[5],
          progDia.entrada,
          progDia.salida
        )

        horasMin+=minutosTrabajados(calc.trabajadas)
        extraTotal+=calc.extraMin
      }
    }

    detalle.push(
      `• ${f[2]}: Ent ${f[3]} | Sal ${f[5]||'SIN SALIDA'} | Trab ${f[10]||'0'} | Extra ${f[11]||'0'} | ${f[12]||''}`
    )
  }

  let esperados=0

  for(
    let d=new Date(lunes);
    d<=domingo;
    d.setDate(d.getDate()+1)
  ){
    if(!info.descansos.has(d.getDay()))esperados++
  }

  const faltas=Math.max(0,esperados-diasSem)

  const diasNom=[
    'Dom',
    'Lun',
    'Mar',
    'Mie',
    'Jue',
    'Vie',
    'Sab'
  ]

  const descTxt=
    [...info.descansos]
      .map(d=>diasNom[d])
      .join(', ')||
    'ninguno'

  const txt=
`📊 *${buscar.toUpperCase()}*
Descansos: ${descTxt}
${rangoTxt}

*SEMANA ${tipo.toUpperCase()}*

- Trabajados: ${diasSem}/${esperados}
- Faltas: ${faltas}
- Retardos: ${retSem} (${minSem} min)
- Sin salida: ${sin.length}
- Horas Trab: ${formatoHoras(horasMin)}
- Horas Extra: ${formatoHoras(extraTotal)}

*Detalle:*
${detalle.join('\n')||'Sin registros'}`

  await sock.sendMessage(
    jidRespuesta,
    {text:txt}
  )

  if(diasSem>0){
    await generarExcelEmpleado(
      buscar,
      jidRespuesta,
      sock,
      tipo
    )
  }
}

// =====================================================
// EXCEL EMPLEADO
// =====================================================

export async function generarExcelEmpleado(
  nombreBuscarRaw,
  jid,
  sock,
  tipo='actual'
){

  const sClient=await sheetsClient()
  const baseMap=await getHorarioBaseMap()

  const asisRes=await sClient.spreadsheets.values.get({
    spreadsheetId:SPREADSHEET_ID,
    range:'Asistencia!A2:M'
  })

  const filas=asisRes.data.values||[]

  const buscar=normaliza(
    nombreBuscarRaw.replace(
      /actual|pasada|pasado|esta semana|hoy/g,
      ''
    )
  )

  const {lunes,domingo,rangoTxt}=getRangoSemana(tipo)

  const filtradas=filas.filter(f=>{
    if(!normaliza(f[1]||'').includes(buscar))return false
    const fe=parseFechaMX(f[2])
    return fe&&fe>=lunes&&fe<=domingo
  })

  const wb=new ExcelJS.Workbook()
  const ws=wb.addWorksheet('Resumen')

  ws.addRow([
    `REPORTE ${buscar.toUpperCase()} - ${tipo.toUpperCase()}`
  ]).font={bold:true,size:14}

  ws.addRow([rangoTxt])
  ws.addRow([])

  let dias=0
  let ret=0
  let minRet=0
  let sin=0
  let horasMin=0
  let extraMin=0

  for(const f of filtradas){

    if(esDescansoRegistro(f))continue
    if(f[3])dias++

    const m=(f[4]||'').match(/(\d+)\s*min/)

    if(m){
      ret++
      minRet+=parseInt(m[1])
    }

    if(f[3]&&!f[5])sin++

    if(f[3]&&f[5]){

      if(esTrabajoDescanso(f)){

        const mt=minutosTrabajados(
          calcularExtra(
            f[3],
            f[5],
            null,
            null
          ).trabajadas
        )

        horasMin+=mt
        extraMin+=mt

      }else{

        const fe=parseFechaMX(f[2])
        const tel=tel10(f[0])
        const prog=baseMap[tel]?.horas?.[fe?.getDay()]

        const calc=calcularExtra(
          f[3],
          f[5],
          prog?.entrada||null,
          prog?.salida||null
        )

        horasMin+=minutosTrabajados(calc.trabajadas)
        extraMin+=calc.extraMin
      }
    }
  }

  ws.addRow([
    'Nombre',
    'Días',
    'Retardos',
    'Min',
    'Sin salida',
    'Horas Trab',
    'Horas Extra'
  ]).font={bold:true}

  ws.addRow([
    buscar,
    `${dias} días`,
    `Ret ${ret}`,
    minRet,
    sin,
    formatoHoras(horasMin),
    formatoHoras(extraMin)
  ])

  ws.columns.forEach(c=>c.width=22)

  const ws2=wb.addWorksheet('Detalle')

  const header=[
    'Tel',
    'Nombre',
    'Fecha',
    'Entrada',
    'Estatus Entrada',
    'Salida',
    'Suc Entrada',
    'Dist Entr',
    'Suc Salida',
    'Dist Sal',
    'Horas Trabajadas',
    'Horas Extra',
    'Jornada Programada'
  ]

  ws2.addRow(header).font={bold:true}

  for(const f of filtradas){
    ws2.addRow([
      f[0]||'',
      f[1]||'',
      f[2]||'',
      f[3]||'',
      f[4]||'',
      f[5]||'',
      f[6]||'',
      f[7]||'',
      f[8]||'',
      f[9]||'',
      f[10]||'',
      f[11]||'',
      f[12]||''
    ])
  }

  ws2.columns.forEach(c=>c.width=18)

  const fileName=
    `Reporte_${buscar.replace(/\s+/g,'_')}_${tipo}.xlsx`

  const fp=path.join(os.tmpdir(),fileName)

  await wb.xlsx.writeFile(fp)

  await sock.sendMessage(
    jid,
    {
      document:fs.readFileSync(fp),
      mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      fileName
    }
  )

  try{fs.unlinkSync(fp)}catch{}
}

// =====================================================
// REPORTE POR SUCURSAL
// =====================================================

export async function reporteSucursal(
  filtroSucursal,
  jid,
  sock,
  tipo='pasada'
){

  const {asis,emp,base}=await cargarDatos()
  const {lunes,domingo,rangoTxt}=getRangoSemana(tipo)

  const telEmp=mapaEmpleados(emp)
  const telBase=mapaBase(base)
  const datos={}

  for(const f of asis){

    if(esDescansoRegistro(f))continue

    const fe=parseFechaMX(f[2])

    if(!fe||fe<lunes||fe>domingo)continue

    const tel=tel10(f[0])
    const empleado=telEmp[tel]

    const sucAsignada=
      empleado?.[2]||
      telBase[tel]?.[2]||
      ''

    if(!sucursalCoincideAsignada(sucAsignada,filtroSucursal))continue
    if(!f[3])continue

    const nombre=
      empleado?.[3]||
      empleado?.[1]||
      f[1]||
      'Desconocido'

    const key=tel||normaliza(nombre)

    if(!datos[key]){
      datos[key]={
        nombre,
        tel,
        dias:0,
        ret:0,
        min:0,
        sin:0,
        horasMin:0,
        extraMin:0
      }
    }

    const d=datos[key]

    d.dias++

    if(!f[5])d.sin++

    if(esTrabajoDescanso(f)){

      if(f[5]){
        const mt=minutosTrabajados(
          calcularExtra(
            f[3],
            f[5],
            null,
            null
          ).trabajadas
        )

        d.horasMin+=mt
        d.extraMin+=mt
      }

      continue
    }

    const m=(f[4]||'').match(/(\d+)\s*min/)

    if(m){
      d.ret++
      d.min+=parseInt(m[1])
    }

    if(f[3]&&f[5]){

      const horario=
        telBase[tel]
          ?getHorarioDia(
            telBase[tel],
            fe.getDay()
          )
          :''

      const p=parseHorarioRango(horario)

      const calc=calcularExtra(
        f[3],
        f[5],
        p?.entrada||null,
        p?.salida||null
      )

      d.horasMin+=minutosTrabajados(calc.trabajadas)
      d.extraMin+=calc.extraMin
    }
  }

  let txt=
`📊 *REPORTE ${String(filtroSucursal).toUpperCase()}* - ${tipo}
${rangoTxt}

`

  if(!Object.keys(datos).length){

    txt+='Sin registros.'

  }else{

    for(const key of Object.keys(datos)){

      const d=datos[key]

      txt+=
`*${d.nombre}*
Días: ${d.dias}
Retardos: ${d.ret} (${d.min} min)
Trabajadas: ${formatoHoras(d.horasMin)}
Extra: ${formatoHoras(d.extraMin)}
Sin salida: ${d.sin}

`
    }
  }

  await sock.sendMessage(jid,{text:txt})
}

// =====================================================
// EXCEL POR SUCURSAL
// =====================================================

export async function generarExcelSemanaYEnviar(
  filtroSucursal,
  jid,
  sock,
  tipo='pasada'
){

  const {asis,emp,base}=await cargarDatos()
  const {lunes,domingo,rangoTxt}=getRangoSemana(tipo)

  const telEmp=mapaEmpleados(emp)
  const telBase=mapaBase(base)
  const filtradas=[]

  for(const f of asis){

    if(esDescansoRegistro(f))continue

    const fe=parseFechaMX(f[2])

    if(!fe||fe<lunes||fe>domingo)continue

    const tel=tel10(f[0])
    const empleado=telEmp[tel]

    const sucAsignada=
      empleado?.[2]||
      telBase[tel]?.[2]||
      ''

    if(!sucursalCoincideAsignada(sucAsignada,filtroSucursal))continue

    filtradas.push(f)
  }

  const datos={}

  for(const f of filtradas){

    const tel=tel10(f[0])
    const empleado=telEmp[tel]

    const nombre=
      empleado?.[3]||
      empleado?.[1]||
      f[1]||
      'Desconocido'

    const key=tel||normaliza(nombre)

    if(!datos[key]){
      datos[key]={
        nombre,
        tel,
        dias:0,
        ret:0,
        min:0,
        sin:0,
        horasMin:0,
        extraMin:0
      }
    }

    const d=datos[key]

    d.dias++

    if(!f[5])d.sin++

    if(esTrabajoDescanso(f)){

      if(f[5]){
        const mt=minutosTrabajados(
          calcularExtra(
            f[3],
            f[5],
            null,
            null
          ).trabajadas
        )

        d.horasMin+=mt
        d.extraMin+=mt
      }

      continue
    }

    const m=(f[4]||'').match(/(\d+)\s*min/)

    if(m){
      d.ret++
      d.min+=parseInt(m[1])
    }

    if(f[3]&&f[5]){

      const horario=
        telBase[tel]
          ?getHorarioDia(
            telBase[tel],
            parseFechaMX(f[2]).getDay()
          )
          :''

      const p=parseHorarioRango(horario)

      const calc=calcularExtra(
        f[3],
        f[5],
        p?.entrada||null,
        p?.salida||null
      )

      d.horasMin+=minutosTrabajados(calc.trabajadas)
      d.extraMin+=calc.extraMin
    }
  }

  const wb=new ExcelJS.Workbook()
  const ws=wb.addWorksheet('Resumen')

  ws.addRow([
    `REPORTE ${String(filtroSucursal).toUpperCase()} - ${tipo}`
  ]).font={bold:true,size:14}

  ws.addRow([rangoTxt])
  ws.addRow([])

  ws.addRow([
    'Nombre',
    'Tel',
    'Días',
    'Retardos',
    'Min Retardo',
    'Sin salida',
    'Horas Trab',
    'Horas Extra'
  ]).font={bold:true}

  for(const key of Object.keys(datos)){

    const d=datos[key]

    ws.addRow([
      d.nombre,
      d.tel,
      d.dias,
      d.ret,
      d.min,
      d.sin,
      formatoHoras(d.horasMin),
      formatoHoras(d.extraMin)
    ])
  }

  ws.columns.forEach(c=>c.width=22)

  const ws2=wb.addWorksheet('Detalle')

  const header=[
    'Tel',
    'Nombre',
    'Fecha',
    'Entrada',
    'Estatus Entrada',
    'Salida',
    'Suc Entrada',
    'Dist Entr',
    'Suc Salida',
    'Dist Sal',
    'Horas Trabajadas',
    'Horas Extra',
    'Jornada'
  ]

  ws2.addRow(header).font={bold:true}

  for(const f of filtradas){
    ws2.addRow([
      f[0]||'',
      f[1]||'',
      f[2]||'',
      f[3]||'',
      f[4]||'',
      f[5]||'',
      f[6]||'',
      f[7]||'',
      f[8]||'',
      f[9]||'',
      f[10]||'',
      f[11]||'',
      f[12]||''
    ])
  }

  ws2.columns.forEach(c=>c.width=18)

  const fileName=
    `Reporte_${filtroSucursal}_${tipo}_${lunes.toISOString().split('T')[0]}.xlsx`

  const fp=path.join(os.tmpdir(),fileName)

  await wb.xlsx.writeFile(fp)

  await sock.sendMessage(
    jid,
    {
      document:fs.readFileSync(fp),
      mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      fileName
    }
  )

  try{fs.unlinkSync(fp)}catch{}
}

// =====================================================
// FALTAS Y RETARDOS - 45 DÍAS
// =====================================================

async function analizar45Dias(filtroSucursal){

  const {asis,emp,base}=await cargarDatos()

  const telEmp=mapaEmpleados(emp)
  const telBase=mapaBase(base)

  const hasta=new Date(
    new Date().toLocaleString(
      'en-US',
      {timeZone:'America/Mexico_City'}
    )
  )

  hasta.setHours(23,59,59,999)

  const desde=fechaDesdeHoy(DIAS_CRITICOS)

  const faltas={}
  const retardos={}

  for(const r of base){

    const tel=tel10(r[0])
    const suc=r[2]||''

    if(!tel)continue
    if(!sucursalCoincideAsignada(suc,filtroSucursal))continue

    const nombre=r[1]||tel

    for(
      let d=new Date(desde);
      d<=hasta;
      d.setDate(d.getDate()+1)
    ){

      const fecha=d.toISOString().slice(0,10)
      const dia=d.getDay()

      const horario=getHorarioDia(r,dia)
      if(!horario)continue

      if(horario.toLowerCase().includes('descanso'))continue

      const parsed=parseHorarioRango(horario)

      if(!parsed||parsed.entrada==='LIBRE')continue

      const registro=asis.find(a=>
        tel10(a[0])===tel&&
        fechaClave(a[2])===fecha&&
        a[3]
      )

      if(!registro){

        if(!faltas[tel]){
          faltas[tel]={
            nombre,
            sucursal:suc,
            fechas:[]
          }
        }

        faltas[tel].fechas.push(fecha)
        continue
      }

      if(esDescansoRegistro(registro))continue
      if(esTrabajoDescanso(registro))continue

      const entrada=minutos(registro[3])
      const prog=minutos(parsed.entrada)
      const dif=entrada-prog

      if(dif>15){

        if(!retardos[tel]){
          retardos[tel]={
            nombre,
            sucursal:suc,
            fechas:[]
          }
        }

        retardos[tel].fechas.push({
          fecha,
          minutos:dif
        })
      }
    }
  }

  return {faltas,retardos}
}

export async function reporteFaltas(
  filtroSucursal,
  jid,
  sock
){

  const {faltas}=await analizar45Dias(filtroSucursal)

  let txt=
`❌ *FALTAS ${String(filtroSucursal).toUpperCase()}*
Últimos 45 días

`

  let total=0

  for(const tel of Object.keys(faltas)){

    const d=faltas[tel]

    if(!d.fechas.length)continue

    total+=d.fechas.length

    txt+=
`*${d.nombre}*
${d.fechas.map(x=>`• ${x}`).join('\n')}

`
  }

  if(!total)txt+='No se encontraron faltas naturales.'

  await sock.sendMessage(jid,{text:txt})
}

export async function reporteRetardos(
  filtroSucursal,
  jid,
  sock
){

  const {retardos}=await analizar45Dias(filtroSucursal)

  let txt=
`⏰ *RETARDOS ${String(filtroSucursal).toUpperCase()}*
Últimos 45 días

`

  let total=0

  for(const tel of Object.keys(retardos)){

    const d=retardos[tel]

    if(!d.fechas.length)continue

    total+=d.fechas.length

    txt+=
`*${d.nombre}*
${d.fechas.map(x=>`• ${x.fecha} - ${x.minutos} min`).join('\n')}

`
  }

  if(!total)txt+='No se encontraron retardos.'

  await sock.sendMessage(jid,{text:txt})
}

// =====================================================
// CRÍTICOS / GRAVES
// 3 RETARDOS = 1 FALTA
// =====================================================

export async function reporteCriticos(
  filtroSucursal,
  jid,
  sock
){

  const {faltas,retardos}=await analizar45Dias(filtroSucursal)
  const nombres={}

  for(const tel of Object.keys(faltas)){

    nombres[tel]={
      nombre:faltas[tel].nombre,
      sucursal:faltas[tel].sucursal,
      faltas:faltas[tel].fechas.length,
      retardos:retardos[tel]?.fechas.length||0
    }
  }

  for(const tel of Object.keys(retardos)){

    if(!nombres[tel]){
      nombres[tel]={
        nombre:retardos[tel].nombre,
        sucursal:retardos[tel].sucursal,
        faltas:0,
        retardos:retardos[tel].fechas.length
      }
    }
  }

  let txt=
`🚨 *CRÍTICOS / GRAVES ${String(filtroSucursal).toUpperCase()}*
Últimos 45 días

`

  let total=0

  for(const tel of Object.keys(nombres)){

    const d=nombres[tel]

    const equivalentes=
      d.faltas+
      Math.floor(d.retardos/3)

    if(equivalentes<3)continue

    total++

    txt+=
`*${d.nombre}*
Faltas: ${d.faltas}
Retardos: ${d.retardos}
Equivalentes: ${equivalentes}

`
  }

  if(!total){
    txt+='No hay empleados con 3 o más faltas equivalentes.'
  }

  await sock.sendMessage(jid,{text:txt})
}

// =====================================================
// ROUTER PRINCIPAL
// =====================================================

export async function handleReportes({
  texto,
  jid,
  sock,
  filtroGrupo,
  usuario
}){

  const low=normaliza(texto)

  if(low==='lista de reportes'){

    await sock.sendMessage(jid,{
      text:textoListaReportes(filtroGrupo)
    })

    return true
  }

  if(low==='reportes'||low==='reporte'){
    return false
  }

  // ================================================
  // JORNADA SEMANAL
  // ================================================

  if(
    low==='jornada semanal'||
    low==='jornada'||
    low.startsWith('jornada semanal ')||
    low.startsWith('jornada ')
  ){

    const filtros=obtenerFiltrosJornada(
      texto
        .replace(/^jornada\s+semanal\s*/i,'')
        .replace(/^jornada\s*/i,'')
        .trim(),
      filtroGrupo
    )

    await jornadaSemanal(
      filtroGrupo,
      jid,
      sock,
      filtros.sucursal,
      filtros.area,
      filtros.empleado
    )

    return true
  }

  // ================================================
  // ASISTENCIA HOY
  // ================================================
if(low.startsWith('asistencia hoy')){
  let resto=low.replace('asistencia hoy','').trim()
  let area=''
  if(resto.includes('cocina'))area='cocina'
  else if(resto.includes('barra'))area='barra'
  else if(resto.includes('salon')||resto.includes('salón'))area='salon'
  let suc=''
  if(resto.includes('bucareli'))suc='bucareli'
  else if(resto.includes('coyoacan')||resto.includes('coyoacán')||resto.includes('hotel')||resto.includes('trinidad'))suc='coyoacan'
  const esReportes=!filtroGrupo
  if(esReportes&&!suc&&!area)suc='todas'
  if(!esReportes&&!suc)suc=filtroGrupo||'coyoacan'
  if(esReportes&&area&&!suc){
    const mensaje=iniciarLista({
      jid,
      usuario,
      contexto:{
        proceso:'ASISTENCIA_HOY',
        area
      },
      opciones:{
        1:{texto:'Coyoacán',valor:'coyoacan'},
        2:{texto:'Bucareli',valor:'bucareli'}
      },
      mensaje:'🏢 ¿De qué sucursal quieres la asistencia de hoy?'
    })
    await sock.sendMessage(jid,{text:mensaje})
    return true
  }
  await asistenciaHoy(
    suc,
    jid,
    sock,
    area
  )
  return true
}
  // ================================================
  // FALTAS
  // ================================================
  if(low.startsWith('faltas')){

    let suc=low
      .replace('faltas','')
      .trim()

    if(!suc)suc=filtroGrupo||'coyoacan'

    await reporteFaltas(
      suc,
      jid,
      sock
    )

    return true
  }

  // ================================================
  // RETARDOS
  // ================================================

  if(low.startsWith('retardos')){

    let suc=low
      .replace('retardos','')
      .trim()

    if(!suc)suc=filtroGrupo||'coyoacan'

    await reporteRetardos(
      suc,
      jid,
      sock
    )

    return true
  }

  // ================================================
  // CRÍTICOS / GRAVES
  // ================================================

  if(
    low.startsWith('criticos')||
    low.startsWith('graves')
  ){

    let suc=low
      .replace(/^criticos/,'')
      .replace(/^graves/,'')
      .trim()

    if(!suc)suc=filtroGrupo||'coyoacan'

    await reporteCriticos(
      suc,
      jid,
      sock
    )

    return true
  }

  // ================================================
  // RESUMEN
  // ================================================

  if(low.startsWith('resumen ')){

    let tipo='actual'

    if(
      low.includes('pasada')||
      low.includes('pasado')
    ){
      tipo='pasada'
    }

    let limpio=texto
      .slice(8)
      .toLowerCase()
      .trim()
      .replace(
        /pasada|pasado|actual|esta semana|hoy/g,
        ''
      )
      .trim()

    if(
      limpio.includes('bucareli')||
      limpio.includes('coyo')||
      limpio.includes('hotel')
    ){

      let suc='coyoacan'

      if(limpio.includes('bucareli')){
        suc='bucareli'
      }

      await reporteSucursal(
        suc,
        jid,
        sock,
        tipo
      )

      await generarExcelSemanaYEnviar(
        suc,
        jid,
        sock,
        tipo
      )

      return true
    }

    if(!limpio){

      await sock.sendMessage(jid,{
        text:'Escribe: resumen [nombre] pasada o actual'
      })

      return true
    }

    await resumenEmpleado(
      limpio,
      jid,
      sock,
      tipo
    )

    return true
  }

  // ================================================
  // REPORTES
  // ================================================

  if(
    /^reporte\s+.+/i.test(texto)||
    /^checador\s+.+/i.test(texto)||
    /^reporte x unidad\s+.+/i.test(texto)
  ){

    const tipo=
      low.includes('pasada')||
      low.includes('pasado')
        ?'pasada'
        :low.includes('actual')||
         low.includes('esta')||
         low.includes('hoy')
          ?'actual'
          :'pasada'

    let suc='coyoacan'

    if(low.includes('bucareli')){
      suc='bucareli'
    }else if(low.includes('hotel')){
      suc='hotel'
    }else if(low.includes('coyo')){
      suc='coyoacan'
    }

    if(
      low.includes('x unidad')||
      low.includes('por unidad')
    ){

      if(!filtroGrupo){

        await reporteSucursal(
          'coyoacan',
          jid,
          sock,
          tipo
        )

        await generarExcelSemanaYEnviar(
          'coyoacan',
          jid,
          sock,
          tipo
        )

        await reporteSucursal(
          'bucareli',
          jid,
          sock,
          tipo
        )

        await generarExcelSemanaYEnviar(
          'bucareli',
          jid,
          sock,
          tipo
        )

      }else{

        await reporteSucursal(
          filtroGrupo,
          jid,
          sock,
          tipo
        )

        await generarExcelSemanaYEnviar(
          filtroGrupo,
          jid,
          sock,
          tipo
        )
      }

      return true
    }

    await reporteSucursal(
      suc,
      jid,
      sock,
      tipo
    )

    await generarExcelSemanaYEnviar(
      suc,
      jid,
      sock,
      tipo
    )

    return true
  }

  // ================================================
  // DATOS BANCARIOS
  // ================================================

  if(
    /^(cuenta bancaria|cuenta empleado|datos bancarios|banco|clave|clabe|cuenta)\b/i.test(texto)
  ){

    let buscar=texto
      .replace(
        /^(cuenta bancaria|cuenta empleado|datos bancarios|banco|clave|clabe|cuenta)\s*/i,
        ''
      )
      .trim()

    if(!buscar){

      await sock.sendMessage(jid,{
        text:'Escribe por ejemplo: banco Daniel'
      })

      return true
    }

    await enviarDatosBancarios(
      buscar,
      jid,
      sock,
      filtroGrupo
    )

    return true
  }

  // ================================================
  // DATOS GENERALES
  // ================================================

  if(
    /^(numero|número|num|tel|telefono|teléfono|info|ficha|datos|dato)\s*(de)?/i.test(texto)
  ){

    let buscar=texto
      .replace(
        /^(numero|número|num|tel|telefono|teléfono|info|ficha|datos|dato)\s*(de)?\s*/i,
        ''
      )
      .trim()

    if(!buscar){

      await sock.sendMessage(jid,{
        text:'Escribe: datos Daniel'
      })

      return true
    }

    await enviarDatosEmpleado(
      buscar,
      jid,
      sock,
      filtroGrupo
    )

    return true
  }

  return false
}

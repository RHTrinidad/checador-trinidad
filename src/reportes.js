import ExcelJS from 'exceljs'
import fs from 'fs'
import path from 'path'
import os from 'os'
import { SPREADSHEET_ID } from './config.js'
import { getRows, sheetsClient, getHorarioBaseMap } from './sheets.js'
import {
  fechaLaboral,
  horaMX,
  minutos,
  parseFechaMX,
  normaliza,
  parseHorarioRango,
  getRangoSemana,
  sucursalCoincideConFiltro,
  scoreEmpleado,
  calcularExtra
} from './utils.js'

const DIAS_LABORALES=[1,2,3,4,5]

function esSinSalida(f){
  return String(f?.[8]||'').toUpperCase().includes('SIN SALIDA') ||
         String(f?.[9]||'').toUpperCase().includes('AUTO 16H')
}

function minutosTrabajados(v){
  if(!v) return 0
  if(v==='8') return 480
  const m=String(v).match(/^(\d+):(\d+)/)
  if(!m) return 0
  return parseInt(m[1])*60+parseInt(m[2])
}

function fechaClave(d){
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`
}

function inicio45Dias(){
  const ahora=new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'}))
  ahora.setHours(0,0,0,0)
  const inicio=new Date(ahora)
  inicio.setDate(inicio.getDate()-44)
  inicio.setHours(0,0,0,0)
  return {inicio,fin:ahora}
}

function inicioMesActual(){
  const ahora=new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'}))
  return new Date(ahora.getFullYear(),ahora.getMonth(),1,0,0,0,0)
}

function nombreDia(n){
  return ['Dom','Lun','Mar','Mie','Jue','Vie','Sab'][n]||''
}

function normalizaSucursal(s){
  const x=normaliza(s)
  if(x.includes('coyo')||x.includes('hotel')||x.includes('trinidad')) return 'coyoacan'
  if(x.includes('bucareli')||x.includes('juarez')) return 'bucareli'
  return x
}

function empleadoActivo(r){
  const status=normaliza(r?.[16]||'')
  return !status || status.includes('activo')
}

function empleadoPerteneceSucursal(r,suc){
  if(!suc) return true
  const x=normalizaSucursal(r?.[2]||'')
  if(suc==='coyoacan') return x==='coyoacan'
  if(suc==='bucareli') return x==='bucareli'
  return x===normalizaSucursal(suc)
}

function obtenerJornada(base,tel,fecha){
  const info=base[tel]||null
  if(!info) return null
  const h=info.horas?.[fecha.getDay()]
  return h||null
}

function esDiaProgramado(base,tel,fecha){
  const h=obtenerJornada(base,tel,fecha)
  if(!h) return false
  if(h.entrada==='LIBRE') return false
  if(!h.entrada) return false
  return DIAS_LABORALES.includes(fecha.getDay())
}

function estaEnRango(fe,inicio,fin){
  if(!fe) return false
  const d=new Date(fe)
  d.setHours(0,0,0,0)
  return d>=inicio&&d<=fin
}

function agruparAsistencia(filas){
  const map=new Map()
  for(const f of filas){
    const tel=(f[0]||'').replace(/\D/g,'').slice(-10)
    const fecha=f[2]||''
    if(!tel||!fecha) continue
    map.set(`${tel}|${fecha}`,f)
  }
  return map
}

function buscarEmpleados(empRows,buscar,filtroGrupo){
  const b=normaliza(buscar)
  const candidatos=[]

  for(const r of empRows.slice(1)){
    if(!empleadoPerteneceSucursal(r,filtroGrupo)) continue
    const corto=r[1]||''
    const completo=r[3]||''
    const score=scoreEmpleado(corto,completo,b)

    if(score>=0){
      candidatos.push({r,score})
    }
  }

  candidatos.sort((a,b)=>b.score-a.score)
  return candidatos
}

async function obtenerDatos(){
  const sClient=await sheetsClient()
  const [asisRes,empRows,base]=await Promise.all([
    sClient.spreadsheets.values.get({
      spreadsheetId:SPREADSHEET_ID,
      range:'Asistencia!A2:M'
    }),
    getRows('Empleados!A:T'),
    getHorarioBaseMap()
  ])

  return {
    filas:asisRes.data.values||[],
    empRows:empRows||[],
    base:base||{}
  }
}

export async function asistenciaHoy(filtroSucursal,jid,sock){
  const sClient=await sheetsClient()

  const [baseRows,asisRows]=await Promise.all([
    getRows('Horario_Base!A2:K'),
    sClient.spreadsheets.values.get({
      spreadsheetId:SPREADSHEET_ID,
      range:'Asistencia!A2:M'
    }).then(r=>r.data.values||[])
  ])

  const fLab=fechaLaboral()
  const ahoraMin=minutos(horaMX())
  const ahoraMXDate=new Date(
    new Date().toLocaleString(
      'en-US',
      {timeZone:'America/Mexico_City'}
    )
  )
  const diaNum=ahoraMXDate.getDay()

  const filtro=filtroSucursal.toLowerCase()
  const esCoyo=filtro.includes('coyo')||filtro.includes('hotel')
  const esBuc=filtro.includes('bucareli')||filtro.includes('juarez')

  let llego=[]
  let retardo=[]
  let falta=[]
  let futuro=[]

  for(const r of baseRows){
    const suc=(r[2]||'').toLowerCase()
    let inc=false

    if(esCoyo){
      inc=suc.includes('coyo')||suc.includes('hotel')||suc.includes('trinidad')
    }else if(esBuc){
      inc=suc.includes('bucareli')||suc.includes('juarez')
    }else{
      inc=suc.includes(filtro)
    }

    if(!inc) continue

    const nombre=r[1]||''
    const tel=(r[0]||'').replace(/\D/g,'').slice(-10)

    const mapa={
      1:r[3],
      2:r[4],
      3:r[5],
      4:r[6],
      5:r[7],
      6:r[8],
      0:r[9]
    }

    const v=(mapa[diaNum]||'').toString().trim()
    const parsed=parseHorarioRango(v)

    if(!parsed) continue

    const esLibre=parsed.entrada==='LIBRE'
    const horaProg=esLibre?'LIBRE':parsed.entrada

    const registro=asisRows.find(a=>
      (a[0]||'').replace(/\D/g,'').slice(-10)===tel &&
      a[2]===fLab
    )

    if(registro&&registro[3]){
      const entrada=registro[3]
      const sucEnt=registro[6]||''
      const dif=esLibre?0:minutos(entrada)-minutos(horaProg)

      if(esLibre){
        llego.push(
          `• ${nombre} - Entró ${entrada} en ${sucEnt} ✅`
        )
      }else if(dif>15){
        retardo.push(
          `• ${nombre} - [${r[2]||''}] Prog ${horaProg} - Entró ${entrada} - ⏰ ${dif}m tarde - ${sucEnt}`
        )
      }else{
        llego.push(
          `• ${nombre} - [${r[2]||''}] Prog ${horaProg} - Entró ${entrada} ✅ - ${sucEnt}`
        )
      }
    }else{
      if(esLibre) continue

      const dif=ahoraMin-minutos(horaProg)

      if(dif<0){
        futuro.push(
          `• ${nombre} - [${r[2]||''}] - Prog ${horaProg}`
        )
      }else{
        falta.push(
          `• ${nombre} - [${r[2]||''}] - Prog ${horaProg} - ❌ ${dif}m sin llegar`
        )
      }
    }
  }

  const txt=
    `📍 *ASISTENCIA HOY ${fLab} - ${filtroSucursal.toUpperCase()}* ${horaMX()}\n\n`+
    `✅ *A TIEMPO (${llego.length}):*\n${llego.join('\n')||'-'}\n\n`+
    `⏰ *RETARDOS (${retardo.length}):*\n${retardo.join('\n')||'-'}\n\n`+
    `❌ *NO HAN LLEGADO (${falta.length}):*\n${falta.join('\n')||'Todos llegaron'}\n\n`+
    `⏳ *PRÓXIMOS (${futuro.length}):*\n${futuro.join('\n')||'-'}`

  await sock.sendMessage(jid,{text:txt})
}

export async function generarExcelEmpleado(nombreBuscarRaw,jid,sock,tipo='actual'){
  const {filas,base}=await obtenerDatos()

  const buscar=normaliza(nombreBuscarRaw)
    .replace(/actual|pasada|pasado|esta semana|hoy/g,'')
    .trim()

  const {lunes,domingo,rangoTxt}=getRangoSemana(tipo)

  const filtradas=filas.filter(f=>{
    const n=normaliza(f[1]||'')
    if(!n.includes(buscar)) return false
    const fe=parseFechaMX(f[2])
    return fe&&fe>=lunes&&fe<=domingo
  })

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

  const wb=new ExcelJS.Workbook()
  const ws=wb.addWorksheet('Resumen')

  ws.addRow([
    `REPORTE ${buscar.toUpperCase()} - ${tipo.toUpperCase()}`
  ]).font={bold:true,size:14}

  ws.addRow([rangoTxt])
  ws.addRow([])

  let dias=0
  let ret=0
  let min=0
  let sin=0
  let horasMin=0
  let extraMin=0

  for(const f of filtradas){
    if(f[3]) dias++

    const m=(f[4]||'').match(/(\d+)\s*min/)
    if(m){
      ret++
      min+=parseInt(m[1])
    }

    if(esSinSalida(f)) sin++

    if(f[3]&&f[5]&&!esSinSalida(f)){
      const fe=parseFechaMX(f[2])
      const prog=base[
        (f[0]||'').replace(/\D/g,'').slice(-10)
      ]?.horas?.[fe?.getDay()]

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
    `Ret ${ret} (${min}m)`,
    sin,
    `${Math.floor(horasMin/60)}:${String(horasMin%60).padStart(2,'0')}h`,
    `${Math.floor(extraMin/60)}:${String(extraMin%60).padStart(2,'0')}h`
  ])

  ws.columns.forEach(c=>c.width=22)

  const ws2=wb.addWorksheet('Detalle')
  ws2.addRow(header).font={bold:true}

  for(const f of filtradas){
    ws2.addRow(f)
  }

  ws2.columns.forEach(c=>c.width=18)

  const fileName=
    `Reporte_${buscar.replace(/\s+/g,'_')}_${tipo}.xlsx`

  const fp=path.join(os.tmpdir(),fileName)

  await wb.xlsx.writeFile(fp)

  await sock.sendMessage(jid,{
    document:fs.readFileSync(fp),
    mimetype:
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    fileName
  })

  fs.unlinkSync(fp)
}

export async function resumenEmpleado(
  nombreBuscar,
  jidRespuesta,
  sock,
  tipo='actual'
){
  const {filas,empRows,base}=await obtenerDatos()

  let buscar=normaliza(nombreBuscar)
    .replace(/actual|pasada|pasado|esta semana|hoy/g,'')
    .trim()

  const candidatos=buscarEmpleados(empRows,buscar,null)

  let empleado=null

  if(candidatos.length){
    empleado=candidatos[0].r
  }

  const {lunes,domingo,rangoTxt}=getRangoSemana(tipo)

  let filasEmpleado=filas.filter(f=>{
    const n=normaliza(f[1]||'')
    const tel=(f[0]||'').replace(/\D/g,'').slice(-10)

    if(empleado){
      const telEmp=(empleado[0]||'').replace(/\D/g,'').slice(-10)
      if(telEmp) return tel===telEmp
    }

    return n.includes(buscar)
  })

  const info=empleado
    ? base[(empleado[0]||'').replace(/\D/g,'').slice(-10)]
    : Object.values(base).find(v=>
        normaliza(v?.nombreOriginal||'').includes(buscar)
      )

  const descansos=info?.descansos||new Set()
  const horas=info?.horas||{}

  let diasSem=0
  let retSem=0
  let minSem=0
  let horasMin=0
  let extraTotal=0
  let sin=[]
  let detalle=[]

  for(const f of filasEmpleado){
    const fe=parseFechaMX(f[2])

    if(!fe||fe<lunes||fe>domingo) continue

    if(f[3]){
      diasSem++

      const ret=(f[4]||'').match(/(\d+)\s*min/)

      if(ret){
        retSem++
        minSem+=parseInt(ret[1])
      }

      if(esSinSalida(f)){
        sin.push(f[2])
      }

      const prog=horas[fe.getDay()]

      if(f[3]&&f[5]&&!esSinSalida(f)&&prog){
        const calc=calcularExtra(
          f[3],
          f[5],
          prog.entrada,
          prog.salida
        )

        horasMin+=minutosTrabajados(calc.trabajadas)
        extraTotal+=calc.extraMin
      }

      detalle.push(
        `• ${f[2]}: Ent ${f[3]} | Sal ${esSinSalida(f)?'SIN SALIDA':(f[5]||'SIN')} | Trab ${f[10]||'0'} | Extra ${f[11]||'0'} | ${f[12]||''}`
      )
    }
  }

  let esperados=0

  for(
    let d=new Date(lunes);
    d<=domingo;
    d.setDate(d.getDate()+1)
  ){
    if(!descansos.has(d.getDay())&&DIAS_LABORALES.includes(d.getDay())){
      esperados++
    }
  }

  const faltasSemana=Math.max(0,esperados-diasSem)

  // Faltas del mes actual
  const inicioMes=inicioMesActual()
  const hoyMX=new Date(
    new Date().toLocaleString(
      'en-US',
      {timeZone:'America/Mexico_City'}
    )
  )
  hoyMX.setHours(23,59,59,999)

  let faltasMes=0

  const telEmpleado=empleado
    ? (empleado[0]||'').replace(/\D/g,'').slice(-10)
    : ''

  if(telEmpleado){
    const registros=new Set()

    for(const f of filas){
      const tel=(f[0]||'').replace(/\D/g,'').slice(-10)
      const fe=parseFechaMX(f[2])

      if(tel!==telEmpleado||!fe) continue
      if(!estaEnRango(fe,inicioMes,hoyMX)) continue

      if(f[3]) registros.add(fechaClave(fe))
    }

    for(
      let d=new Date(inicioMes);
      d<=hoyMX;
      d.setDate(d.getDate()+1)
    ){
      if(!DIAS_LABORALES.includes(d.getDay())) continue
      if(descansos.has(d.getDay())) continue

      if(!registros.has(fechaClave(d))){
        faltasMes++
      }
    }
  }

  const nombreMostrar=empleado
    ? (empleado[3]||empleado[1])
    : buscar

  const descTxt=
    [...descansos].map(nombreDia).join(', ')||'ninguno'

  const txt=
    `📊 *${String(nombreMostrar).toUpperCase()}*\n`+
    `${rangoTxt}\n\n`+
    `*SEMANA ${tipo.toUpperCase()}*\n`+
    `- Trabajados: ${diasSem}/${esperados}\n`+
    `- Faltas: ${faltasSemana}\n`+
    `- Faltas en el mes: ${faltasMes}\n`+
    `- Retardos: ${retSem} (${minSem} min)\n`+
    `- Sin salida: ${sin.length}\n`+
    `- Horas Trab: ${Math.floor(horasMin/60)}:${String(horasMin%60).padStart(2,'0')}h\n`+
    `- Horas Extra: ${Math.floor(extraTotal/60)}:${String(extraTotal%60).padStart(2,'0')}h\n`+
    `- Descansos: ${descTxt}\n\n`+
    `*DÍAS SIN SALIDA:*\n`+
    `${sin.map(x=>`• ${x}`).join('\n')||'-'}\n\n`+
    `*DETALLE:*\n`+
    `${detalle.join('\n')||'Sin registros'}`

  await sock.sendMessage(jidRespuesta,{text:txt})

  if(diasSem>0){
    await generarExcelEmpleado(buscar,jidRespuesta,sock,tipo)
  }
}

async function obtenerFaltas45(sucursal){
  const {filas,empRows,base}=await obtenerDatos()
  const {inicio,fin}=inicio45Dias()

  const asistencia=agruparAsistencia(filas)
  const resultado={}

  for(const r of empRows.slice(1)){
    if(!empleadoActivo(r)) continue

    const tel=(r[0]||'').replace(/\D/g,'').slice(-10)
    const nombre=r[3]||r[1]||tel

    if(!tel) continue
    if(!empleadoPerteneceSucursal(r,sucursal)) continue

    const faltas=[]

    for(
      let d=new Date(inicio);
      d<=fin;
      d.setDate(d.getDate()+1)
    ){
      if(!DIAS_LABORALES.includes(d.getDay())) continue
      if(!esDiaProgramado(base,tel,d)) continue

      const fecha=fechaClave(d)
      const reg=asistencia.get(`${tel}|${fecha}`)

      if(!reg||!reg[3]){
        faltas.push(fecha)
      }
    }

    if(faltas.length){
      resultado[tel]={
        nombre,
        tel,
        faltas
      }
    }
  }

  return resultado
}

async function obtenerRetardos45(sucursal){
  const {filas,empRows}=await obtenerDatos()
  const {inicio,fin}=inicio45Dias()

  const resultado={}

  for(const r of empRows.slice(1)){
    if(!empleadoActivo(r)) continue

    const tel=(r[0]||'').replace(/\D/g,'').slice(-10)
    const nombre=r[3]||r[1]||tel

    if(!tel) continue
    if(!empleadoPerteneceSucursal(r,sucursal)) continue

    const retardos=[]

    for(const f of filas){
      const telF=(f[0]||'').replace(/\D/g,'').slice(-10)
      if(telF!==tel) continue

      const fe=parseFechaMX(f[2])
      if(!estaEnRango(fe,inicio,fin)) continue

      const m=(f[4]||'').match(/(\d+)\s*min/)

      if(m){
        retardos.push({
          fecha:f[2],
          minutos:parseInt(m[1])
        })
      }
    }

    if(retardos.length){
      resultado[tel]={
        nombre,
        tel,
        retardos
      }
    }
  }

  return resultado
}

async function enviarFaltas45(sucursal,jid,sock){
  const datos=await obtenerFaltas45(sucursal)
  const {inicio,fin}=inicio45Dias()

  let txt=
    `❌ *FALTAS ${sucursal.toUpperCase()}*\n`+
    `${fechaClave(inicio)} al ${fechaClave(fin)}\n\n`

  const keys=Object.keys(datos)

  if(!keys.length){
    txt+='No hay faltas en el periodo.'
  }else{
    for(const k of keys){
      const d=datos[k]
      txt+=
        `*${d.nombre}* - ${d.faltas.length} falta(s)\n`+
        d.faltas.map(x=>`• ${x}`).join('\n')+
        `\n\n`
    }
  }

  await sock.sendMessage(jid,{text:txt})
}

async function enviarRetardos45(sucursal,jid,sock){
  const datos=await obtenerRetardos45(sucursal)
  const {inicio,fin}=inicio45Dias()

  let txt=
    `⏰ *RETARDOS ${sucursal.toUpperCase()}*\n`+
    `${fechaClave(inicio)} al ${fechaClave(fin)}\n\n`

  const keys=Object.keys(datos)

  if(!keys.length){
    txt+='No hay retardos en el periodo.'
  }else{
    for(const k of keys){
      const d=datos[k]
      const total=d.retardos.reduce((a,b)=>a+b.minutos,0)

      txt+=
        `*${d.nombre}* - ${d.retardos.length} retardo(s) - ${total} min\n`+
        d.retardos
          .map(x=>`• ${x.fecha} - ${x.minutos} min`)
          .join('\n')+
        `\n\n`
    }
  }

  await sock.sendMessage(jid,{text:txt})
}

async function enviarCriticos(jid,sock){
  const faltasCoyo=await obtenerFaltas45('coyoacan')
  const faltasBuc=await obtenerFaltas45('bucareli')
  const retCoyo=await obtenerRetardos45('coyoacan')
  const retBuc=await obtenerRetardos45('bucareli')

  const mapa=new Map()

  function agregarFaltas(datos,suc){
    for(const tel of Object.keys(datos)){
      if(!mapa.has(tel)){
        mapa.set(tel,{
          nombre:datos[tel].nombre,
          suc,
          faltas:[],
          retardos:[]
        })
      }

      mapa.get(tel).faltas.push(...datos[tel].faltas)
    }
  }

  function agregarRetardos(datos,suc){
    for(const tel of Object.keys(datos)){
      if(!mapa.has(tel)){
        mapa.set(tel,{
          nombre:datos[tel].nombre,
          suc,
          faltas:[],
          retardos:[]
        })
      }

      mapa.get(tel).retardos.push(...datos[tel].retardos)
    }
  }

  agregarFaltas(faltasCoyo,'Coyoacán')
  agregarFaltas(faltasBuc,'Bucareli')
  agregarRetardos(retCoyo,'Coyoacán')
  agregarRetardos(retBuc,'Bucareli')

  let txt='🚨 *CRÍTICOS / GRAVES - 45 DÍAS*\n\n'
  let encontrados=0

  for(const d of mapa.values()){
    const faltasNaturales=d.faltas.length
    const retardos=d.retardos.length
    const faltasPorRetardos=Math.floor(retardos/3)
    const faltasEquivalentes=faltasNaturales+faltasPorRetardos

    if(faltasEquivalentes<3) continue

    encontrados++

    txt+=
      `*${d.nombre}* - ${d.suc}\n`+
      `Faltas naturales: ${faltasNaturales}\n`+
      `Retardos: ${retardos}\n`+
      `Retardos equivalentes: ${faltasPorRetardos} falta(s)\n`+
      `Total equivalente: ${faltasEquivalentes} falta(s)\n\n`+

      `Faltas:\n`+
      `${d.faltas.map(x=>`• ${x}`).join('\n')||'-'}\n`+

      `Retardos:\n`+
      `${d.retardos.map(x=>`• ${x.fecha} - ${x.minutos} min`).join('\n')||'-'}\n\n`
  }

  if(!encontrados){
    txt+='No hay empleados con 3 faltas equivalentes o más.'
  }

  await sock.sendMessage(jid,{text:txt})
}

export async function reporteSucursal(
  filtroSucursal,
  jid,
  sock,
  tipo='pasada'
){
  const [asisRes,empRows]=await Promise.all([
    (await sheetsClient()).spreadsheets.values.get({
      spreadsheetId:SPREADSHEET_ID,
      range:'Asistencia!A2:M'
    }),
    getRows('Empleados!A:K')
  ])

  const filas=asisRes.data.values||[]
  const {lunes,domingo,rangoTxt}=getRangoSemana(tipo)

  const suc=normalizaSucursal(filtroSucursal)
  const base=await getHorarioBaseMap()

  const telToNombre={}

  empRows.forEach(r=>{
    const tel=(r[0]||'').replace(/\D/g,'').slice(-10)
    if(tel){
      telToNombre[tel]=(r[3]||r[1]||'').trim()
    }
  })

  const datos={}

  for(const f of filas){
    const fe=parseFechaMX(f[2])
    if(!fe||fe<lunes||fe>domingo) continue

    const sucRegistro=normalizaSucursal(
      `${f[6]||''} ${f[8]||''}`
    )

    if(sucRegistro!==suc) continue

    const tel=(f[0]||'').replace(/\D/g,'').slice(-10)
    const nombre=
      telToNombre[tel]||
      (f[1]||'').trim()||
      'Desconocido'

    const key=tel||normaliza(nombre)

    if(!datos[key]){
      datos[key]={
        nombre,
        dias:0,
        ret:0,
        min:0,
        sin:0,
        horasMin:0,
        extraMin:0
      }
    }

    if(f[3]){
      datos[key].dias++

      const m=(f[4]||'').match(/(\d+)\s*min/)

      if(m){
        datos[key].ret++
        datos[key].min+=parseInt(m[1])
      }

      if(esSinSalida(f)){
        datos[key].sin++
      }

      if(f[3]&&f[5]&&!esSinSalida(f)){
        const prog=base[tel]?.horas?.[fe.getDay()]
        const calc=calcularExtra(
          f[3],
          f[5],
          prog?.entrada||null,
          prog?.salida||null
        )

        datos[key].horasMin+=minutosTrabajados(calc.trabajadas)
        datos[key].extraMin+=calc.extraMin
      }
    }
  }

  let txt=
    `📊 *REPORTE ${filtroSucursal.toUpperCase()}* - ${tipo}\n`+
    `${rangoTxt}\n\n`

  const keys=Object.keys(datos)

  if(!keys.length){
    txt+='Sin registros.'
  }else{
    for(const k of keys){
      const d=datos[k]

      txt+=
        `*${d.nombre}*\n`+
        `Días: ${d.dias}\n`+
        `Retardos: ${d.ret} (${d.min}m)\n`+
        `Trab: ${Math.floor(d.horasMin/60)}:${String(d.horasMin%60).padStart(2,'0')}h\n`+
        `Extra: ${Math.floor(d.extraMin/60)}:${String(d.extraMin%60).padStart(2,'0')}h\n`+
        `Sin salida: ${d.sin}\n\n`
    }
  }

  await sock.sendMessage(jid,{text:txt})
}

export async function generarExcelSemanaYEnviar(
  filtroSucursal,
  jid,
  sock,
  tipo='pasada'
){
  const sClient=await sheetsClient()
  const base=await getHorarioBaseMap()

  const [asisRes,empRows]=await Promise.all([
    sClient.spreadsheets.values.get({
      spreadsheetId:SPREADSHEET_ID,
      range:'Asistencia!A2:M'
    }),
    getRows('Empleados!A:K')
  ])

  const filas=asisRes.data.values||[]
  const {lunes,domingo,rangoTxt}=getRangoSemana(tipo)
  const suc=normalizaSucursal(filtroSucursal)

  const telToNombre={}

  empRows.forEach(r=>{
    const tel=(r[0]||'').replace(/\D/g,'').slice(-10)

    if(tel){
      telToNombre[tel]=(r[3]||r[1]||'').trim()
    }
  })

  const filtradas=filas.filter(f=>{
    const fe=parseFechaMX(f[2])
    if(!fe||fe<lunes||fe>domingo) return false

    const sucRegistro=normalizaSucursal(
      `${f[6]||''} ${f[8]||''}`
    )

    return sucRegistro===suc
  })

  const datos={}

  for(const f of filtradas){
    const tel=(f[0]||'').replace(/\D/g,'').slice(-10)
    const nombre=
      telToNombre[tel]||
      (f[1]||'Desconocido').trim()

    const key=tel||normaliza(nombre)

    if(!datos[key]){
      datos[key]={
        nombre,
        dias:0,
        ret:0,
        min:0,
        sin:0,
        horasMin:0,
        extraMin:0,
        tel
      }
    }

    const fe=parseFechaMX(f[2])

    if(fe&&f[3]&&f[5]&&!esSinSalida(f)){
      const prog=base[tel]?.horas?.[fe.getDay()]

      const calc=calcularExtra(
        f[3],
        f[5],
        prog?.entrada||null,
        prog?.salida||null
      )

      datos[key].horasMin+=minutosTrabajados(calc.trabajadas)
      datos[key].extraMin+=calc.extraMin
    }

    if(f[3]) datos[key].dias++

    const m=(f[4]||'').match(/(\d+)\s*min/)

    if(m){
      datos[key].ret++
      datos[key].min+=parseInt(m[1])
    }

    if(esSinSalida(f)){
      datos[key].sin++
    }
  }

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

  const wb=new ExcelJS.Workbook()

  const ws=wb.addWorksheet('Resumen')

  ws.addRow([
    `REPORTE ${filtroSucursal.toUpperCase()} - ${tipo}`
  ]).font={bold:true,size:14}

  ws.addRow([rangoTxt])
  ws.addRow([])

  ws.addRow([
    'Nombre',
    'Tel',
    'Días',
    'Retardos',
    'Min',
    'Sin salida',
    'Horas Trab',
    'Horas Extra'
  ]).font={bold:true}

  for(const k of Object.keys(datos)){
    const d=datos[k]

    ws.addRow([
      d.nombre,
      d.tel,
      `${d.dias} días`,
      `Ret ${d.ret} (${d.min}m)`,
      d.sin,
      `${Math.floor(d.horasMin/60)}:${String(d.horasMin%60).padStart(2,'0')}h`,
      `${Math.floor(d.extraMin/60)}:${String(d.extraMin%60).padStart(2,'0')}h`
    ])
  }

  ws.columns.forEach(c=>c.width=22)

  const ws2=wb.addWorksheet('Detalle')

  ws2.addRow(header).font={bold:true}

  filtradas.forEach(f=>{
    ws2.addRow(f)
  })

  ws2.columns.forEach(c=>c.width=18)

  const fileName=
    `Reporte_${filtroSucursal}_${tipo}_${lunes.toISOString().split('T')[0]}.xlsx`

  const fp=path.join(os.tmpdir(),fileName)

  await wb.xlsx.writeFile(fp)

  await sock.sendMessage(jid,{
    document:fs.readFileSync(fp),
    mimetype:
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    fileName
  })

  fs.unlinkSync(fp)
}

async function enviarDatosEmpleado(buscar,jid,sock,filtroGrupo){
  const empRows=await getRows('Empleados!A:T')
  const candidatos=buscarEmpleados(empRows,buscar,filtroGrupo)

  if(!candidatos.length){
    await sock.sendMessage(jid,{
      text:'No se encontró empleado en esta sucursal.'
    })
    return true
  }

  if(candidatos.length>1&&candidatos[0].score<100){
    const lista=candidatos.slice(0,5).map((x,i)=>
      `${i+1}. ${x.r[3]||x.r[1]} - ${x.r[2]||'-'}`
    ).join('\n')

    await sock.sendMessage(jid,{
      text:
        `Encontré varios empleados que coinciden con "${buscar}":\n\n`+
        `${lista}\n\n`+
        `Escribe el nombre más completo.`
    })

    return true
  }

  const r=candidatos[0].r
  const nombre=r[3]||r[1]||buscar
  const status=normaliza(r[16]||'')

  if(status&&!status.includes('activo')){
    const txt=
      `📋 *${nombre}*\n`+
      `📍 ${r[2]||'-'}\n`+
      `📅 Ingreso: ${r[5]||'-'}\n`+
      `📅 Baja: ${r[17]||'-'}\n`+
      `📝 Motivo de baja: ${r[18]||'-'}\n`+
      `🔄 Reingreso: ${r[19]||'-'}`

    await sock.sendMessage(jid,{text:txt})
    return true
  }

  const txt=
    `📋 *${nombre}*\n`+
    `📍 ${r[2]||'-'}\n`+
    `📱 ${r[0]||'-'}\n`+
    `💼 Puesto: ${r[4]||'-'}\n`+
    `🚨 Contacto de emergencia: ${r[6]||'-'}\n`+
    `📞 Tel. emergencia: ${r[7]||'-'}`

  await sock.sendMessage(jid,{text:txt})
  return true
}

async function enviarBanco(buscar,jid,sock,filtroGrupo){
  const empRows=await getRows('Empleados!A:T')
  const candidatos=buscarEmpleados(empRows,buscar,filtroGrupo)

  if(!candidatos.length){
    await sock.sendMessage(jid,{
      text:'No se encontró empleado en esta sucursal.'
    })
    return true
  }

  if(candidatos.length>1&&candidatos[0].score<100){
    const lista=candidatos.slice(0,5).map((x,i)=>
      `${i+1}. ${x.r[3]||x.r[1]} - ${x.r[2]||'-'}`
    ).join('\n')

    await sock.sendMessage(jid,{
      text:
        `Encontré varios empleados que coinciden con "${buscar}":\n\n`+
        `${lista}\n\n`+
        `Escribe el nombre más completo.`
    })

    return true
  }

  const r=candidatos[0].r
  const nombre=r[3]||r[1]||buscar

  const banco=r[12]||''
  const clabe=r[13]||''
  const cuenta=r[14]||''
  const tarjeta=r[15]||''

  let tipo=''
  let valor=''

  if(clabe){
    tipo='CLABE'
    valor=clabe
  }else if(cuenta){
    tipo='Cuenta'
    valor=cuenta
  }else if(tarjeta){
    tipo='Tarjeta'
    valor=tarjeta
  }

  if(!banco&&!valor){
    await sock.sendMessage(jid,{
      text:
        `🏦 *${nombre}*\n`+
        `No hay datos bancarios registrados.`
    })
    return true
  }

  let txt=`🏦 *${nombre}*\n`

  if(banco){
    txt+=`Banco: ${banco}\n`
  }

  if(valor){
    txt+=`${tipo}: ${valor}`
  }else{
    txt+='No hay CLABE, cuenta ni tarjeta registrada.'
  }

  await sock.sendMessage(jid,{text:txt})
  return true
}

export async function handleReportes({
  texto,
  jid,
  sock,
  filtroGrupo
}){
  const low=normaliza(texto)

  // =========================
  // ASISTENCIA HOY
  // =========================
  if(low.startsWith('asistencia hoy')){
    let suc=low.replace('asistencia hoy','').trim()

    if(!suc){
      suc=filtroGrupo||'coyoacan'
    }

    await asistenciaHoy(suc,jid,sock)
    return true
  }

  // =========================
  // FALTAS 45 DIAS
  // =========================
  if(/^faltas\b/.test(low)){
    let suc=''

    if(low.includes('bucareli')||low.includes('juarez')){
      suc='bucareli'
    }else if(low.includes('coyo')||low.includes('hotel')){
      suc='coyoacan'
    }else{
      suc=filtroGrupo
    }

    if(!suc){
      await sock.sendMessage(jid,{
        text:'Especifica: faltas Coyoacán o faltas Bucareli'
      })
      return true
    }

    await enviarFaltas45(suc,jid,sock)
    return true
  }

  // =========================
  // RETARDOS 45 DIAS
  // =========================
  if(/^retardos\b/.test(low)){
    let suc=''

    if(low.includes('bucareli')||low.includes('juarez')){
      suc='bucareli'
    }else if(low.includes('coyo')||low.includes('hotel')){
      suc='coyoacan'
    }else{
      suc=filtroGrupo
    }

    if(!suc){
      await sock.sendMessage(jid,{
        text:'Especifica: retardos Coyoacán o retardos Bucareli'
      })
      return true
    }

    await enviarRetardos45(suc,jid,sock)
    return true
  }

  // =========================
  // CRITICOS / GRAVES
  // SOLO REPORTES / PRUEBAS
  // =========================
  if(/^(críticos|criticos|graves)\b/.test(low)){
    const esPermitido=
      filtroGrupo===null||
      filtroGrupo===undefined

    if(!esPermitido){
      return false
    }

    await enviarCriticos(jid,sock)
    return true
  }

  // =========================
  // RESUMEN
  // =========================
  if(low.startsWith('resumen ')){
    let tipo='actual'

    if(low.includes('pasada')||low.includes('pasado')){
      tipo='pasada'
    }

    let limpio=low
      .slice(8)
      .trim()
      .replace(/pasada|pasado|actual|esta semana|hoy/g,'')
      .trim()

    if(
      limpio.includes('bucareli')||
      limpio.includes('juarez')||
      limpio.includes('coyo')||
      limpio.includes('hotel')
    ){
      let suc='coyoacan'

      if(
        limpio.includes('bucareli')||
        limpio.includes('juarez')
      ){
        suc='bucareli'
      }

      await reporteSucursal(suc,jid,sock,tipo)
      await generarExcelSemanaYEnviar(suc,jid,sock,tipo)
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

  // =========================
  // REPORTE / CHECADOR
  // =========================
  if(/^(reporte|checador|reporte x unidad)\b/i.test(texto)){
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

    if(
      low.includes('juarez')||
      low.includes('bucareli')
    ){
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

  // =========================
  // DATOS BANCARIOS
  // =========================
  if(
    /^(banco|clabe|clave|cuenta bancaria|cuenta empleado|cuenta|datos bancarios)\b/i.test(texto)
  ){
    let buscar=texto.replace(
      /^(banco|clabe|clave|cuenta bancaria|cuenta empleado|cuenta|datos bancarios)\s*(de)?\s*/i,
      ''
    ).trim()

    if(!buscar){
      await sock.sendMessage(jid,{
        text:'Escribe: banco Daniel, clabe Daniel, clave Daniel o cuenta Daniel'
      })
      return true
    }

    await enviarBanco(
      buscar,
      jid,
      sock,
      filtroGrupo
    )

    return true
  }

  // =========================
  // DATOS GENERALES
  // =========================
  if(
    /^(numero|número|num|tel|telefono|teléfono|info|ficha|datos|dato)\s*(de)?/i.test(texto)
  ){
    let buscar=texto.replace(
      /^(numero|número|num|tel|telefono|teléfono|info|ficha|datos|dato)\s*(de)?\s*/i,
      ''
    ).trim()

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

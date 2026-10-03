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
  calcularHorasTrabajadas
} from './utils.js'

// ============================================================
// AUXILIARES
// ============================================================

function tel10De(v){ return (v||'').toString().replace(/\D/g,'').slice(-10) }

function horasAMinutos(v){
  if(v===null||v===undefined||v==='') return 0
  const s=v.toString().trim()
  if(!s) return 0
  if(/^\d+(\.\d+)?$/.test(s)) return Math.round(parseFloat(s)*60)
  const m=s.match(/^(\d{1,3}):(\d{2})(?::(\d{2}))?$/)
  if(!m) return 0
  return parseInt(m[1],10)*60+parseInt(m[2],10)
}

function minutosAHoras(totalMin){
  const min=Math.max(0,Math.round(totalMin||0))
  return `${Math.floor(min/60)}:${String(min%60).padStart(2,'0')}`
}

function duracionProgramadaMin(prog){
  if(!prog?.entrada||!prog?.salida||prog.entrada==='LIBRE') return 0
  let ini=minutos(prog.entrada), fin=minutos(prog.salida), diff=fin-ini
  if(diff<0) diff+=1440
  return diff
}

// ============================================================
// HORARIO PROGRAMADO
// Primero usamos Asistencia!M porque el checador ya guardó
// la jornada que correspondía ese día. Si no existe, usamos
// Horario_Base como respaldo.
// ============================================================

function horarioProgramadoDeFila(f,baseMap){
  const jornada=(f[12]||'').toString().trim()
  const parsed=parseHorarioRango(jornada)
  if(parsed) return parsed

  const tel=tel10De(f[0])
  const info=baseMap[tel]
  if(!info) return null

  const fecha=parseFechaMX(f[2])
  if(!fecha) return null

  return info.horas?.[fecha.getDay()]||null
}

// ============================================================
// Minutos de entrada comparando correctamente jornadas que
// cruzan medianoche y usando 05:00 como inicio del día laboral.
// ============================================================

function minutosHorarioLaboral(v){
  let m=minutos(v)
  if(m<300) m+=1440
  return m
}

function calcularRetardo(entrada,prog){
  if(!entrada||!prog?.entrada||prog.entrada==='LIBRE') return 0
  const ent=minutosHorarioLaboral(entrada)
  const pro=minutosHorarioLaboral(prog.entrada)
  return Math.max(0,ent-pro)
}

// ============================================================
// MÉTRICAS DE UNA FILA
// ============================================================

function metricasFila(f,prog){
  const entrada=(f[3]||'').toString().trim()
  const salida=(f[5]||'').toString().trim()

  const esLibre=
    prog?.entrada==='LIBRE' ||
    (f[12]||'').toString().toUpperCase().includes('LIBRE')

  let trabajadasMin=0
  let extraMin=0
  let faltanteMin=0

  if(!entrada||!salida){
    return {trabajadasMin:0,extraMin:0,faltanteMin:0,tieneSalida:false,esLibre}
  }

  const trab=calcularHorasTrabajadas(entrada,salida)
  trabajadasMin=horasAMinutos(trab)

  // LIBRE: todo lo trabajado cuenta como trabajado; no hay extra.
  if(esLibre){
    return {trabajadasMin,extraMin:0,faltanteMin:0,tieneSalida:true,esLibre:true}
  }

  const jornadaMin=duracionProgramadaMin(prog)

  // Si no conocemos la jornada, no convertimos todo en extra.
  if(jornadaMin<=0){
    return {trabajadasMin,extraMin:0,faltanteMin:0,tieneSalida:true,esLibre:false}
  }

  faltanteMin=Math.max(0,jornadaMin-trabajadasMin)
  extraMin=Math.max(0,trabajadasMin-jornadaMin)

  return {trabajadasMin,extraMin,faltanteMin,tieneSalida:true,esLibre:false}
}

// ============================================================
// 1. ASISTENCIA HOY
// ============================================================

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
  const ahora=horaMX()
  const ahoraMin=minutosHorarioLaboral(ahora)

  const fechaLabDate=parseFechaMX(fLab)
  const diaNum=fechaLabDate?.getDay() ?? new Date().getDay()

  const filtro=(filtroSucursal||'').toLowerCase()
  const esCoyo=filtro.includes('coyo')||filtro.includes('hotel')
  const esJuarez=filtro.includes('juarez')||filtro.includes('bucareli')

  const llego=[],retardo=[],falta=[],futuro=[]

  for(const r of baseRows){
    const sucBaseLower=(r[2]||'').toLowerCase()
    const sucBaseOriginal=r[2]||''
    let inc=false

    if(esCoyo) inc=sucBaseLower.includes('coyo')||sucBaseLower.includes('hotel')||sucBaseLower.includes('trinidad')
    else if(esJuarez) inc=sucBaseLower.includes('juarez')||sucBaseLower.includes('bucareli')
    else inc=sucBaseLower.includes(filtro)

    if(!inc) continue

    const nombre=r[1]||''
    const tel=tel10De(r[0])

    const mapa={1:r[3],2:r[4],3:r[5],4:r[6],5:r[7],6:r[8],0:r[9]}
    const v=(mapa[diaNum]||'').toString().trim()
    const parsed=parseHorarioRango(v)
    if(!parsed) continue

    const esLibre=parsed.entrada==='LIBRE'
    const horaProg=esLibre?'LIBRE':parsed.entrada

    const registros=asisRows.filter(a=>
      tel10De(a[0])===tel &&
      a[2]===fLab &&
      a[3]
    )

    // LIBRE: acumulamos todos los intervalos del día.
    if(esLibre){
      if(!registros.length){
        if(ahoraMin>=1200) falta.push(`• ${nombre} - [${sucBaseOriginal}] - ❌ sin llegar`)
        continue
      }

      let totalMin=0
      let abierto=false

      for(const reg of registros){
        if(reg[3]&&reg[5]){
          totalMin+=horasAMinutos(calcularHorasTrabajadas(reg[3],reg[5]))
        }else if(reg[3]&&!reg[5]){
          abierto=true
        }
      }

      llego.push(
        `• ${nombre} - [${sucBaseOriginal}] - Trab: ${minutosAHoras(totalMin)}h`+
        `${abierto?' - 🟢 En turno':' ✅'}`
      )
      continue
    }

    const registro=registros[0]

    if(registro?.[3]){
      const entrada=registro[3]
      const sucEnt=registro[6]||''
      const dif=calcularRetardo(entrada,parsed)
      const metricas=metricasFila(registro,parsed)

      if(dif>15){
        retardo.push(
          `• ${nombre} - [${sucBaseOriginal}] Prog ${horaProg} - Entró ${entrada} - ⏰ ${dif}m tarde`+
          ` - Trab: ${minutosAHoras(metricas.trabajadasMin)}h`+
          `${metricas.faltanteMin>0?` - Falt: ${minutosAHoras(metricas.faltanteMin)}h`:''}`+
          ` - ${sucEnt}`
        )
      }else{
        llego.push(
          `• ${nombre} - [${sucBaseOriginal}] Prog ${horaProg} - Entró ${entrada} ✅ - ${sucEnt}`+
          `${metricas.tieneSalida?` - Trab: ${minutosAHoras(metricas.trabajadasMin)}h`:''}`
        )
      }
    }else{
      const progMin=minutosHorarioLaboral(horaProg)

      if(ahoraMin<progMin){
        futuro.push(`• ${nombre} - [${sucBaseOriginal}] - Prog ${horaProg}`)
      }else{
        falta.push(`• ${nombre} - [${sucBaseOriginal}] - Prog ${horaProg} - ❌ ${ahoraMin-progMin}m sin llegar`)
      }
    }
  }

  const txt=
    `📍 *ASISTENCIA HOY ${fLab} - ${filtroSucursal.toUpperCase()}* ${ahora}\n\n`+
    `✅ *A TIEMPO (${llego.length}):*\n${llego.join('\n')||'-'}\n\n`+
    `⏰ *RETARDOS (${retardo.length}):*\n${retardo.join('\n')||'-'}\n\n`+
    `❌ *NO HAN LLEGADO (${falta.length}):*\n${falta.join('\n')||'Todos llegaron'}\n\n`+
    `⏳ *PRÓXIMOS (${futuro.length}):*\n${futuro.join('\n')||'-'}`

  await sock.sendMessage(jid,{text:txt})
}

// ============================================================
// 2. EXCEL INDIVIDUAL
// ============================================================

export async function generarExcelEmpleado(nombreBuscarRaw,jid,sock,tipo='actual'){
  const sClient=await sheetsClient()
  const baseMap=await getHorarioBaseMap()

  const asisRes=await sClient.spreadsheets.values.get({
    spreadsheetId:SPREADSHEET_ID,
    range:'Asistencia!A2:M'
  })

  const filas=asisRes.data.values||[]

  const buscar=nombreBuscarRaw.toLowerCase()
    .replace(/actual|pasada|pasado|esta semana|hoy/g,'')
    .trim()

  const {lunes,domingo,rangoTxt}=getRangoSemana(tipo)

  const filtradas=filas.filter(f=>{
    const n=(f[1]||'').toLowerCase()
    if(!n.includes(buscar)) return false
    const fe=parseFechaMX(f[2])
    return fe&&fe>=lunes&&fe<=domingo
  })

  const header=[
    'Tel','Nombre','Fecha','Entrada','Estatus Entrada','Salida',
    'Suc Entrada','Dist Entr','Suc Salida','Dist Sal',
    'Horas Trabajadas','Horas Extra','Faltante','Jornada Programada'
  ]

  const wb=new ExcelJS.Workbook()
  const ws=wb.addWorksheet('Resumen')

  ws.addRow([`REPORTE ${buscar.toUpperCase()} - ${tipo.toUpperCase()}`]).font={bold:true,size:14}
  ws.addRow([rangoTxt])
  ws.addRow([])

  let dias=0,ret=0,minRet=0,sin=0
  let horasMin=0,extraMin=0,faltanteMin=0
  const diasUnicos=new Set()

  for(const f of filtradas){
    const fecha=parseFechaMX(f[2])
    const fechaClave=fecha
      ? `${fecha.getFullYear()}-${String(fecha.getMonth()+1).padStart(2,'0')}-${String(fecha.getDate()).padStart(2,'0')}`
      : f[2]||''

    const prog=horarioProgramadoDeFila(f,baseMap)
    const metricas=metricasFila(f,prog)

    if(f[3]&&!diasUnicos.has(fechaClave)){
      diasUnicos.add(fechaClave)
      dias++
    }

    const mRet=calcularRetardo(f[3],prog)
    if(mRet>15){
      ret++
      minRet+=mRet
    }

    if(f[3]&&!f[5]) sin++

    if(f[3]&&f[5]){
      horasMin+=metricas.trabajadasMin
      extraMin+=metricas.extraMin
      faltanteMin+=metricas.faltanteMin
    }
  }

  ws.addRow([
    'Nombre','Días','Retardos','Min Retardo','Sin salida',
    'Horas Trab','Horas Extra','Horas Faltantes'
  ]).font={bold:true}

  ws.addRow([
    buscar,`${dias} días`,`Ret ${ret}`,`${minRet} min`,sin,
    `${minutosAHoras(horasMin)}h`,
    `${minutosAHoras(extraMin)}h`,
    `${minutosAHoras(faltanteMin)}h`
  ])

  ws.columns.forEach(c=>{c.width=22})

  const ws2=wb.addWorksheet('Detalle')
  ws2.addRow(header).font={bold:true}

  for(const f of filtradas){
    const prog=horarioProgramadoDeFila(f,baseMap)
    const metricas=metricasFila(f,prog)

    ws2.addRow([
      f[0]||'',f[1]||'',f[2]||'',f[3]||'',f[4]||'',f[5]||'',
      f[6]||'',f[7]||'',f[8]||'',f[9]||'',
      f[10]||minutosAHoras(metricas.trabajadasMin),
      f[11]||(metricas.extraMin>0?`${minutosAHoras(metricas.extraMin)}:00`:'0'),
      metricas.faltanteMin>0?`${minutosAHoras(metricas.faltanteMin)}:00`:'0',
      f[12]||''
    ])
  }

  ws2.columns.forEach(c=>{c.width=18})

  const fileName=`Reporte_${buscar.replace(/\s+/g,'_')}_${tipo}.xlsx`
  const fp=path.join(os.tmpdir(),fileName)

  await wb.xlsx.writeFile(fp)

  await sock.sendMessage(jid,{
    document:fs.readFileSync(fp),
    mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    fileName
  })

  fs.unlinkSync(fp)
}

// ============================================================
// 3. RESUMEN DEL EMPLEADO
// ============================================================

export async function resumenEmpleado(nombreBuscar,jidRespuesta,sock,tipo='actual'){
  const sClient=await sheetsClient()

  const [asisRes,baseMap]=await Promise.all([
    sClient.spreadsheets.values.get({
      spreadsheetId:SPREADSHEET_ID,
      range:'Asistencia!A2:M'
    }),
    getHorarioBaseMap()
  ])

  const filas=asisRes.data.values||[]

  const buscar=nombreBuscar.toLowerCase()
    .replace(/actual|pasada|pasado|esta semana|hoy/g,'')
    .trim()

  const {lunes,domingo,rangoTxt}=getRangoSemana(tipo)

  const info=
    Object.values(baseMap).find(v=>
      normaliza(v.nombreOriginal||'').includes(normaliza(buscar))
    )||
    {descansos:new Set(),horas:{}}

  const porDia=new Map()

  for(const f of filas){
    const n=(f[1]||'').toLowerCase()
    if(!n.includes(buscar)) continue

    const fe=parseFechaMX(f[2])
    if(!fe) continue

    fe.setHours(0,0,0,0)
    if(fe<lunes||fe>domingo) continue

    const fechaKey=
      `${fe.getFullYear()}-${String(fe.getMonth()+1).padStart(2,'0')}-${String(fe.getDate()).padStart(2,'0')}`

    if(!porDia.has(fechaKey)){
      porDia.set(fechaKey,{fecha:f[2],filas:[]})
    }

    porDia.get(fechaKey).filas.push(f)
  }

  let diasSem=0,retSem=0,minSem=0,sin=0
  let horasMin=0,extraTotal=0,faltanteTotal=0
  const detalle=[]

  for(const dia of porDia.values()){
    const filasDia=dia.filas

    if(filasDia.some(f=>f[3])) diasSem++

    let diaTrab=0,diaExtra=0,diaFaltante=0,diaRetMin=0
    let haySinSalida=false
    let esLibre=false

    for(const f of filasDia){
      const prog=horarioProgramadoDeFila(f,baseMap)
      const metricas=metricasFila(f,prog)

      if(metricas.esLibre) esLibre=true

      diaTrab+=metricas.trabajadasMin
      diaExtra+=metricas.extraMin
      diaFaltante+=metricas.faltanteMin

      if(f[3]&&!f[5]) haySinSalida=true

      const mRet=calcularRetardo(f[3],prog)
      if(mRet>15) diaRetMin+=mRet
    }

    if(diaRetMin>0){
      retSem++
      minSem+=diaRetMin
    }

    if(haySinSalida) sin++

    horasMin+=diaTrab
    extraTotal+=diaExtra
    faltanteTotal+=diaFaltante

    // LIBRE: muestra intervalos y total.
    if(esLibre){
      const intervalos=filasDia
        .filter(f=>f[3])
        .map(f=>`${f[3]} - ${f[5]||'SIN'}`)

      detalle.push(
        `• ${dia.fecha}: LIBRE ${intervalos.join(' | ')} | Total ${minutosAHoras(diaTrab)}h`
      )
    }else{
      const primera=filasDia.find(f=>f[3])
      const prog=primera?horarioProgramadoDeFila(primera,baseMap):null
      const entrada=primera?.[3]||''
      const salida=primera?.[5]||'SIN'
      const txtProg=prog?.entrada
        ? (prog.salida?`${prog.entrada}-${prog.salida}`:prog.entrada)
        : (primera?.[12]||'')

      detalle.push(
        `• ${dia.fecha}: Prog ${txtProg}`+
        ` | Ent ${entrada||'SIN'}`+
        ` | Sal ${salida}`+
        ` | Trab ${minutosAHoras(diaTrab)}h`+
        ` | Extra ${minutosAHoras(diaExtra)}h`+
        ` | Falt ${minutosAHoras(diaFaltante)}h`+
        `${diaRetMin>0?` | Ret ${diaRetMin}m`:''}`
      )
    }
  }

  let esperados=0

  for(let d=new Date(lunes);d<=domingo;d.setDate(d.getDate()+1)){
    if(!info.descansos.has(d.getDay())) esperados++
  }

  const faltas=Math.max(0,esperados-diasSem)

  const diasNom=['Dom','Lun','Mar','Mie','Jue','Vie','Sab']
  const descTxt=[...info.descansos].map(d=>diasNom[d]).join(', ')||'ninguno'

  const txt=
    `📊 *${buscar.toUpperCase()}* - Desc: ${descTxt}\n`+
    `${rangoTxt}\n\n`+
    `*SEMANA ${tipo.toUpperCase()}*\n`+
    `- Trabajados: ${diasSem}/${esperados}\n`+
    `- Faltas: ${faltas}\n`+
    `- Retardos: ${retSem} (${minSem} min)\n`+
    `- Sin salida: ${sin}\n`+
    `- Horas Trab: ${minutosAHoras(horasMin)}h\n`+
    `- Horas Extra: ${minutosAHoras(extraTotal)}h\n`+
    `- Horas Faltantes: ${minutosAHoras(faltanteTotal)}h\n\n`+
    `*Detalle:*\n${detalle.join('\n')||'Sin registros'}`

  await sock.sendMessage(jidRespuesta,{text:txt})

  if(diasSem>0){
    await generarExcelEmpleado(buscar,jidRespuesta,sock,tipo)
  }
}

// ============================================================
// 4. REPORTE DE SUCURSAL
// ============================================================

export async function reporteSucursal(filtroSucursal,jid,sock,tipo='pasada'){
  const [asisRes,empRows]=await Promise.all([
    (await sheetsClient()).spreadsheets.values.get({
      spreadsheetId:SPREADSHEET_ID,
      range:'Asistencia!A2:M'
    }),
    getRows('Empleados!A:K')
  ])

  const filas=asisRes.data.values||[]
  const {lunes,domingo,rangoTxt}=getRangoSemana(tipo)
  const filtro=(filtroSucursal||'').toLowerCase()

  const esCoyo=filtro.includes('coyo')||filtro.includes('hotel')
  const esJuarez=filtro.includes('juarez')||filtro.includes('bucareli')
  const baseMap=await getHorarioBaseMap()

  const telToNombre={}

  empRows.forEach(r=>{
    const tel=tel10De(r[0])
    if(tel) telToNombre[tel]=(r[3]||r[1]||'').trim()
  })

  const filtradas=filas.filter(f=>{
    const fe=parseFechaMX(f[2])
    if(!fe||fe<lunes||fe>domingo) return false

    const suc=((f[6]||'')+' '+(f[8]||'')).toLowerCase()

    if(esCoyo) return suc.includes('coyo')||suc.includes('hotel')
    if(esJuarez) return suc.includes('juarez')||suc.includes('bucareli')
    return suc.includes(filtro)
  })

  const diasMap=new Map()

  for(const f of filtradas){
    const tel=tel10De(f[0])
    const nombre=telToNombre[tel]||(f[1]||'').trim()||'Desconocido'
    const fe=parseFechaMX(f[2])
    if(!fe) continue

    const fechaKey=
      `${fe.getFullYear()}-${String(fe.getMonth()+1).padStart(2,'0')}-${String(fe.getDate()).padStart(2,'0')}`

    const key=`${tel||normaliza(nombre)}_${fechaKey}`

    if(!diasMap.has(key)){
      diasMap.set(key,{tel,nombre,fecha:f[2],filas:[]})
    }

    diasMap.get(key).filas.push(f)
  }

  const datos={}

  for(const dia of diasMap.values()){
    const key=dia.tel||normaliza(dia.nombre)

    if(!datos[key]){
      datos[key]={
        nombre:dia.nombre,tel:dia.tel,dias:0,ret:0,min:0,sin:0,
        horasMin:0,extraMin:0,faltanteMin:0
      }
    }

    datos[key].dias++

    let diaRetMin=0,diaHoras=0,diaExtra=0,diaFaltante=0,diaSinSalida=false

    for(const f of dia.filas){
      const prog=horarioProgramadoDeFila(f,baseMap)
      const metricas=metricasFila(f,prog)

      diaHoras+=metricas.trabajadasMin
      diaExtra+=metricas.extraMin
      diaFaltante+=metricas.faltanteMin

      if(f[3]&&!f[5]) diaSinSalida=true

      const mRet=calcularRetardo(f[3],prog)
      if(mRet>15) diaRetMin+=mRet
    }

    if(diaRetMin>0){
      datos[key].ret++
      datos[key].min+=diaRetMin
    }

    if(diaSinSalida) datos[key].sin++

    datos[key].horasMin+=diaHoras
    datos[key].extraMin+=diaExtra
    datos[key].faltanteMin+=diaFaltante
  }

  let txt=`📊 *REPORTE ${filtroSucursal.toUpperCase()}* - ${tipo}\n${rangoTxt}\n\n`

  if(!Object.keys(datos).length){
    txt+='Sin registros\n'
  }else{
    for(const k in datos){
      const d=datos[k]
      txt+=
        `*${d.nombre}*: ${d.dias} días`+
        ` | Ret ${d.ret} (${d.min}m)`+
        ` | Trab: ${minutosAHoras(d.horasMin)}h`+
        ` | Extra: ${minutosAHoras(d.extraMin)}h`+
        ` | Falt: ${minutosAHoras(d.faltanteMin)}h`+
        ` | Sin salida: ${d.sin}\n\n`
    }
  }

  await sock.sendMessage(jid,{text:txt})
}

// ============================================================
// 5. EXCEL SEMANAL DE SUCURSAL
// ============================================================

export async function generarExcelSemanaYEnviar(filtroSucursal,jid,sock,tipo='pasada'){
  const sClient=await sheetsClient()
  const baseMap=await getHorarioBaseMap()

  const [asisRes,empRows]=await Promise.all([
    sClient.spreadsheets.values.get({
      spreadsheetId:SPREADSHEET_ID,
      range:'Asistencia!A2:M'
    }),
    getRows('Empleados!A:K')
  ])

  const filas=asisRes.data.values||[]
  const {lunes,domingo,rangoTxt}=getRangoSemana(tipo)
  const filtro=(filtroSucursal||'').toLowerCase()

  const esCoyo=filtro.includes('coyo')||filtro.includes('hotel')
  const esJuarez=filtro.includes('juarez')||filtro.includes('bucareli')

  const telToNombre={}

  empRows.forEach(r=>{
    const tel=tel10De(r[0])
    if(tel) telToNombre[tel]=(r[3]||r[1]||'').trim()
  })

  const filtradas=filas.filter(f=>{
    const fe=parseFechaMX(f[2])
    if(!fe||fe<lunes||fe>domingo) return false

    const suc=((f[6]||'')+' '+(f[8]||'')).toLowerCase()

    if(esCoyo) return suc.includes('coyo')||suc.includes('hotel')
    if(esJuarez) return suc.includes('juarez')||suc.includes('bucareli')
    return suc.includes(filtro)
  })

  const diasMap=new Map()

  for(const f of filtradas){
    const tel=tel10De(f[0])
    const fe=parseFechaMX(f[2])
    if(!fe) continue

    const fechaKey=
      `${fe.getFullYear()}-${String(fe.getMonth()+1).padStart(2,'0')}-${String(fe.getDate()).padStart(2,'0')}`

    const nombre=telToNombre[tel]||(f[1]||'').trim()||'Desconocido'
    const key=`${tel||normaliza(nombre)}_${fechaKey}`

    if(!diasMap.has(key)){
      diasMap.set(key,{tel,nombre,fecha:f[2],filas:[]})
    }

    diasMap.get(key).filas.push(f)
  }

  const resumenDiario=[]

  for(const dia of diasMap.values()){
    let horasMin=0,extraMin=0,faltanteMin=0,retMin=0
    let sinSalida=false

    for(const f of dia.filas){
      const prog=horarioProgramadoDeFila(f,baseMap)
      const metricas=metricasFila(f,prog)

      horasMin+=metricas.trabajadasMin
      extraMin+=metricas.extraMin
      faltanteMin+=metricas.faltanteMin

      if(f[3]&&!f[5]) sinSalida=true

      const mRet=calcularRetardo(f[3],prog)
      if(mRet>15) retMin+=mRet
    }

    resumenDiario.push({
      tel:dia.tel,
      nombre:dia.nombre,
      fecha:dia.fecha,
      horasMin,
      extraMin,
      faltanteMin,
      retMin,
      sinSalida
    })
  }

  const datos={}

  for(const d of resumenDiario){
    const key=d.tel||normaliza(d.nombre)

    if(!datos[key]){
      datos[key]={
        nombre:d.nombre,tel:d.tel,dias:0,ret:0,min:0,sin:0,
        horasMin:0,extraMin:0,faltanteMin:0
      }
    }

    datos[key].dias++

    if(d.retMin>0){
      datos[key].ret++
      datos[key].min+=d.retMin
    }

    if(d.sinSalida) datos[key].sin++

    datos[key].horasMin+=d.horasMin
    datos[key].extraMin+=d.extraMin
    datos[key].faltanteMin+=d.faltanteMin
  }

  const wb=new ExcelJS.Workbook()

  const ws=wb.addWorksheet('Resumen')
  ws.addRow([`REPORTE ${filtroSucursal.toUpperCase()} - ${tipo}`]).font={bold:true,size:14}
  ws.addRow([rangoTxt])
  ws.addRow([])

  ws.addRow([
    'Nombre','Tel','Días','Retardos','Min Retardo','Sin salida',
    'Horas Trab','Horas Extra','Horas Faltantes'
  ]).font={bold:true}

  for(const k of Object.keys(datos)){
    const d=datos[k]

    ws.addRow([
      d.nombre,d.tel,`${d.dias} días`,d.ret,`${d.min} min`,d.sin,
      `${minutosAHoras(d.horasMin)}h`,
      `${minutosAHoras(d.extraMin)}h`,
      `${minutosAHoras(d.faltanteMin)}h`
    ])
  }

  ws.columns.forEach(c=>{c.width=22})

  // Diario: un registro por empleado y día.
  // LIBRE queda resumido por total diario, no por movimiento.
  const ws2=wb.addWorksheet('Diario')

  ws2.addRow([
    'Tel','Nombre','Fecha','Horas Trabajadas','Horas Extra',
    'Horas Faltantes','Min Retardo','Sin salida'
  ]).font={bold:true}

  resumenDiario
    .sort((a,b)=>`${a.nombre}-${a.fecha}`.localeCompare(`${b.nombre}-${b.fecha}`))
    .forEach(d=>{
      ws2.addRow([
        d.tel,
        d.nombre,
        d.fecha,
        `${minutosAHoras(d.horasMin)}h`,
        `${minutosAHoras(d.extraMin)}h`,
        `${minutosAHoras(d.faltanteMin)}h`,
        d.retMin,
        d.sinSalida?'SI':'NO'
      ])
    })

  ws2.columns.forEach(c=>{c.width=22})

  const fileName=
    `Reporte_${filtroSucursal}_${tipo}_${lunes.toISOString().split('T')[0]}.xlsx`

  const fp=path.join(os.tmpdir(),fileName)

  await wb.xlsx.writeFile(fp)

  await sock.sendMessage(jid,{
    document:fs.readFileSync(fp),
    mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    fileName
  })

  fs.unlinkSync(fp)
}

// ============================================================
// 6. MANEJO DE COMANDOS
// ============================================================

export async function handleReportes({texto,jid,sock,filtroGrupo}){
  const low=texto.toLowerCase()

  // ==========================================================
  // ASISTENCIA HOY
  // ==========================================================

  if(low.startsWith('asistencia hoy')){
    let suc=low.replace('asistencia hoy','').trim()
    if(!suc) suc=filtroGrupo||'coyoacan'

    await asistenciaHoy(suc,jid,sock)
    return true
  }

  // ==========================================================
  // RESUMEN
  // ==========================================================

  if(low.startsWith('resumen ')){
    let tipo='actual'

    if(low.includes('pasada')||low.includes('pasado')) tipo='pasada'

    let limpio=texto
      .slice(8)
      .toLowerCase()
      .trim()
      .replace(/pasada|pasado|actual|esta semana|hoy/g,'')
      .trim()

    // Resumen por sucursal.
    if(limpio.includes('bucareli')||limpio.includes('juarez')||limpio.includes('coyo')||limpio.includes('hotel')){
      let suc='coyoacan'

      if(limpio.includes('juarez')||limpio.includes('bucareli')) suc='juarez'

      await reporteSucursal(suc,jid,sock,tipo)
      await generarExcelSemanaYEnviar(suc,jid,sock,tipo)
      return true
    }

    // Resumen por empleado.
    if(!limpio){
      await sock.sendMessage(jid,{text:'Escribe: resumen fer pasada o actual'})
      return true
    }

    await resumenEmpleado(limpio,jid,sock,tipo)
    return true
  }

  // ==========================================================
  // REPORTE / CHECADOR
  // ==========================================================

  if(/^(reporte|checador|reporte x unidad)/i.test(texto)){
    const tipo=
      low.includes('pasada')||low.includes('pasado')
        ? 'pasada'
        : low.includes('actual')||low.includes('esta')||low.includes('hoy')
          ? 'actual'
          : 'pasada'

    let suc='coyoacan'

    if(low.includes('juarez')||low.includes('bucareli')) suc='juarez'
    else if(low.includes('hotel')) suc='hotel'
    else if(low.includes('coyo')) suc='coyoacan'

    // POR UNIDAD.
    if(low.includes('x unidad')||low.includes('por unidad')){
      if(!filtroGrupo){
        await reporteSucursal('coyoacan',jid,sock,tipo)
        await generarExcelSemanaYEnviar('coyoacan',jid,sock,tipo)
        await reporteSucursal('juarez',jid,sock,tipo)
        await generarExcelSemanaYEnviar('juarez',jid,sock,tipo)
      }else{
        await reporteSucursal(filtroGrupo,jid,sock,tipo)
        await generarExcelSemanaYEnviar(filtroGrupo,jid,sock,tipo)
      }
      return true
    }

    await reporteSucursal(suc,jid,sock,tipo)
    await generarExcelSemanaYEnviar(suc,jid,sock,tipo)
    return true
  }

  // ==========================================================
  // INFORMACIÓN DE EMPLEADO
  // ==========================================================

  if(/^(numero|número|num|tel|telefono|teléfono|info|ficha|datos|dato)\s*(de)?/i.test(texto)){
    const buscar=texto
      .toLowerCase()
      .replace(/^(numero|número|num|tel|telefono|teléfono|info|ficha|datos|dato)\s*(de)?\s*/i,'')
      .trim()

    if(!buscar){
      await sock.sendMessage(jid,{text:'Escribe: info fer'})
      return true
    }

    const buscarNorm=normaliza(buscar)
    const empRows=await getRows('Empleados!A:K')
    const candidatos=[]

    for(const r of empRows.slice(1)){
      const corto=r[1]||''
      const completo=r[3]||''
      const suc=r[2]||''

      if(!sucursalCoincideConFiltro(suc,filtroGrupo)) continue

      const s=scoreEmpleado(corto,completo,buscarNorm)
      if(s>=0) candidatos.push({r,s})
    }

    candidatos.sort((a,b)=>b.s-a.s)

    if(!candidatos.length){
      await sock.sendMessage(jid,{text:`No encontré a "${buscar}"`})
      return true
    }

    const r=candidatos[0].r

    const ficha=
      `📋 *${r[3]||r[1]}*\n`+
      `👤 Corto: ${r[1]}\n`+
      `📍 ${r[2]||'-'}\n`+
      `📱 ${r[0]||'-'}`

    await sock.sendMessage(jid,{text:ficha})
    return true
  }

  return false
}

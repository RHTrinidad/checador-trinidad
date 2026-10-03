import { SUCURSALES, GRUPO_COYOACAN_ID, GRUPO_BUCARELI_ID, SPREADSHEET_ID } from './config.js'
import { distM, fechaLaboral, horaMX, minutos, normaliza, calcularHorasTrabajadas, calcularExtra, parseHorarioRango } from './utils.js'
import { getRows, sheetsClient } from './sheets.js'

const TOLERANCIA_MIN = 15
const AVISO_FALTA_MIN = 20
const CIERRE_AUTO_HORAS = 16

let avisosHoy = new Set()
let fechaAvisos = ''
let ultimoRegistroDescansos = ''

function esDescanso(v){
  const s=(v||'').toString().trim().toLowerCase()
  return s.includes('descanso')
}

function diaMexico(){
  return new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'})).getDay()
}

function horarioDelDia(r,dia){
  const mapa={
    1:r[3],
    2:r[4],
    3:r[5],
    4:r[6],
    5:r[7],
    6:r[8],
    0:r[9]
  }
  return (mapa[dia]||'').toString().trim()
}

export async function handleChecador({ sock, jid, m, loc, rawLid, tel10, tel }){
  const lat=loc.degreesLatitude
  const lng=loc.degreesLongitude

  let cercana=null
  let dMin=Infinity

  for(const s of SUCURSALES){
    const d=distM(lat,lng,s.lat,s.lng)
    if(d<dMin){
      dMin=d
      cercana=s
    }
  }

  const empRows=await getRows('Empleados!A:K')

  const getLid=r=>(r.find(x=>String(x).includes('@lid'))||'').trim()

  let emp=null

  if(tel10.length>=10){
    const i=empRows.findIndex((r,idx)=>
      idx>0 &&
      r[0] &&
      r[0].replace(/\D/g,'').slice(-10)===tel10
    )

    if(i>-1)emp=empRows[i]
  }

  if(!emp&&rawLid.includes('@lid')){
    const i=empRows.findIndex((r,idx)=>
      idx>0 &&
      getLid(r)===rawLid
    )

    if(i>-1){
      emp=empRows[i]
      tel10=(emp[0]||'').replace(/\D/g,'').slice(-10)
    }
  }

  const nombre=emp?(emp[3]||emp[1]):(m.pushName||tel10||'Desconocido')
  const telF=emp?(emp[0]||'').replace(/\D/g,''):tel
  const tel10F=telF.slice(-10)||tel10

  const fLab=fechaLaboral()

  const asis=await getRows('Asistencia!A:M')

  const idx=asis.findIndex((r,i)=>
    i>0 &&
    r[0] &&
    r[0].replace(/\D/g,'').slice(-10)===tel10F &&
    r[2]===fLab
  )

  let hoy=idx>-1?asis[idx]:null

  const sClient=await sheetsClient()

  const baseRows=await getRows('Horario_Base!A2:K')

  const baseRow=baseRows.find(r=>
    (r[0]||'').replace(/\D/g,'').slice(-10)===tel10F
  )

  const dia=diaMexico()
  const horario=baseRow?horarioDelDia(baseRow,dia):''
  const esDesc=esDescanso(horario)

  let hObj=null

  if(!esDesc){
    const parsed=parseHorarioRango(horario)

    if(parsed&&parsed.entrada!=='LIBRE'&&parsed.entrada){
      hObj=parsed
    }
  }

  let estatus='A TIEMPO'

  if(esDesc){
    estatus='TRABAJO EN DESCANSO'
  }else if(hObj?.entrada){
    const dif=minutos(horaMX())-minutos(hObj.entrada)

    if(dif>TOLERANCIA_MIN){
      estatus=`RETARDO ${dif}min (Prog ${hObj.entrada})`
    }else{
      estatus=`A TIEMPO (Prog ${hObj.entrada})`
    }
  }

  try{

    // ==============================
    // PRIMER REGISTRO / ENTRADA
    // ==============================

    if(!hoy||(!hoy[3]&&String(hoy[3]||'').toUpperCase()!=='DESCANSO')){

      if(dMin>cercana.rEnt){
        await sock.sendMessage(
          jid,
          {text:`Debes estar a max ${cercana.rEnt}m de ${cercana.nombre}`},
          {quoted:m}
        )
        return
      }

      const h=horaMX()

      // ==============================
      // TRABAJO EN DÍA DE DESCANSO
      // ==============================

      if(esDesc){

        const jTxt='DESCANSO'

        const row=[
          tel10F,
          nombre,
          fLab,
          h,
          'TRABAJO EN DESCANSO',
          '',
          cercana.nombre,
          Math.round(dMin).toString(),
          '',
          '',
          '0',
          '0',
          jTxt
        ]

        if(idx===-1){
          await sClient.spreadsheets.values.append({
            spreadsheetId:SPREADSHEET_ID,
            range:'Asistencia!A:M',
            valueInputOption:'USER_ENTERED',
            requestBody:{values:[row]}
          })
        }else{
          await sClient.spreadsheets.values.update({
            spreadsheetId:SPREADSHEET_ID,
            range:`Asistencia!A${idx+1}:M${idx+1}`,
            valueInputOption:'USER_ENTERED',
            requestBody:{values:[row]}
          })
        }

        await sock.sendMessage(jid,{
          text:`✅ TRABAJO EN DESCANSO - ${nombre} en ${cercana.nombre} - ${h}`
        })

        return
      }

      // ==============================
      // ENTRADA NORMAL
      // ==============================

      const jTxt=hObj
        ?`${hObj.entrada}${hObj.salida?` - ${hObj.salida}`:''}`
        :'8h'

      const row=[
        tel10F,
        nombre,
        fLab,
        h,
        estatus,
        '',
        cercana.nombre,
        Math.round(dMin).toString(),
        '',
        '',
        calcularHorasTrabajadas(h,''),
        '0',
        jTxt
      ]

      if(idx===-1){
        await sClient.spreadsheets.values.append({
          spreadsheetId:SPREADSHEET_ID,
          range:'Asistencia!A:M',
          valueInputOption:'USER_ENTERED',
          requestBody:{values:[row]}
        })
      }else{
        await sClient.spreadsheets.values.update({
          spreadsheetId:SPREADSHEET_ID,
          range:`Asistencia!A${idx+1}:M${idx+1}`,
          valueInputOption:'USER_ENTERED',
          requestBody:{values:[row]}
        })
      }

      await sock.sendMessage(jid,{
        text:`✅ ${estatus} - ${nombre} en ${cercana.nombre} - ${h}`
      })

      return
    }

    // ==============================
    // CONVERTIR DESCANSO EN TRABAJO
    // ==============================

    if(hoy&&String(hoy[3]||'').toUpperCase()==='DESCANSO'){

      if(dMin>cercana.rEnt){
        await sock.sendMessage(
          jid,
          {text:`Debes estar a max ${cercana.rEnt}m de ${cercana.nombre}`},
          {quoted:m}
        )
        return
      }

      const h=horaMX()

      const row=[
        tel10F,
        nombre,
        fLab,
        h,
        'TRABAJO EN DESCANSO',
        '',
        cercana.nombre,
        Math.round(dMin).toString(),
        '',
        '',
        '0',
        '0',
        'DESCANSO'
      ]

      await sClient.spreadsheets.values.update({
        spreadsheetId:SPREADSHEET_ID,
        range:`Asistencia!A${idx+1}:M${idx+1}`,
        valueInputOption:'USER_ENTERED',
        requestBody:{values:[row]}
      })

      await sock.sendMessage(jid,{
        text:`✅ TRABAJO EN DESCANSO - ${nombre} en ${cercana.nombre} - ${h}`
      })

      return
    }

    // ==============================
    // SALIDA YA REGISTRADA
    // ==============================

    if(hoy[5]){
      await sock.sendMessage(jid,{
        text:`Salida ya registrada`
      })
      return
    }

    // ==============================
    // VALIDAR DISTANCIA DE SALIDA
    // ==============================

    if(dMin>cercana.rSal){
      await sock.sendMessage(
        jid,
        {text:`No puedes checar salida a ${Math.round(dMin)}m`},
        {quoted:m}
      )
      return
    }

    const h=horaMX()
    const jTxt=hoy[12]||'8h'

    let trabajadas='0'
    let extra='0'

    // ==============================
    // SALIDA TRABAJO EN DESCANSO
    // ==============================

    if(String(hoy[4]||'').toUpperCase()==='TRABAJO EN DESCANSO'){

      trabajadas=calcularHorasTrabajadas(hoy[3],h)
      extra=trabajadas

    }else{

      const resultado=calcularExtra(
        hoy[3],
        h,
        hObj?.entrada||null,
        hObj?.salida||null
      )

      trabajadas=resultado.trabajadas
      extra=resultado.extra
    }

    await sClient.spreadsheets.values.update({
      spreadsheetId:SPREADSHEET_ID,
      range:`Asistencia!F${idx+1}:M${idx+1}`,
      valueInputOption:'USER_ENTERED',
      requestBody:{
        values:[[
          h,
          hoy[6]||'',
          hoy[7]||'',
          cercana.nombre,
          Math.round(dMin).toString(),
          trabajadas,
          extra,
          jTxt
        ]]
      }
    })

    await sock.sendMessage(jid,{
      text:`✅ Salida - ${nombre} en ${cercana.nombre} - ${h}`
    })

  }catch(e){
    console.error('Error checador:',e)
  }
}


// =====================================================
// REGISTRAR DESCANSOS AUTOMÁTICAMENTE
// =====================================================

export async function registrarDescansos(){

  const ahora=horaMX()
  const [hh]=ahora.split(':').map(Number)

  // No registrar descansos antes de las 06:00
  if(hh<6)return

  const hoy=fechaLaboral()

  if(ultimoRegistroDescansos===hoy)return

  const baseRows=await getRows('Horario_Base!A2:K')
  const asisRows=await getRows('Asistencia!A2:M')

  const dia=diaMexico()
  const sClient=await sheetsClient()

  for(const r of baseRows){

    const tel=(r[0]||'').replace(/\D/g,'').slice(-10)

    if(!tel)continue

    const horario=horarioDelDia(r,dia)

    if(!esDescanso(horario))continue

    const existe=asisRows.findIndex(a=>
      (a[0]||'').replace(/\D/g,'').slice(-10)===tel &&
      a[2]===hoy
    )

    // Si ya existe registro de hoy, no crear otro
    if(existe!==-1){
      continue
    }

    const nombre=r[1]||tel

    const row=[
      tel,
      nombre,
      hoy,
      'DESCANSO',
      'DESCANSO',
      'DESCANSO',
      '',
      '',
      '',
      '',
      '0',
      '0',
      'DESCANSO'
    ]

    await sClient.spreadsheets.values.append({
      spreadsheetId:SPREADSHEET_ID,
      range:'Asistencia!A:M',
      valueInputOption:'USER_ENTERED',
      requestBody:{values:[row]}
    })
  }

  ultimoRegistroDescansos=hoy
}


// =====================================================
// CIERRE AUTOMÁTICO A LAS 16 HORAS
// =====================================================

export async function cerrarSalidasPendientes(){

  const asis=await getRows('Asistencia!A:M')
  const sClient=await sheetsClient()
  const ahora=new Date()

  for(let i=1;i<asis.length;i++){

    const r=asis[i]

    if(!r[2]||!r[3]||r[5])continue

    const est=String(r[4]||'').toUpperCase()

    if(est==='DESCANSO')continue
    if(String(r[3]).toUpperCase()==='DESCANSO')continue

    const fecha=String(r[2]).trim()
    const entrada=String(r[3]).trim()

    const fechaHora=new Date(`${fecha}T${entrada}`)

    if(isNaN(fechaHora.getTime()))continue

    const horas=(ahora-fechaHora)/3600000

    if(horas<CIERRE_AUTO_HORAS)continue

    let trabajadas='0'
    let extra='0'

    const jornada=String(r[12]||'').toLowerCase()

    // ==============================
    // LIBRE / FLEX
    // ==============================

    if(jornada.includes('libre')||jornada.includes('flex')){

      trabajadas='0'
      extra='0'

    }else{

      const entradaProg=(r[12]||'').match(/(\d{1,2}:\d{2})/)
      const salidaProg=(r[12]||'').match(/(?:-|a)\s*(\d{1,2}:\d{2})/)

      const pe=entradaProg?.[1]||null
      const ps=salidaProg?.[1]||null

      if(pe&&ps){

        const resultado=calcularExtra(
          entrada,
          horaMX(),
          pe,
          ps
        )

        trabajadas=resultado.trabajadas
        extra=resultado.extra

      }else{

        trabajadas=calcularHorasTrabajadas(
          entrada,
          horaMX()
        )

        extra='0'
      }
    }

    // La salida automática es exactamente 16 horas
    // después de la entrada real.
    const fechaCierre=new Date(
      fechaHora.getTime()+CIERRE_AUTO_HORAS*3600000
    )

    const cierre=horaMX(fechaCierre)

    await sClient.spreadsheets.values.update({
      spreadsheetId:SPREADSHEET_ID,
      range:`Asistencia!F${i+1}:M${i+1}`,
      valueInputOption:'USER_ENTERED',
      requestBody:{
        values:[[
          cierre,
          r[6]||'',
          r[7]||'',
          r[8]||'',
          r[9]||'',
          trabajadas,
          extra,
          r[12]||''
        ]]
      }
    })

    // Marcar como salida automática
    await sClient.spreadsheets.values.update({
      spreadsheetId:SPREADSHEET_ID,
      range:`Asistencia!I${i+1}:J${i+1}`,
      valueInputOption:'USER_ENTERED',
      requestBody:{
        values:[[
          'SIN SALIDA',
          'AUTO 16H'
        ]]
      }
    })
  }
}


// =====================================================
// AVISO DE PERSONAS QUE NO HAN LLEGADO
// =====================================================

export async function checkNoLlegaron(sock){

  const hoy=fechaLaboral()

  if(fechaAvisos!==hoy){
    avisosHoy.clear()
    fechaAvisos=hoy
  }

  const ahoraMin=minutos(horaMX())

  const baseRows=await getRows('Horario_Base!A2:K')
  const asisRows=await getRows('Asistencia!A2:M')

  const diaNum=diaMexico()

  for(const r of baseRows){

    const tel=(r[0]||'').replace(/\D/g,'').slice(-10)

    if(!tel)continue

    const key=`${hoy}_${tel}`

    if(avisosHoy.has(key))continue

    const v=horarioDelDia(r,diaNum)

    // Día de descanso: no generar aviso
    if(esDescanso(v))continue

    const parsed=parseHorarioRango(v)

    // LIBRE: no generar aviso
    if(!parsed||parsed.entrada==='LIBRE'||!parsed.entrada)continue

    const dif=ahoraMin-minutos(parsed.entrada)

    // Solo después de 20 minutos de la hora programada
    if(dif<AVISO_FALTA_MIN)continue

    const registro=asisRows.find(a=>
      (a[0]||'').replace(/\D/g,'').slice(-10)===tel &&
      a[2]===hoy &&
      a[3] &&
      String(a[3]).toUpperCase()!=='DESCANSO'
    )

    if(registro)continue

    avisosHoy.add(key)

    const nombre=r[1]||tel
    const sucursalAsignada=(r[2]||'').toString().trim()
    const sucId=sucursalAsignada.toLowerCase()

    // =================================================
    // IMPORTANTE:
    // EL AVISO VA AL GRUPO DE GERENTES,
    // NO AL GRUPO DEL CHECADOR.
    //
    // Se toma la sucursal ASIGNADA del empleado.
    // =================================================

    let grupoAviso=null

    if(
      sucId.includes('coyo') ||
      sucId.includes('hotel') ||
      sucId.includes('trinidad')
    ){
      grupoAviso=GRUPO_COYOACAN_ID
    }else{
      grupoAviso=GRUPO_BUCARELI_ID
    }

    if(!grupoAviso)continue

    try{

      await sock.sendMessage(
        grupoAviso,
        {
          text:
`⚠️ NO HA LLEGADO - ${nombre}
Prog: ${parsed.entrada}
Más de 20 min sin registrar entrada
[${sucursalAsignada}]`
        }
      )

    }catch(e){
      console.error('Error enviando aviso de no llegada:',e)
    }
  }
}

import {
  SUCURSALES,
  GRUPO_CHECADOR_COYOACAN_ID,
  GRUPO_CHECADOR_BUCARELI_ID,
  SPREADSHEET_ID
} from './config.js'
import {
  distM,
  fechaLaboral,
  horaMX,
  minutos,
  normaliza,
  calcularHorasTrabajadas,
  calcularExtra,
  parseHorarioRango
} from './utils.js'
import { getRows, sheetsClient, getHorarioBaseMap } from './sheets.js'

const TOLERANCIA_MIN = 15
const AVISO_FALTA_MIN = 20
const CIERRE_AUTO_HORAS = 16

let avisosHoy = new Set()
let fechaAvisos = ''

function esLibre(jornada){
  return /libre|flex/i.test(String(jornada || ''))
}

function fechaHoraMXDesdeFechaHora(fecha, hora){
  const [hh, mm, ss] = String(hora || '00:00:00').split(':').map(Number)
  const d = new Date(`${fecha}T${String(hh || 0).padStart(2,'0')}:${String(mm || 0).padStart(2,'0')}:${String(ss || 0).padStart(2,'0')}`)
  return d
}

function formatearHora(d){
  return d.toLocaleTimeString('es-MX',{
    hour12:false,
    timeZone:'America/Mexico_City'
  })
}

function minutosEntreFechas(a,b){
  return Math.floor((b.getTime()-a.getTime())/60000)
}

function esSinSalida(row){
  return String(row?.[8] || '').toUpperCase().includes('SIN SALIDA') ||
         String(row?.[9] || '').toUpperCase().includes('AUTO 16H')
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
    if(i>-1) emp=empRows[i]
  }

  if(!emp && rawLid.includes('@lid')){
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

  const hoy=idx>-1?asis[idx]:null

  const sClient=await sheetsClient()
  const baseMap=await getHorarioBaseMap()
  const base=
    baseMap[tel10F]||
    baseMap[normaliza(nombre).split(' ')[0]]||
    null

  let estatus='A TIEMPO'
  let hObj=null
  let esRet=false
  let minRet=0

  if(base){
    const fe=new Date(fLab+'T12:00:00')
    const dn=fe.getDay()

    if(base.descansos.has(dn)){
      estatus='DESCANSO'
    }else if(base.horas[dn]){
      hObj=base.horas[dn]

      if(hObj.entrada==='LIBRE'){
        estatus='LIBRE'
      }else if(hObj.entrada){
        const dif=minutos(horaMX())-minutos(hObj.entrada)

        if(dif>TOLERANCIA_MIN){
          estatus=`RETARDO ${dif}min (Prog ${hObj.entrada})`
          esRet=true
          minRet=dif
        }else{
          estatus=`A TIEMPO (Prog ${hObj.entrada})`
        }
      }
    }
  }

  try{
    // =========================
    // ENTRADA
    // =========================
    if(!hoy || !hoy[3]){

      if(dMin>cercana.rEnt){
        await sock.sendMessage(
          jid,
          {text:`Debes estar a max ${cercana.rEnt}m de ${cercana.nombre}`},
          {quoted:m}
        )
        return
      }

      const h=horaMX()

      const jTxt=
        hObj
          ? `${hObj.entrada}${hObj.salida?` - ${hObj.salida}`:''}`
          : '8h'

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
        esLibre(jTxt)?'0':calcularHorasTrabajadas(h,''),
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

      await sock.sendMessage(
        jid,
        {text:`✅ ${estatus} - ${nombre} en ${cercana.nombre} - ${h}`}
      )

      return
    }

    // =========================
    // SALIDA
    // =========================
    if(hoy[5]){
      await sock.sendMessage(
        jid,
        {text:`Salida ya registrada`}
      )
      return
    }

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

    if(!esLibre(jTxt)){
      const calc=calcularExtra(
        hoy[3],
        h,
        hObj?.entrada||null,
        hObj?.salida||null
      )

      trabajadas=calc.trabajadas
      extra=calc.extra
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

    // No mostramos horas ni extra en WhatsApp
    await sock.sendMessage(
      jid,
      {text:`✅ Salida - ${nombre} en ${cercana.nombre} - ${h}`}
    )

  }catch(e){
    console.error('ERROR CHECADOR:',e)
  }
}

// ============================================================
// CIERRE AUTOMÁTICO DESPUÉS DE 16 HORAS
// ============================================================
export async function cerrarSalidasPendientes(){
  try{
    const ahora=new Date()
    const asis=await getRows('Asistencia!A:M')
    const baseMap=await getHorarioBaseMap()
    const sClient=await sheetsClient()

    if(!asis || asis.length<=1) return

    let cerrados=0

    for(let i=1;i<asis.length;i++){
      const r=asis[i]

      if(!r || !r[0] || !r[2] || !r[3]) continue

      // Ya tiene salida real o cierre automático
      if(r[5] || esSinSalida(r)) continue

      const tel=(r[0]||'').replace(/\D/g,'').slice(-10)
      const nombre=r[1]||tel
      const fecha=r[2]
      const entrada=r[3]
      const jornada=r[12]||''

      if(!tel || !fecha || !entrada) continue

      const entradaDate=fechaHoraMXDesdeFechaHora(fecha,entrada)

      if(isNaN(entradaDate.getTime())) continue

      const minutosTranscurridos=minutosEntreFechas(
        entradaDate,
        ahora
      )

      if(minutosTranscurridos < CIERRE_AUTO_HORAS*60) continue

      // ======================================================
      // LIBRE:
      // SIN SALIDA pero 0 horas y 0 extra
      // ======================================================
      let trabajadas='0'
      let extra='0'

      if(!esLibre(jornada)){
        const base=
          baseMap[tel]||
          baseMap[normaliza(nombre).split(' ')[0]]||
          null

        let horaProgramadaEntrada=null
        let horaProgramadaSalida=null

        if(base){
          const fe=new Date(fecha+'T12:00:00')
          const dn=fe.getDay()
          const hObj=base.horas?.[dn]

          if(hObj){
            horaProgramadaEntrada=hObj.entrada||null
            horaProgramadaSalida=hObj.salida||null
          }
        }

        const cierreDate=new Date(
          entradaDate.getTime()+
          CIERRE_AUTO_HORAS*60*60*1000
        )

        const cierreHora=formatearHora(cierreDate)

        const calc=calcularExtra(
          entrada,
          cierreHora,
          horaProgramadaEntrada,
          horaProgramadaSalida
        )

        trabajadas=calc.trabajadas
        extra=calc.extra
      }

      const cierreDate=new Date(
        entradaDate.getTime()+
        CIERRE_AUTO_HORAS*60*60*1000
      )

      const cierreHora=formatearHora(cierreDate)

      // F = salida automática
      // I = SIN SALIDA
      // J = AUTO 16H
      // K = horas trabajadas
      // L = extra
      await sClient.spreadsheets.values.update({
        spreadsheetId:SPREADSHEET_ID,
        range:`Asistencia!F${i+1}:M${i+1}`,
        valueInputOption:'USER_ENTERED',
        requestBody:{
          values:[[
            cierreHora,
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

      // Marcar SIN SALIDA / AUTO 16H
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

      cerrados++

      console.log(
        `AUTO CIERRE 16H: ${nombre} ${fecha} entrada ${entrada} -> salida ${cierreHora} | Trab ${trabajadas} | Extra ${extra}`
      )
    }

    if(cerrados){
      console.log(`AUTO CIERRES REALIZADOS: ${cerrados}`)
    }

  }catch(e){
    console.error('ERROR CIERRE AUTO 16H:',e)
  }
}

// ============================================================
// AVISO DE PERSONAS QUE NO HAN LLEGADO
// ============================================================
export async function checkNoLlegaron(sock){
  const hoy=fechaLaboral()

  if(fechaAvisos!==hoy){
    avisosHoy.clear()
    fechaAvisos=hoy
  }

  const ahoraMin=minutos(horaMX())

  const baseRows=await getRows('Horario_Base!A2:K')
  const asisRows=await getRows('Asistencia!A2:M')

  const mxNow=new Date(
    new Date().toLocaleString(
      'en-US',
      {timeZone:'America/Mexico_City'}
    )
  )

  const diaNum=mxNow.getDay()

  for(const r of baseRows){
    const tel=(r[0]||'').replace(/\D/g,'').slice(-10)

    if(!tel) continue

    const key=`${hoy}_${tel}`

    if(avisosHoy.has(key)) continue

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
    if(parsed.entrada==='LIBRE') continue
    if(!parsed.entrada) continue

    const dif=ahoraMin-minutos(parsed.entrada)

    // Nunca avisar antes de la hora programada
    if(dif<AVISO_FALTA_MIN) continue

    const registro=asisRows.find(a=>
      (a[0]||'').replace(/\D/g,'').slice(-10)===tel &&
      a[2]===hoy &&
      a[3]
    )

    if(registro) continue

    avisosHoy.add(key)

    const nombre=r[1]||tel
    const sucId=(r[2]||'').toLowerCase()

    let grupoAviso=null

    if(
      sucId.includes('coyo')||
      sucId.includes('hotel')||
      sucId.includes('trinidad')
    ){
      grupoAviso=GRUPO_CHECADOR_COYOACAN_ID
    }else{
      grupoAviso=GRUPO_CHECADOR_BUCARELI_ID
    }

    if(grupoAviso){
      try{
        await sock.sendMessage(
          grupoAviso,
          {
            text:
              `⚠️ *NO HA LLEGADO* - ${nombre}\n`+
              `Prog: ${parsed.entrada}\n`+
              `Más de 20 min sin registrar entrada\n`+
              `[${r[2]||''}]`
          }
        )
      }catch(e){
        console.error('ERROR AVISO NO LLEGADA:',e)
      }
    }
  }
}

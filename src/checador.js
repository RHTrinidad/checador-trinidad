import { SUCURSALES, GRUPO_COYOACAN_ID, GRUPO_BUCARELI_ID, GRUPO_CHECADOR_COYOACAN_ID, GRUPO_CHECADOR_BUCARELI_ID, SPREADSHEET_ID } from './config.js'
import { distM, fechaLaboral, horaMX, minutos, normaliza, calcularHorasTrabajadas, calcularExtra, parseHorarioRango } from './utils.js'
import { getRows, sheetsClient, getHorarioBaseMap } from './sheets.js'

const TOLERANCIA_MIN = 15
const AVISO_FALTA_MIN = 20
const AVISO_RETARDO_60_MIN = 60
const FECHA_CORTE_RETARDOS = '2026-10-02'
const HOJA_ALERTAS_RETARDOS = 'Alertas_Retardos'

let avisosHoy = new Set()
let fechaAvisos = ''
let avisos60Hoy = new Set()
let fechaAvisos60 = ''

// Determina el grupo de gerentes según la sucursal asignada al empleado.
function grupoGerentesPorSucursal(sucursal){
  const x = normaliza(sucursal)
  if(x.includes('coyo') || x.includes('hotel') || x.includes('trinidad')) return GRUPO_COYOACAN_ID
  if(x.includes('bucareli') || x.includes('juarez')) return GRUPO_BUCARELI_ID
  return null
}

// Crea la hoja de control para conservar los avisos de 3 retardos aunque Railway reinicie.
async function asegurarHojaAlertasRetardos(){
  const client = await sheetsClient()
  const meta = await client.spreadsheets.get({spreadsheetId:SPREADSHEET_ID})
  const existe = (meta.data.sheets||[]).some(s=>s.properties?.title===HOJA_ALERTAS_RETARDOS)
  if(!existe){
    await client.spreadsheets.batchUpdate({
      spreadsheetId:SPREADSHEET_ID,
      requestBody:{requests:[{addSheet:{properties:{title:HOJA_ALERTAS_RETARDOS}}}]}
    })
    await client.spreadsheets.values.update({
      spreadsheetId:SPREADSHEET_ID,
      range:`${HOJA_ALERTAS_RETARDOS}!A1:E1`,
      valueInputOption:'USER_ENTERED',
      requestBody:{values:[['FechaAviso','Telefono','Nombre','Retardos','Detalle']]}
    })
  }
}

// Convierte una fecha laboral en número de día para poder calcular días hábiles.
function fechaObj(s){
  const d = new Date(`${s}T12:00:00`)
  return isNaN(d) ? null : d
}

function diasHabilesEntre(inicio,fin){
  const a=fechaObj(inicio), b=fechaObj(fin)
  if(!a||!b) return 0
  let n=0
  for(let d=new Date(a);d<=b;d.setDate(d.getDate()+1)){
    const dia=d.getDay()
    if(dia!==0&&dia!==6) n++
  }
  return n
}

function esDentroDe30DiasHabiles(fecha,hoy){
  const a=fechaObj(fecha), b=fechaObj(hoy)
  if(!a||!b||a>b) return false
  let n=0
  for(let d=new Date(a);d<=b;d.setDate(d.getDate()+1)){
    const dia=d.getDay()
    if(dia!==0&&dia!==6) n++
  }
  return n<=30
}

// Revisa los retardos acumulados y genera el aviso al llegar a 3.
async function revisarTresRetardos(sock,tel,nombre,sucursal){
  if(!tel) return
  const hoy=fechaLaboral()
  if(hoy<FECHA_CORTE_RETARDOS) return

  const rows=await getRows('Asistencia!A2:M')
  const retardos=rows.filter(r=>{
    const t=(r[0]||'').replace(/\D/g,'').slice(-10)
    const f=(r[2]||'').toString()
    const m=(r[4]||'').toString().match(/RETARDO\s+(\d+)\s*min/i)
    return t===tel&&f>=FECHA_CORTE_RETARDOS&&f<=hoy&&esDentroDe30DiasHabiles(f,hoy)&&m
  }).map(r=>({fecha:r[2],min:Number((r[4].toString().match(/RETARDO\s+(\d+)\s*min/i)||[])[1]||0)}))

  if(retardos.length<3) return

  const alertaRows=await getRows(`${HOJA_ALERTAS_RETARDOS}!A2:E`)
  const anteriores=alertaRows.filter(r=>(r[1]||'').replace(/\D/g,'').slice(-10)===tel)
  const ultima=anteriores.length?anteriores[anteriores.length-1]:null
  const fechaUltima=ultima?.[0]||''

  // El último aviso funciona como punto de reinicio: solo cuentan retardos posteriores.
  const nuevos=fechaUltima?retardos.filter(x=>x.fecha>fechaUltima):retardos
  if(nuevos.length<3) return

  const grupo=grupoGerentesPorSucursal(sucursal)
  if(!grupo) return

  const detalle=nuevos.slice(0,3).map(x=>`${x.fecha}: ${x.min}min`).join(' | ')
  try{
    await sock.sendMessage(grupo,{text:`⚠️ 3 RETARDOS - ${nombre}\nSe acumularon 3 retardos dentro de los últimos 30 días hábiles.\n${detalle}\n[${sucursal}]`})
    const client=await sheetsClient()
    await client.spreadsheets.values.append({
      spreadsheetId:SPREADSHEET_ID,
      range:`${HOJA_ALERTAS_RETARDOS}!A:E`,
      valueInputOption:'USER_ENTERED',
      requestBody:{values:[[hoy,tel,nombre,'3',detalle]]}
    })
  }catch(e){ console.error('ERROR AVISO 3 RETARDOS:',e.message) }
}

export async function handleChecador({sock,jid,m,loc,rawLid,tel10,tel}){
  const lat=loc.degreesLatitude,lng=loc.degreesLongitude
  let cercana=null,dMin=Infinity

  // La ubicación recibida determina dónde se registra y a qué grupo se confirma.
  for(const s of SUCURSALES){
    const d=distM(lat,lng,s.lat,s.lng)
    if(d<dMin){dMin=d;cercana=s}
  }

  const empRows=await getRows('Empleados!A:K')
  const getLid=r=>(r.find(x=>String(x).includes('@lid'))||'').trim()
  let emp=null

  if(tel10.length>=10){
    const i=empRows.findIndex((r,idx)=>idx>0&&r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel10)
    if(i>-1) emp=empRows[i]
  }

  if(!emp&&rawLid.includes('@lid')){
    const i=empRows.findIndex((r,idx)=>idx>0&&getLid(r)===rawLid)
    if(i>-1){
      emp=empRows[i]
      tel10=(emp[0]||'').replace(/\D/g,'').slice(-10)
    }
  }

  const nombre=emp?(emp[3]||emp[1]):(m.pushName||tel10||'Desconocido')
  const telF=emp?(emp[0]||'').replace(/\D/g,''):tel
  const tel10F=telF.slice(-10)||tel10
  const sucursalAsignada=emp?.[2]||cercana?.nombre||''

  // La fecha se determina con jornada laboral 05:00-04:59.
  const fLab=fechaLaboral()
  const asis=await getRows('Asistencia!A:M')
  const idx=asis.findIndex((r,i)=>i>0&&r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel10F&&r[2]===fLab)
  const hoy=idx>-1?asis[idx]:null

  const sClient=await sheetsClient()
  const baseMap=await getHorarioBaseMap()
  const base=baseMap[tel10F]||baseMap[normaliza(nombre).split(' ')[0]]||null

  let estatus='A TIEMPO',hObj=null,esRet=false,minRet=0

  if(base){
    const fe=new Date(`${fLab}T12:00:00`)
    const dn=fe.getDay()

    if(base.descansos.has(dn)) estatus='DESCANSO'
    else if(base.horas[dn]){
      hObj=base.horas[dn]

      if(hObj.entrada!=='LIBRE'&&hObj.entrada){
        let ahora=minutos(horaMX())
        let prog=minutos(hObj.entrada)

        // Si la hora actual está entre 00:00 y 04:59, pertenece al final de la jornada laboral.
        if(ahora<300) ahora+=1440
        if(prog<300) prog+=1440

        const dif=ahora-prog

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
    if(!hoy||!hoy[3]){
      if(dMin>cercana.rEnt){
        await sock.sendMessage(jid,{text:`Debes estar a max ${cercana.rEnt}m de ${cercana.nombre}`},{quoted:m})
        return
      }

      const h=horaMX()
      const jTxt=hObj?`${hObj.entrada}${hObj.salida?` - ${hObj.salida}`:''}`:'8h'
      const row=[tel10F,nombre,fLab,h,estatus,'',cercana.nombre,Math.round(dMin).toString(),'','',calcularHorasTrabajadas(h,''),'0',jTxt]

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

      await sock.sendMessage(jid,{text:`✅ ${estatus} - ${nombre} en ${cercana.nombre} - ${h}`})

      // Todo retardo mayor a 15 min genera aviso a gerentes.
      if(esRet){
        const grupo=grupoGerentesPorSucursal(sucursalAsignada)
        if(grupo){
          let texto=`⚠️ RETARDO ${minRet}min - ${nombre}\nProg: ${hObj?.entrada||''}\nLlegó: ${h}\n[${sucursalAsignada}]`

          if(minRet>=AVISO_RETARDO_60_MIN){
            texto=`⚠️ RETARDO ${minRet}min - ${nombre}\nProg: ${hObj?.entrada||''}\nLlegó: ${h}\n⚠️ Favor de confirmar si el empleado se queda / trabaja su turno.\n[${sucursalAsignada}]`
          }

          const key60=`${fLab}_${tel10F}`
          if(minRet>=AVISO_RETARDO_60_MIN){
            if(fechaAvisos60!==fLab){avisos60Hoy.clear();fechaAvisos60=fLab}
            if(!avisos60Hoy.has(key60)){
              avisos60Hoy.add(key60)
              await sock.sendMessage(grupo,{text:texto})
            }
          }else{
            await sock.sendMessage(grupo,{text:texto})
          }
        }

        await revisarTresRetardos(sock,tel10F,nombre,sucursalAsignada)
      }
    }else{
      if(hoy[5]){
        await sock.sendMessage(jid,{text:`Salida ya registrada`})
        return
      }

      if(dMin>cercana.rSal){
        await sock.sendMessage(jid,{text:`No puedes checar salida a ${Math.round(dMin)}m`},{quoted:m})
        return
      }

      const h=horaMX()
      const jTxt=hoy[12]||'8h'
      const {trabajadas,extra}=calcularExtra(hoy[3],h,hObj?.entrada||null,hObj?.salida||null)

      await sClient.spreadsheets.values.update({
        spreadsheetId:SPREADSHEET_ID,
        range:`Asistencia!F${idx+1}:M${idx+1}`,
        valueInputOption:'USER_ENTERED',
        requestBody:{values:[[h,hoy[6]||'',hoy[7]||'',cercana.nombre,Math.round(dMin).toString(),trabajadas,extra,jTxt]]}
      })

      // WhatsApp solo confirma la salida. Trab/Extra quedan en Sheets.
      await sock.sendMessage(jid,{text:`✅ Salida - ${nombre} en ${cercana.nombre} - ${h}`})
    }
  }catch(e){
    console.error('ERROR CHECADOR:',e)
  }
}

export async function checkNoLlegaron(sock){
  const hoy=fechaLaboral()

  if(fechaAvisos!==hoy){
    avisosHoy.clear()
    fechaAvisos=hoy
  }

  const ahoraMinBase=minutos(horaMX())
  const ahoraMin=ahoraMinBase<300?ahoraMinBase+1440:ahoraMinBase

  const baseRows=await getRows('Horario_Base!A2:K')
  const asisRows=await getRows('Asistencia!A2:M')

  // El día de horario se toma de la fecha laboral, no de la fecha calendario.
  const diaNum=new Date(`${hoy}T12:00:00`).getDay()

  for(const r of baseRows){
    const tel=(r[0]||'').replace(/\D/g,'').slice(-10)
    if(!tel) continue

    const key=`${hoy}_${tel}`
    if(avisosHoy.has(key)) continue

    const mapa={1:r[3],2:r[4],3:r[5],4:r[6],5:r[7],6:r[8],0:r[9]}
    const v=(mapa[diaNum]||'').toString().trim()
    const parsed=parseHorarioRango(v)

    if(!parsed||parsed.entrada==='LIBRE'||!parsed.entrada) continue

    let progMin=minutos(parsed.entrada)
    if(progMin<300) progMin+=1440

    const dif=ahoraMin-progMin

    // No avisar antes de la hora programada.
    if(dif<AVISO_FALTA_MIN) continue

    const registro=asisRows.find(a=>
      (a[0]||'').replace(/\D/g,'').slice(-10)===tel &&
      a[2]===hoy &&
      a[3]
    )

    if(registro) continue

    avisosHoy.add(key)

    const nombre=r[1]||tel
    const sucursal=r[2]||''
    const grupoAviso=grupoGerentesPorSucursal(sucursal)

    if(grupoAviso){
      try{
        await sock.sendMessage(grupoAviso,{
          text:`⚠️ NO HA LLEGADO - ${nombre}\nProg: ${parsed.entrada}\nMás de 20 min sin registrar entrada\n[${sucursal}]`
        })
      }catch(e){
        console.error('ERROR AVISO NO LLEGADA:',e.message)
      }
    }
  }
}

import { SUCURSALES, GRUPO_CHECADOR_COYOACAN_ID, GRUPO_CHECADOR_BUCARELI_ID, GRUPO_COYOACAN_ID, GRUPO_BUCARELI_ID, SPREADSHEET_ID } from './config.js'
import { distM, fechaLaboral, horaMX, minutos, normaliza, calcularHorasTrabajadas, calcularExtra, parseHorarioRango } from './utils.js'
import { getRows, sheetsClient, getHorarioBaseMap } from './sheets.js'

const TOLERANCIA_MIN=15
const AVISO_FALTA_MIN=20
const AVISO_RETARDO_60_MIN=60
const FECHA_CORTE_RETARDOS='2026-10-02'
const HOJA_ALERTAS_RETARDOS='Alertas_Retardos'
const HOJA_AVISOS='Avisos'

const normalizaTel=v=>(v||'').toString().replace(/\D/g,'').slice(-10)

function grupoGerentesPorSucursal(sucursal){
  const x=normaliza(sucursal)
  if(x.includes('coyo')||x.includes('hotel')||x.includes('trinidad'))return GRUPO_COYOACAN_ID
  if(x.includes('juarez')||x.includes('bucareli'))return GRUPO_BUCARELI_ID
  return null
}

function horarioLaboralMin(h){
  const m=minutos(h)
  return m<300?m+1440:m
}

function calcularRetardo(hEntrada,hProg){
  if(!hEntrada||!hProg||hProg==='LIBRE')return 0
  return horarioLaboralMin(hEntrada)-horarioLaboralMin(hProg)
}

async function asegurarHojaAvisos(sClient){
  try{
    await sClient.spreadsheets.values.get({spreadsheetId:SPREADSHEET_ID,range:`${HOJA_AVISOS}!A:D`})
  }catch(e){
    try{
      await sClient.spreadsheets.batchUpdate({
        spreadsheetId:SPREADSHEET_ID,
        requestBody:{requests:[{addSheet:{properties:{title:HOJA_AVISOS}}}]}
      })
      await sClient.spreadsheets.values.update({
        spreadsheetId:SPREADSHEET_ID,
        range:`${HOJA_AVISOS}!A1:D1`,
        valueInputOption:'USER_ENTERED',
        requestBody:{values:[['Tel','Fecha','HoraAviso','Nombre Corto']]}
      })
    }catch(err){
      if(!String(err.message||'').toLowerCase().includes('already exists'))throw err
    }
  }
}

async function yaAvisadoNoLlegada(sClient,tel,fecha){
  await asegurarHojaAvisos(sClient)
  const r=await sClient.spreadsheets.values.get({spreadsheetId:SPREADSHEET_ID,range:`${HOJA_AVISOS}!A:D`})
  const rows=r.data.values||[]
  const tel10=normalizaTel(tel)
  return rows.slice(1).some(x=>normalizaTel(x[0])===tel10&&String(x[1]||'').trim()===fecha)
}

async function registrarAvisoNoLlegada(sClient,tel,fecha,hora,nombreCorto){
  await asegurarHojaAvisos(sClient)
  await sClient.spreadsheets.values.append({
    spreadsheetId:SPREADSHEET_ID,
    range:`${HOJA_AVISOS}!A:D`,
    valueInputOption:'USER_ENTERED',
    requestBody:{values:[[normalizaTel(tel),fecha,hora,nombreCorto||'']]}
  })
}

async function asegurarHojaAlertasRetardos(sClient){
  try{
    await sClient.spreadsheets.values.get({spreadsheetId:SPREADSHEET_ID,range:`${HOJA_ALERTAS_RETARDOS}!A:E`})
  }catch(e){
    try{
      await sClient.spreadsheets.batchUpdate({
        spreadsheetId:SPREADSHEET_ID,
        requestBody:{requests:[{addSheet:{properties:{title:HOJA_ALERTAS_RETARDOS}}}]}
      })
      await sClient.spreadsheets.values.update({
        spreadsheetId:SPREADSHEET_ID,
        range:`${HOJA_ALERTAS_RETARDOS}!A1:E1`,
        valueInputOption:'USER_ENTERED',
        requestBody:{values:[['Tel','Nombre Corto','Fecha Aviso','Retardos Contados','Fechas Retardos']]}
      })
    }catch(err){
      if(!String(err.message||'').toLowerCase().includes('already exists'))throw err
    }
  }
}

async function revisarTresRetardos(sock){
  try{
    const sClient=await sheetsClient()
    await asegurarHojaAlertasRetardos(sClient)
    const [asisRes,baseMap,alertRes,empRows]=await Promise.all([
      sClient.spreadsheets.values.get({spreadsheetId:SPREADSHEET_ID,range:'Asistencia!A2:M'}),
      getHorarioBaseMap(),
      sClient.spreadsheets.values.get({spreadsheetId:SPREADSHEET_ID,range:`${HOJA_ALERTAS_RETARDOS}!A:E`}),
      getRows('Empleados!A:K')
    ])
    const asis=asisRes.data.values||[], alertas=alertRes.data.values||[], empleados={}
    empRows.slice(1).forEach(r=>{
      const tel=normalizaTel(r[0])
      if(tel)empleados[tel]={corto:r[1]||'',completo:r[3]||r[1]||'',suc:r[2]||''}
    })
    const porEmpleado={}
    for(const f of asis){
      const tel=normalizaTel(f[0]),fecha=String(f[2]||'').trim(),entrada=f[3]||''
      if(!tel||!fecha||fecha<FECHA_CORTE_RETARDOS||!entrada)continue
      const info=empleados[tel]||{}
      const prog=baseMap[tel]?.horas?.[parseFechaDia(fecha)]
      const ret=calcularRetardo(entrada,prog?.entrada)
      if(ret>15){
        if(!porEmpleado[tel])porEmpleado[tel]=[]
        porEmpleado[tel].push({fecha,min:ret,nombre:info.corto||f[1]||tel,suc:info.suc||f[6]||''})
      }
    }
    for(const tel of Object.keys(porEmpleado)){
      const lista=porEmpleado[tel].sort((a,b)=>a.fecha.localeCompare(b.fecha))
      const prev=alertas.slice(1).filter(r=>normalizaTel(r[0])===tel)
      const ultimo=prev.length?prev[prev.length-1]:null
      const desde=ultimo?.[4]?String(ultimo[4]).split(',').map(x=>x.trim()).filter(Boolean):[]
      const nuevos=lista.filter(x=>!desde.includes(x.fecha))
      if(nuevos.length<3)continue
      const grupo=grupoGerentesPorSucursal(lista[0].suc)
      if(!grupo)continue
      const ultimos=nuevos.slice(0,3)
      const nombre=lista[0].nombre
      const detalle=ultimos.map(x=>`${x.fecha}: ${x.min} min`).join('\n')
      try{
        await sock.sendMessage(grupo,{text:`⚠️ *3 RETARDOS*\n\n${nombre}\n${detalle}\n\nFavor de revisar y dar seguimiento.`})
        const fechas=[...desde,...ultimos.map(x=>x.fecha)]
        await sClient.spreadsheets.values.append({
          spreadsheetId:SPREADSHEET_ID,
          range:`${HOJA_ALERTAS_RETARDOS}!A:E`,
          valueInputOption:'USER_ENTERED',
          requestBody:{values:[[tel,nombre,fechaActualMX(),3,fechas.join(', ')]]}
        })
      }catch(e){console.error('ERROR ALERTA 3 RETARDOS:',e.message)}
    }
  }catch(e){console.error('ERROR REVISAR 3 RETARDOS:',e)}
}

function parseFechaDia(fecha){
  const d=new Date(`${fecha}T12:00:00`)
  return d.getDay()
}

function fechaActualMX(){
  return fechaLaboral(new Date())
}

export async function handleChecador({sock,jid,m,loc,rawLid,tel10,tel}){
  const lat=loc.degreesLatitude,lng=loc.degreesLongitude
  let cercana=null,dMin=Infinity
  for(const s of SUCURSALES){
    const d=distM(lat,lng,s.lat,s.lng)
    if(d<dMin){dMin=d;cercana=s}
  }

  const empRows=await getRows('Empleados!A:K')
  const getLid=r=>(r.find(x=>String(x).includes('@lid'))||'').trim()
  let emp=null
  tel10=normalizaTel(tel10||tel)

  if(tel10.length>=10){
    const i=empRows.findIndex((r,idx)=>idx>0&&normalizaTel(r[0])===tel10)
    if(i>-1)emp=empRows[i]
  }

  if(!emp&&rawLid?.includes('@lid')){
    const i=empRows.findIndex((r,idx)=>idx>0&&getLid(r)===rawLid)
    if(i>-1){
      emp=empRows[i]
      tel10=normalizaTel(emp[0])
    }
  }

  const nombreCorto=emp?(emp[1]||emp[3]||tel10):(m.pushName||tel10||'Desconocido')
  const nombreCompleto=emp?(emp[3]||emp[1]||nombreCorto):nombreCorto
  const telF=emp?normalizaTel(emp[0]):normalizaTel(tel10||tel)
  const fLab=fechaLaboral()
  const asis=await getRows('Asistencia!A:M')
  const idx=asis.findIndex((r,i)=>i>0&&normalizaTel(r[0])===telF&&r[2]===fLab)
  const hoy=idx>-1?asis[idx]:null

  const sClient=await sheetsClient()
  const baseMap=await getHorarioBaseMap()
  const base=baseMap[telF]||baseMap[normaliza(nombreCompleto).split(' ')[0]]||null

  let estatus='A TIEMPO',hObj=null,esRet=false,minRet=0,hProg=''
  if(base){
    const fe=new Date(fLab+'T12:00:00'),dn=fe.getDay()
    if(base.descansos.has(dn))estatus='DESCANSO'
    else if(base.horas[dn]){
      hObj=base.horas[dn];hProg=hObj.entrada||''
      if(hObj.entrada!=='LIBRE'&&hObj.entrada){
        let ahora=minutos(horaMX()),prog=minutos(hObj.entrada)
        if(ahora<300)ahora+=1440
        if(prog<300)prog+=1440
        const dif=ahora-prog
        if(dif>TOLERANCIA_MIN){
          estatus=`RETARDO ${dif}min (Prog ${hObj.entrada})`
          esRet=true;minRet=dif
        }else estatus=`A TIEMPO (Prog ${hObj.entrada})`
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
      const row=[telF,nombreCompleto,fLab,h,estatus,'',cercana.nombre,Math.round(dMin).toString(),'','',calcularHorasTrabajadas(h,''),'0',jTxt]

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

      await sock.sendMessage(jid,{text:`✅ ${estatus} - ${nombreCorto} en ${cercana.nombre} - ${h}`})

      if(esRet){
        const grupo=grupoGerentesPorSucursal(emp?.[2]||cercana.nombre)
        if(grupo){
          let extra60=''
          if(minRet>=AVISO_RETARDO_60_MIN)extra60=`\n\n⚠️ Lleva ${minRet} min de retardo. Favor de confirmar si permanece a trabajar.`
          await sock.sendMessage(grupo,{text:`⚠️ *RETARDO* - ${nombreCorto}\nProg: ${hProg}\nEntrada: ${h}\nRetardo: ${minRet} min\n[${emp?.[2]||cercana.nombre}]${extra60}`})
        }
      }

      if(esRet)await revisarTresRetardos(sock)
    }else{
      if(hoy[5]){
        await sock.sendMessage(jid,{text:`Salida ya registrada`})
        return
      }

      if(dMin>cercana.rSal){
        await sock.sendMessage(jid,{text:`No puedes checar salida a ${Math.round(dMin)}m`},{quoted:m})
        return
      }

      const h=horaMX(),jTxt=hoy[12]||'8h'
      const {trabajadas,extra}=calcularExtra(hoy[3],h,hObj?.entrada||null,hObj?.salida||null)

      await sClient.spreadsheets.values.update({
        spreadsheetId:SPREADSHEET_ID,
        range:`Asistencia!F${idx+1}:M${idx+1}`,
        valueInputOption:'USER_ENTERED',
        requestBody:{values:[[h,hoy[6]||'',hoy[7]||'',cercana.nombre,Math.round(dMin).toString(),trabajadas,extra,jTxt]]}
      })

      await sock.sendMessage(jid,{text:`✅ Salida - ${nombreCorto} en ${cercana.nombre} - ${h}`})
    }
  }catch(e){
    console.error('ERROR CHECADOR:',e)
  }
}

export async function checkNoLlegaron(sock){
  try{
    const hoy=fechaLaboral()
    const ahora=horaMX()
    const ahoraMinRaw=minutos(ahora)
    let ahoraMin=ahoraMinRaw
    if(ahoraMin<300)ahoraMin+=1440

    const sClient=await sheetsClient()
    const [baseRows,asisRows,empRows]=await Promise.all([
      getRows('Horario_Base!A2:K'),
      getRows('Asistencia!A2:M'),
      getRows('Empleados!A:K')
    ])

    await asegurarHojaAvisos(sClient)

    const empMap={}
    empRows.slice(1).forEach(r=>{
      const tel=normalizaTel(r[0])
      if(tel)empMap[tel]={corto:r[1]||'',completo:r[3]||r[1]||'',suc:r[2]||''}
    })

    const diaNum=new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'})).getDay()

    for(const r of baseRows){
      const tel=normalizaTel(r[0])
      if(!tel)continue

      const mapa={1:r[3],2:r[4],3:r[5],4:r[6],5:r[7],6:r[8],0:r[9]}
      const v=(mapa[diaNum]||'').toString().trim()
      const parsed=parseHorarioRango(v)
      if(!parsed||parsed.entrada==='LIBRE'||!parsed.entrada)continue

      let progMin=minutos(parsed.entrada)
      if(progMin<300)progMin+=1440

      const dif=ahoraMin-progMin
      if(dif<AVISO_FALTA_MIN)continue

      const registro=asisRows.find(a=>normalizaTel(a[0])===tel&&a[2]===hoy&&a[3])
      if(registro)continue

      const yaAvisado=await yaAvisadoNoLlegada(sClient,tel,hoy)
      if(yaAvisado)continue

      const info=empMap[tel]||{}
      const nombreCorto=info.corto||r[1]||tel
      const suc=info.suc||r[2]||''
      const grupo=grupoGerentesPorSucursal(suc)
      if(!grupo)continue

      const horaAviso=horaMX()

      try{
        await sock.sendMessage(grupo,{
          text:`⚠️ *NO HA LLEGADO* - ${nombreCorto}\nProg: ${parsed.entrada}\nMás de 20 min sin registrar entrada\n[${suc}]`
        })
        await registrarAvisoNoLlegada(sClient,tel,hoy,horaAviso,nombreCorto)
      }catch(e){
        console.error('ERROR AVISO NO LLEGADA:',e.message)
      }
    }
  }catch(e){
    console.error('ERROR CHECK NO LLEGARON:',e)
  }
}

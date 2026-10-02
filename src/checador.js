import { SUCURSALES, GRUPO_CHECADOR_COYOACAN_ID, GRUPO_CHECADOR_BUCARELI_ID, SPREADSHEET_ID } from './config.js' // <- FIX 1: importar checadores, no gerentes
import { distM, fechaLaboral, horaMX, minutos, normaliza, calcularHorasTrabajadas, calcularExtra, parseHorarioRango } from './utils.js'
import { getRows, sheetsClient, getHorarioBaseMap } from './sheets.js'

const TOLERANCIA_MIN = 15
const AVISO_FALTA_MIN = 20

let avisosHoy = new Set()
let fechaAvisos = ''

export async function handleChecador({ sock, jid, m, loc, rawLid, tel10, tel }){
  const lat=loc.degreesLatitude,lng=loc.degreesLongitude; let cercana=null,dMin=Infinity;
  for(const s of SUCURSALES){ const d=distM(lat,lng,s.lat,s.lng); if(d<dMin){dMin=d; cercana=s} }
  const empRows=await getRows('Empleados!A:K'); const getLid=r=>(r.find(x=>String(x).includes('@lid'))||'').trim();
  let emp=null; if(tel10.length>=10){ const i=empRows.findIndex((r,idx)=>idx>0&&r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel10); if(i>-1) emp=empRows[i] }
  if(!emp&&rawLid.includes('@lid')){ const i=empRows.findIndex((r,idx)=>idx>0&&getLid(r)===rawLid); if(i>-1){emp=empRows[i]; tel10=(emp[0]||'').replace(/\D/g,'').slice(-10)} }
  const nombre=emp?(emp[3]||emp[1]):(m.pushName||tel10||'Desconocido'); const telF=emp?(emp[0]||'').replace(/\D/g,''):tel; const tel10F=telF.slice(-10)||tel10;
  const fLab=fechaLaboral(); const asis=await getRows('Asistencia!A:M'); const idx=asis.findIndex((r,i)=>i>0&&r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel10F&&r[2]===fLab); const hoy=idx>-1?asis[idx]:null;
  const sClient=await sheetsClient(); const baseMap=await getHorarioBaseMap(); const base=baseMap[tel10F]||baseMap[normaliza(nombre).split(' ')[0]]||null;
  let estatus='A TIEMPO', hObj=null, esRet=false, minRet=0, hProg='';
  if(base){
    const fe=new Date(fLab+'T12:00:00'); const dn=fe.getDay();
    if(base.descansos.has(dn)) estatus='DESCANSO';
    else if(base.horas[dn]){
      hObj=base.horas[dn]; hProg=hObj.entrada||'';
      if(hObj.entrada!=='LIBRE' && hObj.entrada){
        const dif=minutos(horaMX())-minutos(hObj.entrada);
        if(dif > TOLERANCIA_MIN){
          estatus=`RETARDO ${dif}min (Prog ${hObj.entrada})`;
          esRet=true;
          minRet=dif
        } else {
          estatus=`A TIEMPO (Prog ${hObj.entrada})`
        }
      }
    }
  }
  try{
    if(!hoy||!hoy[3]){
      if(dMin>cercana.rEnt){ await sock.sendMessage(jid,{text:`Debes estar a max ${cercana.rEnt}m de ${cercana.nombre}`},{quoted:m}); return }
      const h=horaMX(); const jTxt=hObj?`${hObj.entrada}${hObj.salida?` - ${hObj.salida}`:''}`:"8h"; const row=[tel10F,nombre,fLab,h,estatus,'',cercana.nombre,Math.round(dMin).toString(),'','',calcularHorasTrabajadas(h,""),"0",jTxt];
      if(idx===-1) await sClient.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Asistencia!A:M',valueInputOption:'USER_ENTERED',requestBody:{values:[row]}}); else await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!A${idx+1}:M${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[row]}});
      await sock.sendMessage(jid,{text:`✅ ${estatus} - ${nombre} en ${cercana.nombre} - ${h}`});
    }else{
      if(hoy[5]){ await sock.sendMessage(jid,{text:`Salida ya registrada`}); return }
      if(dMin>cercana.rSal){ await sock.sendMessage(jid,{text:`No puedes checar salida a ${Math.round(dMin)}m`},{quoted:m}); return }
      const h=horaMX(); const jTxt=hoy[12]||"8h"; const {trabajadas,extra}=calcularExtra(hoy[3],h,hObj?.entrada||null,hObj?.salida||null);
      await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!F${idx+1}:M${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[[h,hoy[6]||'',hoy[7]||'',cercana.nombre,Math.round(dMin).toString(),trabajadas,extra,jTxt]]}});
      await sock.sendMessage(jid,{text:`✅ Salida - ${nombre} en ${cercana.nombre} - ${h} - Trab ${trabajadas} Extra ${extra}`});
    }
  }catch(e){ console.error(e) }
}

export async function checkNoLlegaron(sock){
  const hoy = fechaLaboral()
  if(fechaAvisos!== hoy){ avisosHoy.clear(); fechaAvisos = hoy }
  const ahoraMin = minutos(horaMX())
  const baseRows = await getRows('Horario_Base!A2:K')
  const asisRows = await getRows('Asistencia!A2:M')
  const diaNum = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'})).getDay()
  for(const r of baseRows){
    const tel = (r[0]||'').replace(/\D/g,'').slice(-10)
    if(!tel) continue
    const key = `${hoy}_${tel}`
    if(avisosHoy.has(key)) continue
    const mapa = {1:r[3],2:r[4],3:r[5],4:r[6],5:r[7],6:r[8],0:r[9]}
    const v = (mapa[diaNum]||'').toString().trim()
    const parsed = parseHorarioRango(v)
    if(!parsed || parsed.entrada === 'LIBRE' ||!parsed.entrada) continue
    const dif = ahoraMin - minutos(parsed.entrada)
    if(dif >= AVISO_FALTA_MIN){
      const registro = asisRows.find(a => (a[0]||'').replace(/\D/g,'').slice(-10)===tel && a[2]===hoy && a[3])
      if(!registro){
        avisosHoy.add(key)
        const nombre = r[1]||tel
        const sucId = (r[2]||'').toLowerCase()
        let grupoAviso = null
        if(sucId.includes('coyo') || sucId.includes('hotel') || sucId.includes('trinidad')) grupoAviso = GRUPO_CHECADOR_COYOACAN_ID // <- FIX 2
        else grupoAviso = GRUPO_CHECADOR_BUCARELI_ID // <- FIX 3
        if(grupoAviso){
          try{
            await sock.sendMessage(grupoAviso, {text: `⚠️ NO HA LLEGADO - ${nombre} - Prog ${parsed.entrada} - ${dif}min tarde - [${r[2]}]`})
          }catch{}
        }
      }
    }
  }
}

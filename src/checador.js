import { SUCURSALES, GRUPO_COYOACAN_ID, GRUPO_BUCARELI_ID, SPREADSHEET_ID } from './config.js'
import { distM, fechaLaboral, horaMX, minutos, normaliza, calcularHorasTrabajadas, calcularExtra } from './utils.js'
import { getRows, sheetsClient, getHorarioBaseMap } from './sheets.js'

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
  if(base){ const fe=new Date(fLab+'T12:00:00'); const dn=fe.getDay(); if(base.descansos.has(dn)) estatus='DESCANSO'; else if(base.horas[dn]){ hObj=base.horas[dn]; hProg=hObj.entrada||''; if(hObj.entrada!=='LIBRE'){ const dif=minutos(horaMX())-minutos(hObj.entrada); if(dif>1){estatus=`RETARDO ${dif}min (Prog ${hObj.entrada})`; esRet=true; minRet=dif} else estatus=`A TIEMPO Prog ${hObj.entrada}`} } }
  try{
    if(!hoy||!hoy[3]){
      if(dMin>cercana.rEnt){ await sock.sendMessage(jid,{text:`Debes estar a max ${cercana.rEnt}m de ${cercana.nombre}`},{quoted:m}); return }
      const h=horaMX(); const jTxt=hObj?`${hObj.entrada}${hObj.salida?` - ${hObj.salida}`:''}`:"8h"; const row=[tel10F,nombre,fLab,h,estatus,'',cercana.nombre,Math.round(dMin).toString(),'','',calcularHorasTrabajadas(h,""),"0",jTxt];
      if(idx===-1) await sClient.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Asistencia!A:M',valueInputOption:'USER_ENTERED',requestBody:{values:[row]}}); else await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!A${idx+1}:M${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[row]}});
      await sock.sendMessage(jid,{text:`✅ ${estatus} - ${nombre} en ${cercana.nombre} - ${h}`});
      if(esRet){ const g=(cercana.id==='COYOACAN'||cercana.id==='HOTEL')?GRUPO_COYOACAN_ID:GRUPO_BUCARELI_ID; if(g) try{ await sock.sendMessage(g,{text:`⏰ RETARDO ${minRet}min (Prog ${hProg}) - ${nombre} en ${cercana.nombre} - ${h}`}) }catch{} }
    }else{
      if(hoy[5]){ await sock.sendMessage(jid,{text:`Salida ya registrada`}); return }
      if(dMin>cercana.rSal){ await sock.sendMessage(jid,{text:`No puedes checar salida a ${Math.round(dMin)}m`},{quoted:m}); return }
      const h=horaMX(); const jTxt=hoy[12]||"8h"; const {trabajadas,extra}=calcularExtra(hoy[3],h,hObj?.entrada||null,hObj?.salida||null);
      await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!F${idx+1}:M${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[[h,hoy[6]||'',hoy[7]||'',cercana.nombre,Math.round(dMin).toString(),trabajadas,extra,jTxt]]}});
      await sock.sendMessage(jid,{text:`✅ Salida - ${nombre} en ${cercana.nombre} - ${h} - Trab ${trabajadas} Extra ${extra}`});
    }
  }catch(e){ console.error(e) }
}

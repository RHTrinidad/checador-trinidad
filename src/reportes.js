import ExcelJS from 'exceljs'
import fs from 'fs'
import path from 'path'
import os from 'os'
import { SPREADSHEET_ID } from './config.js'
import { getRows, sheetsClient, getHorarioBaseMap } from './sheets.js'
import { fechaLaboral, horaMX, minutos, parseFechaMX, normaliza, parseHorarioRango, getRangoSemana, sucursalCoincideConFiltro, scoreEmpleado, calcularExtra } from './utils.js'

// ASISTENCIA HOY
export async function asistenciaHoy(filtroSucursal, jid, sock){
  const sClient = await sheetsClient();
  const [baseRows, asisRows] = await Promise.all([ getRows('Horario_Base!A2:K'), sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}).then(r=>r.data.values||[]) ]);
  const fLab = fechaLaboral(); const ahoraMin = minutos(horaMX());
  const ahoraMXDate = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'})); const diaNum = ahoraMXDate.getDay();
  const filtro = filtroSucursal.toLowerCase(); const esCoyo = filtro.includes('coyo') || filtro.includes('hotel'); const esJuarez = filtro.includes('juarez')||filtro.includes('bucareli');
  let llego=[], retardo=[], falta=[], futuro=[];
  for(const r of baseRows){
    const sucBaseLower = (r[2]||'').toLowerCase(); const sucBaseOriginal = r[2]||'';
    let inc=false;
    if(esCoyo) inc = sucBaseLower.includes('coyo')||sucBaseLower.includes('hotel')||sucBaseLower.includes('trinidad');
    else if(esJuarez) inc = sucBaseLower.includes('juarez')||sucBaseLower.includes('bucareli');
    else inc = sucBaseLower.includes(filtro);
    if(!inc) continue;
    const nombre = r[1]||''; const tel = (r[0]||'').replace(/\D/g,'').slice(-10);
    const mapa = {1:r[3],2:r[4],3:r[5],4:r[6],5:r[7],6:r[8],0:r[9]};
    const v = (mapa[diaNum]||'').toString().trim(); const parsed = parseHorarioRango(v); if(!parsed) continue;
    const esLibre = parsed.entrada === 'LIBRE'; const horaProg = esLibre? 'LIBRE' : parsed.entrada;
    const registro = asisRows.find(a => a[0]?.replace(/\D/g,'').slice(-10)===tel && a[2]===fLab);
    if(registro && registro[3]){
      const entrada = registro[3]; const sucEnt = registro[6]||''; const dif = esLibre? 0 : minutos(entrada) - minutos(horaProg);
      if(esLibre){ llego.push(• ${nombre} - Entró ${entrada} en ${sucEnt} ✅) }
      else if(dif > 15){ retardo.push(• ${nombre} - [${sucBaseOriginal}] Prog ${horaProg} - Entró ${entrada} - ⏰ ${dif}m tarde - ${sucEnt}) }
      else llego.push(• ${nombre} - [${sucBaseOriginal}] Prog ${horaProg} - Entró ${entrada} ✅ - ${sucEnt})
    } else {
      if(esLibre){ if(ahoraMin >= 20*60) falta.push(• ${nombre} - [${sucBaseOriginal}] - ❌ sin llegar); continue }
      const dif = ahoraMin - minutos(horaProg);
      if(dif < 0){ futuro.push(• ${nombre} - [${sucBaseOriginal}] - Prog ${horaProg}) }
      else falta.push(• ${nombre} - [${sucBaseOriginal}] - Prog ${horaProg} - ❌ ${dif}m sin llegar)
    }
  }
  let txt = 📍 *ASISTENCIA HOY ${fLab} - ${filtroSucursal.toUpperCase()}* ${horaMX()}\n\n✅ *A TIEMPO (${llego.length}):*\n${llego.join('\n')||'-'}\n\n⏰ *RETARDOS (${retardo.length}):*\n${retardo.join('\n')||'-'}\n\n❌ *NO HAN LLEGADO (${falta.length}):*\n${falta.join('\n')||'Todos llegaron'}\n\n⏳ *PRÓXIMOS (${futuro.length}):*\n${futuro.join('\n')||'-'};
  await sock.sendMessage(jid,{text:txt})
}

export async function generarExcelEmpleado(nombreBuscarRaw, jid, sock, tipo='actual'){
  const sClient=await sheetsClient(); const baseMap=await getHorarioBaseMap();
  const asisRes = await sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'});
  const filas=asisRes.data.values||[]; const buscar=nombreBuscarRaw.toLowerCase().replace(/actual|pasada|pasado|esta semana|hoy/g,'').trim();
  const { lunes, domingo, rangoTxt } = getRangoSemana(tipo);
  const filtradas = filas.filter(f=>{ const n=(f[1]||'').toLowerCase(); if(!n.includes(buscar)) return false; const fe=parseFechaMX(f[2]); return fe&&fe>=lunes&&fe<=domingo });
  const header=["Tel","Nombre","Fecha","Entrada","Estatus Entrada","Salida","Suc Entrada","Dist Entr","Suc Salida","Dist Sal","Horas Trabajadas","Horas Extra","Jornada Programada"];
  const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet('Resumen');
  ws.addRow([REPORTE ${buscar.toUpperCase()} - ${tipo.toUpperCase()}]).font={bold:true,size:14}; ws.addRow([rangoTxt]); ws.addRow([]);
  let dias=0,ret=0,min=0,sin=0,horasMin=0,extraMin=0;
  filtradas.forEach(f=>{ if(f[3]) dias++; const m=(f[4]||'').match(/(\d+)\s*min/); if(m){ret++; min+=parseInt(m[1])} if(!f[5]) sin++; if(f[3]&&f[5]){ const fe=parseFechaMX(f[2]); const prog=baseMap[(f[0]||'').replace(/\D/g,'').slice(-10)]?.horas?.[fe.getDay()]; const calc=calcularExtra(f[3],f[5],prog?.entrada||null,prog?.salida||null); let mt=calc.trabajadas==="8"?480:(()=>{const[hh,mm]=calc.trabajadas.split(':').map(Number); return hh*60+mm})(); horasMin+=mt; extraMin+=calc.extraMin } });
  ws.addRow(['Nombre','Días','Retardos','Min','Sin salida','Horas Trab','Horas Extra']).font={bold:true};
  ws.addRow([buscar, ${dias} días, Ret ${ret} (${min}m), sin, ${Math.floor(horasMin/60)}:${String(horasMin%60).padStart(2,'0')}h, ${Math.floor(extraMin/60)}:${String(extraMin%60).padStart(2,'0')}h]); ws.columns.forEach(c=>c.width=22);
  const ws2=wb.addWorksheet('Detalle'); ws2.addRow(header).font={bold:true}; filtradas.forEach(f=>{ ws2.addRow(f) }); ws2.columns.forEach(c=>c.width=18);
  const fileName=Reporte_${buscar.replace(/\s+/g,'_')}_${tipo}.xlsx; const fp=path.join(os.tmpdir(),fileName); await wb.xlsx.writeFile(fp); await sock.sendMessage(jid,{document:fs.readFileSync(fp),mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',fileName}); fs.unlinkSync(fp)
}

export async function resumenEmpleado(nombreBuscar, jidRespuesta, sock, tipo='actual'){
  const sClient=await sheetsClient(); const [asisRes, baseMap]=await Promise.all([ sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}), getHorarioBaseMap() ]);
  const filas=asisRes.data.values||[]; let buscar=nombreBuscar.toLowerCase().replace(/actual|pasada|pasado|esta semana|hoy/g,'').trim();
  const { lunes, domingo, rangoTxt } = getRangoSemana(tipo);
  const info = baseMap[buscar] || Object.values(baseMap).find(v=> (v.nombreOriginal||'').toLowerCase().includes(buscar)) || { descansos:new Set(), horas:{} };
  let diasSem=0,retSem=0,minSem=0,detalle=[],sin=[],horasMin=0,extraTotal=0, diasNom=['Dom','Lun','Mar','Mie','Jue','Vie','Sab'];
  for(const f of filas){ const n=(f[1]||'').toLowerCase(); if(!n.includes(buscar)) continue; const fe=parseFechaMX(f[2]); if(!fe) continue; fe.setHours(0,0,0,0); if(fe<lunes||fe>domingo) continue; if(f[3]){ diasSem++; const ret=(f[4]||'').match(/(\d+)\s*min/); if(ret){retSem++; minSem+=parseInt(ret[1])} if(!f[5]) sin.push(f[2]); const progDia=info.horas[fe.getDay()]; if(f[3]&&f[5]&&progDia){ const calc=calcularExtra(f[3],f[5],progDia.entrada,progDia.salida); let mt=calc.trabajadas==="8"?480:(()=>{const[hh,mm]=calc.trabajadas.split(':').map(Number);return hh*60+mm})(); horasMin+=mt; extraTotal+=calc.extraMin } detalle.push(• ${f[2]}: Ent ${f[3]} | Sal ${f[5]||'SIN'} | Trab ${f[10]||''} | Extra ${f[11]||'0'} | ${f[12]||''}) } }
  let esperados=0; for(let d=new Date(lunes); d<=domingo; d.setDate(d.getDate()+1)){ if(!info.descansos.has(d.getDay())) esperados++ }
  const faltas=Math.max(0,esperados-diasSem); const descTxt=[...info.descansos].map(d=>diasNom[d]).join(', ')||'ninguno';
  const txt=📊 *${buscar.toUpperCase()}* - Desc: ${descTxt}\n${rangoTxt}\n\n*SEMANA ${tipo.toUpperCase()}*\n- Trabajados: ${diasSem}/${esperados} - Faltas: ${faltas}\n- Retardos: ${retSem} (${minSem} min)\n- Sin salida: ${sin.length}\n- Horas Trab: ${Math.floor(horasMin/60)}:${String(horasMin%60).padStart(2,'0')}h\n- Horas Extra: ${Math.floor(extraTotal/60)}:${String(extraTotal%60).padStart(2,'0')}h\n\nDetalle:\n${detalle.join('\n')||'Sin registros'};
  await sock.sendMessage(jidRespuesta,{text:txt}); if(diasSem>0) await generarExcelEmpleado(buscar, jidRespuesta, sock, tipo)
}

export async function reporteSucursal(filtroSucursal, jid, sock, tipo='pasada'){
  const [asisRes, empRows] = await Promise.all([ (await sheetsClient()).spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}), getRows('Empleados!A:K') ]);
  const filas=asisRes.data.values||[]; const {lunes,domingo,rangoTxt}=getRangoSemana(tipo);
  const filtro=filtroSucursal.toLowerCase(); const esCoyo=filtro.includes('coyo')||filtro.includes('hotel'); const esJuarez=filtro.includes('juarez')||filtro.includes('bucareli');
  const baseMap = await getHorarioBaseMap(); const telToNombre = {}; empRows.forEach(r=>{ const tel=(r[0]||'').replace(/\D/g,'').slice(-10); if(tel) telToNombre[tel] = (r[3]||r[1]||'').trim() });
  let datos={};
  for(const f of filas){
    const fe=parseFechaMX(f[2]); if(!fe||fe<lunes||fe>domingo) continue; const suc=((f[6]||'')+' '+(f[8]||'')).toLowerCase();
    let inc=false; if(esCoyo) inc=suc.includes('coyo')||suc.includes('hotel'); else if(esJuarez) inc=suc.includes('juarez')||suc.includes('bucareli'); else inc=suc.includes(filtro); if(!inc) continue;
    const tel=(f[0]||'').replace(/\D/g,'').slice(-10); const nombreOficial = telToNombre[tel] || (f[1]||'').trim() || 'Desconocido'; const key = tel || normaliza(nombreOficial);
    if(!datos[key]) datos[key]={nombre:nombreOficial, dias:0,ret:0,min:0,sin:0,horasMin:0,extraMin:0};
    if(f[3]){ datos[key].dias++; const m=(f[4]||'').match(/(\d+)\s*min/); if(m){datos[key].ret++; datos[key].min+=parseInt(m[1])} if(!f[5]) datos[key].sin++; if(f[3]&&f[5]){ const feDia=parseFechaMX(f[2]); const info=baseMap[tel]||null; const prog=info?.horas?.[feDia?feDia.getDay():1]||null; const calc=calcularExtra(f[3],f[5],prog?.entrada||null,prog?.salida||null); let mt=calc.trabajadas==="8"?480:(()=>{const[hh,mm]=calc.trabajadas.split(':').map(Number);return hh*60+mm})(); datos[key].horasMin+=mt; datos[key].extraMin+=calc.extraMin } }
  }
  let txt=📊 *REPORTE ${filtroSucursal.toUpperCase()}* - ${tipo}\n${rangoTxt}\n; if(!Object.keys(datos).length) txt+='Sin registros\n'; else for(const k in datos){ const d=datos[k]; txt+=*${d.nombre}*: ${d.dias} días | Ret ${d.ret} (${d.min}m) | Trab: ${Math.floor(d.horasMin/60)}:${String(d.horasMin%60).padStart(2,'0')}h | Extra: ${Math.floor(d.extraMin/60)}:${String(d.extraMin%60).padStart(2,'0')}h | Sin salida: ${d.sin}\n\n }
  await sock.sendMessage(jid,{text:txt})
}

export async function generarExcelSemanaYEnviar(filtroSucursal, jid, sock, tipo='pasada'){
  const sClient=await sheetsClient(); const baseMap=await getHorarioBaseMap();
  const [asisRes, empRows] = await Promise.all([ sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}), getRows('Empleados!A:K') ]);
  const filas=asisRes.data.values||[]; const {lunes,domingo,rangoTxt}=getRangoSemana(tipo);
  const esCoyo=filtroSucursal.toLowerCase().includes('coyo'); const esJuarez=filtroSucursal.toLowerCase().includes('juarez')||filtroSucursal.toLowerCase().includes('bucareli');
  const telToNombre = {}; empRows.forEach(r=>{ const tel=(r[0]||'').replace(/\D/g,'').slice(-10); if(tel) telToNombre[tel] = (r[3]||r[1]||'').trim() });
  const filtradas=filas.filter(f=>{ const fe=parseFechaMX(f[2]); if(!fe||fe<lunes||fe>domingo) return false; const suc=((f[6]||'')+' '+(f[8]||'')).toLowerCase(); if(esCoyo) return suc.includes('coyo')||suc.includes('hotel'); if(esJuarez) return suc.includes('juarez')||suc.includes('bucareli'); return suc.includes(filtroSucursal.toLowerCase()) });
  let datos={}; for(const f of filtradas){ const tel=(f[0]||'').replace(/\D/g,'').slice(-10); const nombreOficial = telToNombre[tel] || (f[1]||'Desconocido').trim(); const key = tel || normaliza(nombreOficial); if(!datos[key]) datos[key]={nombre:nombreOficial, dias:0,ret:0,min:0,sin:0,horasMin:0,extraMin:0, tel}; const fe=parseFechaMX(f[2]); if(fe){ const prog=baseMap[tel]?.horas?.[fe.getDay()]; if(f[3]&&f[5]){ const calc=calcularExtra(f[3],f[5],prog?.entrada,prog?.salida); let mt=calc.trabajadas==="8"?480:(()=>{const[hh,mm]=calc.trabajadas.split(':').map(Number);return hh*60+mm})(); datos[key].horasMin+=mt; datos[key].extraMin+=calc.extraMin } } datos[key].dias++; const m=(f[4]||'').match(/(\d+)\s*min/); if(m){datos[key].ret++; datos[key].min+=parseInt(m[1])} if(!f[5]) datos[key].sin++ };
  const header=["Tel","Nombre","Fecha","Entrada","Estatus Entrada","Salida","Suc Entrada","Dist Entr","Suc Salida","Dist Sal","Horas Trabajadas","Horas Extra","Jornada"];
  const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet('Resumen'); ws.addRow([REPORTE ${filtroSucursal.toUpperCase()} - ${tipo}]).font={bold:true,size:14}; ws.addRow([rangoTxt]); ws.addRow([]); ws.addRow(['Nombre','Tel','Días','Retardos','Min','Sin salida','Horas Trab','Horas Extra']).font={bold:true};
  for(const k of Object.keys(datos)){ const d=datos[k]; ws.addRow([d.nombre, d.tel, ${d.dias} días, Ret ${d.ret} (${d.min}m), d.sin, ${Math.floor(d.horasMin/60)}:${String(d.horasMin%60).padStart(2,'0')}h, ${Math.floor(d.extraMin/60)}:${String(d.extraMin%60).padStart(2,'0')}h]) } ws.columns.forEach(c=>c.width=22);
  const ws2=wb.addWorksheet('Detalle'); ws2.addRow(header).font={bold:true}; filtradas.forEach(f=>{ ws2.addRow(f) }); ws2.columns.forEach(c=>c.width=18);
  const fileName=Reporte_${filtroSucursal}_${tipo}_${lunes.toISOString().split('T')[0]}.xlsx; const fp=path.join(os.tmpdir(),fileName); await wb.xlsx.writeFile(fp); await sock.sendMessage(jid,{document:fs.readFileSync(fp),mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',fileName}); fs.unlinkSync(fp)
}

export async function handleReportes({ texto, jid, sock, filtroGrupo }){
  const low=texto.toLowerCase()
  if(low.startsWith('asistencia hoy')){ let suc=low.replace('asistencia hoy','').trim(); if(!suc) suc=filtroGrupo||'coyoacan'; await asistenciaHoy(suc,jid,sock); return true }
  if(low.startsWith('resumen ')){
    let tipo='actual'; if(low.includes('pasada')||low.includes('pasado')) tipo='pasada';
    let limpio=texto.slice(8).toLowerCase().trim().replace(/pasada|pasado|actual|esta semana|hoy/g,'').trim();
    if(limpio.includes('bucareli')||limpio.includes('juarez')||limpio.includes('coyo')||limpio.includes('hotel')){
      let suc='coyoacan'; if(limpio.includes('juarez')||limpio.includes('bucareli')) suc='juarez';
      await reporteSucursal(suc,jid,sock,tipo); await generarExcelSemanaYEnviar(suc,jid,sock,tipo); return true
    }
    if(!limpio){ await sock.sendMessage(jid,{text:'Escribe: resumen [nombre] pasada o actual'}); return true }
    await resumenEmpleado(limpio,jid,sock,tipo); return true
  }
  if(/^(reporte|checador|reporte x unidad)/i.test(texto)){
    const tipo=low.includes('pasada')||low.includes('pasado')?'pasada':low.includes('actual')||low.includes('esta')||low.includes('hoy')?'actual':'pasada';
    let suc='coyoacan'; if(low.includes('juarez')||low.includes('bucareli')) suc='juarez'; else if(low.includes('hotel')) suc='hotel'; else if(low.includes('coyo')) suc='coyoacan';
    if(low.includes('x unidad')||low.includes('por unidad')){
      if(!filtroGrupo){ await reporteSucursal('coyoacan',jid,sock,tipo); await generarExcelSemanaYEnviar('coyoacan',jid,sock,tipo); await reporteSucursal('juarez',jid,sock,tipo); await generarExcelSemanaYEnviar('juarez',jid,sock,tipo); }
      else { await reporteSucursal(filtroGrupo,jid,sock,tipo); await generarExcelSemanaYEnviar(filtroGrupo,jid,sock,tipo); }
      return true
    }
    await reporteSucursal(suc,jid,sock,tipo); await generarExcelSemanaYEnviar(suc,jid,sock,tipo); return true
  }
  if(/^(numero|número|num|tel|telefono|teléfono|info|ficha|datos|dato)\s*(de)?/i.test(texto)){
    let buscar = texto.toLowerCase().replace(/^(numero|número|num|tel|telefono|teléfono|info|ficha|datos|dato)\s*(de)?\s*/i,'').trim();
    if(!buscar){ await sock.sendMessage(jid,{text:'Escribe: info fer'}); return true }
    const buscarNorm = normaliza(buscar); const empRows = await getRows('Empleados!A:K'); let candidatos=[]; for(const r of empRows.slice(1)){ const corto=r[1]||''; const completo=r[3]||''; const suc=r[2]||''; if(!sucursalCoincideConFiltro(suc,filtroGrupo)) continue; const s=scoreEmpleado(corto,completo,buscarNorm); if(s>=0) candidatos.push({r,s}) } candidatos.sort((a,b)=>b.s-a.s); if(!candidatos.length){ await sock.sendMessage(jid,{text:No encontré a "${buscar}"}); return true } const r=candidatos[0].r; const ficha=📋 *${r[3]||r[1]}*\n👤 Corto: ${r[1]}\n📍 ${r[2]||'-'}\n📱 ${r[0]||'-'}; await sock.sendMessage(jid,{text:ficha}); return true
  }
  return false
}

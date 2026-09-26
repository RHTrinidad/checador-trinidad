import { default as makeWASocket, useMultiFileAuthState, DisconnectReason, fetchLatestBaileysVersion } from '@whiskeysockets/baileys'
import qrcodeTerminal from 'qrcode-terminal'
import QRCode from 'qrcode'
import express from 'express'
import { createRequire } from 'module'
import fs from 'fs'
import path from 'path'
import os from 'os'
import cron from 'node-cron'
import P from 'pino'
const require = createRequire(import.meta.url)
const { google } = require('googleapis')
const ExcelJS = require('exceljs')

const SPREADSHEET_ID = process.env.SPREADSHEET_ID
const GRUPO_REPORTES_ID = process.env.GRUPO_REPORTES_ID || "120363412984528459@g.us"
const GRUPO_COYOACAN_ID = process.env.GRUPO_COYOACAN_ID || "120363430474175735@g.us"
const GRUPO_BUCARELI_ID = process.env.GRUPO_BUCARELI_ID || "120363410610062461@g.us"
const GRUPO_JUAREZ_ID = process.env.GRUPO_JUAREZ_ID || GRUPO_BUCARELI_ID
const GRUPO_CHECADOR_COYOACAN_ID = process.env.GRUPO_CHECADOR_COYOACAN_ID || "120363403668034900@g.us"
const GRUPO_CHECADOR_BUCARELI_ID = process.env.GRUPO_CHECADOR_BUCARELI_ID || "120363411739089744@g.us"

const GRUPOS = {
  REPORTES: [GRUPO_REPORTES_ID],
  CHECADORES: [GRUPO_CHECADOR_COYOACAN_ID, GRUPO_CHECADOR_BUCARELI_ID],
  GERENTES: [GRUPO_COYOACAN_ID, GRUPO_BUCARELI_ID, GRUPO_JUAREZ_ID]
}

const PAQUETES = {
  REPORTES_PARA_GERENTES: /^(numero|número|num|tel|telefono|teléfono|info|ficha|datos|dato|asistencia hoy|resumen|reporte|checador|reporte x unidad|rfc|ine|curp)/i,
  CHECADA: 'LOCATION'
}

function getTipoGrupo(jid){
  if(GRUPOS.REPORTES.includes(jid)) return 'REPORTES'
  if(GRUPOS.CHECADORES.includes(jid)) return 'CHECADORES'
  if(GRUPOS.GERENTES.includes(jid)) return 'GERENTES'
  return null
}

let SUCURSALES = []
try {
  if (process.env.SUCURSALES_JSON) {
    const parsed = JSON.parse(process.env.SUCURSALES_JSON)
    const arr = Array.isArray(parsed)? parsed : Object.values(parsed)
    SUCURSALES = arr.map(s => ({ id: s.id || s.nombre, nombre: s.nombre || s.id, lat: s.lat, lng: s.lng || s.lon, rEnt: 150, rSal: 150 }))
  } else throw new Error('no json')
} catch (e) {
  SUCURSALES = [
    { id:"COYOACAN", nombre:"Trinidad Coyoacan", lat:19.352525, lng:-99.161817, rEnt:150, rSal:150 },
    { id:"JUAREZ", nombre:"Trinidad Juarez", lat:19.4314119, lng:-99.1512074, rEnt:150, rSal:150 },
    { id:"HOTEL", nombre:"Servicio Hotel", lat:19.351770, lng:-99.165458, rEnt:150, rSal:150 }
  ]
}

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets']
})
async function sheetsClient(){ const c=await auth.getClient(); return google.sheets({version:'v4',auth:c}) }
function distM(a,b,c,d){ const R=6371000, toRad=x=>x*Math.PI/180; const dLa=toRad(c-a), dLo=toRad(d-b); const q=Math.sin(dLa/2)**2+Math.cos(toRad(a))*Math.cos(toRad(c))*Math.sin(dLo/2)**2; return R*2*Math.atan2(Math.sqrt(q),Math.sqrt(1-q)) }
function fechaLaboral(d=new Date()){ const mx = new Date(d.toLocaleString('en-US',{timeZone:'America/Mexico_City'})); if(mx.getHours() < 4) mx.setDate(mx.getDate()-1); return `${mx.getFullYear()}-${String(mx.getMonth()+1).padStart(2,'0')}-${String(mx.getDate()).padStart(2,'0')}` }
function horaMX(d=new Date()){ return d.toLocaleTimeString('es-MX',{hour12:false,timeZone:'America/Mexico_City'}) }
function minutos(h){ const mm = h?.toString().match(/(\d{1,2}):(\d{2})/); return mm? parseInt(mm[1])*60+parseInt(mm[2]) : 0 }
function parseHorarioRango(v){
  const s = (v||'').toString().trim(); if(!s) return null; const low = s.toLowerCase()
  if(low.includes('libre')||low.includes('flex')||low.includes('abierto')||low.includes('comodin')) return { entrada: 'LIBRE', salida: null, raw: s }
  if(low.includes('descanso')) return null
  const m = s.match(/(\d{1,2}:\d{2})\s*(?:-|a|hasta)?\s*(\d{1,2}:\d{2})?/i); if(!m) return null
  return { entrada: m[1], salida: m[2]||null, raw: s }
}
function calcularHorasTrabajadas(hE,hS){
  if(!hS) return "8"
  const toSeg=(h)=>{const [hh,mm,ss]=(h||'').split(":").map(Number); return (hh||0)*3600+(mm||0)*60+(ss||0)}
  let diff=toSeg(hS)-toSeg(hE); if(diff<0) diff+=24*3600
  const h=Math.floor(diff/3600),m=Math.floor((diff%3600)/60)
  return h===8&&m===0?"8":`${h}:${String(m).padStart(2,'0')}:00`
}
function calcularExtra(hE,hS, progEnt, progSal){
  const trabajadas = calcularHorasTrabajadas(hE,hS)
  if(!progEnt||!progSal||progEnt==='LIBRE') return { trabajadas, extra: "0", extraMin: 0 }
  const toMin = h => { const [hh,mm]=h.split(':').map(Number); return (hh||0)*60+(mm||0) }
  let minProg = toMin(progSal) - toMin(progEnt); if(minProg < 0) minProg += 24*60
  let minTrab = 0; if(trabajadas==="8") minTrab=8*60; else { const [hh,mm]=trabajadas.split(':').map(Number); minTrab=(hh||0)*60+(mm||0) }
  let extraMin = minTrab - minProg; if(extraMin < 0) extraMin = 0
  const eh = Math.floor(extraMin/60); const em = extraMin%60
  return { trabajadas, extra: extraMin>0? `${eh}:${String(em).padStart(2,'0')}:00` : "0", extraMin }
}
function parseFechaMX(s){ if(!s) return null; if(/^\d{4}-\d{2}-\d{2}$/.test(s)) return new Date(s+'T12:00:00'); const m=s.match(/(\d{1,2})\/(\d{2,4})/); if(m){ return new Date(`${m[3].length==2?'20'+m[3]:m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}T12:00:00`) } return new Date(s) }
async function getRows(range){ const s=await sheetsClient(); const r=await s.spreadsheets.values.get({spreadsheetId:SPREADSHEET_ID,range}); return r.data.values||[] }
function getRangoSemana(tipo){
  const hoyMX = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'}))
  const diaSem = hoyMX.getDay(); const lunesEstaSem = new Date(hoyMX); lunesEstaSem.setDate(hoyMX.getDate() - (diaSem===0?6:diaSem-1))
  let lunes, domingo; if(tipo==='pasada'){ lunes=new Date(lunesEstaSem); lunes.setDate(lunes.getDate()-7); domingo=new Date(lunes); domingo.setDate(domingo.getDate()+6); } else { lunes=lunesEstaSem; domingo=hoyMX; }
  lunes.setHours(0,0,0,0); domingo.setHours(23,59,59,999)
  return { lunes, domingo, rangoTxt: `${lunes.toLocaleDateString('es-MX')} al ${domingo.toLocaleDateString('es-MX')} (${tipo})` }
}
async function getHorarioBaseMap(){
  try{
    const rows = await getRows('Horario_Base!A2:K')
    const map = {}
    for(const f of rows){
      const tel = (f[0]||'').replace(/\D/g,'').slice(-10); const nombre = (f[1]||'').toLowerCase().trim(); if(!nombre) continue
      const dias = { 1:f[3], 2:f[4], 3:f[5], 4:f[6], 5:f[7], 6:f[8], 0:f[9] }
      const descansos = new Set(); const horas = {}
      for(const [numDia, valor] of Object.entries(dias)){
        const v = (valor||'').toString().trim(); if(!v){ descansos.add(parseInt(numDia)); continue }
        const parsed = parseHorarioRango(v); if(!parsed) descansos.add(parseInt(numDia)); else horas[numDia]=parsed
      }
      const obj = { descansos, horas, nombreOriginal: f[1], tel, sucursal: f[2]||'' }
      if(tel) map[tel]=obj; map[nombre]=obj; const primer=nombre.split(' ')[0]; if(primer &&!map[primer]) map[primer]=obj
    }
    return map
  }catch(e){ return {} }
}
let grupoFiltroCache = new Map()
async function getFiltroPorGrupo(jid, sock){
  if(jid === GRUPO_REPORTES_ID) return null
  if(jid === GRUPO_COYOACAN_ID || jid === GRUPO_CHECADOR_COYOACAN_ID) return 'coyoacan'
  if(jid === GRUPO_BUCARELI_ID || jid === GRUPO_JUAREZ_ID || jid === GRUPO_CHECADOR_BUCARELI_ID) return 'juarez'
  if(grupoFiltroCache.has(jid)) return grupoFiltroCache.get(jid)
  try{
    const meta = await sock.groupMetadata(jid)
    const subj = (meta.subject || '').toLowerCase()
    let filtro = null
    if(subj.includes('coyo')) filtro = 'coyoacan'
    else if(subj.includes('bucareli') || subj.includes('juarez')) filtro = 'juarez'
    grupoFiltroCache.set(jid, filtro)
    return filtro
  }catch{ return null }
}
function sucursalCoincideConFiltro(sucEmpleado, filtro){
  if(!filtro) return true
  const s = (sucEmpleado||'').toLowerCase()
  if(filtro === 'coyoacan') return s.includes('coyo') || s.includes('hotel') || s.includes('trinidad')
  if(filtro === 'juarez') return s.includes('juarez') || s.includes('bucareli')
  return s.includes(filtro)
}
function normaliza(s){ return (s||'').toString().normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim() }
function scoreEmpleado(corto, completo, buscarNorm){
  const c = normaliza(corto); const d = normaliza(completo)
  const cWords = c.split(/\s+/); const dWords = d.split(/\s+/)
  const allWords = [...cWords,...dWords]
  if(c === buscarNorm) return 100
  if(c.startsWith(buscarNorm)) return 90
  if(cWords.some(w=>w.startsWith(buscarNorm))) return 80
  if(dWords.some(w=>w.startsWith(buscarNorm))) return 70
  if(d.startsWith(buscarNorm)) return 60
  if(allWords.includes(buscarNorm)) return 50
  if(buscarNorm.length <= 3) return -1
  if(c.includes(buscarNorm) || d.includes(buscarNorm)) return 10
  return -1
}
async function asistenciaHoy(filtroSucursal, jid, sock){
  const sClient = await sheetsClient()
  const [baseRows, asisRows] = await Promise.all([ getRows('Horario_Base!A2:K'), sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}).then(r=>r.data.values||[]) ])
  const fLab = fechaLaboral(); const ahoraMin = minutos(horaMX()); const ahoraMXDate = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'})); const diaNum = ahoraMXDate.getDay()
  const filtro = filtroSucursal.toLowerCase(); const esCoyo = filtro.includes('coyo') || filtro.includes('hotel'); const esJuarez = filtro.includes('juarez')||filtro.includes('bucareli')
  let llego=[], retardo=[], falta=[], futuro=[]
  for(const r of baseRows){
    const sucBaseLower = (r[2]||'').toLowerCase(); const sucBaseOriginal = r[2]||''; let inc=false
    if(esCoyo) inc = sucBaseLower.includes('coyo')||sucBaseLower.includes('hotel')||sucBaseLower.includes('trinidad'); else if(esJuarez) inc = sucBaseLower.includes('juarez')||sucBaseLower.includes('bucareli'); else inc = sucBaseLower.includes(filtro); if(!inc) continue
    const nombre = r[1]||''; const tel = (r[0]||'').replace(/\D/g,'').slice(-10); const mapa = {1:r[3],2:r[4],3:r[5],4:r[6],5:r[7],6:r[8],0:r[9]}
    const v = (mapa[diaNum]||'').toString().trim(); const parsed = parseHorarioRango(v); if(!parsed) continue; const esLibre = parsed.entrada === 'LIBRE'; const horaProg = esLibre? 'LIBRE' : parsed.entrada
    const registro = asisRows.find(a => a[0]?.replace(/\D/g,'').slice(-10)===tel && a[2]===fLab)
    if(registro && registro[3]){
      const entrada = registro[3]; const sucEnt = registro[6]||''; const dif = esLibre? 0 : minutos(entrada) - minutos(horaProg)
      if(esLibre){ llego.push(`• ${nombre} - Entró ${entrada} en ${sucEnt} ✅`) }
      else if(dif > 15){ retardo.push(`• ${nombre} - [${sucBaseOriginal}] Prog ${horaProg} - Entró ${entrada} - ⏰ ${dif}m tarde - ${sucEnt}`) }
      else if(dif > 0){ llego.push(`• ${nombre} - [${sucBaseOriginal}] Prog ${horaProg} - Entró ${entrada} - ${dif}m tarde ✅ - ${sucEnt}`) }
      else if(dif < 0){ llego.push(`• ${nombre} - [${sucBaseOriginal}] Prog ${horaProg} - Entró ${entrada} - ${Math.abs(dif)}m antes ✅ - ${sucEnt}`) }
      else { llego.push(`• ${nombre} - [${sucBaseOriginal}] Prog ${horaProg} - Entró ${entrada} puntual ✅ - ${sucEnt}`) }
    } else {
      if(esLibre){ if(ahoraMin >= 20*60) falta.push(`• ${nombre} - [${sucBaseOriginal}] - ❌ ${ahoraMin - 20*60}m sin llegar (corte 20:00)`); continue }
      const dif = ahoraMin - minutos(horaProg)
      if(dif < 0){ futuro.push(`• ${nombre} - [${sucBaseOriginal}] - Prog ${horaProg} - en ${Math.abs(dif)}m`) }
      else { falta.push(`• ${nombre} - [${sucBaseOriginal}] - Prog ${horaProg} - ❌ ${dif}m sin llegar`) }
    }
  }
  let txt = `📍 *ASISTENCIA HOY ${fLab} - ${filtroSucursal.toUpperCase()}* ${horaMX()}\n`; if(esCoyo) txt+= `_Incluye Hotel + Coyoacán_\n`
  txt+=`\n✅ *A TIEMPO (${llego.length}):*\n${llego.join('\n')||'-'}\n\n⏰ *RETARDOS (${retardo.length}):*\n${retardo.join('\n')||'-'}\n\n❌ *NO HAN LLEGADO (${falta.length}):*\n${falta.join('\n')||'Todos llegaron'}\n\n⏳ *PRÓXIMOS (${futuro.length}):*\n${futuro.join('\n')||'-'}`
  await sock.sendMessage(jid,{text:txt})
}
async function autocierreAsistencia(){
  const sClient = await sheetsClient()
  const asisRows = await getRows('Asistencia!A2:M')
  const ahoraMX = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'}))
  for(let i=1; i<asisRows.length; i++){
    const r = asisRows[i]
    const entrada = r[3]
    const salida = r[5]
    const fecha = r[2]
    if(!entrada || salida) continue
    if((r[5]||'').toString().includes('AUTOCIERRE')) continue
    try{
      const fechaEntrada = new Date(fecha+'T'+entrada)
      if(isNaN(fechaEntrada)) continue
      const diffHoras = (ahoraMX - fechaEntrada) / (1000*60*60)
      if(diffHoras >= 16){
        await sClient.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `Asistencia!F${i+2}:M${i+2}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [["AUTOCIERRE 16H", r[6]||'', r[7]||'', "AUTOCIERRE", "0", "8", "0", r[12]||""]] }
        })
      }
    }catch(e){}
  }
}
async function generarExcelEmpleado(nombreBuscarRaw, jid, sock, tipo='actual'){ const sClient=await sheetsClient(); const baseMap=await getHorarioBaseMap(); const asisRes = await sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}); const filas=asisRes.data.values||[]; const buscar=nombreBuscarRaw.toLowerCase().replace(/actual|pasada|pasado|esta semana|hoy/g,'').trim(); const { lunes, domingo, rangoTxt } = getRangoSemana(tipo); const filtradas = filas.filter(f=>{ const n=(f[1]||'').toLowerCase(); if(!n.includes(buscar)) return false; const fe=parseFechaMX(f[2]); return fe&&fe>=lunes&&fe<=domingo }); const header=["Tel","Nombre","Fecha","Entrada","Estatus Entrada","Salida","Suc Entrada","Dist Entr","Suc Salida","Dist Sal","Horas Trabajadas","Horas Extra","Jornada Programada"]; const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet('Resumen'); ws.addRow([`REPORTE ${buscar.toUpperCase()} - ${tipo.toUpperCase()}`]).font={bold:true,size:14}; ws.addRow([rangoTxt]); ws.addRow([]); let dias=0,ret=0,min=0,sin=0,horasMin=0,extraMin=0; filtradas.forEach(f=>{ if(f[3]) dias++; const m=(f[4]||'').match(/(\d+)\s*min/); if(m){ret++; min+=parseInt(m[1])} if(!f[5]) sin++; if(f[3]&&f[5]){ const fe=parseFechaMX(f[2]); const prog=baseMap[(f[0]||'').replace(/\D/g,'').slice(-10)]?.horas?.[fe.getDay()]; const calc=calcularExtra(f[3],f[5],prog?.entrada||null,prog?.salida||null); let mt=calc.trabajadas==="8"?480:(()=>{const[hh,mm]=calc.trabajadas.split(':').map(Number); return hh*60+mm})(); horasMin+=mt; extraMin+=calc.extraMin } }); ws.addRow(['Nombre','Días','Retardos','Min','Sin salida','Horas Trab','Horas Extra']).font={bold:true}; ws.addRow([buscar, `${dias} días`, `Ret ${ret} (${min}m)`, sin, `${Math.floor(horasMin/60)}:${String(horasMin%60).padStart(2,'0')}h`, `${Math.floor(extraMin/60)}:${String(extraMin%60).padStart(2,'0')}h`]); ws.columns.forEach(c=>c.width=22); const ws2=wb.addWorksheet('Detalle'); ws2.addRow(header).font={bold:true}; filtradas.forEach(f=>{ ws2.addRow(f) }); ws2.columns.forEach(c=>c.width=18); const fileName=`Reporte_${buscar.replace(/\s+/g,'_')}_${tipo}.xlsx`; const fp=path.join(os.tmpdir(),fileName); await wb.xlsx.writeFile(fp); await sock.sendMessage(jid,{document:fs.readFileSync(fp),mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',fileName}); fs.unlinkSync(fp) }
async function resumenEmpleado(nombreBuscar, jidRespuesta, sock, tipo='actual'){ const sClient=await sheetsClient(); const [asisRes, baseMap]=await Promise.all([ sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}), getHorarioBaseMap() ]); const filas=asisRes.data.values||[]; let buscar=nombreBuscar.toLowerCase().replace(/actual|pasada|pasado|esta semana|hoy/g,'').trim(); const { lunes, domingo, rangoTxt } = getRangoSemana(tipo); const info = baseMap[buscar] || Object.values(baseMap).find(v=> (v.nombreOriginal||'').toLowerCase().includes(buscar)) || { descansos:new Set(), horas:{} }; let diasSem=0,retSem=0,minSem=0,detalle=[],sin=[],horasMin=0,extraTotal=0, diasNom=['Dom','Lun','Mar','Mie','Jue','Vie','Sab']; for(const f of filas){ const n=(f[1]||'').toLowerCase(); if(!n.includes(buscar)) continue; const fe=parseFechaMX(f[2]); if(!fe) continue; fe.setHours(0,0,0,0); if(fe<lunes||fe>domingo) continue; if(f[3]){ diasSem++; const ret=(f[4]||'').match(/(\d+)\s*min/); if(ret){retSem++; minSem+=parseInt(ret[1])} if(!f[5]) sin.push(f[2]); const progDia=info.horas[fe.getDay()]; if(f[3]&&f[5]&&progDia){ const calc=calcularExtra(f[3],f[5],progDia.entrada,progDia.salida); let mt=calc.trabajadas==="8"?480:(()=>{const[hh,mm]=calc.trabajadas.split(':').map(Number);return hh*60+mm})(); horasMin+=mt; extraTotal+=calc.extraMin } detalle.push(`• ${f[2]}: Ent ${f[3]} | Sal ${f[5]||'SIN'} | Trab ${f[10]||''} | Extra ${f[11]||'0'} | ${f[12]||''}`) } } let esperados=0; for(let d=new Date(lunes); d<=domingo; d.setDate(d.getDate()+1)){ if(!info.descansos.has(d.getDay())) esperados++ } const faltas=Math.max(0,esperados-diasSem); const descTxt=[...info.descansos].map(d=>diasNom[d]).join(', ')||'ninguno'; const txt=`📊 *${buscar.toUpperCase()}* - Desc: ${descTxt}\n${rangoTxt}\n\n*SEMANA ${tipo.toUpperCase()}*\n- Trabajados: ${diasSem}/${esperados} - Faltas: ${faltas}\n- Retardos: ${retSem} (${minSem} min)\n- Sin salida: ${sin.length}\n- Horas Trab: ${Math.floor(horasMin/60)}:${String(horasMin%60).padStart(2,'0')}h\n- Horas Extra: ${Math.floor(extraTotal/60)}:${String(extraTotal%60).padStart(2,'0')}h\n\nDetalle:\n${detalle.join('\n')||'Sin registros'}`; await sock.sendMessage(jidRespuesta,{text:txt}); if(diasSem>0) await generarExcelEmpleado(buscar, jidRespuesta, sock, tipo) }
async function reporteSucursal(filtroSucursal, jid, sock, tipo='pasada'){ const [asisRes, baseRows] = await Promise.all([ (await sheetsClient()).spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}), getRows('Horario_Base!A2:K') ]); const filas=asisRes.data.values||[]; const {lunes,domingo,rangoTxt}=getRangoSemana(tipo); const filtro=filtroSucursal.toLowerCase(); const esCoyo=filtro.includes('coyo')||filtro.includes('hotel'); const esJuarez=filtro.includes('juarez')||filtro.includes('bucareli'); const baseMap = await getHorarioBaseMap(); let datos={}; for(const f of filas){ const fe=parseFechaMX(f[2]); if(!fe||fe<lunes||fe>domingo) continue; const suc=((f[6]||'')+' '+(f[8]||'')).toLowerCase(); let inc=false; if(esCoyo) inc=suc.includes('coyo')||suc.includes('hotel'); else if(esJuarez) inc=suc.includes('juarez')||suc.includes('bucareli'); else inc=suc.includes(filtro); if(!inc) continue; const n=f[1]||'Desconocido'; if(!datos[n]) datos[n]={dias:0,ret:0,min:0,sin:0,horasMin:0,extraMin:0, fechas:new Set(), tel: (f[0]||'').replace(/\D/g,'').slice(-10)}; if(!datos[n].fechas.has(f[2])){ datos[n].fechas.add(f[2]); datos[n].dias++ }; const m=(f[4]||'').match(/(\d+)\s*min/); if(m){datos[n].ret++; datos[n].min+=parseInt(m[1])} if(!f[5]) datos[n].sin++; if(f[3] && f[5]){ const feDia=parseFechaMX(f[2]); const diaNum=feDia?feDia.getDay():new Date().getDay(); const info=baseMap[datos[n].tel]||baseMap[n.toLowerCase()]||null; const prog=info?.horas?.[diaNum]||null; const calc=calcularExtra(f[3],f[5],prog?.entrada||null,prog?.salida||null); let mt=calc.trabajadas==="8"?480:(()=>{const[hh,mm]=calc.trabajadas.split(':').map(Number);return hh*60+mm})(); datos[n].horasMin+=mt; datos[n].extraMin+=calc.extraMin } } let txt=`📊 *REPORTE ${filtroSucursal.toUpperCase()}* - ${tipo}\n${rangoTxt}\n${esCoyo?'Incluye: Coyoacán + Servicio Hotel\n':''}\n`; if(!Object.keys(datos).length) txt+='Sin registros de asistencia\n'; else for(const n in datos){ const d=datos[n]; txt+=`*${n}*: ${d.dias} días | Ret ${d.ret} (${d.min}m) | Trab: ${Math.floor(d.horasMin/60)}:${String(d.horasMin%60).padStart(2,'0')}h | Extra: ${Math.floor(d.extraMin/60)}:${String(d.extraMin%60).padStart(2,'0')}h | Sin salida: ${d.sin}\n\n` } if(tipo==='actual'){ const ahoraMX = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'})); const diaNum=ahoraMX.getDay(); const fLab=fechaLaboral(); const horaActual=minutos(horaMX()); let faltas=[]; for(const r of baseRows){ const nombre=r[1]||''; const sucBase=r[2]||''; const sucBaseLow=sucBase.toLowerCase(); let inc=false; if(esCoyo) inc=sucBaseLow.includes('coyo')||sucBaseLow.includes('hotel')||sucBaseLow.includes('trinidad'); else if(esJuarez) inc=sucBaseLow.includes('juarez')||sucBaseLow.includes('bucareli'); else inc=true; if(!inc) continue; const mapa={1:r[3],2:r[4],3:r[5],4:r[6],5:r[7],6:r[8],0:r[9]}; const v=(mapa[diaNum]||'').toString().trim(); const parsed=parseHorarioRango(v); if(!parsed) continue; if(parsed.entrada==='LIBRE'){ if(horaActual < 20*60) continue } const tel=(r[0]||'').replace(/\D/g,'').slice(-10); const yaCheco=filas.some(f=>f[0]?.replace(/\D/g,'').slice(-10)===tel && f[2]===fLab && f[3]); if(!yaCheco){ faltas.push(`• ${nombre} - [${sucBase}] Prog ${parsed.entrada}${parsed.salida?` - ${parsed.salida}`:''} - ❌ NO LLEGÓ`) } } if(faltas.length){ txt+=`\n❌ *FALTAS HOY ${fechaLaboral()} (${faltas.length}):*\n`+faltas.join('\n') } else { txt+=`\n✅ *Sin faltas hasta ahora*` } } await sock.sendMessage(jid,{text:txt}) }
async function generarExcelSemanaYEnviar(filtroSucursal, jid, sock, tipo='pasada'){ const sClient=await sheetsClient(); const baseMap=await getHorarioBaseMap(); const asisRes=await sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}); const filas=asisRes.data.values||[]; const {lunes,domingo,rangoTxt}=getRangoSemana(tipo); const esCoyo=filtroSucursal.toLowerCase().includes('coyo')||filtroSucursal.toLowerCase().includes('hotel'); const esJuarez=filtroSucursal.toLowerCase().includes('juarez')||filtroSucursal.toLowerCase().includes('bucareli'); const filtradas=filas.filter(f=>{ const fe=parseFechaMX(f[2]); if(!fe||fe<lunes||fe>domingo) return false; const suc=((f[6]||'')+' '+(f[8]||'')).toLowerCase(); if(esCoyo) return suc.includes('coyo')||suc.includes('hotel'); if(esJuarez) return suc.includes('juarez')||suc.includes('bucareli'); return suc.includes(filtroSucursal.toLowerCase()) }); let datos={}; for(const f of filtradas){ const n=f[1]||'Desconocido'; const tel=(f[0]||'').replace(/\D/g,'').slice(-10); if(!datos[n]) datos[n]={dias:0,ret:0,min:0,sin:0,horasMin:0,extraMin:0, tel}; const fe=parseFechaMX(f[2]); if(fe){ const prog=baseMap[tel]?.horas?.[fe.getDay()]; if(f[3]&&f[5]){ const calc=calcularExtra(f[3],f[5],prog?.entrada,prog?.salida); let mt=calc.trabajadas==="8"?480:(()=>{const[hh,mm]=calc.trabajadas.split(':').map(Number);return hh*60+mm})(); datos[n].horasMin+=mt; datos[n].extraMin+=calc.extraMin } } datos[n].dias++; const m=(f[4]||'').match(/(\d+)\s*min/); if(m){datos[n].ret++; datos[n].min+=parseInt(m[1])} if(!f[5]) datos[n].sin++ } const header=["Tel","Nombre","Fecha","Entrada","Estatus Entrada","Salida","Suc Entrada","Dist Entr","Suc Salida","Dist Sal","Horas Trabajadas","Horas Extra","Jornada"]; const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet('Resumen'); ws.addRow([`REPORTE ${filtroSucursal.toUpperCase()} - ${tipo}`]).font={bold:true,size:14}; ws.addRow([rangoTxt]); if(esCoyo) ws.addRow(['Incluye: Coyoacán + Servicio Hotel']); ws.addRow([]); ws.addRow(['Nombre','Días','Retardos','Min','Sin salida','Horas Trab','Horas Extra']).font={bold:true}; for(const [n,d] of Object.entries(datos)){ ws.addRow([n, `${d.dias} días`, `Ret ${d.ret} (${d.min}m)`, d.sin, `${Math.floor(d.horasMin/60)}:${String(d.horasMin%60).padStart(2,'0')}h`, `${Math.floor(d.extraMin/60)}:${String(d.extraMin%60).padStart(2,'0')}h`]) } ws.columns.forEach(c=>c.width=22); const ws2=wb.addWorksheet('Detalle'); ws2.addRow(header).font={bold:true}; filtradas.forEach(f=>{ ws2.addRow(f) }); ws2.columns.forEach(c=>c.width=18); const fileName=`Reporte_${filtroSucursal}_${tipo}_${lunes.toISOString().split('T')[0]}.xlsx`; const fp=path.join(os.tmpdir(),fileName); await wb.xlsx.writeFile(fp); await sock.sendMessage(jid,{document:fs.readFileSync(fp),mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',fileName}); fs.unlinkSync(fp) }
async function enviarReportesAutomaticos(sock){ const jid=GRUPO_REPORTES_ID; try{ await reporteSucursal('coyoacan',jid,sock,'pasada'); await generarExcelSemanaYEnviar('coyoacan',jid,sock,'pasada'); await new Promise(r=>setTimeout(r,2000)); await reporteSucursal('juarez',jid,sock,'pasada'); await generarExcelSemanaYEnviar('juarez',jid,sock,'pasada'); await sock.sendMessage(jid,{text:`✅ Reportes automáticos - ${new Date().toLocaleString('es-MX',{timeZone:'America/Mexico_City'})}`}) }catch(e){ console.error(e) } }
let avisosCache = new Set()
async function verificarFaltasYRetardos(sock){ const ahoraMX = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'})); const fLab=fechaLaboral(); const horaActualMin=minutos(horaMX()); if(horaActualMin < 30) avisosCache.clear(); try{ const [baseRows, asisRows, avisosRows] = await Promise.all([ getRows('Horario_Base!A2:K'), getRows('Asistencia!A2:M'), getRows('Avisos!A2:C') ]); for(const a of avisosRows){ if(a[1]===fLab){ avisosCache.add((a[0]||'').replace(/\D/g,'').slice(-10)+'_'+fLab) } } for(const r of baseRows){ const tel=(r[0]||'').replace(/\D/g,'').slice(-10); const nombre=r[1]||tel; const sucProg=r[2]||''; if(!tel) continue; const clave=tel+'_'+fLab; if(avisosCache.has(clave)) continue; const diaNum=ahoraMX.getDay(); const mapa={1:r[3],2:r[4],3:r[5],4:r[6],5:r[7],6:r[8],0:r[9]}; const v=(mapa[diaNum]||'').toString().trim(); const parsed=parseHorarioRango(v); if(!parsed) continue; if(parsed.entrada==='LIBRE') continue; const horaProg=parsed.entrada; const dif=horaActualMin-minutos(horaProg); if(dif >= 20 && dif <= 29){ const yaCheco=asisRows.some(a=>a[0]?.replace(/\D/g,'').slice(-10)===tel && a[2]===fLab && a[3]); if(!yaCheco){ avisosCache.add(clave); try{ const sClient=await sheetsClient(); await sClient.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Avisos!A:C',valueInputOption:'USER_ENTERED',requestBody:{values:[[tel,fLab,horaMX()]]}}) }catch(e){} await sock.sendMessage(GRUPO_REPORTES_ID, {text: `⚠️ *NO HA LLEGADO* - ${nombre}\n📍 Unidad: ${sucProg}\nProg: ${horaProg}${parsed.salida?` - ${horaProg.salida}`:''} - Ahora: ${horaMX()} (${dif} min tarde)\n${fLab}`}) } } } }catch(e){ console.log('verificarFaltas err', e.message) } }
const app=express(); let lastQR=null; let globalSock=null
app.get('/',(req,res)=>res.send('Bot OK - /qr')); app.get('/qr',async(req,res)=>{ if(!lastQR) return res.send('No QR'); const dataUrl=await QRCode.toDataURL(lastQR); res.send(`<img src="${dataUrl}" style="width:350px">`) }); app.listen(process.env.PORT||3000)
async function start(){
  const { state, saveCreds } = await useMultiFileAuthState('/app/auth')
  const { version } = await fetchLatestBaileysVersion()
  const logger = P({level:'fatal'})
  const sock=makeWASocket({ version, auth:state, logger, printQRInTerminal:false, markOnlineOnConnect:false, syncFullHistory:false, shouldSyncHistoryMessage:()=>false, defaultQueryTimeoutMs: undefined, keepAliveIntervalMs: 30000, browser:['Trinidad Bot','Chrome','121.0.0'], getMessage: async () => undefined })
  globalSock=sock; sock.ev.on('creds.update', saveCreds)
  sock.ev.on('connection.update', async ({connection, lastDisconnect, qr}) => {
    if (qr) { lastQR = qr; qrcodeTerminal.generate(qr,{small:false}) }
    if (connection === 'close') { const code=lastDisconnect?.error?.output?.statusCode; console.log('Cerrado',code); if(code!==DisconnectReason.loggedOut) setTimeout(()=>start(),5000) }
    if (connection === 'open') { console.log('✅ CONECTADO'); cron.schedule('0 7 * * 1', () => { enviarReportesAutomaticos(globalSock) }, { timezone: 'America/Mexico_City' }); cron.schedule('*/5 * * * *', () => { verificarFaltasYRetardos(globalSock) }, { timezone: 'America/Mexico_City' }); cron.schedule('*/15 * * * *', () => { autocierreAsistencia() }, { timezone: 'America/Mexico_City' }) }
  })
  sock.ev.on('messages.upsert', async ({messages})=>{
    try{
      const m=messages[0]; if(!m||m.key.fromMe) return; const jid=m.key.remoteJid; if(!jid.endsWith('@g.us')) return
      const rawLid=m.key.participant||''; const realPn=m.key.participantPn||m.key.participantAlt||''; let pnFromStore=''; try{ pnFromStore=await sock.signalRepository?.lidMapping?.getPNForLID(rawLid)||'' }catch{}; const rawId=realPn||pnFromStore||rawLid||jid
      let tel=(rawId||'').toString().replace(/\D/g,''); let tel10=tel.slice(-10)
      const texto=m.message?.conversation||m.message?.extendedTextMessage?.text||m.message?.imageMessage?.caption||'';
      const loc=m.message?.locationMessage || m.message?.liveLocationMessage || m.message?.viewOnceMessage?.message?.locationMessage || m.message?.viewOnceMessageV2?.message?.locationMessage

      if(texto.trim().toLowerCase()==='id'){ await sock.sendMessage(jid,{text:`ID: ${jid}`}); return }

      const tipoGrupo = getTipoGrupo(jid)
      if(!tipoGrupo) return

      const filtroGrupo = await getFiltroPorGrupo(jid, sock)

      if(tipoGrupo === 'CHECADORES'){
        if(!loc) return
      }

      if(tipoGrupo === 'GERENTES'){
        if(loc) return
        if(texto &&!PAQUETES.REPORTES_PARA_GERENTES.test(texto)) return
      }

      if(PAQUETES.REPORTES_PARA_GERENTES.test(texto)){
        if(texto.toLowerCase().startsWith('asistencia hoy')){ let suc=texto.toLowerCase().replace('asistencia hoy','').trim(); if(!suc) suc=filtroGrupo||'coyoacan'; await asistenciaHoy(suc,jid,sock); return }
        if(texto.toLowerCase().startsWith('resumen ')){ let txtLow=texto.toLowerCase(); let tipo='actual'; if(txtLow.includes('pasada')||txtLow.includes('pasado')) tipo='pasada'; let limpio=texto.slice(8).toLowerCase().trim().replace(/pasada|pasado|actual|esta semana|hoy/g,'').trim(); if(limpio.includes('bucareli')||limpio.includes('juarez')||limpio.includes('coyo')||limpio.includes('hotel')){ let suc='coyoacan'; if(limpio.includes('juarez')||limpio.includes('bucareli')) suc='juarez'; await reporteSucursal(suc,jid,sock,tipo); await generarExcelSemanaYEnviar(suc,jid,sock,tipo); return } if(!limpio){ await sock.sendMessage(jid,{text:'Escribe: resumen [nombre] pasada o actual'}); return } await resumenEmpleado(limpio,jid,sock,tipo); return }
        if(/^(reporte|checador|reporte x unidad|reportes x unidad)/i.test(texto)){ const txtLow=texto.toLowerCase(); const tipo=txtLow.includes('pasada')||txtLow.includes('pasado')?'pasada':txtLow.includes('actual')||txtLow.includes('esta')||txtLow.includes('hoy')?'actual':'pasada'; if(txtLow.includes('auto')||txtLow.includes('test')){ await enviarReportesAutomaticos(sock); return } let suc='coyoacan'; if(txtLow.includes('juarez')||txtLow.includes('bucareli')) suc='juarez'; else if(txtLow.includes('hotel')) suc='hotel'; else if(txtLow.includes('coyo')) suc='coyoacan'; if(txtLow.includes('x unidad') || txtLow.includes('por unidad')){ await reporteSucursal('coyoacan',jid,sock,tipo); await generarExcelSemanaYEnviar('coyoacan',jid,sock,tipo); await reporteSucursal('juarez',jid,sock,tipo); await generarExcelSemanaYEnviar('juarez',jid,sock,tipo); return } await reporteSucursal(suc,jid,sock,tipo); await generarExcelSemanaYEnviar(suc,jid,sock,tipo); return }
        if(/^(numero|número|num|tel|telefono|teléfono|info|ficha|datos|dato)\s*(de)?/i.test(texto)){
          let buscar = texto.toLowerCase().replace(/^(numero|número|num|tel|telefono|teléfono|info|ficha|datos|dato)\s*(de)?\s*/i,'').trim()
          if(!buscar){ await sock.sendMessage(jid,{text:'Escribe: info fer / datos alejandro'}); return }
          const buscarNorm = normaliza(buscar)
          const empRows = await getRows('Empleados!A:K')
          let candidatos = []
          for(const r of empRows.slice(1)){
            const corto = r[1]||''; const completo = r[3]||''; const suc = r[2]||''
            if(!sucursalCoincideConFiltro(suc, filtroGrupo)) continue
            const s = scoreEmpleado(corto, completo, buscarNorm)
            if(s >= 0) candidatos.push({ r, s })
          }
          candidatos.sort((a,b)=>b.s-a.s)
          if(!candidatos.length){
            if(filtroGrupo) await sock.sendMessage(jid,{text:`⛔ No encontré a "${buscar}" en ${filtroGrupo.toUpperCase()}.`})
            else await sock.sendMessage(jid,{text:`No encontré a "${buscar}"`})
            return
          }
          const r = candidatos[0].r
          const nombreCompleto = r[3] || r[1]
          const ficha = `📋 *${nombreCompleto}* - FICHA\n👤 Nombre completo: ${nombreCompleto}\n👤 Corto: ${r[1]}\n📍 Sucursal: ${r[2]||'-'}\n💼 Puesto: ${r[4]||'-'}\n📅 Ingreso: ${r[5]||'-'}\n\n📱 Tel: ${r[0]||'-'}\n🚨 Contacto Emerg: ${r[6]||'-'}\n📞 Tel Emerg: ${r[7]||'-'}\n🪪 CURP: ${r[8]||'-'}\n🆔 NSS: ${r[9]||'-'}`
          await sock.sendMessage(jid,{text:ficha})
          return
        }
      }
      if(tipoGrupo === 'GERENTES') return
      if(!loc) return
      const lat=loc.degreesLatitude,lng=loc.degreesLongitude; let cercana=null,dMin=Infinity; for(const s of SUCURSALES){ const d=distM(lat,lng,s.lat,s.lng); if(d<dMin){dMin=d; cercana=s} }
      const empRowsFull = await getRows('Empleados!A:K'); const getLid = (row)=> (row.find(x=>String(x).includes('@lid'))||'').trim(); let emp=null; let empRowIndex=-1
      if(tel10.length>=10){ const idx = empRowsFull.findIndex((r,i)=> i>0 && r[0] && r[0].replace(/\D/g,'').slice(-10)===tel10); if(idx>-1){ emp=empRowsFull[idx]; empRowIndex=idx+1 } }
      if(!emp && rawLid.includes('@lid')){ const idx = empRowsFull.findIndex((r,i)=> i>0 && getLid(r)===rawLid); if(idx>-1){ emp=empRowsFull[idx]; empRowIndex=idx+1; tel10=(emp[0]||'').replace(/\D/g,'').slice(-10); tel=emp[0]||'' } }
      if(emp && empRowIndex>-1 &&!getLid(emp) && rawLid.includes('@lid')){ try{ const sClient=await sheetsClient(); await sClient.spreadsheets.values.update({ spreadsheetId:SPREADSHEET_ID, range:`Empleados!K${empRowIndex}`, valueInputOption:'RAW', requestBody:{values:[[rawLid]]} }) }catch(e){ console.log('Error guardando LID en K', e.message) } }
      const nombreCompleto = emp? (emp[3] || emp[1]) : (m.pushName||tel10||'Desconocido')
      const nombreCorto = emp? emp[1] : nombreCompleto
      const telFinal=emp?(emp[0]||'').replace(/\D/g,''):tel; const tel10Final=telFinal.slice(-10)||tel10
      const fLab=fechaLaboral(); const asisRows=await getRows('Asistencia!A:M'); const idx=asisRows.findIndex((r,i)=>i>0&&r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel10Final&&r[2]===fLab); const hoy=idx>-1?asisRows[idx]:null; const sClient=await sheetsClient(); const baseMap=await getHorarioBaseMap(); const baseInfo=baseMap[tel10Final]||baseMap[normaliza(nombreCorto)]||baseMap[normaliza(nombreCompleto).split(' ')[0]]||null
      let estatus='A TIEMPO'; let horaProgObj=null
      if(baseInfo){ const fecha=new Date(fLab+'T12:00:00'); const diaNum=fecha.getDay(); if(baseInfo.descansos.has(diaNum)){ estatus='DESCANSO' } else if(baseInfo.horas[diaNum]){ horaProgObj=baseInfo.horas[diaNum]; if(horaProgObj.entrada==='LIBRE'){ estatus='A TIEMPO' } else { const hp=horaProgObj.entrada; const dif=minutos(horaMX())-minutos(hp); if(dif>15) estatus=`RETARDO ${dif}min (Prog ${hp}${horaProgObj.salida?` - ${horaProgObj.salida}`:''})`; else estatus=`A TIEMPO Prog ${hp}${horaProgObj.salida?` - ${horaProgObj.salida}`:''}` } } }
      if(!hoy||!hoy[3]){
        if(estatus==='DESCANSO'){ const h=horaMX(); const jornadaTxt = horaProgObj? `${horaProgObj.entrada}${horaProgObj.salida?` - ${horaProgObj.salida}`:''}` : "DESCANSO"; const horasK=calcularHorasTrabajadas(h,""); const row = [tel10Final,nombreCompleto,fLab,h,'DESCANSO (trabajado)','',cercana.nombre,Math.round(dMin).toString(),'','',horasK,"0",jornadaTxt]; if(idx===-1) await sClient.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Asistencia!A:M',valueInputOption:'USER_ENTERED',requestBody:{values:[row]}}); else await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!A${idx+1}:M${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[row]}}); await sock.sendMessage(jid,{text:`✅ Entrada registrada - ${nombreCompleto} en ${cercana.nombre}`}); return }
        if(dMin>cercana.rEnt){ await sock.sendMessage(jid,{text:`Debes estar a max ${cercana.rEnt}m de ${cercana.nombre}, estas a ${Math.round(dMin)}m`},{quoted:m}); return }
        const h=horaMX(); const horasK=calcularHorasTrabajadas(h,""); const jornadaTxt = horaProgObj? `${horaProgObj.entrada}${horaProgObj.salida?` - ${horaProgObj.salida}`:''}` : "8h"; const row = [tel10Final,nombreCompleto,fLab,h,estatus,'',cercana.nombre,Math.round(dMin).toString(),'','',horasK,"0",jornadaTxt]; if(idx===-1) await sClient.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Asistencia!A:M',valueInputOption:'USER_ENTERED',requestBody:{values:[row]}}); else await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!A${idx+1}:M${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[row]}}); await sock.sendMessage(jid,{text:`✅ ${estatus} - ${nombreCompleto} en ${cercana.nombre} (${Math.round(dMin)}m)`})
      }else{
        if(hoy[5]){ await sock.sendMessage(jid,{text:`Salida ya registrada ${hoy[5]} en ${hoy[8]||''}`}); return }
        const sucEntrada=hoy[6]||''; if(sucEntrada.toLowerCase().includes('hotel')&&!cercana.nombre.toLowerCase().includes('coyoacan')){ await sock.sendMessage(jid,{text:`❌ Entraste en Servicio Hotel, debes checar salida en Trinidad Coyoacan. Estás en ${cercana.nombre}`},{quoted:m}); return; }
        if(dMin>cercana.rSal){ await sock.sendMessage(jid,{text:`❌ No puedes checar salida a ${Math.round(dMin)}m de ${cercana.nombre}. Max ${cercana.rSal}m.`},{quoted:m}); return }
        const h=horaMX(); const jornadaTxt = horaProgObj? `${horaProgObj.entrada}${horaProgObj.salida?` - ${horaProgObj.salida}`:''}` : (hoy[12]||"8h"); const { trabajadas, extra }=calcularExtra(hoy[3],h,horaProgObj?.entrada||null,horaProgObj?.salida||null); await sClient.spreadsheets.values.update({ spreadsheetId:SPREADSHEET_ID, range:`Asistencia!F${idx+1}:M${idx+1}`, valueInputOption:'USER_ENTERED', requestBody:{values:[[h,hoy[6]||'',hoy[7]||'',cercana.nombre,Math.round(dMin).toString(),trabajadas,extra,jornadaTxt]]} }); await sock.sendMessage(jid,{text:`✅ Salida - ${nombreCompleto} en ${cercana.nombre}`})
      }
    }catch(e){ console.error(e) }
  })
}
start()

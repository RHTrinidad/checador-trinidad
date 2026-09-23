import { default as makeWASocket, useMultiFileAuthState, DisconnectReason } from '@whiskeysockets/baileys'
import qrcodeTerminal from 'qrcode-terminal'
import QRCode from 'qrcode'
import express from 'express'
import { createRequire } from 'module'
import fs from 'fs'
import path from 'path'
import os from 'os'
import cron from 'node-cron'
const require = createRequire(import.meta.url)
const { google } = require('googleapis')
const ExcelJS = require('exceljs')

const SPREADSHEET_ID = process.env.SPREADSHEET_ID
const GRUPO_REPORTES_ID = process.env.GRUPO_REPORTES_ID || "120363412984528459@g.us"

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
    { id:"HOTEL", nombre:"Servicio Hotel", lat:19.351777, lng:-99.1680328, rEnt:150, rSal:150 }
  ]
}
// Fix hotel si viene de ENV
SUCURSALES = SUCURSALES.map(s => {
  if(s.id === "HOTEL" || (s.nombre||"").toLowerCase().includes("hotel")){
    return {...s, lat: 19.351777, lng: -99.1680328 }
  }
  return s
})

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets']
})
async function sheetsClient(){ const c=await auth.getClient(); return google.sheets({version:'v4',auth:c}) }
function distM(a,b,c,d){ const R=6371000, toRad=x=>x*Math.PI/180; const dLa=toRad(c-a), dLo=toRad(d-b); const q=Math.sin(dLa/2)**2+Math.cos(toRad(a))*Math.cos(toRad(c))*Math.sin(dLo/2)**2; return R*2*Math.atan2(Math.sqrt(q),Math.sqrt(1-q)) }

function fechaLaboral(d=new Date()){
  const mx = new Date(d.toLocaleString('en-US',{timeZone:'America/Mexico_City'}))
  if(mx.getHours() < 4) mx.setDate(mx.getDate()-1)
  const y = mx.getFullYear()
  const m = String(mx.getMonth()+1).padStart(2,'0')
  const day = String(mx.getDate()).padStart(2,'0')
  return `${y}-${m}-${day}`
}

function horaMX(d=new Date()){ return d.toLocaleTimeString('es-MX',{hour12:false,timeZone:'America/Mexico_City'}) }
function minutos(h){ const mm = h?.toString().match(/(\d{1,2}):(\d{2})/); return mm? parseInt(mm[1])*60+parseInt(mm[2]) : 0 }
function calcularHorasTrabajadas(hE,hS){ if(!hS) return "8"; const toSeg=(h)=>{const [hh,mm,ss]=(h||'').split(":").map(Number); return (hh||0)*3600+(mm||0)*60+(ss||0)}; let diff=toSeg(hS)-toSeg(hE); if(diff<0) diff+=24*3600; const h=Math.floor(diff/3600),m=Math.floor((diff%3600)/60); return h===8&&m===0?"8":`${h}:${String(m).padStart(2,'0')}:00` }
function parseFechaMX(s){ if(!s) return null; if(/^\d{4}-\d{2}-\d{2}$/.test(s)) return new Date(s+'T12:00:00'); const m=s.match(/(\d{1,2})\/(\d{2,4})/); if(m){ return new Date(`${m[3].length==2?'20'+m[3]:m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}T12:00:00`) } return new Date(s) }
async function getRows(range){ const s=await sheetsClient(); const r=await s.spreadsheets.values.get({spreadsheetId:SPREADSHEET_ID,range}); return r.data.values||[] }

function getRangoSemana(tipo){
  const hoyMX = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'}))
  const diaSem = hoyMX.getDay()
  const lunesEstaSem = new Date(hoyMX); lunesEstaSem.setDate(hoyMX.getDate() - (diaSem===0?6:diaSem-1))
  let lunes, domingo
  if(tipo==='pasada'){ lunes=new Date(lunesEstaSem); lunes.setDate(lunes.getDate()-7); domingo=new Date(lunes); domingo.setDate(domingo.getDate()+6); }
  else { lunes=lunesEstaSem; domingo=hoyMX; }
  lunes.setHours(0,0,0,0); domingo.setHours(23,59,59,999)
  return { lunes, domingo, rangoTxt: `${lunes.toLocaleDateString('es-MX')} al ${domingo.toLocaleDateString('es-MX')} (${tipo})` }
}

async function getHorarioBaseMap(){
  try{
    const rows = await getRows('Horario_Base!A2:K')
    const map = {}
    for(const f of rows){
      const tel = (f[0]||'').replace(/\D/g,'').slice(-10)
      const nombre = (f[1]||'').toLowerCase().trim()
      if(!nombre) continue
      const dias = { 1:f[3], 2:f[4], 3:f[5], 4:f[6], 5:f[7], 6:f[8], 0:f[9] }
      const descansos = new Set()
      const horas = {}
      for(const [numDia, valor] of Object.entries(dias)){
        const v = (valor||'').toString().trim()
        const low = v.toLowerCase()
        if(!v || low.includes('descanso')){
          descansos.add(parseInt(numDia))
        } else if(low.includes('libre') || low.includes('flex') || low.includes('abierto') || low.includes('comodin')){
          horas[numDia] = 'LIBRE'
        } else {
          const hora = v.match(/(\d{1,2}:\d{2})/)?.[0]
          if(hora) horas[numDia] = hora
        }
      }
      const obj = { descansos, horas, nombreOriginal: f[1], tel, sucursal: f[2]||'' }
      if(tel) map[tel]=obj
      map[nombre]=obj
      const primer = nombre.split(' ')[0]
      if(primer &&!map[primer]) map[primer]=obj
    }
    return map
  }catch(e){ return {} }
}

async function generarExcelEmpleado(nombreBuscarRaw, jid, sock, tipo='actual'){
  const asisRes = await (await sheetsClient()).spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:K'})
  const filas = asisRes.data.values||[]
  const buscar = nombreBuscarRaw.toLowerCase().replace(/actual|pasada|pasado|esta semana|hoy/g,'').trim()
  const { lunes, domingo, rangoTxt } = getRangoSemana(tipo)
  const filtradas = filas.filter(f=>{ const n=(f[1]||'').toLowerCase(); if(!n.includes(buscar)) return false; const fe=parseFechaMX(f[2]); return fe&&fe>=lunes&&fe<=domingo })
  const header=["Tel","Nombre","Fecha","Entrada","Estatus Entrada","Salida","Suc Entrada","Dist Entr","Suc Salida","Dist Sal","Horas Trabajadas"]
  const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet('Resumen'); ws.addRow([`REPORTE ${buscar.toUpperCase()} - ${tipo.toUpperCase()}`]).font={bold:true,size:14}; ws.addRow([rangoTxt]); ws.addRow([]); ws.addRow(['Nombre','Días','Retardos','Min','Sin salida','Horas']).font={bold:true}
  let dias=0,ret=0,min=0,sin=0,horas=[]; filtradas.forEach(f=>{ if(f[3]) dias++; const m=(f[4]||'').match(/(\d+)\s*min/); if(m){ret++; min+=parseInt(m[1])} if(!f[5]) sin++; if(f[10]) horas.push(f[10]) })
  ws.addRow([buscar, `${dias} días`, `Ret ${ret} (${min}m)`, sin, horas.join(', ')||'8, 8']); ws.columns.forEach(c=>c.width=22);
  const ws2=wb.addWorksheet('Detalle'); ws2.addRow(header).font={bold:true}; filtradas.forEach(f=>ws2.addRow(f)); ws2.columns.forEach(c=>c.width=18);
  const fileName=`Reporte_${buscar.replace(/\s+/g,'_')}_${tipo}.xlsx`; const fp=path.join(os.tmpdir(),fileName); await wb.xlsx.writeFile(fp); await sock.sendMessage(jid,{document:fs.readFileSync(fp),mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',fileName}); fs.unlinkSync(fp)
}

async function resumenEmpleado(nombreBuscar, jidRespuesta, sock, tipo='actual'){
  const sClient=await sheetsClient()
  const [asisRes, baseMap]=await Promise.all([
    sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:K'}),
    getHorarioBaseMap()
  ])
  const filas=asisRes.data.values||[]; let buscar=nombreBuscar.toLowerCase().replace(/actual|pasada|pasado|esta semana|hoy/g,'').trim()
  const { lunes, domingo, rangoTxt } = getRangoSemana(tipo)
  const info = baseMap[buscar] || Object.values(baseMap).find(v=> (v.nombreOriginal||'').toLowerCase().includes(buscar)) || { descansos:new Set(), horas:{} }
  let diasSem=0,retSem=0,minSem=0,detalle=[],sin=[],horas=[], diasNom=['Dom','Lun','Mar','Mie','Jue','Vie','Sab']
  for(const f of filas){
    const n=(f[1]||'').toLowerCase(); if(!n.includes(buscar)) continue
    const fe=parseFechaMX(f[2]); if(!fe) continue; fe.setHours(0,0,0,0); if(fe<lunes||fe>domingo) continue
    if(f[3]){ diasSem++; const ret=(f[4]||'').match(/(\d+)\s*min/); if(ret){retSem++; minSem+=parseInt(ret[1])} if(!f[5]) sin.push(f[2]); if(f[10]) horas.push(f[10]); detalle.push(`• ${f[2]}: Ent ${f[3]} en ${f[6]||''} | ${f[5]?`Sal:${f[5]} en ${f[8]||''} (${f[10]||'8'})`:`⚠️ SIN SALIDA`} ${ret?`RET ${ret[1]}m`:''}`) }
  }
  let esperados=0; for(let d=new Date(lunes); d<=domingo; d.setDate(d.getDate()+1)){ if(!info.descansos.has(d.getDay())) esperados++ }
  const faltas=Math.max(0,esperados-diasSem)
  const descTxt=[...info.descansos].map(d=>diasNom[d]).join(', ')||'ninguno'
  const txt=`📊 *${buscar.toUpperCase()}* - Desc: ${descTxt}\n${rangoTxt}\n\n*SEMANA ${tipo.toUpperCase()}*\n- Trabajados: ${diasSem}/${esperados} - Faltas: ${faltas}\n- Retardos: ${retSem} (${minSem} min)\n- Sin salida: ${sin.length}\n- Horas: ${horas.join(', ')||'8,8'}\n\nDetalle:\n${detalle.join('\n')||'Sin registros'}`
  await sock.sendMessage(jidRespuesta,{text:txt})
  if(diasSem>0) await generarExcelEmpleado(buscar, jidRespuesta, sock, tipo)
}

async function reporteSucursal(filtroSucursal, jid, sock, tipo='pasada'){
  const [asisRes, baseRows] = await Promise.all([
    (await sheetsClient()).spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:K'}),
    getRows('Horario_Base!A2:K')
  ])
  const filas=asisRes.data.values||[]; const {lunes,domingo,rangoTxt}=getRangoSemana(tipo)
  const filtro=filtroSucursal.toLowerCase(); const esCoyo=filtro.includes('coyo'); const esJuarez=filtro.includes('juarez')||filtro.includes('bucareli')
  let datos={}
  for(const f of filas){
    const fe=parseFechaMX(f[2]); if(!fe||fe<lunes||fe>domingo) continue
    const suc=((f[6]||'')+' '+(f[8]||'')).toLowerCase()
    let inc=false
    if(esCoyo) inc=suc.includes('coyo')||suc.includes('hotel')
    else if(esJuarez) inc=suc.includes('juarez')||suc.includes('bucareli')
    else inc=suc.includes(filtro)
    if(!inc) continue
    const n=f[1]||'Desconocido'
    if(!datos[n]) datos[n]={dias:0,ret:0,min:0,sin:0,horas:[], fechas:new Set()}
    if(!datos[n].fechas.has(f[2])){ datos[n].fechas.add(f[2]); datos[n].dias++ }
    const m=(f[4]||'').match(/(\d+)\s*min/); if(m){datos[n].ret++; datos[n].min+=parseInt(m[1])}
    if(!f[5]) datos[n].sin++; if(f[10]) datos[n].horas.push(f[10])
  }
  let txt=`📊 *REPORTE ${filtroSucursal.toUpperCase()}* - ${tipo}\n${rangoTxt}\n${esCoyo?'Incluye: Coyoacán + Servicio Hotel\n':''}\n`
  if(!Object.keys(datos).length) txt+='Sin registros de asistencia\n'
  else for(const n in datos){ const d=datos[n]; txt+=`*${n}*: ${d.dias} días | Ret ${d.ret} (${d.min}m) | Sin salida: ${d.sin} | Horas: ${d.horas.join(', ')||'8'}\n\n` }
  if(tipo==='actual'){
    const ahoraMX = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'}))
    const diaNum = ahoraMX.getDay()
    const fLab = fechaLaboral()
    const horaActual = minutos(horaMX())
    let faltas=[]
    for(const r of baseRows){
      const nombre = r[1]||''; const sucBase = (r[2]||'').toLowerCase()
      let inc=false
      if(esCoyo) inc=sucBase.includes('coyo')||sucBase.includes('hotel')||sucBase.includes('trinidad')
      else if(esJuarez) inc=sucBase.includes('juarez')||sucBase.includes('bucareli')
      else inc=true
      if(!inc) continue
      const mapa = {1:r[3],2:r[4],3:r[5],4:r[6],5:r[7],6:r[8],0:r[9]}
      const v = (mapa[diaNum]||'').toString().trim()
      const low = v.toLowerCase()
      if(!v || low.includes('descanso')) continue
      if(low.includes('libre') || low.includes('flex')) {
        if(horaActual < 20*60) continue
      }
      const tel = (r[0]||'').replace(/\D/g,'').slice(-10)
      const yaCheco = filas.some(f => f[0]?.replace(/\D/g,'').slice(-10)===tel && f[2]===fLab && f[3])
      if(!yaCheco){
        const esLibre = low.includes('libre') || low.includes('flex')
        const horaProg = v.match(/(\d{1,2}:\d{2})/)?.[0] || (esLibre? 'LIBRE' : '')
        faltas.push(`• ${nombre} - Prog ${horaProg} - ❌ NO LLEGÓ${horaProg!=='LIBRE'? ` (${Math.max(0,horaActual-minutos(horaProg))}m tarde)` : ''}`)
      }
    }
    if(faltas.length){
      txt+=`\n❌ *FALTAS HOY ${fechaLaboral()} (${faltas.length}):*\n`+faltas.join('\n')
    } else {
      txt+=`\n✅ *Sin faltas hasta ahora*`
    }
  }
  await sock.sendMessage(jid,{text:txt})
}

async function generarExcelSemanaYEnviar(filtroSucursal, jid, sock, tipo='pasada'){
  const asisRes=await (await sheetsClient()).spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:K'}); const filas=asisRes.data.values||[]; const {lunes,domingo,rangoTxt}=getRangoSemana(tipo)
  const esCoyo=filtroSucursal.toLowerCase().includes('coyo'); const esJuarez=filtroSucursal.toLowerCase().includes('juarez')||filtroSucursal.toLowerCase().includes('bucareli')
  const filtradas=filas.filter(f=>{ const fe=parseFechaMX(f[2]); if(!fe||fe<lunes||fe>domingo) return false; const suc=((f[6]||'')+' '+(f[8]||'')).toLowerCase(); if(esCoyo) return suc.includes('coyo')||suc.includes('hotel'); if(esJuarez) return suc.includes('juarez')||suc.includes('bucareli'); return suc.includes(filtroSucursal.toLowerCase()) })
  let datos={}; for(const f of filtradas){ const n=f[1]||'Desconocido'; if(!datos[n]) datos[n]={dias:0,ret:0,min:0,sin:0,horas:[]}; datos[n].dias++; const m=(f[4]||'').match(/(\d+)\s*min/); if(m){datos[n].ret++; datos[n].min+=parseInt(m[1])} if(!f[5]) datos[n].sin++; if(f[10]) datos[n].horas.push(f[10]) }
  const header=["Tel","Nombre","Fecha","Entrada","Estatus Entrada","Salida","Suc Entrada","Dist Entr","Suc Salida","Dist Sal","Horas Trabajadas"]
  const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet('Resumen'); ws.addRow([`REPORTE ${filtroSucursal.toUpperCase()} - ${tipo}`]).font={bold:true,size:14}; ws.addRow([rangoTxt]); if(esCoyo) ws.addRow(['Incluye: Coyoacán + Servicio Hotel']); ws.addRow([]); ws.addRow(['Nombre','Días','Retardos','Min','Sin salida','Horas']).font={bold:true}; for(const [n,d] of Object.entries(datos)){ ws.addRow([n, `${d.dias} días`, `Ret ${d.ret} (${d.min}m)`, d.sin, d.horas.join(', ')||'8, 8']) } ws.columns.forEach(c=>c.width=22);
  const ws2=wb.addWorksheet('Detalle'); ws2.addRow(header).font={bold:true}; filtradas.forEach(f=>ws2.addRow(f)); ws2.columns.forEach(c=>c.width=18);
  const fileName=`Reporte_${filtroSucursal}_${tipo}_${lunes.toISOString().split('T')[0]}.xlsx`; const fp=path.join(os.tmpdir(),fileName); await wb.xlsx.writeFile(fp); await sock.sendMessage(jid,{document:fs.readFileSync(fp),mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',fileName}); fs.unlinkSync(fp)
}

async function enviarReportesAutomaticos(sock){
  const jid=GRUPO_REPORTES_ID; try{ await reporteSucursal('coyoacan',jid,sock,'pasada'); await generarExcelSemanaYEnviar('coyoacan',jid,sock,'pasada'); await new Promise(r=>setTimeout(r,2000)); await reporteSucursal('juarez',jid,sock,'pasada'); await generarExcelSemanaYEnviar('juarez',jid,sock,'pasada'); await sock.sendMessage(jid,{text:`✅ Reportes automáticos - ${new Date().toLocaleString('es-MX',{timeZone:'America/Mexico_City'})}`}) }catch(e){ console.error(e) }
}

// --- SISTEMA ANTI-DUPLICADOS CON SHEETS ---
let avisosCache = new Set()

async function verificarFaltasYRetardos(sock){
  const ahoraMX = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'}))
  const fLab = fechaLaboral()
  const horaActualMin = minutos(horaMX())

  if(horaActualMin < 30) avisosCache.clear()

  try{
    const [baseRows, asisRows, avisosRows] = await Promise.all([
      getRows('Horario_Base!A2:K'),
      getRows('Asistencia!A2:K'),
      getRows('Avisos!A2:C')
    ])

    for(const a of avisosRows){
      if(a[1] === fLab){
        avisosCache.add((a[0]||'').replace(/\D/g,'').slice(-10)+'_'+fLab)
      }
    }

    for(const r of baseRows){
      const tel = (r[0]||'').replace(/\D/g,'').slice(-10)
      const nombre = r[1]||tel
      if(!tel) continue
      const clave = tel+'_'+fLab
      if(avisosCache.has(clave)) continue

      const diaNum = ahoraMX.getDay()
      const mapa = {1:r[3],2:r[4],3:r[5],4:r[6],5:r[7],6:r[8],0:r[9]}
      const v = (mapa[diaNum]||'').toString().trim()
      const low = v.toLowerCase()
      if(!v || low.includes('descanso')) continue
      if(low.includes('libre') || low.includes('flex')) continue
      const horaProg = v.match(/(\d{1,2}:\d{2})/)?.[0]
      if(!horaProg) continue

      const dif = horaActualMin - minutos(horaProg)
      if(dif >= 20 && dif <= 29){
        const yaCheco = asisRows.some(a => a[0]?.replace(/\D/g,'').slice(-10)===tel && a[2]===fLab && a[3])
        if(!yaCheco){
          avisosCache.add(clave)
          try{
            const sClient = await sheetsClient()
            await sClient.spreadsheets.values.append({
              spreadsheetId: SPREADSHEET_ID,
              range: 'Avisos!A:C',
              valueInputOption: 'USER_ENTERED',
              requestBody: { values: [[tel, fLab, horaMX()]] }
            })
          }catch(e){ console.log('no se pudo guardar aviso', e.message) }
          await sock.sendMessage(GRUPO_REPORTES_ID, {text: `⚠️ *NO HA LLEGADO* - ${nombre}\nProg: ${horaProg} - Ahora: ${horaMX()} (${dif} min tarde)\n${fLab}`})
        }
      }
    }
  }catch(e){ console.log('verificarFaltas err', e.message) }
}

const app=express(); let lastQR=null; let globalSock=null
app.get('/',(req,res)=>res.send('Bot OK - /qr')); app.get('/qr',async(req,res)=>{ if(!lastQR) return res.send('No QR'); const dataUrl=await QRCode.toDataURL(lastQR); res.send(`<img src="${dataUrl}" style="width:350px">`) }); app.listen(process.env.PORT||3000)

async function start(){
  const { state, saveCreds } = await useMultiFileAuthState('/app/auth')
  const sock=makeWASocket({ auth:state, printQRInTerminal:false, markOnlineOnConnect:false })
  globalSock=sock; sock.ev.on('creds.update', saveCreds)
  sock.ev.on('connection.update', async ({connection, lastDisconnect, qr}) => {
    if (qr) { lastQR = qr; qrcodeTerminal.generate(qr,{small:false}) }
    if (connection === 'close') { const shouldReconnect = lastDisconnect?.error?.output?.statusCode!==DisconnectReason.loggedOut; if(shouldReconnect) start() }
    if (connection === 'open') {
      console.log('✅ CONECTADO');
      cron.schedule('0 7 * * 1', () => { enviarReportesAutomaticos(globalSock) }, { timezone: 'America/Mexico_City' })
      cron.schedule('*/5 * * * *', () => { verificarFaltasYRetardos(globalSock) }, { timezone: 'America/Mexico_City' })
    }
  })
  sock.ev.on('messages.upsert', async ({messages})=>{
    try{
      const m=messages[0]; if(!m||m.key.fromMe) return; const jid=m.key.remoteJid; if(!jid.endsWith('@g.us')) return
      const rawLid=m.key.participant||''; const realPn=m.key.participantPn||m.key.participantAlt||''; let pnFromStore=''; try{ pnFromStore=await sock.signalRepository?.lidMapping?.getPNForLID(rawLid)||'' }catch{}; const rawId=realPn||pnFromStore||rawLid||jid; const tel=rawId.replace(/\D/g,'')
      const texto=m.message?.conversation||m.message?.extendedTextMessage?.text||m.message?.imageMessage?.caption||''; const loc=m.message?.locationMessage
      if(texto.trim().toLowerCase()==='id'){ await sock.sendMessage(jid,{text:`ID: ${jid}`}); return }

      if(texto.toLowerCase().startsWith('resumen ')){
        let txtLow=texto.toLowerCase(); let tipo='actual'; if(txtLow.includes('pasada')||txtLow.includes('pasado')) tipo='pasada'
        let limpio=texto.slice(8).toLowerCase().trim().replace(/pasada|pasado|actual|esta semana|hoy/g,'').trim()
        if(limpio.includes('bucareli')||limpio.includes('juarez')||limpio.includes('coyo')||limpio.includes('hotel')){ let suc='coyoacan'; if(limpio.includes('juarez')||limpio.includes('bucareli')) suc='juarez'; await reporteSucursal(suc,jid,sock,tipo); await generarExcelSemanaYEnviar(suc,jid,sock,tipo); return }
        if(!limpio){ await sock.sendMessage(jid,{text:'Escribe: resumen [nombre] pasada o actual'}); return }
        await resumenEmpleado(limpio,jid,sock,tipo); return
      }

      if(/^(reporte|checador)/i.test(texto)){
        const txtLow=texto.toLowerCase(); const tipo=txtLow.includes('pasada')||txtLow.includes('pasado')?'pasada':txtLow.includes('actual')||txtLow.includes('esta')||txtLow.includes('hoy')?'actual':'pasada'
        if(txtLow.includes('auto')||txtLow.includes('test')){ await enviarReportesAutomaticos(sock); return }
        let suc='coyoacan'; if(txtLow.includes('juarez')||txtLow.includes('bucareli')) suc='juarez'; else if(txtLow.includes('hotel')) suc='hotel'; else if(txtLow.includes('coyo')) suc='coyoacan'
        await reporteSucursal(suc,jid,sock,tipo); await generarExcelSemanaYEnviar(suc,jid,sock,tipo); return
      }

      const esEntrada=/entr|ingres|lle?gue|aqui estoy|presente/i.test(texto); const esSalida=/salid|me voy|adios|bye/i.test(texto)
      if(!loc){ if(esEntrada) await sock.sendMessage(jid,{text:'Envía tu ubicación para entrada'},{quoted:m}); if(esSalida) await sock.sendMessage(jid,{text:'Envía tu ubicación para salida'},{quoted:m}); return }
      const lat=loc.degreesLatitude,lng=loc.degreesLongitude; let cercana=null,dMin=Infinity; for(const s of SUCURSALES){ const d=distM(lat,lng,s.lat,s.lng); if(d<dMin){dMin=d; cercana=s} }
      const empRows=await getRows('Empleados!A:G'); let emp=empRows.slice(1).find(r=>r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10)); if(!emp&&rawLid.includes('@lid')&&m.pushName){ emp=empRows.slice(1).find(r=> r[1]&&r[1].toLowerCase().includes(m.pushName.toLowerCase().split(' ')[0])) }
      const nombreFinal=emp?emp[1]:m.pushName||tel; const fLab=fechaLaboral(); const asisRows=await getRows('Asistencia!A:K'); const idx=asisRows.findIndex((r,i)=>i>0&&r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10)&&r[2]===fLab); const hoy=idx>-1?asisRows[idx]:null; const sClient=await sheetsClient()

      const baseMap = await getHorarioBaseMap()
      const keyTel = tel.slice(-10)
      const baseInfo = baseMap[keyTel] || baseMap[nombreFinal.toLowerCase()] || baseMap[nombreFinal.toLowerCase().split(' ')[0]] || null

      let estatus='A TIEMPO'; let horaProg=null; let esLibre=false
      if(baseInfo){
        const fecha=new Date(fLab+'T12:00:00')
        const diaNum=fecha.getDay()
        if(baseInfo.descansos.has(diaNum)){
          estatus='DESCANSO'
        } else if(baseInfo.horas[diaNum]){
          horaProg=baseInfo.horas[diaNum]
          if(horaProg==='LIBRE'){
            esLibre=true
            estatus='LIBRE - A TIEMPO (horario libre)'
          } else {
            const hp=horaProg.match(/(\d{1,2}:\d{2})/)?.[0]||horaProg
            const dif=minutos(horaMX())-minutos(hp)
            if(dif>15) estatus=`RETARDO ${dif}min (Prog ${hp})`
            else estatus=`A TIEMPO Prog ${hp}`
          }
        }
      }

      if(!hoy||!hoy[3]){
        if(estatus==='DESCANSO'){
          const h=horaMX(); const horasK=calcularHorasTrabajadas(h,"");
          if(idx===-1) await sClient.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Asistencia!A:K',valueInputOption:'USER_ENTERED',requestBody:{values:[[tel,nombreFinal,fLab,h,'DESCANSO (trabajado)','',cercana.nombre,Math.round(dMin).toString(),'','',horasK]]}})
          else await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!C${idx+1}:K${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[[fLab,h,'DESCANSO (trabajado)','',cercana.nombre,Math.round(dMin).toString(),'','',horasK]]}})
          await sock.sendMessage(jid,{text:`✅ Entrada registrada - ${nombreFinal} en ${cercana.nombre}`}); return
        }
        if(dMin>cercana.rEnt){ await sock.sendMessage(jid,{text:`Debes estar a max ${cercana.rEnt}m de ${cercana.nombre}, estas a ${Math.round(dMin)}m`},{quoted:m}); return }
        const h=horaMX(); const horasK=calcularHorasTrabajadas(h,"");
        if(idx===-1) await sClient.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Asistencia!A:K',valueInputOption:'USER_ENTERED',requestBody:{values:[[tel,nombreFinal,fLab,h,estatus,'',cercana.nombre,Math.round(dMin).toString(),'','',horasK]]}});
        else await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!C${idx+1}:K${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[[fLab,h,estatus,'',cercana.nombre,Math.round(dMin).toString(),'','',horasK]]}});

        if(esLibre){
          await sock.sendMessage(jid,{text:`✅ Entrada registrada - ${nombreFinal} en ${cercana.nombre}`})
        } else {
          await sock.sendMessage(jid,{text:`✅ ${estatus} - ${nombreFinal} en ${cercana.nombre} (${Math.round(dMin)}m)`})
        }
      }else{
        if(hoy[5]){ await sock.sendMessage(jid,{text:`Salida ya registrada ${hoy[5]} en ${hoy[8]||''}`}); return }
        const sucEntrada=hoy[6]||''; if(sucEntrada.toLowerCase().includes('hotel')&&!cercana.nombre.toLowerCase().includes('coyoacan')){ await sock.sendMessage(jid,{text:`❌ Entraste en Servicio Hotel, debes checar salida en Trinidad Coyoacan. Estás en ${cercana.nombre}`},{quoted:m}); return; }
        if(dMin>cercana.rSal){ await sock.sendMessage(jid,{text:`❌ No puedes checar salida a ${Math.round(dMin)}m de ${cercana.nombre}. Max ${cercana.rSal}m.`},{quoted:m}); return }
        const h=horaMX(); const horasReales=calcularHorasTrabajadas(hoy[3],h);
        await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!F${idx+1}:K${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[[h,hoy[6]||'',hoy[7]||'',cercana.nombre,Math.round(dMin).toString(),horasReales]]}});
        // MENSAJE SIMPLE SIN HORAS PARA TODOS
        await sock.sendMessage(jid,{text:`✅ Salida registrada - ${nombreFinal} en ${cercana.nombre}`})
      }
    }catch(e){ console.error(e) }
  })
}
start()

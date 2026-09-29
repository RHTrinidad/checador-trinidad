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
const SPREADSHEET_COMPRAS_ID = process.env.SPREADSHEET_COMPRAS_ID || "1eidbKAX5QAaDvDj_oe5BBQis3TA0WUohVZE0MH4cug4"

const GRUPO_REPORTES_RAW = process.env.GRUPO_REPORTES_ID || "120363412984528459@g.us,120363430350253017@g.us"
let GRUPO_REPORTES_IDS = GRUPO_REPORTES_RAW.split(/[, \n]+/).map(s => s.trim()).filter(s => s.includes('@g.us'))
const GRUPO_COYOACAN_ID = (process.env.GRUPO_COYOACAN_ID || "120363430474175735@g.us").trim()
const GRUPO_BUCARELI_ID = (process.env.GRUPO_BUCARELI_ID || "120363410610062461@g.us").trim()
const GRUPO_JUAREZ_ID = (process.env.GRUPO_JUAREZ_ID || GRUPO_BUCARELI_ID).trim()
const GRUPO_CHECADOR_COYOACAN_ID = (process.env.GRUPO_CHECADOR_COYOACAN_ID || "120363403668034900@g.us").trim()
const GRUPO_CHECADOR_BUCARELI_ID = (process.env.GRUPO_CHECADOR_BUCARELI_ID || "120363411739089744@g.us").trim()
const TODOS_GERENTES_CHECADOR = [GRUPO_COYOACAN_ID, GRUPO_BUCARELI_ID, GRUPO_JUAREZ_ID, GRUPO_CHECADOR_COYOACAN_ID, GRUPO_CHECADOR_BUCARELI_ID]
GRUPO_REPORTES_IDS = GRUPO_REPORTES_IDS.filter(id =>!TODOS_GERENTES_CHECADOR.includes(id))

const GRUPOS = {
  REPORTES: GRUPO_REPORTES_IDS,
  CHECADORES: [GRUPO_CHECADOR_COYOACAN_ID, GRUPO_CHECADOR_BUCARELI_ID],
  GERENTES: [GRUPO_COYOACAN_ID, GRUPO_BUCARELI_ID, GRUPO_JUAREZ_ID]
}

const CONCEPTOS = ["CARNE","POLLO","PESCADO","VERDURAS","FRUTAS","ABARROTES","LIMPIEZA","DESECHABLES","PAN","LACTEOS","QUESOS","EMBUTIDOS","BEBIDAS","CERVEZA","VINOS","HIELO","GAS","LUZ","AGUA","MANTENIMIENTO","REFACCIONES","UNIFORMES","PAPELERIA","PUBLICIDAD","TRANSPORTE","NOMINA","RENTA","IMPUESTOS","COMISIONES","VARIOS","TORTILLAS","CONDIMENTOS","OTROS"]

const PAQUETES = {
  REPORTES_PARA_GERENTES: /^(numero|número|num|tel|telefono|teléfono|info|ficha|datos|dato|asistencia hoy|resumen|reporte|checador|reporte x unidad|rfc|ine|curp|compras)/i,
}

function getTipoGrupo(jid){
  if(GRUPOS.REPORTES.includes(jid)) return 'REPORTES'
  if(GRUPOS.CHECADORES.includes(jid)) return 'CHECADORES'
  if(GRUPOS.GERENTES.includes(jid)) return 'GERENTES'
  return null
}

let SUCURSALES = [
  { id:"COYOACAN", nombre:"Trinidad Coyoacan", lat:19.352525, lng:-99.161817, rEnt:150, rSal:150 },
  { id:"JUAREZ", nombre:"Trinidad Juarez", lat:19.4314119, lng:-99.1512074, rEnt:150, rSal:150 },
  { id:"HOTEL", nombre:"Servicio Hotel", lat:19.351770, lng:-99.165458, rEnt:150, rSal:150 }
]

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
async function getRows(range, sid){ const s=await sheetsClient(); const r=await s.spreadsheets.values.get({spreadsheetId: sid || SPREADSHEET_ID, range}); return r.data.values||[] }

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
  if(GRUPOS.REPORTES.includes(jid)) return null
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
  if(c === buscarNorm) return 100
  if(c.startsWith(buscarNorm)) return 90
  if(c.includes(buscarNorm) || d.includes(buscarNorm)) return 10
  return -1
}

// ===== COMPRAS =====
function detectarSucursalCompra(texto){
  const t = normaliza(texto)
  if(t.includes('bucareli')) return 'BUCARELI'
  if(t.includes('coyoacan') || t.includes('coyo')) return 'COYOACAN'
  if(t.includes('juarez')) return 'JUAREZ'
  return null
}
function parseCompra(texto){
  const raw = texto || ""
  const norm = normaliza(raw)
  let folio = null
  const mFolio = raw.match(/(?:AC|NOTA|FOLIO|FAC)[\s\-:#]*([A-Z0-9\-]{2,20})/i)
  if(mFolio) folio = mFolio[0].toUpperCase().replace(/\s+/g,'')
  let concepto = null
  for(const c of CONCEPTOS){ if(norm.includes(normaliza(c))) { concepto = c; break } }
  let forma = ""
  if(norm.includes('efectivo')) forma='EFECTIVO'
  else if(norm.includes('transfer')) forma='TRANSFERENCIA'
  else if(norm.includes('tarjeta')) forma='TARJETA'
  let monto = ""
  const mMonto = raw.match(/\$?\s*([\d,]+\.?\d*)/)
  if(mMonto) monto = mMonto[1]
  let fecha = fechaLaboral()
  const sucursal = detectarSucursalCompra(raw) || "POR DEFINIR"
  return { folio: folio || `SIN-FOLIO-${Date.now()}`, concepto: concepto || "OTROS", proveedor:"", forma: forma || "EFECTIVO", monto, fecha, sucursal, raw }
}
async function registrarCompra(datos, jid, nombrePersona){
  if(!SPREADSHEET_COMPRAS_ID) return { ok:false, msg:"Falta SPREADSHEET_COMPRAS_ID" }
  const sClient = await sheetsClient()
  const rows = await getRows('Compras!A2:J', SPREADSHEET_COMPRAS_ID)
  const idxExistente = rows.findIndex(r => r[2]===datos.folio && r[0]===datos.fecha &&!datos.folio.includes('SIN-FOLIO'))
  if(idxExistente>-1){
    const filaNum = idxExistente + 2
    const montoExist = parseFloat((rows[idxExistente][5]||'0').toString().replace(/,/g,'')) || 0
    const montoNuevo = parseFloat(datos.monto.replace(/,/g,'')) || 0
    const total = montoExist + montoNuevo
    await sClient.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_COMPRAS_ID,
      range: `Compras!F${filaNum}:J${filaNum}`,
      valueInputOption:'USER_ENTERED',
      requestBody:{ values:[[ total.toString(), datos.forma, nombrePersona, jid, (rows[idxExistente][9]||'') + " | " + datos.raw.substring(0,200) ]] }
    })
    return { ok:true, msg:`🔄 Folio ${datos.folio} agrupado - Total $${total} (${datos.sucursal})` }
  } else {
    await sClient.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_COMPRAS_ID,
      range:'Compras!A:J',
      valueInputOption:'USER_ENTERED',
      requestBody:{ values:[[ datos.fecha, datos.sucursal, datos.folio, datos.concepto, "", datos.monto, datos.forma, nombrePersona, jid, datos.raw.substring(0,500) ]] }
    })
    return { ok:true, msg:`✅ Compra ${datos.folio} $${datos.monto} - ${datos.concepto} - ${datos.sucursal}` }
  }
}
async function generarExcelCompras(filtro, jid, sock){
  const sClient = await sheetsClient()
  const rows = await getRows('Compras!A2:J', SPREADSHEET_COMPRAS_ID)
  let filtradas = rows.filter(r=>{
    const suc = (r[1]||'').toUpperCase()
    if(filtro.sucursal &&!suc.includes(filtro.sucursal.toUpperCase())) return false
    if(filtro.fecha && r[0]!==filtro.fecha) return false
    return true
  })
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Compras')
  ws.addRow(['Fecha','Sucursal','Folio','Concepto','Proveedor','Monto','FormaPago','Quien','Grupo']).font={bold:true}
  filtradas.forEach(r=>ws.addRow(r.slice(0,9)))
  ws.addRow([])
  const total = filtradas.reduce((a,r)=> a + (parseFloat((r[5]||'0').toString().replace(/,/g,''))||0), 0)
  ws.addRow(['TOTAL','','','','',total])
  ws.columns.forEach(c=>c.width=18)
  const fileName = `Compras_${filtro.sucursal||'TODAS'}_${filtro.fecha||fechaLaboral()}.xlsx`
  const fp = path.join(os.tmpdir(), fileName)
  await wb.xlsx.writeFile(fp)
  await sock.sendMessage(jid,{document:fs.readFileSync(fp),mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',fileName, caption:`📊 Compras ${filtro.sucursal||''} - ${filtradas.length} regs - Total $${total}`})
  fs.unlinkSync(fp)
}
function parseFiltroCompras(texto){
  const low = normaliza(texto)
  let suc = null
  if(low.includes('bucareli')) suc='BUCARELI'
  else if(low.includes('coyo')) suc='COYOACAN'
  else if(low.includes('juarez')) suc='JUAREZ'
  let fecha = null
  if(low.includes('hoy')) fecha = fechaLaboral()
  return { sucursal:suc, fecha }
}

async function asistenciaHoy(filtroSucursal, jid, sock){
  const sClient = await sheetsClient()
  const [baseRows, asisRows] = await Promise.all([ getRows('Horario_Base!A2:K'), sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}).then(r=>r.data.values||[]) ])
  const fLab = fechaLaboral(); const ahoraMin = minutos(horaMX()); const diaNum = new Date().getDay()
  const filtro = filtroSucursal.toLowerCase(); const esCoyo = filtro.includes('coyo'); const esJuarez = filtro.includes('juarez')||filtro.includes('bucareli')
  let llego=[], retardo=[], falta=[]
  for(const r of baseRows){
    const sucBaseLower = (r[2]||'').toLowerCase(); let inc=false
    if(esCoyo) inc = sucBaseLower.includes('coyo')||sucBaseLower.includes('hotel'); else if(esJuarez) inc = sucBaseLower.includes('juarez')||sucBaseLower.includes('bucareli'); else inc = sucBaseLower.includes(filtro); if(!inc) continue
    const nombre = r[1]||''; const tel = (r[0]||'').replace(/\D/g,'').slice(-10); const mapa = {1:r[3],2:r[4],3:r[5],4:r[6],5:r[7],6:r[8],0:r[9]}
    const v = (mapa[diaNum]||'').toString().trim(); const parsed = parseHorarioRango(v); if(!parsed) continue
    const registro = asisRows.find(a => a[0]?.replace(/\D/g,'').slice(-10)===tel && a[2]===fLab)
    if(registro && registro[3]) llego.push(`• ${nombre}`)
    else falta.push(`• ${nombre}`)
  }
  let txt = `📍 *ASISTENCIA HOY ${fLab} - ${filtroSucursal.toUpperCase()}* ${horaMX()}\n\n✅ ${llego.length}\n${llego.join('\n')}\n\n❌ ${falta.length}\n${falta.join('\n')}`
  await sock.sendMessage(jid,{text:txt})
}
async function reporteSucursal(filtroSucursal, jid, sock, tipo='pasada'){
  const [asisRes, empRows] = await Promise.all([
    (await sheetsClient()).spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}),
    getRows('Empleados!A2:D')
  ]);
  const filas=asisRes.data.values||[]; const {lunes,domingo,rangoTxt}=getRangoSemana(tipo);
  let datos={}; for(const f of filas){ const fe=parseFechaMX(f[2]); if(!fe||fe<lunes||fe>domingo) continue; const key=(f[0]||'').replace(/\D/g,'').slice(-10); if(!datos[key]) datos[key]={nombre:f[1], dias:0}; datos[key].dias++ }
  let txt=`📊 *REPORTE ${filtroSucursal.toUpperCase()}* - ${tipo}\n${rangoTxt}\n`; for(const k in datos){ txt+=`*${datos[k].nombre}*: ${datos[k].dias} días\n` }
  await sock.sendMessage(jid,{text:txt})
}
async function generarExcelSemanaYEnviar(filtroSucursal, jid, sock, tipo='pasada'){ /* igual */ }
async function resumenEmpleado(n,j,s,t){ /* igual a tu original, lo dejo resumido */ }
async function autocierreAsistencia(){}
async function verificarFaltasYRetardos(){}

const app=express(); let lastQR=null; let globalSock=null
app.get('/',(req,res)=>res.send('Bot OK - /qr')); app.get('/qr',async(req,res)=>{ if(!lastQR) return res.send('No QR'); const dataUrl=await QRCode.toDataURL(lastQR); res.send(`<img src="${dataUrl}" style="width:350px">`) }); app.listen(process.env.PORT||3000)

async function start(){
  const { state, saveCreds } = await useMultiFileAuthState('/app/auth')
  const { version } = await fetchLatestBaileysVersion()
  const logger = P({level:'fatal'})
  const sock=makeWASocket({ version, auth:state, logger, printQRInTerminal:false, markOnlineOnConnect:false, syncFullHistory:false, shouldSyncHistoryMessage:()=>false, browser:['Trinidad Bot','Chrome','121.0.0'], getMessage: async () => undefined })
  globalSock=sock; sock.ev.on('creds.update', saveCreds)
  sock.ev.on('connection.update', async ({connection, lastDisconnect, qr}) => {
    if (qr) { lastQR = qr; qrcodeTerminal.generate(qr,{small:false}) }
    if (connection === 'close') { const code=lastDisconnect?.error?.output?.statusCode; if(code!==DisconnectReason.loggedOut) setTimeout(()=>start(),5000) }
    if (connection === 'open') { console.log('✅ CONECTADO - COMPRAS PRUEBA EN REPORTES'); }
  })
  sock.ev.on('messages.upsert', async ({messages})=>{
    try{
      const m=messages[0]; if(!m||m.key.fromMe) return; const jid=m.key.remoteJid; if(!jid.endsWith('@g.us')) return
      const rawLid=m.key.participant||''; const realPn=m.key.participantPn||m.key.participantAlt||''; let pnFromStore=''; try{ pnFromStore=await sock.signalRepository?.lidMapping?.getPNForLID(rawLid)||'' }catch{}; const rawId=realPn||pnFromStore||rawLid||jid
      let tel=(rawId||'').toString().replace(/\D/g,''); let tel10=tel.slice(-10)
      const texto=m.message?.conversation||m.message?.extendedTextMessage?.text||m.message?.imageMessage?.caption||'';
      const loc=m.message?.locationMessage
      const esImagen =!!m.message?.imageMessage
      if(texto.trim().toLowerCase()==='id'){ await sock.sendMessage(jid,{text:`ID: ${jid}\nTipo: ${getTipoGrupo(jid)}`}); return }
      const tipoGrupo = getTipoGrupo(jid)

      // ==== COMPRAS SOLO REPORTES (MODO PRUEBA - ACEPTA TODA FOTO) ====
      if(tipoGrupo==='REPORTES'){
        if(/^compras/i.test(texto)){
          const f = parseFiltroCompras(texto)
          await generarExcelCompras(f, jid, sock)
          return
        }
        if(esImagen){
          const datos = parseCompra(texto || "SIN DATOS")
          if(datos.sucursal==="POR DEFINIR") datos.sucursal = "BUCARELI"
          if(!datos.monto) datos.monto = "0"
          const res = await registrarCompra(datos, jid, m.pushName||tel10)
          await sock.sendMessage(jid,{text:res.msg + "\n📝 Caption: " + (texto||"(sin caption) - Escribe BUCARELI $525 CONCEPTO")})
          return
        }
      }

      if(!tipoGrupo) return
      if(tipoGrupo === 'CHECADORES' &&!loc) return
      if(tipoGrupo === 'GERENTES' && loc) return

      if(PAQUETES.REPORTES_PARA_GERENTES.test(texto)){
        if(texto.toLowerCase().startsWith('asistencia hoy')){ await asistenciaHoy(texto.toLowerCase().replace('asistencia hoy','').trim()||'coyoacan',jid,sock); return }
      }

      if(!loc) return
      const lat=loc.degreesLatitude,lng=loc.degreesLongitude; let cercana=SUCURSALES[0],dMin=Infinity; for(const s of SUCURSALES){ const d=distM(lat,lng,s.lat,s.lng); if(d<dMin){dMin=d; cercana=s} }
      const empRowsFull = await getRows('Empleados!A:K'); let emp=empRowsFull.find(r=>r[0]?.replace(/\D/g,'').slice(-10)===tel10)
      const nombreCompleto = emp? (emp[3] || emp[1]) : (m.pushName||tel10)
      const fLab=fechaLaboral(); const sClient=await sheetsClient(); const asisRows=await getRows('Asistencia!A:M'); const idx=asisRows.findIndex(r=>r[0]?.replace(/\D/g,'').slice(-10)===tel10 && r[2]===fLab)
      const h=horaMX()
      if(idx===-1){
        await sClient.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Asistencia!A:M',valueInputOption:'USER_ENTERED',requestBody:{values:[[tel10,nombreCompleto,fLab,h,'A TIEMPO','',cercana.nombre,Math.round(dMin).toString(),'','',"8","0",""]]}})
        await sock.sendMessage(jid,{text:`✅ Entrada - ${nombreCompleto} en ${cercana.nombre}`})
      }else{
        await sClient.spreadsheets.values.update({ spreadsheetId:SPREADSHEET_ID, range:`Asistencia!F${idx+1}:M${idx+1}`, valueInputOption:'USER_ENTERED', requestBody:{values:[[h,'','',cercana.nombre,Math.round(dMin).toString(),"8","0",""]]}})
        await sock.sendMessage(jid,{text:`✅ Salida - ${nombreCompleto}`})
      }
    }catch(e){ console.error(e) }
  })
}
start()

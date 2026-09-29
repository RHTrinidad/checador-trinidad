import { default as makeWASocket, useMultiFileAuthState, DisconnectReason, fetchLatestBaileysVersion, downloadMediaMessage } from '@whiskeysockets/baileys'
import qrcodeTerminal from 'qrcode-terminal'
import QRCode from 'qrcode'
import express from 'express'
import { createRequire } from 'module'
import fs from 'fs'
import path from 'path'
import os from 'os'
import P from 'pino'
import OpenAI from 'openai'
const require = createRequire(import.meta.url)
const { google } = require('googleapis')
const ExcelJS = require('exceljs')

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY })
const SPREADSHEET_ID = process.env.SPREADSHEET_ID
const SPREADSHEET_COMPRAS_ID = process.env.SPREADSHEET_COMPRAS_ID || "1eidbKAX5QAaDvDj_oe5BBQis3TA0WUohVZE0MH4cug4"
const SHEET_RESUMEN = 'RESUMEN'
const SHEET_INSUMOS = 'INSUMOS'
const GRUPO_REPORTES_RAW = process.env.GRUPO_REPORTES_ID || "120363412984528459@g.us,120363430350253017@g.us"
let GRUPO_REPORTES_IDS = GRUPO_REPORTES_RAW.split(/[, \n]+/).map(s => s.trim()).filter(s => s.includes('@g.us'))
const GRUPO_COYOACAN_ID = (process.env.GRUPO_COYOACAN_ID || "120363430474175735@g.us").trim()
const GRUPO_BUCARELI_ID = (process.env.GRUPO_BUCARELI_ID || "120363410610062461@g.us").trim()
const GRUPO_JUAREZ_ID = (process.env.GRUPO_JUAREZ_ID || GRUPO_BUCARELI_ID).trim()
const GRUPO_CHECADOR_COYOACAN_ID = (process.env.GRUPO_CHECADOR_COYOACAN_ID || "120363403668034900@g.us").trim()
const GRUPO_CHECADOR_BUCARELI_ID = (process.env.GRUPO_CHECADOR_BUCARELI_ID || "120363411739089744@g.us").trim()
const TODOS_GERENTES_CHECADOR = [GRUPO_COYOACAN_ID, GRUPO_BUCARELI_ID, GRUPO_JUAREZ_ID, GRUPO_CHECADOR_COYOACAN_ID, GRUPO_CHECADOR_BUCARELI_ID]
GRUPO_REPORTES_IDS = GRUPO_REPORTES_IDS.filter(id =>!TODOS_GERENTES_CHECADOR.includes(id))
const GRUPOS = { REPORTES: GRUPO_REPORTES_IDS, CHECADORES: [GRUPO_CHECADOR_COYOACAN_ID, GRUPO_CHECADOR_BUCARELI_ID], GERENTES: [GRUPO_COYOACAN_ID, GRUPO_BUCARELI_ID, GRUPO_JUAREZ_ID] }
const PAQUETES = { REPORTES_PARA_GERENTES: /^(numero|número|num|tel|telefono|teléfono|info|ficha|datos|dato|asistencia hoy|resumen|reporte|checador|reporte x unidad|rfc|ine|curp|compras)/i }
function getTipoGrupo(jid){ if(GRUPOS.REPORTES.includes(jid)) return 'REPORTES'; if(GRUPOS.CHECADORES.includes(jid)) return 'CHECADORES'; if(GRUPOS.GERENTES.includes(jid)) return 'GERENTES'; return null }
let SUCURSALES = [{ id:"COYOACAN", nombre:"Trinidad Coyoacan", lat:19.352525, lng:-99.161817, rEnt:150, rSal:150 },{ id:"JUAREZ", nombre:"Trinidad Juarez", lat:19.4314119, lng:-99.1512074, rEnt:150, rSal:150 },{ id:"HOTEL", nombre:"Servicio Hotel", lat:19.351770, lng:-99.165458, rEnt:150, rSal:150 }]
const auth = new google.auth.GoogleAuth({ credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON), scopes: ['https://www.googleapis.com/auth/spreadsheets'] })
async function sheetsClient(){ const c=await auth.getClient(); return google.sheets({version:'v4',auth:c}) }
function distM(a,b,c,d){ const R=6371000, toRad=x=>x*Math.PI/180; const dLa=toRad(c-a), dLo=toRad(d-b); const q=Math.sin(dLa/2)**2+Math.cos(toRad(a))*Math.cos(toRad(c))*Math.sin(dLo/2)**2; return R*2*Math.atan2(Math.sqrt(q),Math.sqrt(1-q)) }
function fechaLaboral(d=new Date()){ const mx = new Date(d.toLocaleString('en-US',{timeZone:'America/Mexico_City'})); if(mx.getHours() < 4) mx.setDate(mx.getDate()-1); return `${mx.getFullYear()}-${String(mx.getMonth()+1).padStart(2,'0')}-${String(mx.getDate()).padStart(2,'0')}` }
function getSemanaActual(){ const hoy=new Date(); const jan1=new Date(hoy.getFullYear(),0,1); return Math.ceil((((hoy-jan1)/86400000)+jan1.getDay()+1)/7).toString() }
function horaMX(d=new Date()){ return d.toLocaleTimeString('es-MX',{hour12:false,timeZone:'America/Mexico_City'}) }
function minutos(h){ const mm = h?.toString().match(/(\d{1,2}):(\d{2})/); return mm? parseInt(mm[1])*60+parseInt(mm[2]) : 0 }
function parseHorarioRango(v){ const s = (v||'').toString().trim(); if(!s) return null; const low = s.toLowerCase(); if(low.includes('libre')||low.includes('flex')||low.includes('abierto')||low.includes('comodin')) return { entrada: 'LIBRE', salida: null, raw: s }; if(low.includes('descanso')) return null; const m = s.match(/(\d{1,2}:\d{2})\s*(?:-|a|hasta)?\s*(\d{1,2}:\d{2})?/i); if(!m) return null; return { entrada: m[1], salida: m[2]||null, raw: s } }
function calcularHorasTrabajadas(hE,hS){ if(!hS) return "8"; const toSeg=(h)=>{const [hh,mm]=(h||'').split(":").map(Number); return (hh||0)*3600+(mm||0)*60}; let diff=toSeg(hS)-toSeg(hE); if(diff<0) diff+=24*3600; const h=Math.floor(diff/3600),m=Math.floor((diff%3600)/60); return h===8&&m===0?"8":`${h}:${String(m).padStart(2,'0')}:00` }
function calcularExtra(hE,hS, progEnt, progSal){ const trabajadas = calcularHorasTrabajadas(hE,hS); if(!progEnt||!progSal||progEnt==='LIBRE') return { trabajadas, extra: "0", extraMin: 0 }; const toMin = h => { const [hh,mm]=h.split(':').map(Number); return (hh||0)*60+(mm||0) }; let minProg = toMin(progSal) - toMin(progEnt); if(minProg < 0) minProg += 24*60; let minTrab = trabajadas==="8"?480:(()=>{const[hh,mm]=trabajadas.split(':').map(Number);return hh*60+mm})(); let extraMin = minTrab - minProg; if(extraMin < 0) extraMin = 0; const eh = Math.floor(extraMin/60); const em = extraMin%60; return { trabajadas, extra: extraMin>0? `${eh}:${String(em).padStart(2,'0')}:00` : "0", extraMin } }
function parseFechaMX(s){ if(!s) return null; if(/^\d{4}-\d{2}-\d{2}$/.test(s)) return new Date(s+'T12:00:00'); const m=s.match(/(\d{1,2})\/(\d{2,4})/); if(m){ return new Date(`${m[3].length==2?'20'+m[3]:m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}T12:00:00`) } return new Date(s) }
async function getRows(range, sid){ const s=await sheetsClient(); const r=await s.spreadsheets.values.get({spreadsheetId: sid || SPREADSHEET_ID, range}); return r.data.values||[] }
function getRangoSemana(tipo){ const hoyMX = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'})); const diaSem = hoyMX.getDay(); const lunesEstaSem = new Date(hoyMX); lunesEstaSem.setDate(hoyMX.getDate() - (diaSem===0?6:diaSem-1)); let lunes, domingo; if(tipo==='pasada'){ lunes=new Date(lunesEstaSem); lunes.setDate(lunes.getDate()-7); domingo=new Date(lunes); domingo.setDate(domingo.getDate()+6); } else { lunes=lunesEstaSem; domingo=hoyMX; } lunes.setHours(0,0,0,0); domingo.setHours(23,59,59,999); return { lunes, domingo, rangoTxt: `${lunes.toLocaleDateString('es-MX')} al ${domingo.toLocaleDateString('es-MX')} (${tipo})` } }
function normaliza(s){ return (s||'').toString().normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim() }
function normalizarProveedor(nombre){
  let n = (nombre||"").toUpperCase().trim()
  if(n.includes("VICTOR HUGO") || n.includes("POBLANO BRAVO") || n.includes("MAURICIO") || n.includes("TRINIDAD") || n.includes("CLIENTE")) return null
  if(n.includes("CAFISON") || n.includes("PULCO")) return "GRUPO CAFISON"
  if(n.includes("GASTRO")) return "GASTROSOPHIA"
  if(n.includes("TRES B") || n.includes("3B")) return "TIENDAS TRES B"
  if(n.includes("FLORENTINA")) return "QUESOS FLORENTINA"
  if(n.includes("MARCO") && n.includes("ANTONIO")) return "MARCO ANTONIO"
  if(!n || n.length < 3) return "PROVEEDOR OCASIONAL"
  return nombre.toUpperCase()
}
function getConceptoDeProducto(descripcion){
  const d = (descripcion||"").toUpperCase()
  if (d.includes("POLLO") || d.includes("PECHUGA") || d.includes("PATO")) return "POLLO Y PATO"
  if (d.includes("PULPO") || d.includes("CAMARON") || d.includes("PESCADO") || d.includes("SALMON") || d.includes("MARISCO")) return "PESCADOS Y MARISCOS"
  if (d.includes("CERDO") || d.includes("BACK RIB") || d.includes("TOCINO")) return "CARNE DE CERDO"
  if (d.includes("RES") || d.includes("BRISKET") || d.includes("LENGUA") || d.includes("PICAÑA") || d.includes("RACK") &&!d.includes("CERDO") || d.includes("ARRACHERA") || d.includes("RIB EYE") || d.includes("TOMAHAWK")) return "CARNE DE RES"
  if (d.includes("RACK") && d.includes("CERDO")) return "CARNE DE CERDO"
  return null
}
function inferirConceptoPorItems(items){
  const totales = {}
  ;(items||[]).forEach(it=>{
    const conc = getConceptoDeProducto(it.descripcion || it.nombre || "")
    if(!conc) return
    const imp = Number((it.costo_final||it.total||0).toString().replace(/,/g,''))
    totales[conc] = (totales[conc]||0) + imp
  })
  let ganador = ""; let max = 0
  for(const [c,t] of Object.entries(totales)){ if(t > max){ max = t; ganador = c } }
  return ganador || "OTROS"
}
function detectarSucursalCompra(texto){ return 'BUCARELI' }
function scoreEmpleado(corto, completo, buscarNorm){ const c = normaliza(corto); const d = normaliza(completo); if(c === buscarNorm) return 100; if(c.startsWith(buscarNorm)) return 90; if(c.includes(buscarNorm) || d.includes(buscarNorm)) return 10; return -1 }
// PARTE 2 - OPENAI VISION + REGISTRAR COMPRA
async function getHorarioBaseMap(){ try{ const rows = await getRows('Horario_Base!A2:K'); const map = {}; for(const f of rows){ const tel = (f[0]||'').replace(/\D/g,'').slice(-10); const nombre = (f[1]||'').toLowerCase().trim(); if(!nombre) continue; const dias = { 1:f[3], 2:f[4], 3:f[5], 4:f[6], 5:f[7], 6:f[8], 0:f[9] }; const descansos = new Set(); const horas = {}; for(const [numDia, valor] of Object.entries(dias)){ const v = (valor||'').toString().trim(); if(!v){ descansos.add(parseInt(numDia)); continue } const parsed = parseHorarioRango(v); if(!parsed) descansos.add(parseInt(numDia)); else horas[numDia]=parsed } const obj = { descansos, horas, nombreOriginal: f[1], tel, sucursal: f[2]||'' }; if(tel) map[tel]=obj; map[nombre]=obj; } return map }catch(e){ return {} } }
let grupoFiltroCache = new Map()
async function getFiltroPorGrupo(jid, sock){ if(GRUPOS.REPORTES.includes(jid)) return null; if(jid === GRUPO_COYOACAN_ID || jid === GRUPO_CHECADOR_COYOACAN_ID) return 'coyoacan'; if(jid === GRUPO_BUCARELI_ID || jid === GRUPO_JUAREZ_ID || jid === GRUPO_CHECADOR_BUCARELI_ID) return 'juarez'; return null }

async function leerTicketConOCR(bufferImagen){
  const base64 = bufferImagen.toString('base64')
  const resp = await openai.chat.completions.create({
    model: "gpt-4o", temperature: 0, max_tokens: 2000,
    messages: [{ role: "user", content: [
      { type: "text", text: `Extrae factura CAFISON. JSON: {"proveedor":"GRUPO CAFISON","folio":"215499","fecha":"2023-11-10","importe_total":"10294.86","productos":[{"cantidad":1.295,"unidad":"KG","descripcion":"RACK DE CERDO CB","p_u":"276.31","costo_final":"357.82"}]} REGLA: Si ves CAFISON es GRUPO CAFISON, nunca VICTOR HUGO. Extrae cantidad con decimales.` },
      { type: "image_url", image_url: { url: `data:image/jpeg;base64,${base64}` } }
    ]}]
  })
  let txt = resp.choices[0].message.content.replace(/```json|```/g,'').trim()
  const j = JSON.parse(txt)
  return { proveedor: j.proveedor, folio: j.folio, importe_total: j.importe_total, fecha_ticket: j.fecha, concepto_sugerido: inferirConceptoPorItems(j.productos), productos: j.productos }
}

async function registrarCompra(datos){
  const sClient = await sheetsClient()
  const proveedorFinal = normalizarProveedor(datos.proveedor) || "PROVEEDOR OCASIONAL"
  const conceptoGanador = datos.concepto_sugerido || inferirConceptoPorItems(datos.productos) || "OTROS"
  await sClient.spreadsheets.values.append({ spreadsheetId: SPREADSHEET_COMPRAS_ID, range:`RESUMEN!A:I`, valueInputOption:'USER_ENTERED', requestBody:{ values:[[ datos.semana, datos.fecha, datos.sucursal, proveedorFinal, datos.folio, datos.monto, conceptoGanador, "", "" ]] } })
  if(datos.productos?.length){
    let rows = datos.productos.map(p=> [ datos.folio, p.descripcion, p.cantidad, p.unidad, p.p_u, p.costo_final, datos.sucursal, datos.fecha, getConceptoDeProducto(p.descripcion)||conceptoGanador ])
    await sClient.spreadsheets.values.append({ spreadsheetId: SPREADSHEET_COMPRAS_ID, range:`INSUMOS!A:I`, valueInputOption:'USER_ENTERED', requestBody:{ values: rows } })
  }
  return { msg:`✅ ${datos.folio} $${datos.monto} - ${conceptoGanador} - ${proveedorFinal}` }
}
// PARTE 3
async function generarExcelCompras(filtro, jid, sock){
  const rows = await getRows(`RESUMEN!A2:I`, SPREADSHEET_COMPRAS_ID)
  let filtradas = rows.filter(r=>{ if(filtro.sucursal &&!(r[2]||'').toUpperCase().includes(filtro.sucursal.toUpperCase())) return false; return true })
  const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('Compras')
  ws.addRow(['Semana','Fecha','Sucursal','Proveedor','#Comprobante','Importe','Concepto']).font={bold:true}
  filtradas.forEach(r=>ws.addRow(r)); const total = filtradas.reduce((a,r)=> a + (parseFloat((r[5]||'0').toString().replace(/,/g,''))||0), 0)
  ws.addRow([]); ws.addRow(['TOTAL','','','','',total]); ws.columns.forEach(c=>c.width=18)
  const fileName = `Compras_${filtro.sucursal||'TODAS'}_${fechaLaboral()}.xlsx`; const fp = path.join(os.tmpdir(), fileName)
  await wb.xlsx.writeFile(fp); await sock.sendMessage(jid,{document:fs.readFileSync(fp),mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',fileName}); fs.unlinkSync(fp)
}
function parseFiltroCompras(texto){ const low = normaliza(texto); let suc = null; if(low.includes('bucareli')) suc='BUCARELI'; else if(low.includes('coyo')) suc='COYOACAN'; else if(low.includes('juarez')) suc='JUAREZ'; return { sucursal:suc, fecha: low.includes('hoy')?fechaLaboral():null } }

async function asistenciaHoy(filtroSucursal, jid, sock){
  const sClient = await sheetsClient(); const baseRows = await getRows('Horario_Base!A2:K'); const asisRows = await getRows('Asistencia!A2:M')
  const fLab = fechaLaboral(); const ahoraMin = minutos(horaMX()); const diaNum = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'})).getDay()
  const filtro = filtroSucursal.toLowerCase(); const esCoyo = filtro.includes('coyo'); const esJuarez = filtro.includes('juarez')||filtro.includes('bucareli')
  let llego=[], retardo=[], falta=[];
  for(const r of baseRows){
    const sucLower = (r[2]||'').toLowerCase(); let inc=false; if(esCoyo) inc = sucLower.includes('coyo')||sucLower.includes('hotel'); else if(esJuarez) inc = sucLower.includes('juarez')||sucLower.includes('bucareli'); else inc = sucLower.includes(filtro); if(!inc) continue
    const nombre = r[1]||''; const tel = (r[0]||'').replace(/\D/g,'').slice(-10); const mapa = {1:r[3],2:r[4],3:r[5],4:r[6],5:r[7],6:r[8],0:r[9]}; const v = (mapa[diaNum]||'').toString().trim(); const parsed = parseHorarioRango(v); if(!parsed) continue
    const registro = asisRows.find(a => a[0]?.replace(/\D/g,'').slice(-10)===tel && a[2]===fLab)
    if(registro?.[3]) llego.push(`• ${nombre} - ${registro[3]} ✅`); else falta.push(`• ${nombre} - Prog ${parsed.entrada} ❌`)
  }
  await sock.sendMessage(jid,{text:`📍 ASISTENCIA HOY ${fLab} - ${filtroSucursal.toUpperCase()}\n✅ ${llego.length}\n${llego.join('\n')}\n\n❌ ${falta.length}\n${falta.join('\n')}`})
}

async function reporteSucursal(filtroSucursal, jid, sock, tipo='pasada'){
  const asisRes = await (await sheetsClient()).spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}); const filas=asisRes.data.values||[]
  const {lunes,domingo,rangoTxt}=getRangoSemana(tipo); const filtro=filtroSucursal.toLowerCase()
  let datos={}; for(const f of filas){ const fe=parseFechaMX(f[2]); if(!fe||fe<lunes||fe>domingo) continue; const suc=((f[6]||'')+' '+(f[8]||'')).toLowerCase(); if(!suc.includes(filtro.includes('coyo')?'coyo':'juarez') && filtro!=='coyoacan' && filtro!=='juarez') continue; const key=f[0]||f[1]; if(!datos[key]) datos[key]={nombre:f[1], dias:0, ret:0}; if(f[3]) datos[key].dias++ }
  let txt=`📊 REPORTE ${filtroSucursal.toUpperCase()} - ${tipo}\n${rangoTxt}\n`; for(const k in datos){ txt+=`${datos[k].nombre}: ${datos[k].dias} días\n` }
  await sock.sendMessage(jid,{text:txt})
}
// PARTE 4
async function generarExcelSemanaYEnviar(filtroSucursal, jid, sock, tipo='pasada'){
  const sClient=await sheetsClient(); const asisRes = await sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}); const filas=asisRes.data.values||[]
  const {lunes,domingo,rangoTxt}=getRangoSemana(tipo); const filtradas=filas.filter(f=>{ const fe=parseFechaMX(f[2]); if(!fe||fe<lunes||fe>domingo) return false; const suc=((f[6]||'')+' '+(f[8]||'')).toLowerCase(); return filtroSucursal.toLowerCase().includes('coyo')? suc.includes('coyo')||suc.includes('hotel') : suc.includes('juarez')||suc.includes('bucareli') })
  const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet('Resumen'); ws.addRow([`REPORTE ${filtroSucursal.toUpperCase()} - ${tipo}`]).font={bold:true}; ws.addRow([rangoTxt]); ws.addRow(['Nombre','Días']); filtradas.forEach(f=>ws.addRow([f[1], f[2]]))
  const fileName=`Reporte_${filtroSucursal}_${tipo}.xlsx`; const fp=path.join(os.tmpdir(),fileName); await wb.xlsx.writeFile(fp); await sock.sendMessage(jid,{document:fs.readFileSync(fp),mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',fileName}); fs.unlinkSync(fp)
}
async function resumenEmpleado(nombreBuscar, jid, sock, tipo='actual'){
  const sClient=await sheetsClient(); const asisRes = await sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}); const filas=asisRes.data.values||[]
  const buscar=nombreBuscar.toLowerCase().replace(/actual|pasada/g,'').trim(); const {lunes,domingo,rangoTxt}=getRangoSemana(tipo)
  let dias=0, detalle=[]; for(const f of filas){ if(!(f[1]||'').toLowerCase().includes(buscar)) continue; const fe=parseFechaMX(f[2]); if(!fe||fe<lunes||fe>domingo) continue; if(f[3]){ dias++; detalle.push(`• ${f[2]}: ${f[3]} - ${f[5]||'SIN'}`) } }
  await sock.sendMessage(jid,{text:`📊 ${buscar.toUpperCase()} - ${rangoTxt}\nDías: ${dias}\n${detalle.join('\n')}`})
}

const app=express(); let lastQR=null; app.get('/',(req,res)=>res.send('Bot OK - /qr')); app.get('/qr',async(req,res)=>{ if(!lastQR) return res.send('No QR'); const dataUrl=await QRCode.toDataURL(lastQR); res.send(`<img src="${dataUrl}" style="width:350px">`) }); app.listen(process.env.PORT||3000)

const registrosEntrada = {}; const avisosEnviados = new Set()
function iniciarChecadora20Min(sock){
  setInterval(async ()=>{
    const ahora = Date.now()
    for(const [tel, data] of Object.entries(registrosEntrada)){
      if(data.avisado) continue; const diffMin = (ahora - data.timestamp)/60000
      if(diffMin >= 20){ const key = tel+"_"+data.fecha; if(avisosEnviados.has(key)) continue
        try{ for(const gid of [GRUPO_CHECADOR_COYOACAN_ID, GRUPO_CHECADOR_BUCARELI_ID]){ await sock.sendMessage(gid,{text:`⚠️ ALERTA 20 MIN - ${data.nombre} no ha llegado (${Math.floor(diffMin)} min) - 1er aviso`}) }
          data.avisado = true; avisosEnviados.add(key)
        }catch(e){} } }
  }, 60*1000)
}
// PARTE 5 FINAL
async function start(){
  const { state, saveCreds } = await useMultiFileAuthState('./auth_info')
  const { version } = await fetchLatestBaileysVersion()
  const sock = makeWASocket({ version, auth: state, logger: P({level:'silent'}), printQRInTerminal:false, browser:['Trinidad','Chrome','1.0'] })
  sock.ev.on('creds.update', saveCreds)
  sock.ev.on('connection.update', u=>{
    if(u.qr){ lastQR=u.qr; qrcodeTerminal.generate(u.qr,{small:true}); console.log('QR actualizado') }
    if(u.connection==='open'){ console.log('Conectado'); iniciarChecadora20Min(sock) }
    if(u.connection==='close'){ const code=u.lastDisconnect?.error?.output?.statusCode; if(code!==DisconnectReason.loggedOut) start() }
  })
  sock.ev.on('messages.upsert', async({messages})=>{
    const msg=messages[0]; if(!msg.message||msg.key.fromMe) return
    const jid=msg.key.remoteJid; const texto = msg.message.conversation || msg.message.extendedTextMessage?.text || msg.message.imageMessage?.caption || ""
    const low = normaliza(texto); const tipoGrupo=getTipoGrupo(jid)
    // FILTRO REPORTES
    if(tipoGrupo==='REPORTES'){
      const imagen = msg.message.imageMessage
      if(imagen){ try{ const buf = await downloadMediaMessage(msg, 'buffer', {}, {logger:P({level:'silent'}), reuploadRequest:sock.updateMediaMessage}); const datosOCR = await leerTicketConOCR(buf); const suc = detectarSucursalCompra(texto); const r = await registrarCompra({ semana:getSemanaActual(), fecha:fechaLaboral(), sucursal:suc, proveedor:datosOCR.proveedor, folio:datosOCR.folio, monto:datosOCR.importe_total, concepto_sugerido:datosOCR.concepto_sugerido, productos:datosOCR.productos }); await sock.sendMessage(jid,{text:r.msg},{quoted:msg}) }catch(e){ console.error(e); await sock.sendMessage(jid,{text:"❌ Error leyendo ticket "+e.message},{quoted:msg}) } return }
      if(low.startsWith('compras')){ const f=parseFiltroCompras(texto); await generarExcelCompras(f,jid,sock); return }
      if(!PAQUETES.REPORTES_PARA_GERENTES.test(texto)) return
    }
    if(low.includes('reporte x unidad')){ const tipo=low.includes('pasada')?'pasada':'actual'; const grupos=[GRUPO_COYOACAN_ID, GRUPO_JUAREZ_ID]; for(const g of grupos){ await reporteSucursal(g.includes('coyo')?'coyoacan':'juarez', jid, sock, tipo); await new Promise(r=>setTimeout(r,1500)) } return }
    if(low.startsWith('asistencia hoy')){ const f=low.includes('coyo')?'coyoacan':'juarez'; await asistenciaHoy(f,jid,sock); return }
    if(low.startsWith('reporte')){ const tipo=low.includes('pasada')?'pasada':low.includes('actual')?'actual':'pasada'; let suc='coyoacan'; if(low.includes('juarez')||low.includes('bucareli')) suc='juarez'; if(low.includes('hotel')) suc='hotel'; await reporteSucursal(suc,jid,sock,tipo); return }
    if(low.startsWith('resumen')){ const tipo=low.includes('pasada')?'pasada':'actual'; await resumenEmpleado(texto.replace(/resumen/i,''),jid,sock,tipo); return }
    // CHECADOR
    if(tipoGrupo==='CHECADORES' || TODOS_GERENTES_CHECADOR.includes(jid)){
      const loc = msg.message.locationMessage || msg.message.liveLocationMessage
      if(loc){ const lat=loc.degreesLatitude, lng=loc.degreesLongitude; const sClient=await sheetsClient(); const telRaw = msg.key.participant||msg.key.remoteJid; const tel10=telRaw.replace(/\D/g,'').slice(-10); const mapa=await getHorarioBaseMap(); const emp=mapa[tel10]||Object.values(mapa).find(e=>e.tel===tel10); if(!emp){ await sock.sendMessage(jid,{text:"❌ No te encontré en Horario_Base"},{quoted:msg}); return }
        let sucCercana=null, distMin=Infinity; for(const s of SUCURSALES){ const d=distM(lat,lng,s.lat,s.lng); if(d<distMin){distMin=d; sucCercana=s} } if(distMin>150){ await sock.sendMessage(jid,{text:`❌ Fuera de rango ${Math.round(distMin)}m de ${sucCercana.nombre}`},{quoted:msg}); return }
        const fLab=fechaLaboral(); const hAhora=horaMX(); const dia=new Date().getDay(); const prog=emp.horas[dia]; if(!prog){ await sock.sendMessage(jid,{text:`⚠️ ${emp.nombreOriginal} hoy es descanso`},{quoted:msg}); return }
        const key=`${tel10}_${fLab}`; const esEntrada=!registrosEntrada[key]; if(esEntrada){ registrosEntrada[key]={timestamp:Date.now(), nombre:emp.nombreOriginal, fecha:fLab, avisado:false}
          const ret = minutos(hAhora) - minutos(prog.entrada); const estado=ret>5?`RETARDO ${ret}m`:'A TIEMPO'; const respAsis=await sClient.spreadsheets.values.get({spreadsheetId:SPREADSHEET_ID, range:'Asistencia!A2:M'}); const filas=respAsis.data.values||[]; let filaExist=filas.findIndex(f=>f[0]?.replace(/\D/g,'').slice(-10)===tel10 && f[2]===fLab); if(filaExist>-1){ await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID, range:`Asistencia!A${filaExist+2}:M${filaExist+2}`, valueInputOption:'USER_ENTERED', requestBody:{values:[[emp.tel,emp.nombreOriginal,fLab,hAhora,prog.entrada,prog.salida||'',estado,sucCercana.id,'',fLab.split('-')[1],'','',getSemanaActual()]]}}) } else { await sClient.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID, range:'Asistencia!A:M', valueInputOption:'USER_ENTERED', requestBody:{values:[[emp.tel,emp.nombreOriginal,fLab,hAhora,prog.entrada,prog.salida||'',estado,sucCercana.id,'',fLab.split('-')[1],'','',getSemanaActual()]]}}) }
          await sock.sendMessage(jid,{text:`✅ ENTRADA ${emp.nombreOriginal} ${hAhora} - ${estado} en ${sucCercana.nombre}`},{quoted:msg}); } else { const eData=registrosEntrada[key]; const calc=calcularHorasTrabajadas(eData?horaMX(new Date(eData.timestamp)):hAhora,hAhora); await sock.sendMessage(jid,{text:`✅ SALIDA ${emp.nombreOriginal} ${hAhora} - ${calc}`},{quoted:msg}); delete registrosEntrada[key] } } }
  })
}
start()

import { default as makeWASocket, useMultiFileAuthState, DisconnectReason, fetchLatestBaileysVersion } from '@whiskeysockets/baileys'
import qrcodeTerminal from 'qrcode-terminal'
import QRCode from 'qrcode'
import express from 'express'
import { createRequire } from 'module'
import fs from 'fs'
import path from 'path'
import os from 'os'
import P from 'pino'
const require = createRequire(import.meta.url)
const { google } = require('googleapis')
const ExcelJS = require('exceljs')

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
const CONCEPTOS = ["CARNE","POLLO","PESCADO","VERDURAS","FRUTAS","ABARROTES","LIMPIEZA","DESECHABLES","PAN","LACTEOS","QUESOS","EMBUTIDOS","BEBIDAS","CERVEZA","VINOS","HIELO","GAS","LUZ","AGUA","MANTENIMIENTO","REFACCIONES","UNIFORMES","PAPELERIA","PUBLICIDAD","TRANSPORTE","NOMINA","RENTA","IMPUESTOS","COMISIONES","VARIOS","TORTILLAS","CONDIMENTOS","OTROS"]
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
function calcularHorasTrabajadas(hE,hS){ if(!hS) return "8"; const toSeg=(h)=>{const [hh,mm,ss]=(h||'').split(":").map(Number); return (hh||0)*3600+(mm||0)*60+(ss||0)}; let diff=toSeg(hS)-toSeg(hE); if(diff<0) diff+=24*3600; const h=Math.floor(diff/3600),m=Math.floor((diff%3600)/60); return h===8&&m===0?"8":`${h}:${String(m).padStart(2,'0')}:00` }
function calcularExtra(hE,hS, progEnt, progSal){ const trabajadas = calcularHorasTrabajadas(hE,hS); if(!progEnt||!progSal||progEnt==='LIBRE') return { trabajadas, extra: "0", extraMin: 0 }; const toMin = h => { const [hh,mm]=h.split(':').map(Number); return (hh||0)*60+(mm||0) }; let minProg = toMin(progSal) - toMin(progEnt); if(minProg < 0) minProg += 24*60; let minTrab = 0; if(trabajadas==="8") minTrab=8*60; else { const [hh,mm]=trabajadas.split(':').map(Number); minTrab=(hh||0)*60+(mm||0) }; let extraMin = minTrab - minProg; if(extraMin < 0) extraMin = 0; const eh = Math.floor(extraMin/60); const em = extraMin%60; return { trabajadas, extra: extraMin>0? `${eh}:${String(em).padStart(2,'0')}:00` : "0", extraMin } }
function parseFechaMX(s){ if(!s) return null; if(/^\d{4}-\d{2}-\d{2}$/.test(s)) return new Date(s+'T12:00:00'); const m=s.match(/(\d{1,2})\/(\d{2,4})/); if(m){ return new Date(`${m[3].length==2?'20'+m[3]:m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}T12:00:00`) } return new Date(s) }
async function getRows(range, sid){ const s=await sheetsClient(); const r=await s.spreadsheets.values.get({spreadsheetId: sid || SPREADSHEET_ID, range}); return r.data.values||[] }
function getRangoSemana(tipo){ const hoyMX = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'})); const diaSem = hoyMX.getDay(); const lunesEstaSem = new Date(hoyMX); lunesEstaSem.setDate(hoyMX.getDate() - (diaSem===0?6:diaSem-1)); let lunes, domingo; if(tipo==='pasada'){ lunes=new Date(lunesEstaSem); lunes.setDate(lunes.getDate()-7); domingo=new Date(lunes); domingo.setDate(domingo.getDate()+6); } else { lunes=lunesEstaSem; domingo=hoyMX; } lunes.setHours(0,0,0,0); domingo.setHours(23,59,59,999); return { lunes, domingo, rangoTxt: `${lunes.toLocaleDateString('es-MX')} al ${domingo.toLocaleDateString('es-MX')} (${tipo})` } }
function normaliza(s){ return (s||'').toString().normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim() }

// === CORRECCIÓN FINAL: USTEDES SON VICTOR HUGO, PROVEEDOR ES GASTROSOPHIA ===
function normalizarProveedor(nombre){
  let n = (nombre||"").toUpperCase().trim()
  // Ustedes son VICTOR HUGO POBLANO BRAVO, el proveedor que les vende es GASTROSOPHIA
  if(n.includes("GASTRO") || n.includes("VICTOR HUGO") || n.includes("POBLANO BRAVO")) return "GASTROSOPHIA"
  if(n.includes("TRES B") || n.includes("3B") || n.includes("TIENDAS TRES")) return "TIENDAS TRES B"
  if(n.includes("FLORENTINA")) return "QUESOS FLORENTINA"
  if(!n || n.length < 3) return "PROVEEDOR OCASIONAL"
  return nombre.toUpperCase()
}
function detectarSucursalCompra(texto){ const t = normaliza(texto); if(t.includes('bucareli')) return 'BUCARELI'; return 'BUCARELI' }
function scoreEmpleado(corto, completo, buscarNorm){ const c = normaliza(corto); const d = normaliza(completo); if(c === buscarNorm) return 100; if(c.startsWith(buscarNorm)) return 90; if(c.includes(buscarNorm) || d.includes(buscarNorm)) return 10; return -1 }
async function getHorarioBaseMap(){ try{ const rows = await getRows('Horario_Base!A2:K'); const map = {}; for(const f of rows){ const tel = (f[0]||'').replace(/\D/g,'').slice(-10); const nombre = (f[1]||'').toLowerCase().trim(); if(!nombre) continue; const dias = { 1:f[3], 2:f[4], 3:f[5], 4:f[6], 5:f[7], 6:f[8], 0:f[9] }; const descansos = new Set(); const horas = {}; for(const [numDia, valor] of Object.entries(dias)){ const v = (valor||'').toString().trim(); if(!v){ descansos.add(parseInt(numDia)); continue } const parsed = parseHorarioRango(v); if(!parsed) descansos.add(parseInt(numDia)); else horas[numDia]=parsed } const obj = { descansos, horas, nombreOriginal: f[1], tel, sucursal: f[2]||'' }; if(tel) map[tel]=obj; map[nombre]=obj; const primer=nombre.split(' ')[0]; if(primer &&!map[primer]) map[primer]=obj } return map }catch(e){ return {} } }
let grupoFiltroCache = new Map()
async function getFiltroPorGrupo(jid, sock){ if(GRUPOS.REPORTES.includes(jid)) return null; if(jid === GRUPO_COYOACAN_ID || jid === GRUPO_CHECADOR_COYOACAN_ID) return 'coyoacan'; if(jid === GRUPO_BUCARELI_ID || jid === GRUPO_JUAREZ_ID || jid === GRUPO_CHECADOR_BUCARELI_ID) return 'juarez'; if(grupoFiltroCache.has(jid)) return grupoFiltroCache.get(jid); try{ const meta = await sock.groupMetadata(jid); const subj = (meta.subject || '').toLowerCase(); let filtro = null; if(subj.includes('coyo')) filtro = 'coyoacan'; else if(subj.includes('bucareli') || subj.includes('juarez')) filtro = 'juarez'; grupoFiltroCache.set(jid, filtro); return filtro }catch{ return null } }
function sucursalCoincideConFiltro(sucEmpleado, filtro){ if(!filtro) return true; const s = (sucEmpleado||'').toLowerCase(); if(filtro === 'coyoacan') return s.includes('coyo') || s.includes('hotel') || s.includes('trinidad'); if(filtro === 'juarez') return s.includes('juarez') || s.includes('bucareli'); return s.includes(filtro) }

async function leerTicketConIA(bufferImagen){
  if(!process.env.OPENAI_API_KEY) return null
  try{
    const base64 = bufferImagen.toString('base64')
    const resp = await fetch('https://api.openai.com/v1/chat/completions',{
      method:'POST',
      headers:{'Authorization':`Bearer ${process.env.OPENAI_API_KEY}`,'Content-Type':'application/json'},
      body: JSON.stringify({
        model:'gpt-4o-mini',
        messages:[{ role:'user', content:[
          {type:'text', text:`Lee este ticket/factura mexicana. Devuelve SOLO JSON: {proveedor, folio, importe_total, fecha_ticket: YYYY-MM-DD, concepto_sugerido, productos:[{nombre,cantidad,precio,total}]}. Si ves VICTOR HUGO POBLANO BRAVO y GASTROSOPHIA juntos, proveedor=GASTROSOPHIA. Si no hay proveedor legible, proveedor=PROVEEDOR OCASIONAL. concepto_sugerido de: ${CONCEPTOS.join(',')}.`},
          {type:'image_url', image_url:{url:`data:image/jpeg;base64,${base64}`}}
        ]}],
        max_tokens:700
      })
    })
    const data = await resp.json()
    let txt = data.choices?.[0]?.message?.content || ""
    txt = txt.replace(/```json|```/g,'').trim()
    return JSON.parse(txt)
  }catch(e){ console.log('IA fail', e.message); return null }
}

async function registrarCompra(datos, jid, nombrePersona){
  const sClient = await sheetsClient()
  const proveedorFinal = normalizarProveedor(datos.proveedor)
  await sClient.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_COMPRAS_ID,
    range:`${SHEET_RESUMEN}!A:I`,
    valueInputOption:'USER_ENTERED',
    requestBody:{ values:[[ datos.semana, datos.fecha, datos.sucursal, proveedorFinal, datos.folio, datos.monto, datos.concepto, "", "" ]] }
  })
  try{
    if(datos.productos && Array.isArray(datos.productos) && datos.productos.length>0){
      let rows = datos.productos.map(p=> [ datos.folio, p.nombre||datos.concepto, p.cantidad||"1", p.precio||"", p.total||datos.monto, datos.sucursal, datos.fecha ])
      await sClient.spreadsheets.values.append({ spreadsheetId: SPREADSHEET_COMPRAS_ID, range:`${SHEET_INSUMOS}!A:G`, valueInputOption:'USER_ENTERED', requestBody:{ values: rows } })
    } else {
      await sClient.spreadsheets.values.append({ spreadsheetId: SPREADSHEET_COMPRAS_ID, range:`${SHEET_INSUMOS}!A:G`, valueInputOption:'USER_ENTERED', requestBody:{ values:[[ datos.folio, datos.concepto, "1", datos.monto, datos.monto, datos.sucursal, datos.fecha ]] } })
    }
  }catch(e){ console.log(e) }
  return { ok:true, msg:`✅ ${datos.folio} $${datos.monto} - ${datos.concepto} - ${proveedorFinal} guardado.` }
}
async function generarExcelCompras(filtro, jid, sock){
  const rows = await getRows(`${SHEET_RESUMEN}!A2:I`, SPREADSHEET_COMPRAS_ID)
  let filtradas = rows.filter(r=>{ if(filtro.sucursal &&!(r[2]||'').toUpperCase().includes(filtro.sucursal.toUpperCase())) return false; if(filtro.fecha && r[1]!==filtro.fecha) return false; return true })
  const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('Compras'); ws.addRow(['Semana','Fecha','Sucursal','Proveedor','#Comprobante','Importe','Concepto','Area','FormaPago']).font={bold:true}; filtradas.forEach(r=>ws.addRow(r)); const total = filtradas.reduce((a,r)=> a + (parseFloat((r[5]||'0').toString().replace(/,/g,''))||0), 0); ws.addRow([]); ws.addRow(['TOTAL','','','','',total]); ws.columns.forEach(c=>c.width=18); const fileName = `Compras_${filtro.sucursal||'TODAS'}_${fechaLaboral()}.xlsx`; const fp = path.join(os.tmpdir(), fileName); await wb.xlsx.writeFile(fp); await sock.sendMessage(jid,{document:fs.readFileSync(fp),mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',fileName, caption:`📊 Compras ${filtro.sucursal||''} - ${filtradas.length} regs - Total $${total}`}); fs.unlinkSync(fp)
}
function parseFiltroCompras(texto){ const low = normaliza(texto); let suc = null; if(low.includes('bucareli')) suc='BUCARELI'; else if(low.includes('coyo')) suc='COYOACAN'; else if(low.includes('juarez')) suc='JUAREZ'; let fecha = null; if(low.includes('hoy')) fecha = fechaLaboral(); return { sucursal:suc, fecha } }
async function asistenciaHoy(filtroSucursal, jid, sock){ const sClient = await sheetsClient(); const [baseRows, asisRows] = await Promise.all([ getRows('Horario_Base!A2:K'), sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}).then(r=>r.data.values||[]) ]); const fLab = fechaLaboral(); const ahoraMin = minutos(horaMX()); const ahoraMXDate = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'})); const diaNum = ahoraMXDate.getDay(); const filtro = filtroSucursal.toLowerCase(); const esCoyo = filtro.includes('coyo') || filtro.includes('hotel'); const esJuarez = filtro.includes('juarez')||filtro.includes('bucareli'); let llego=[], retardo=[], falta=[], futuro=[]; for(const r of baseRows){ const sucBaseLower = (r[2]||'').toLowerCase(); const sucBaseOriginal = r[2]||''; let inc=false; if(esCoyo) inc = sucBaseLower.includes('coyo')||sucBaseLower.includes('hotel')||sucBaseLower.includes('trinidad'); else if(esJuarez) inc = sucBaseLower.includes('juarez')||sucBaseLower.includes('bucareli'); else inc = sucBaseLower.includes(filtro); if(!inc) continue; const nombre = r[1]||''; const tel = (r[0]||'').replace(/\D/g,'').slice(-10); const mapa = {1:r[3],2:r[4],3:r[5],4:r[6],5:r[7],6:r[8],0:r[9]}; const v = (mapa[diaNum]||'').toString().trim(); const parsed = parseHorarioRango(v); if(!parsed) continue; const esLibre = parsed.entrada === 'LIBRE'; const horaProg = esLibre? 'LIBRE' : parsed.entrada; const registro = asisRows.find(a => a[0]?.replace(/\D/g,'').slice(-10)===tel && a[2]===fLab); if(registro && registro[3]){ const entrada = registro[3]; const sucEnt = registro[6]||''; const dif = esLibre? 0 : minutos(entrada) - minutos(horaProg); if(esLibre){ llego.push(`• ${nombre} - Entró ${entrada} en ${sucEnt} ✅`) } else if(dif > 15){ retardo.push(`• ${nombre} - [${sucBaseOriginal}] Prog ${horaProg} - Entró ${entrada} - ⏰ ${dif}m tarde - ${sucEnt}`) } else llego.push(`• ${nombre} - [${sucBaseOriginal}] Prog ${horaProg} - Entró ${entrada} ✅ - ${sucEnt}`) } else { if(esLibre){ if(ahoraMin >= 20*60) falta.push(`• ${nombre} - [${sucBaseOriginal}] - ❌ sin llegar`); continue } const dif = ahoraMin - minutos(horaProg); if(dif < 0){ futuro.push(`• ${nombre} - [${sucBaseOriginal}] - Prog ${horaProg}`) } else falta.push(`• ${nombre} - [${sucBaseOriginal}] - Prog ${horaProg} - ❌ ${dif}m sin llegar`) } } let txt = `📍 *ASISTENCIA HOY ${fLab} - ${filtroSucursal.toUpperCase()}* ${horaMX()}\n\n✅ *A TIEMPO (${llego.length}):*\n${llego.join('\n')||'-'}\n\n⏰ *RETARDOS (${retardo.length}):*\n${retardo.join('\n')||'-'}\n\n❌ *NO HAN LLEGADO (${falta.length}):*\n${falta.join('\n')||'Todos llegaron'}\n\n⏳ *PRÓXIMOS (${futuro.length}):*\n${futuro.join('\n')||'-'}`; await sock.sendMessage(jid,{text:txt}) }
async function generarExcelEmpleado(nombreBuscarRaw, jid, sock, tipo='actual'){ const sClient=await sheetsClient(); const baseMap=await getHorarioBaseMap(); const asisRes = await sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}); const filas=asisRes.data.values||[]; const buscar=nombreBuscarRaw.toLowerCase().replace(/actual|pasada|pasado|esta semana|hoy/g,'').trim(); const { lunes, domingo, rangoTxt } = getRangoSemana(tipo); const filtradas = filas.filter(f=>{ const n=(f[1]||'').toLowerCase(); if(!n.includes(buscar)) return false; const fe=parseFechaMX(f[2]); return fe&&fe>=lunes&&fe<=domingo }); const header=["Tel","Nombre","Fecha","Entrada","Estatus Entrada","Salida","Suc Entrada","Dist Entr","Suc Salida","Dist Sal","Horas Trabajadas","Horas Extra","Jornada Programada"]; const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet('Resumen'); ws.addRow([`REPORTE ${buscar.toUpperCase()} - ${tipo.toUpperCase()}`]).font={bold:true,size:14}; ws.addRow([rangoTxt]); ws.addRow([]); let dias=0,ret=0,min=0,sin=0,horasMin=0,extraMin=0; filtradas.forEach(f=>{ if(f[3]) dias++; const m=(f[4]||'').match(/(\d+)\s*min/); if(m){ret++; min+=parseInt(m[1])} if(!f[5]) sin++; if(f[3]&&f[5]){ const fe=parseFechaMX(f[2]); const prog=baseMap[(f[0]||'').replace(/\D/g,'').slice(-10)]?.horas?.[fe.getDay()]; const calc=calcularExtra(f[3],f[5],prog?.entrada||null,prog?.salida||null); let mt=calc.trabajadas==="8"?480:(()=>{const[hh,mm]=calc.trabajadas.split(':').map(Number); return hh*60+mm})(); horasMin+=mt; extraMin+=calc.extraMin } }); ws.addRow(['Nombre','Días','Retardos','Min','Sin salida','Horas Trab','Horas Extra']).font={bold:true}; ws.addRow([buscar, `${dias} días`, `Ret ${ret} (${min}m)`, sin, `${Math.floor(horasMin/60)}:${String(horasMin%60).padStart(2,'0')}h`, `${Math.floor(extraMin/60)}:${String(extraMin%60).padStart(2,'0')}h`]); ws.columns.forEach(c=>c.width=22); const ws2=wb.addWorksheet('Detalle'); ws2.addRow(header).font={bold:true}; filtradas.forEach(f=>{ ws2.addRow(f) }); ws2.columns.forEach(c=>c.width=18); const fileName=`Reporte_${buscar.replace(/\s+/g,'_')}_${tipo}.xlsx`; const fp=path.join(os.tmpdir(),fileName); await wb.xlsx.writeFile(fp); await sock.sendMessage(jid,{document:fs.readFileSync(fp),mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',fileName}); fs.unlinkSync(fp) }
async function resumenEmpleado(nombreBuscar, jidRespuesta, sock, tipo='actual'){ const sClient=await sheetsClient(); const [asisRes, baseMap]=await Promise.all([ sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}), getHorarioBaseMap() ]); const filas=asisRes.data.values||[]; let buscar=nombreBuscar.toLowerCase().replace(/actual|pasada|pasado|esta semana|hoy/g,'').trim(); const { lunes, domingo, rangoTxt } = getRangoSemana(tipo); const info = baseMap[buscar] || Object.values(baseMap).find(v=> (v.nombreOriginal||'').toLowerCase().includes(buscar)) || { descansos:new Set(), horas:{} }; let diasSem=0,retSem=0,minSem=0,detalle=[],sin=[],horasMin=0,extraTotal=0, diasNom=['Dom','Lun','Mar','Mie','Jue','Vie','Sab']; for(const f of filas){ const n=(f[1]||'').toLowerCase(); if(!n.includes(buscar)) continue; const fe=parseFechaMX(f[2]); if(!fe) continue; fe.setHours(0,0,0,0); if(fe<lunes||fe>domingo) continue; if(f[3]){ diasSem++; const ret=(f[4]||'').match(/(\d+)\s*min/); if(ret){retSem++; minSem+=parseInt(ret[1])} if(!f[5]) sin.push(f[2]); const progDia=info.horas[fe.getDay()]; if(f[3]&&f[5]&&progDia){ const calc=calcularExtra(f[3],f[5],progDia.entrada,progDia.salida); let mt=calc.trabajadas==="8"?480:(()=>{const[hh,mm]=calc.trabajadas.split(':').map(Number);return hh*60+mm})(); horasMin+=mt; extraTotal+=calc.extraMin } detalle.push(`• ${f[2]}: Ent ${f[3]} | Sal ${f[5]||'SIN'} | Trab ${f[10]||''} | Extra ${f[11]||'0'} | ${f[12]||''}`) } } let esperados=0; for(let d=new Date(lunes); d<=domingo; d.setDate(d.getDate()+1)){ if(!info.descansos.has(d.getDay())) esperados++ } const faltas=Math.max(0,esperados-diasSem); const descTxt=[...info.descansos].map(d=>diasNom[d]).join(', ')||'ninguno'; const txt=`📊 *${buscar.toUpperCase()}* - Desc: ${descTxt}\n${rangoTxt}\n\n*SEMANA ${tipo.toUpperCase()}*\n- Trabajados: ${diasSem}/${esperados} - Faltas: ${faltas}\n- Retardos: ${retSem} (${minSem} min)\n- Sin salida: ${sin.length}\n- Horas Trab: ${Math.floor(horasMin/60)}:${String(horasMin%60).padStart(2,'0')}h\n- Horas Extra: ${Math.floor(extraTotal/60)}:${String(extraTotal%60).padStart(2,'0')}h\n\nDetalle:\n${detalle.join('\n')||'Sin registros'}`; await sock.sendMessage(jidRespuesta,{text:txt}); if(diasSem>0) await generarExcelEmpleado(buscar, jidRespuesta, sock, tipo) }
async function reporteSucursal(filtroSucursal, jid, sock, tipo='pasada'){ const [asisRes, empRows] = await Promise.all([ (await sheetsClient()).spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}), getRows('Empleados!A:K') ]); const filas=asisRes.data.values||[]; const {lunes,domingo,rangoTxt}=getRangoSemana(tipo); const filtro=filtroSucursal.toLowerCase(); const esCoyo=filtro.includes('coyo')||filtro.includes('hotel'); const esJuarez=filtro.includes('juarez')||filtro.includes('bucareli'); const baseMap = await getHorarioBaseMap(); const telToNombre = {}; empRows.forEach(r=>{ const tel=(r[0]||'').replace(/\D/g,'').slice(-10); if(tel) telToNombre[tel] = (r[3]||r[1]||'').trim() }); let datos={}; for(const f of filas){ const fe=parseFechaMX(f[2]); if(!fe||fe<lunes||fe>domingo) continue; const suc=((f[6]||'')+' '+(f[8]||'')).toLowerCase(); let inc=false; if(esCoyo) inc=suc.includes('coyo')||suc.includes('hotel'); else if(esJuarez) inc=suc.includes('juarez')||suc.includes('bucareli'); else inc=suc.includes(filtro); if(!inc) continue; const tel=(f[0]||'').replace(/\D/g,'').slice(-10); const nombreOficial = telToNombre[tel] || (f[1]||'').trim() || 'Desconocido'; const key = tel || normaliza(nombreOficial); if(!datos[key]) datos[key]={nombre:nombreOficial, dias:0,ret:0,min:0,sin:0,horasMin:0,extraMin:0}; if(f[3]){ datos[key].dias++; const m=(f[4]||'').match(/(\d+)\s*min/); if(m){datos[key].ret++; datos[key].min+=parseInt(m[1])} if(!f[5]) datos[key].sin++; if(f[3]&&f[5]){ const feDia=parseFechaMX(f[2]); const info=baseMap[tel]||null; const prog=info?.horas?.[feDia?feDia.getDay():1]||null; const calc=calcularExtra(f[3],f[5],prog?.entrada||null,prog?.salida||null); let mt=calc.trabajadas==="8"?480:(()=>{const[hh,mm]=calc.trabajadas.split(':').map(Number);return hh*60+mm})(); datos[key].horasMin+=mt; datos[key].extraMin+=calc.extraMin } } } let txt=`📊 *REPORTE ${filtroSucursal.toUpperCase()}* - ${tipo}\n${rangoTxt}\n`; if(!Object.keys(datos).length) txt+='Sin registros\n'; else for(const k in datos){ const d=datos[k]; txt+=`*${d.nombre}*: ${d.dias} días | Ret ${d.ret} (${d.min}m) | Trab: ${Math.floor(d.horasMin/60)}:${String(d.horasMin%60).padStart(2,'0')}h | Extra: ${Math.floor(d.extraMin/60)}:${String(d.extraMin%60).padStart(2,'0')}h | Sin salida: ${d.sin}\n\n` } await sock.sendMessage(jid,{text:txt}) }
async function generarExcelSemanaYEnviar(filtroSucursal, jid, sock, tipo='pasada'){ const sClient=await sheetsClient(); const baseMap=await getHorarioBaseMap(); const [asisRes, empRows] = await Promise.all([ sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}), getRows('Empleados!A:K') ]); const filas=asisRes.data.values||[]; const {lunes,domingo,rangoTxt}=getRangoSemana(tipo); const esCoyo=filtroSucursal.toLowerCase().includes('coyo'); const esJuarez=filtroSucursal.toLowerCase().includes('juarez')||filtroSucursal.toLowerCase().includes('bucareli'); const telToNombre = {}; empRows.forEach(r=>{ const tel=(r[0]||'').replace(/\D/g,'').slice(-10); if(tel) telToNombre[tel] = (r[3]||r[1]||'').trim() }); const filtradas=filas.filter(f=>{ const fe=parseFechaMX(f[2]); if(!fe||fe<lunes||fe>domingo) return false; const suc=((f[6]||'')+' '+(f[8]||'')).toLowerCase(); if(esCoyo) return suc.includes('coyo')||suc.includes('hotel'); if(esJuarez) return suc.includes('juarez')||suc.includes('bucareli'); return suc.includes(filtroSucursal.toLowerCase()) }); let datos={}; for(const f of filtradas){ const tel=(f[0]||'').replace(/\D/g,'').slice(-10); const nombreOficial = telToNombre[tel] || (f[1]||'Desconocido').trim(); const key = tel || normaliza(nombreOficial); if(!datos[key]) datos[key]={nombre:nombreOficial, dias:0,ret:0,min:0,sin:0,horasMin:0,extraMin:0, tel}; const fe=parseFechaMX(f[2]); if(fe){ const prog=baseMap[tel]?.horas?.[fe.getDay()]; if(f[3]&&f[5]){ const calc=calcularExtra(f[3],f[5],prog?.entrada,prog?.salida); let mt=calc.trabajadas==="8"?480:(()=>{const[hh,mm]=calc.trabajadas.split(':').map(Number);return hh*60+mm})(); datos[key].horasMin+=mt; datos[key].extraMin+=calc.extraMin } } datos[key].dias++; const m=(f[4]||'').match(/(\d+)\s*min/); if(m){datos[key].ret++; datos[key].min+=parseInt(m[1])} if(!f[5]) datos[key].sin++ }; const header=["Tel","Nombre","Fecha","Entrada","Estatus Entrada","Salida","Suc Entrada","Dist Entr","Suc Salida","Dist Sal","Horas Trabajadas","Horas Extra","Jornada"]; const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet('Resumen'); ws.addRow([`REPORTE ${filtroSucursal.toUpperCase()} - ${tipo}`]).font={bold:true,size:14}; ws.addRow([rangoTxt]); ws.addRow([]); ws.addRow(['Nombre','Tel','Días','Retardos','Min','Sin salida','Horas Trab','Horas Extra']).font={bold:true}; for(const k of Object.keys(datos)){ const d=datos[k]; ws.addRow([d.nombre, d.tel, `${d.dias} días`, `Ret ${d.ret} (${d.min}m)`, d.sin, `${Math.floor(d.horasMin/60)}:${String(d.horasMin%60).padStart(2,'0')}h`, `${Math.floor(d.extraMin/60)}:${String(d.extraMin%60).padStart(2,'0')}h`]) } ws.columns.forEach(c=>c.width=22); const ws2=wb.addWorksheet('Detalle'); ws2.addRow(header).font={bold:true}; filtradas.forEach(f=>{ ws2.addRow(f) }); ws2.columns.forEach(c=>c.width=18); const fileName=`Reporte_${filtroSucursal}_${tipo}_${lunes.toISOString().split('T')[0]}.xlsx`; const fp=path.join(os.tmpdir(),fileName); await wb.xlsx.writeFile(fp); await sock.sendMessage(jid,{document:fs.readFileSync(fp),mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',fileName}); fs.unlinkSync(fp) }

const app=express(); let lastQR=null; app.get('/',(req,res)=>res.send('Bot OK - /qr')); app.get('/qr',async(req,res)=>{ if(!lastQR) return res.send('No QR'); const dataUrl=await QRCode.toDataURL(lastQR); res.send(`<img src="${dataUrl}" style="width:350px">`) }); app.listen(process.env.PORT||3000)

async function start(){
  const { state, saveCreds } = await useMultiFileAuthState('/app/auth')
  const { version } = await fetchLatestBaileysVersion()
  const sock=makeWASocket({ version, auth:state, logger:P({level:'fatal'}), printQRInTerminal:false, markOnlineOnConnect:false, syncFullHistory:false, shouldSyncHistoryMessage:()=>false, browser:['Trinidad Bot','Chrome','121.0.0'], getMessage: async () => undefined })
  sock.ev.on('creds.update', saveCreds)
  sock.ev.on('connection.update', async ({connection, lastDisconnect, qr}) => {
    if (qr) { lastQR = qr; qrcodeTerminal.generate(qr,{small:false}) }
    if (connection === 'close') { const code=lastDisconnect?.error?.output?.statusCode; if(code!==DisconnectReason.loggedOut) setTimeout(()=>start(),5000) }
    if (connection === 'open') { console.log('✅ CONECTADO - CHECADOR + COMPRAS AUTO GASTROSOPHIA PROVEEDOR'); }
  })
  sock.ev.on('messages.upsert', async ({messages})=>{
    try{
      const m=messages[0]; if(!m||m.key.fromMe) return; const jid=m.key.remoteJid; if(!jid.endsWith('@g.us')) return
      const rawLid=m.key.participant||''; const realPn=m.key.participantPn||m.key.participantAlt||''; let pnFromStore=''; try{ pnFromStore=await sock.signalRepository?.lidMapping?.getPNForLID(rawLid)||'' }catch{}; const rawId=realPn||pnFromStore||rawLid||jid
      let tel=(rawId||'').toString().replace(/\D/g,''); let tel10=tel.slice(-10)
      const texto=m.message?.conversation||m.message?.extendedTextMessage?.text||m.message?.imageMessage?.caption||m.message?.documentMessage?.caption||m.message?.extendedTextMessage?.contextInfo?.quotedMessage?.imageMessage?.caption||'';
      const loc=m.message?.locationMessage || m.message?.liveLocationMessage
      const esImagen =!!(m.message?.imageMessage)
      if(texto.trim().toLowerCase()==='id'){ await sock.sendMessage(jid,{text:`ID: ${jid}\nTipo: ${getTipoGrupo(jid)}`}); return }
      const tipoGrupo = getTipoGrupo(jid); const filtroGrupo = await getFiltroPorGrupo(jid, sock)

      if(tipoGrupo==='REPORTES'){
        if(/^compras/i.test(texto)){ const f = parseFiltroCompras(texto); await generarExcelCompras(f, jid, sock); return }
        if(esImagen){
          console.log('📸 FOTO COMPRA - LEYENDO CON IA...')
          try{
            const buffer = await sock.downloadMediaMessage(m, 'buffer', {}, { logger: P({level:'fatal'}) })
            let datosIA = await leerTicketConIA(buffer)
            if(!datosIA){
              await sock.sendMessage(jid,{text:`⚠️ Falta OPENAI_API_KEY, no pude leer auto. Escribe manual: BUCARELI $394 ABARROTES`})
              return
            }
            let datosFinal = {
              folio: datosIA.folio || `AC-${Date.now().toString().slice(-6)}`,
              concepto: datosIA.concepto_sugerido || "OTROS",
              proveedor: normalizarProveedor(datosIA.proveedor),
              monto: (datosIA.importe_total||"0").toString().replace(/,/g,''),
              fecha: datosIA.fecha_ticket || fechaLaboral(),
              sucursal: detectarSucursalCompra(texto) || "BUCARELI",
              semana: getSemanaActual(),
              productos: datosIA.productos || null
            }
            const res = await registrarCompra(datosFinal, jid, m.pushName||tel10)
            await sock.sendMessage(jid,{text: `${res.msg}\n📄 Prov: ${datosFinal.proveedor} | Folio: ${datosFinal.folio}`})
            return
          }catch(e){ console.log(e); await sock.sendMessage(jid,{text:'Error IA: '+e.message}) ; return }
        }
      }

      if(!tipoGrupo) return
      if(tipoGrupo === 'CHECADORES'){ if(!loc){ if(texto && PAQUETES.REPORTES_PARA_GERENTES.test(texto)){ await sock.sendMessage(jid,{text:`⚠️ Este grupo es solo para checar entradas y salidas.`}) } return } }
      if(tipoGrupo === 'GERENTES'){ if(loc) return; if(texto &&!PAQUETES.REPORTES_PARA_GERENTES.test(texto)) return }
      if(PAQUETES.REPORTES_PARA_GERENTES.test(texto)){
        if(texto.toLowerCase().startsWith('asistencia hoy')){ let suc=texto.toLowerCase().replace('asistencia hoy','').trim(); if(!suc) suc=filtroGrupo||'coyoacan'; if(filtroGrupo &&!sucursalCoincideConFiltro(suc, filtroGrupo)){ await sock.sendMessage(jid,{text:`⚠️ No hay registros de *${suc.toUpperCase()}* en esta unidad.`}); return } await asistenciaHoy(suc,jid,sock); return }
        if(texto.toLowerCase().startsWith('resumen ')){ let txtLow=texto.toLowerCase(); let tipo='actual'; if(txtLow.includes('pasada')||txtLow.includes('pasado')) tipo='pasada'; let limpio=texto.slice(8).toLowerCase().trim().replace(/pasada|pasado|actual|esta semana|hoy/g,'').trim(); if(limpio.includes('bucareli')||limpio.includes('juarez')||limpio.includes('coyo')||limpio.includes('hotel')){ let suc='coyoacan'; if(limpio.includes('juarez')||limpio.includes('bucareli')) suc='juarez'; if(filtroGrupo &&!sucursalCoincideConFiltro(suc, filtroGrupo)){ await sock.sendMessage(jid,{text:`⚠️ No hay registros de *${suc.toUpperCase()}* en esta unidad.`}); return } await reporteSucursal(suc,jid,sock,tipo); await generarExcelSemanaYEnviar(suc,jid,sock,tipo); return } if(!limpio){ await sock.sendMessage(jid,{text:'Escribe: resumen [nombre] pasada o actual'}); return } await resumenEmpleado(limpio,jid,sock,tipo); return }
        if(/^(reporte|checador|reporte x unidad|reportes x unidad)/i.test(texto)){ const txtLow=texto.toLowerCase(); const tipo=txtLow.includes('pasada')||txtLow.includes('pasado')?'pasada':txtLow.includes('actual')||txtLow.includes('esta')||txtLow.includes('hoy')?'actual':'pasada'; let suc='coyoacan'; if(txtLow.includes('juarez')||txtLow.includes('bucareli')) suc='juarez'; else if(txtLow.includes('hotel')) suc='hotel'; else if(txtLow.includes('coyo')) suc='coyoacan'; if(txtLow.includes('x unidad') || txtLow.includes('por unidad')){ if(!filtroGrupo){ await reporteSucursal('coyoacan',jid,sock,tipo); await generarExcelSemanaYEnviar('coyoacan',jid,sock,tipo); await reporteSucursal('juarez',jid,sock,tipo); await generarExcelSemanaYEnviar('juarez',jid,sock,tipo); } else { await reporteSucursal(filtroGrupo,jid,sock,tipo); await generarExcelSemanaYEnviar(filtroGrupo,jid,sock,tipo); } return } if(filtroGrupo &&!sucursalCoincideConFiltro(suc, filtroGrupo)){ await sock.sendMessage(jid,{text:`⚠️ No hay registros de *${suc.toUpperCase()}* en esta unidad.`}); return } await reporteSucursal(suc,jid,sock,tipo); await generarExcelSemanaYEnviar(suc,jid,sock,tipo); return }
        if(/^(numero|número|num|tel|telefono|teléfono|info|ficha|datos|dato)\s*(de)?/i.test(texto)){ let buscar = texto.toLowerCase().replace(/^(numero|número|num|tel|telefono|teléfono|info|ficha|datos|dato)\s*(de)?\s*/i,'').trim(); if(!buscar){ await sock.sendMessage(jid,{text:'Escribe: info fer'}); return } const buscarNorm = normaliza(buscar); const empRows = await getRows('Empleados!A:K'); let candidatos = []; for(const r of empRows.slice(1)){ const corto = r[1]||''; const completo = r[3]||''; const suc = r[2]||''; if(!sucursalCoincideConFiltro(suc, filtroGrupo)) continue; const s = scoreEmpleado(corto, completo, buscarNorm); if(s >= 0) candidatos.push({ r, s }) } candidatos.sort((a,b)=>b.s-a.s); if(!candidatos.length){ await sock.sendMessage(jid,{text:`No encontré a "${buscar}"`}); return } const r = candidatos[0].r; const ficha = `📋 *${r[3] || r[1]}*\n👤 Corto: ${r[1]}\n📍 ${r[2]||'-'}\n📱 ${r[0]||'-'}`; await sock.sendMessage(jid,{text:ficha}); return }
      }
      if(tipoGrupo === 'GERENTES') return
      if(!loc) return
      const lat=loc.degreesLatitude,lng=loc.degreesLongitude; let cercana=null,dMin=Infinity; for(const s of SUCURSALES){ const d=distM(lat,lng,s.lat,s.lng); if(d<dMin){dMin=d; cercana=s} }
      const empRowsFull = await getRows('Empleados!A:K'); const getLid = (row)=> (row.find(x=>String(x).includes('@lid'))||'').trim(); let emp=null; let empRowIndex=-1; if(tel10.length>=10){ const idx = empRowsFull.findIndex((r,i)=> i>0 && r[0] && r[0].replace(/\D/g,'').slice(-10)===tel10); if(idx>-1){ emp=empRowsFull[idx]; empRowIndex=idx+1 } } if(!emp && rawLid.includes('@lid')){ const idx = empRowsFull.findIndex((r,i)=> i>0 && getLid(r)===rawLid); if(idx>-1){ emp=empRowsFull[idx]; empRowIndex=idx+1; tel10=(emp[0]||'').replace(/\D/g,'').slice(-10) } }
      const nombreCompleto = emp? (emp[3] || emp[1]) : (m.pushName||tel10||'Desconocido'); const telFinal=emp?(emp[0]||'').replace(/\D/g,''):tel; const tel10Final=telFinal.slice(-10)||tel10; const fLab=fechaLaboral(); const asisRows=await getRows('Asistencia!A:M'); const idx=asisRows.findIndex((r,i)=>i>0&&r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel10Final&&r[2]===fLab); const hoy=idx>-1?asisRows[idx]:null; const sClient=await sheetsClient(); const baseMap=await getHorarioBaseMap(); const baseInfo=baseMap[tel10Final]||baseMap[normaliza(nombreCompleto).split(' ')[0]]||null; let estatus='A TIEMPO'; let horaProgObj=null; if(baseInfo){ const fecha=new Date(fLab+'T12:00:00'); const diaNum=fecha.getDay(); if(baseInfo.descansos.has(diaNum)){ estatus='DESCANSO' } else if(baseInfo.horas[diaNum]){ horaProgObj=baseInfo.horas[diaNum]; if(horaProgObj.entrada!=='LIBRE'){ const hp=horaProgObj.entrada; const dif=minutos(horaMX())-minutos(hp); if(dif>15) estatus=`RETARDO ${dif}min (Prog ${hp})`; else estatus=`A TIEMPO Prog ${hp}` } } }
      if(!hoy||!hoy[3]){ if(dMin>cercana.rEnt){ await sock.sendMessage(jid,{text:`Debes estar a max ${cercana.rEnt}m de ${cercana.nombre}`},{quoted:m}); return } const h=horaMX(); const jornadaTxt = horaProgObj? `${horaProgObj.entrada}${horaProgObj.salida?` - ${horaProgObj.salida}`:''}` : "8h"; const row = [tel10Final,nombreCompleto,fLab,h,estatus,'',cercana.nombre,Math.round(dMin).toString(),'','',calcularHorasTrabajadas(h,""),"0",jornadaTxt]; if(idx===-1) await sClient.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Asistencia!A:M',valueInputOption:'USER_ENTERED',requestBody:{values:[row]}}); else await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!A${idx+1}:M${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[row]}}); await sock.sendMessage(jid,{text:`✅ ${estatus} - ${nombreCompleto} en ${cercana.nombre}`}) }else{ if(hoy[5]){ await sock.sendMessage(jid,{text:`Salida ya registrada`}); return } if(dMin>cercana.rSal){ await sock.sendMessage(jid,{text:`No puedes checar salida a ${Math.round(dMin)}m`},{quoted:m}); return } const h=horaMX(); const jornadaTxt = hoy[12]||"8h"; const { trabajadas, extra }=calcularExtra(hoy[3],h,horaProgObj?.entrada||null,horaProgObj?.salida||null); await sClient.spreadsheets.values.update({ spreadsheetId:SPREADSHEET_ID, range:`Asistencia!F${idx+1}:M${idx+1}`, valueInputOption:'USER_ENTERED', requestBody:{values:[[h,hoy[6]||'',hoy[7]||'',cercana.nombre,Math.round(dMin).toString(),trabajadas,extra,jornadaTxt]]} }); await sock.sendMessage(jid,{text:`✅ Salida - ${nombreCompleto}`}) }
    }catch(e){ console.error(e) }
  })
}
start()

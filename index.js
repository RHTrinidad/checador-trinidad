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
    { id:"HOTEL", nombre:"Servicio Hotel", lat:19.3517918, lng:-99.1681579, rEnt:150, rSal:150 }
  ]
}

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets']
})
async function sheetsClient(){ const c=await auth.getClient(); return google.sheets({version:'v4',auth:c}) }
function distM(a,b,c,d){ const R=6371000, toRad=x=>x*Math.PI/180; const dLa=toRad(c-a), dLo=toRad(d-b); const q=Math.sin(dLa/2)**2+Math.cos(toRad(a))*Math.cos(toRad(c))*Math.sin(dLo/2)**2; return R*2*Math.atan2(Math.sqrt(q),Math.sqrt(1-q)) }
function fechaLaboral(d=new Date()){ const x=new Date(d); if(x.getHours()<4) x.setDate(x.getDate()-1); return x.toISOString().split('T')[0] }
function horaMX(d=new Date()){ return d.toLocaleTimeString('es-MX',{hour12:false,timeZone:'America/Mexico_City'}) }
function minutos(h){ const mm = h?.toString().match(/(\d{1,2}):(\d{2})/); return mm? parseInt(mm[1])*60+parseInt(mm[2]) : 0 }
function calcularHorasTrabajadas(horaEntrada, horaSalida){
  if(!horaSalida) return "8";
  const toSeg = (h) => { const [hh, mm, ss] = (h||'').split(":").map(Number); return (hh||0)*3600 + (mm||0)*60 + (ss||0); };
  let diff = toSeg(horaSalida) - toSeg(horaEntrada);
  if(diff < 0) diff += 24*3600;
  const h = Math.floor(diff/3600);
  const m = Math.floor((diff%3600)/60);
  if(h===8 && m===0) return "8";
  return `${h}:${String(m).padStart(2,'0')}:00`.replace(/^0+/,'') || "8";
}
function parseFechaMX(s){
  if(!s) return null
  if(/^\d{4}-\d{2}-\d{2}$/.test(s)) return new Date(s+'T12:00:00')
  const m=s.match(/(\d{1,2})\/(\d{2,4})/)
  if(m){ return new Date(`${m[3].length==2?'20'+m[3]:m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}T12:00:00`) }
  return new Date(s)
}
async function getRows(range){ const s=await sheetsClient(); const r=await s.spreadsheets.values.get({spreadsheetId:SPREADSHEET_ID,range}); return r.data.values||[] }
async function getDescansosMap(){
  try{
    const rows=await getRows('Horario_Base!A2:J')
    const map={}
    for(const f of rows){
      const nombre=(f[1]||'').toLowerCase().trim(); if(!nombre) continue
      const descansos=new Set()
      const dias=[f[3],f[4],f[5],f[6],f[7],f[8],f[9]]; const numDia=[1,2,3,4,5,6,0]
      dias.forEach((val,i)=>{ if((val||'').toString().toLowerCase().includes('descanso')) descansos.add(numDia[i]) })
      map[nombre]=descansos
    }
    return map
  }catch{ return {} }
}
async function resumenEmpleado(nombreBuscar, jidRespuesta, sock){
  const sClient=await sheetsClient()
  const [asisRes, descansosMap]=await Promise.all([
    sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:K'}),
    getDescansosMap()
  ])
  const filas=asisRes.data.values||[]; const buscar=nombreBuscar.toLowerCase().trim()
  const ahoraMX=new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'})); ahoraMX.setHours(0,0,0,0)
  const hace30=new Date(ahoraMX); hace30.setDate(hace30.getDate()-30)
  const hace14=new Date(ahoraMX); hace14.setDate(hace14.getDate()-14)
  const diaSem=ahoraMX.getDay(); const lunes=new Date(ahoraMX); lunes.setDate(ahoraMX.getDate() - (diaSem===0?6:diaSem-1)); const domingo=new Date(lunes); domingo.setDate(lunes.getDate()+6)
  let descEmpleado=new Set(); for(const k in descansosMap){ if(k.includes(buscar)||buscar.includes(k)){ descEmpleado=descansosMap[k]; break } }
  let dias30=0,diasSem=0,retardosSem=0,minSem=0,retardos14=0,min14=0,detalleSem=[],sinSalidaSem=[]
  for(const f of filas){
    const nombre=(f[1]||'').toString().toLowerCase(); if(!nombre.includes(buscar)) continue
    const fecha=parseFechaMX(f[2]); if(!fecha) continue; fecha.setHours(0,0,0,0)
    const tieneEntrada=!!f[3]; const tieneSalida=!!f[5]; const retMatch=(f[4]||'').toString().match(/(\d+)\s*min/); const ret=retMatch?parseInt(retMatch[1]):0
    if(fecha>=hace30 && tieneEntrada) dias30++
    if(fecha>=hace14 && ret>0){ retardos14++; min14+=ret }
    if(fecha>=lunes && fecha<=domingo && tieneEntrada){
      diasSem++; const flagSalida=tieneSalida? `Sal:${f[5]} en ${f[8]||''} (${f[10]||''})` : `⚠️ ${f[10]||'SIN SALIDA (8h)'} en ${f[6]||''}`; detalleSem.push(`• ${f[2]}: Ent ${f[3]} en ${f[6]||''} | ${flagSalida} ${ret>0?`RET ${ret}m`:''}`); if(!tieneSalida) sinSalidaSem.push(f[2]); if(ret>0){ retardosSem++; minSem+=ret }
    }
  }
  let esperados30=0; for(let d=new Date(hace30); d<=ahoraMX; d.setDate(d.getDate()+1)){ if(!descEmpleado.has(d.getDay())) esperados30++ }
  let esperadosSem=0; for(let d=new Date(lunes); d<=ahoraMX; d.setDate(d.getDate()+1)){ if(!descEmpleado.has(d.getDay())) esperadosSem++ }
  const faltas30=Math.max(0,esperados30-dias30); const faltasSem=Math.max(0,esperadosSem-diasSem)
  const diasNom=['Dom','Lun','Mar','Mie','Jue','Vie','Sab']; const descTxt=[...descEmpleado].map(d=>diasNom[d]).join(', ')||'ninguno'
  const txt=`📊 *${nombreBuscar.toUpperCase()}* - Descansos: ${descTxt}\nSemana ${lunes.toLocaleDateString('es-MX')} al ${domingo.toLocaleDateString('es-MX')}\n\n*SEMANA*\n- Trabajados: ${diasSem}/${esperadosSem} - Faltas: ${faltasSem}\n- Retardos: ${retardosSem} (${minSem} min)\n\n*30 DÍAS*\n- Esperados: ${esperados30} - Trabajados: ${dias30} - FALTAS: ${faltas30}\n\n*RETARDOS 14D:* ${retardos14} (${min14} min)\n*SIN SALIDA:* ${sinSalidaSem.length?sinSalidaSem.join(', '):'Ninguno ✅'}\n\nDetalle:\n${detalleSem.join('\n')||'Sin registros'}`
  await sock.sendMessage(jidRespuesta,{text:txt})
}
async function reporteSucursal(filtroSucursal, jidRespuesta, sock){
  const sClient = await sheetsClient()
  const asisRes = await sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:K'})
  const filas = asisRes.data.values||[]
  const hoyMX = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'}))
  const diaSem = hoyMX.getDay()
  const lunesEstaSem = new Date(hoyMX); lunesEstaSem.setDate(hoyMX.getDate() - (diaSem===0?6:diaSem-1))
  const lunesPasado = new Date(lunesEstaSem); lunesPasado.setDate(lunesEstaSem.getDate()-7)
  const domingoPasado = new Date(lunesPasado); domingoPasado.setDate(lunesPasado.getDate()+6)
  lunesPasado.setHours(0,0,0,0); domingoPasado.setHours(23,59,59,999)
  const filtro = filtroSucursal.toLowerCase()
  const esCoyo = filtro.includes('coyo')
  const esJuarez = filtro.includes('juarez') || filtro.includes('bucareli')
  let datos = {}
  for(const f of filas){
    const fecha = parseFechaMX(f[2]); if(!fecha) continue
    if(fecha < lunesPasado || fecha > domingoPasado) continue
    const suc = ((f[6]||'') + ' ' + (f[8]||'')).toLowerCase()
    let incluir = false
    if(esCoyo) incluir = suc.includes('coyo') || suc.includes('hotel')
    else if(esJuarez) incluir = suc.includes('juarez') || suc.includes('bucareli')
    else incluir = suc.includes(filtro)
    if(!incluir) continue
    const nombre = f[1]||'Desconocido'
    if(!datos[nombre]) datos[nombre] = { dias:0, ret:0, min:0, sinSalida:0, horas:[] }
    datos[nombre].dias++
    const retMatch = (f[4]||'').match(/(\d+)\s*min/)
    if(retMatch){ datos[nombre].ret++; datos[nombre].min += parseInt(retMatch[1]) }
    if(!f[5]) datos[nombre].sinSalida++
    if(f[10]) datos[nombre].horas.push(f[10])
  }
  const rango = `${lunesPasado.toLocaleDateString('es-MX')} al ${domingoPasado.toLocaleDateString('es-MX')}`
  let txt = `📊 *REPORTE ${filtroSucursal.toUpperCase()}* - Semana pasada\n${rango}\n${esCoyo?'Incluye: Coyoacán + Servicio Hotel\n':''}\n`
  if(Object.keys(datos).length===0) txt += 'Sin registros esa semana'
  else {
    for(const nombre in datos){
      const d = datos[nombre]
      txt += `\n*${nombre}*: ${d.dias} días | Ret ${d.ret} (${d.min}m) | Sin salida: ${d.sinSalida} | Horas: ${d.horas.join(', ')||'8, 8'}\n`
    }
  }
  await sock.sendMessage(jidRespuesta,{text:txt})
  return { lunesPasado, domingoPasado, datos }
}
async function generarExcelSemanaYEnviar(filtroSucursal, jid, sock){
  const sClient = await sheetsClient()
  const asisRes = await sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:K'})
  const filas = asisRes.data.values||[]
  const hoyMX = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'}))
  const diaSem = hoyMX.getDay()
  const lunesEstaSem = new Date(hoyMX); lunesEstaSem.setDate(hoyMX.getDate() - (diaSem===0?6:diaSem-1))
  const lunesPasado = new Date(lunesEstaSem); lunesPasado.setDate(lunesEstaSem.getDate()-7)
  const domingoPasado = new Date(lunesPasado); domingoPasado.setDate(lunesPasado.getDate()+6)
  lunesPasado.setHours(0,0,0,0); domingoPasado.setHours(23,59,59,999)
  const esCoyo = filtroSucursal.toLowerCase().includes('coyo')
  const esJuarez = filtroSucursal.toLowerCase().includes('juarez') || filtroSucursal.toLowerCase().includes('bucareli')
  const filtradas = filas.filter(f=>{
    const fecha = parseFechaMX(f[2]); if(!fecha) return false
    if(fecha < lunesPasado || fecha > domingoPasado) return false
    const suc = ((f[6]||'') + ' ' + (f[8]||'')).toLowerCase()
    if(esCoyo) return suc.includes('coyo') || suc.includes('hotel')
    if(esJuarez) return suc.includes('juarez') || suc.includes('bucareli')
    return suc.includes(filtroSucursal.toLowerCase())
  })
  let datos = {}
  for(const f of filtradas){
    const nombre = f[1]||'Desconocido'
    if(!datos[nombre]) datos[nombre] = { dias:0, ret:0, min:0, sinSalida:0, horas:[] }
    datos[nombre].dias++
    const retMatch = (f[4]||'').match(/(\d+)\s*min/)
    if(retMatch){ datos[nombre].ret++; datos[nombre].min += parseInt(retMatch[1]) }
    if(!f[5]) datos[nombre].sinSalida++
    if(f[10]) datos[nombre].horas.push(f[10])
  }
  const headerDetalle = ["Tel","Nombre","Fecha","Entrada","Estatus Entrada","Salida","Suc Entrada","Dist Entr","Suc Salida","Dist Sal","Horas Trabajadas"]
  const rango = `${lunesPasado.toLocaleDateString('es-MX')} al ${domingoPasado.toLocaleDateString('es-MX')}`

  const workbook = new ExcelJS.Workbook()
  const wsResumen = workbook.addWorksheet('Resumen')
  wsResumen.addRow([`REPORTE ${filtroSucursal.toUpperCase()} - Semana pasada`]).font={bold:true, size:14}
  wsResumen.addRow([rango])
  if(esCoyo) wsResumen.addRow(['Incluye: Coyoacán + Servicio Hotel'])
  wsResumen.addRow([])
  wsResumen.addRow(['Nombre','Días','Retardos','Min','Sin salida','Horas']).font={bold:true}
  for(const [nombre,d] of Object.entries(datos)){
    wsResumen.addRow([nombre, `${d.dias} días`, `Ret ${d.ret} (${d.min}m)`, d.sinSalida, d.horas.join(', ') || '8, 8'])
  }
  wsResumen.columns.forEach(c=> c.width=22)

  const wsDetalle = workbook.addWorksheet('Detalle')
  wsDetalle.addRow(headerDetalle).font={bold:true}
  filtradas.forEach(f=> wsDetalle.addRow(f))
  wsDetalle.columns.forEach(c=> c.width=18)
  wsDetalle.getRow(1).font={bold:true}

  try{
    const sheetData = [
      [`REPORTE ${filtroSucursal.toUpperCase()}`, rango, esCoyo?"Coyoacán + Servicio Hotel":""],
      [],
      ['Nombre','Días','Ret','Min','Sin salida','Horas'],
     ...Object.entries(datos).map(([n,d])=>[n,d.dias,d.ret,d.min,d.sinSalida,d.horas.join(', ')]),
      [],
      headerDetalle,
     ...filtradas
    ]
    await sClient.spreadsheets.values.clear({spreadsheetId: SPREADSHEET_ID, range:'Reporte_Semana!A1:K1000'}).catch(async()=>{ await sClient.spreadsheets.batchUpdate({spreadsheetId: SPREADSHEET_ID, requestBody:{requests:[{addSheet:{properties:{title:'Reporte_Semana'}}}]} }) })
    await sClient.spreadsheets.values.update({spreadsheetId: SPREADSHEET_ID, range:'Reporte_Semana!A1', valueInputOption:'USER_ENTERED', requestBody:{values:sheetData}})
  }catch(e){ console.log('Sheet error', e.message) }

  const fileName = `Reporte_${filtroSucursal}_${lunesPasado.toISOString().split('T')[0]}.xlsx`
  const filePath = path.join(os.tmpdir(), fileName)
  await workbook.xlsx.writeFile(filePath)
  await sock.sendMessage(jid,{ document: fs.readFileSync(filePath), mimetype: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName: fileName })
  fs.unlinkSync(filePath)
  return { total: filtradas.length, rango }
}
async function enviarReportesAutomaticos(sock){
  if(!GRUPO_REPORTES_ID) return
  const jid = GRUPO_REPORTES_ID
  try{
    console.log('⏰ Enviando reportes automáticos...')
    await reporteSucursal('coyoacan', jid, sock)
    await generarExcelSemanaYEnviar('coyoacan', jid, sock)
    await new Promise(r=>setTimeout(r,2000))
    await reporteSucursal('juarez', jid, sock)
    await generarExcelSemanaYEnviar('juarez', jid, sock)
    await sock.sendMessage(jid,{text:`✅ Reportes automáticos enviados - ${new Date().toLocaleString('es-MX',{timeZone:'America/Mexico_City'})}`})
  }catch(e){ console.error('Error reporte auto:', e) }
}

const app = express(); let lastQR=null; let globalSock=null
app.get('/', (req,res)=> res.send('Bot OK - /qr'))
app.get('/qr', async (req,res)=>{ if(!lastQR) return res.send('No QR'); const dataUrl=await QRCode.toDataURL(lastQR); res.send(`<img src="${dataUrl}" style="width:350px"><p>Escanea</p>`) })
app.listen(process.env.PORT||3000)

async function start(){
  const { state, saveCreds } = await useMultiFileAuthState('/app/auth')
  const sock=makeWASocket({ auth:state, printQRInTerminal:false, markOnlineOnConnect:false })
  globalSock=sock
  sock.ev.on('creds.update', saveCreds)
  sock.ev.on('connection.update', async ({connection, lastDisconnect, qr}) => {
    if (qr) { lastQR = qr; qrcodeTerminal.generate(qr,{small:false}) }
    if (connection === 'close') { const shouldReconnect = lastDisconnect?.error?.output?.statusCode!==DisconnectReason.loggedOut; if(shouldReconnect) start() }
    if (connection === 'open') {
      console.log('✅ CONECTADO')
      cron.schedule('0 7 * * 1', () => { enviarReportesAutomaticos(globalSock) }, { timezone: 'America/Mexico_City' })
      console.log('⏰ Cron lunes 7am CDMX activado para '+GRUPO_REPORTES_ID)
    }
  })
  sock.ev.on('messages.upsert', async ({messages})=>{
    try{
      const m=messages[0]; if(!m||m.key.fromMe) return; const jid=m.key.remoteJid; if(!jid.endsWith('@g.us')) return
      const rawLid=m.key.participant||''; const realPn=m.key.participantPn||m.key.participantAlt||''; let pnFromStore=''; try{ pnFromStore=await sock.signalRepository?.lidMapping?.getPNForLID(rawLid)||'' }catch{}
      const rawId=realPn||pnFromStore||rawLid||jid; const tel=rawId.replace(/\D/g,'')
      const texto=m.message?.conversation||m.message?.extendedTextMessage?.text||m.message?.imageMessage?.caption||''; const loc=m.message?.locationMessage
      if(texto.trim().toLowerCase()==='id'){ await sock.sendMessage(jid,{text:`ID: ${jid}`}); return }

      if(texto.toLowerCase().startsWith('resumen ')){
        const txtLow = texto.toLowerCase()
        if(txtLow.includes('bucareli') || txtLow.includes('juarez') || txtLow.includes('coyo') || txtLow.includes('hotel')){
          let suc = 'coyoacan'
          if(txtLow.includes('juarez') || txtLow.includes('bucareli')) suc = 'juarez'
          else if(txtLow.includes('coyo')) suc = 'coyoacan'
          await reporteSucursal(suc, jid, sock)
          const info = await generarExcelSemanaYEnviar(suc, jid, sock)
          await sock.sendMessage(jid,{text:`✅ Excel ${suc} - ${info.total} registros - ${info.rango}\nhttps://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}`})
          return
        }
        await resumenEmpleado(texto.slice(8).trim(), jid, sock); return
      }

      if(/^(reporte|checador)/i.test(texto)){
        const txtLow = texto.toLowerCase()
        if(txtLow.includes('auto') || txtLow.includes('test')){
          await enviarReportesAutomaticos(sock)
          return
        }
        let suc = 'coyoacan'
        if(txtLow.includes('juarez') || txtLow.includes('bucareli')) suc = 'juarez'
        else if(txtLow.includes('hotel')) suc = 'hotel'
        else if(txtLow.includes('coyo')) suc = 'coyoacan'
        await reporteSucursal(suc, jid, sock)
        const info = await generarExcelSemanaYEnviar(suc, jid, sock)
        await sock.sendMessage(jid,{text:`✅ Excel ${suc} - ${info.total} registros - ${info.rango}`})
        return
      }

      const esEntrada=/entr|ingres|lle?gue|aqui estoy|presente/i.test(texto); const esSalida=/salid|me voy|adios|bye/i.test(texto)
      if(!loc){ if(esEntrada) await sock.sendMessage(jid,{text:'Envía tu ubicación para entrada'},{quoted:m}); if(esSalida) await sock.sendMessage(jid,{text:'Envía tu ubicación para salida'},{quoted:m}); return }

      const lat=loc.degreesLatitude, lng=loc.degreesLongitude
      let cercana=null,dMin=Infinity; for(const s of SUCURSALES){ const d=distM(lat,lng,s.lat,s.lng); if(d<dMin){dMin=d; cercana=s} }

      const empRows=await getRows('Empleados!A:G')
      let emp = empRows.slice(1).find(r=>r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10))
      if(!emp && rawLid.includes('@lid') && m.pushName){ emp = empRows.slice(1).find(r=> r[1] && r[1].toLowerCase().includes(m.pushName.toLowerCase().split(' ')[0])) }
      const nombreRegistrado = emp? emp[1] : null
      const nombreFinal = nombreRegistrado || m.pushName || tel

      const fLab=fechaLaboral(); const asisRows=await getRows('Asistencia!A:K'); const idx=asisRows.findIndex((r,i)=>i>0&&r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10)&&r[2]===fLab)
      const hoy=idx>-1?asisRows[idx]:null; const sClient=await sheetsClient()
      let estatus='A TIEMPO'; let horaProg=null
      try{
        const cal=await getRows('Calendario_Horarios!A:E'); const progExc=cal.slice(1).find(r=>r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10)&&r[2]===fLab)
        if(progExc && progExc[3]) horaProg=progExc[3]
        else { const baseRows=await getRows('Horario_Base!A:J'); const baseRow=baseRows.slice(1).find(r=>r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10)); if(baseRow){ const fecha=new Date(fLab+'T12:00:00'); const mapa={1:3,2:4,3:5,4:6,5:7,6:8,0:9}; const valor=baseRow[mapa[fecha.getDay()]]||''; if(valor.toLowerCase().includes('descanso')) estatus='DESCANSO'; else horaProg=valor } }
        if(horaProg){ const hp=horaProg.toString().match(/\d{1,2}:\d{2}/)?.[0]||horaProg; const dif=minutos(horaMX())-minutos(hp); if(dif>15) estatus=`RETARDO ${dif}min (Prog ${hp})`; else estatus=`A TIEMPO Prog ${hp}` }
      }catch{}

      if(!hoy ||!hoy[3]){
        if(dMin>cercana.rEnt){ await sock.sendMessage(jid,{text:`Debes estar a max ${cercana.rEnt}m de ${cercana.nombre}, estas a ${Math.round(dMin)}m`},{quoted:m}); return }
        const h=horaMX()
        const horasK = calcularHorasTrabajadas(h, "");
        if(idx===-1) await sClient.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Asistencia!A:K',valueInputOption:'USER_ENTERED',requestBody:{values:[[tel,nombreFinal,fLab,h,estatus,'',cercana.nombre,Math.round(dMin).toString(),'','',horasK]]}})
        else await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!C${idx+1}:K${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[[fLab,h,estatus,'',cercana.nombre,Math.round(dMin).toString(),'','',horasK]]}})
        await sock.sendMessage(jid,{text:`✅ ${estatus} - ${nombreFinal} en ${cercana.nombre} (${Math.round(dMin)}m)`})
      }else{
        if(hoy[5]){ await sock.sendMessage(jid,{text:`Salida ya registrada ${hoy[5]} en ${hoy[8]||''} - ${hoy[10]||''}`}); return }
        const sucEntrada = hoy[6] || '';
        if(sucEntrada.toLowerCase().includes('hotel') &&!cercana.nombre.toLowerCase().includes('coyoacan')){
          await sock.sendMessage(jid,{text:`❌ Entraste en Servicio Hotel, debes checar salida en Trinidad Coyoacan. Estás en ${cercana.nombre}`},{quoted:m});
          return;
        }
        if(dMin>cercana.rSal){ await sock.sendMessage(jid,{text:`❌ No puedes checar salida a ${Math.round(dMin)}m de ${cercana.nombre}. Max ${cercana.rSal}m.`},{quoted:m}); return }
        const h=horaMX()
        const horasReales = calcularHorasTrabajadas(hoy[3], h);
        await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!F${idx+1}:K${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[[h,hoy[6]||'',hoy[7]||'',cercana.nombre,Math.round(dMin).toString(),horasReales]]}})
        await sock.sendMessage(jid,{text:`Salida ${nombreFinal} - ${cercana.nombre} - Trabajado: ${horasReales}`})
      }
    }catch(e){ console.error(e) }
  })
}
start()

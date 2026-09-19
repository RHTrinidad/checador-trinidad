import { default as makeWASocket, useMultiFileAuthState, DisconnectReason } from '@whiskeysockets/baileys'
import qrcodeTerminal from 'qrcode-terminal'
import QRCode from 'qrcode'
import express from 'express'
import { createRequire } from 'module'
const require = createRequire(import.meta.url)
const { google } = require('googleapis')

const SPREADSHEET_ID = process.env.SPREADSHEET_ID
const GRUPO_REPORTES_ID = process.env.GRUPO_REPORTES_ID

// --- LEE SUCURSALES Y SOPORTA radio_entrada / radio_salida ---
let SUCURSALES = []
try {
  if (process.env.SUCURSALES_JSON) {
    const parsed = JSON.parse(process.env.SUCURSALES_JSON)
    const arr = Array.isArray(parsed)? parsed : Object.values(parsed)
    SUCURSALES = arr.map(s => ({
      id: s.id || s.nombre,
      nombre: s.nombre || s.id,
      lat: s.lat,
      lng: s.lng || s.lon,
      rEnt: s.radio_entrada || s.rEnt || s.radio || 250,
      rSal: s.radio_salida || s.rSal || s.radio || 250
    }))
    console.log('Sucursales cargadas:', SUCURSALES)
  } else throw new Error('No SUCURSALES_JSON')
} catch (e) {
  console.log('Usando default', e.message)
  SUCURSALES = [
    { id:"COYOACAN", nombre:"Trinidad Coyoacan", lat:19.352525, lng:-99.161817, rEnt:250, rSal:250 },
    { id:"BUCARELI", nombre:"Trinidad Bucareli", lat:19.426523, lng:-99.153326, rEnt:250, rSal:250 },
    { id:"JUAREZ", nombre:"Trinidad Juarez", lat:19.4314119, lng:-99.1512074, rEnt:250, rSal:250 }
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
function minutos(h){
  const m = h?.toString().match(/(\d{1,2}):(\d{2})/)
  if(!m) return 0
  return parseInt(m[1])*60 + parseInt(m[2])
}
async function getRows(range){ const s=await sheetsClient(); const r=await s.spreadsheets.values.get({spreadsheetId:SPREADSHEET_ID,range}); return r.data.values||[] }

const app = express()
let lastQR = null
app.get('/', (req,res)=> res.send('Bot Trinidad OK - ve a /qr'))
app.get('/qr', async (req,res)=>{
  if(!lastQR) return res.send('<h2>Aun no hay QR, espera 10s</h2>')
  const dataUrl = await QRCode.toDataURL(lastQR)
  res.send(`<div style="text-align:center"><h2>Escanea</h2><img src="${dataUrl}" style="width:400px"></div>`)
})
app.listen(process.env.PORT||3000, ()=> console.log('Web en puerto '+(process.env.PORT||3000)))

async function start(){
  console.log('--- INICIANDO BOT TRINIDAD ---')
  const { state, saveCreds } = await useMultiFileAuthState('/app/auth')
 const sock=makeWASocket({
  auth:state, 
  printQRInTerminal:false,
  markOnlineOnConnect:false,
  syncFullHistory:false,
  generateHighQualityLinkPreview:false,
  retryRequestDelayMs:500,
  keepAliveIntervalMs:30000
})
  sock.ev.on('creds.update', saveCreds)
  sock.ev.on('connection.update', async (update) => {
    const { connection, lastDisconnect, qr } = update
    if (qr) { lastQR = qr; qrcodeTerminal.generate(qr, { small: false }) }
    if (connection === 'close') {
      const shouldReconnect = lastDisconnect?.error?.output?.statusCode!== DisconnectReason.loggedOut
      if (shouldReconnect) start()
    }
    if (connection === 'open') { console.log('✅ CONECTADO'); lastQR = null }
  })

  sock.ev.on('messages.upsert', async ({messages})=>{
    try{
      const m=messages[0]; if(!m||m.key.fromMe) return
      const jid=m.key.remoteJid; if(!jid.endsWith('@g.us')) return
     // WhatsApp ahora manda LID en participant, el telefono real viene en participantAlt
const rawId = m.key.participantAlt || m.key.participantPn || m.key.participant || jid
const tel = rawId.replace(/\D/g,'')
console.log('DEBUG tel detectado:', tel, 'raw:', rawId, 'pushName:', m.pushName)
      const texto=m.message?.conversation||m.message?.extendedTextMessage?.text||m.message?.imageMessage?.caption||''
      const loc=m.message?.locationMessage

      if(texto.trim().toLowerCase()==='id'){ await sock.sendMessage(jid,{text:`ID: ${jid}`}); return }

      const esEntrada=/entr|ingres|lle?gue|aqui estoy|presente/i.test(texto)
      const esSalida=/salid|me voy|adios|bye/i.test(texto)
      if(!loc){ if(esEntrada) await sock.sendMessage(jid,{text:'Envía tu ubicación para entrada'},{quoted:m}); if(esSalida) await sock.sendMessage(jid,{text:'Envía tu ubicación para salida'},{quoted:m}); return }

      const lat=loc.degreesLatitude, lng=loc.degreesLongitude
      let cercana=null, dMin=Infinity; for(const s of SUCURSALES){ const d=distM(lat,lng,s.lat,s.lng); if(d<dMin){dMin=d; cercana=s} }

      const empRows=await getRows('Empleados!A:G'); const emp=empRows.slice(1).find(r=>r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10))
      const nombre=emp?emp[1]:tel; const fLab=fechaLaboral()
      const asisRows=await getRows('Asistencia!A:H'); const idx=asisRows.findIndex((r,i)=>i>0&&r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10)&&r[2]===fLab)
      const hoy=idx>-1?asisRows[idx]:null; const sClient=await sheetsClient()

      // --- CALCULO DE RETARDO CON Horario_Base ---
      let estatus='A TIEMPO'
      let horaProg = null
      try{
        const cal=await getRows('Calendario_Horarios!A:E');
        const progExc=cal.slice(1).find(r=>r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10)&&r[2]===fLab)
        if(progExc && progExc[3]){ horaProg = progExc[3] }
        else {
          const baseRows = await getRows('Horario_Base!A:J');
          const baseRow = baseRows.slice(1).find(r=>r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10))
          if(baseRow){
            const fecha = new Date(fLab + 'T12:00:00')
            const dia = fecha.getDay() // 0 Dom, 1 Lun
            const mapa = {1:3, 2:4, 3:5, 4:6, 5:7, 6:8, 0:9}
            const valor = baseRow[mapa[dia]] || ''
            if(valor.toLowerCase().includes('descanso')){
              estatus = 'DESCANSO'
            } else {
              horaProg = valor
            }
          }
        }
        if(horaProg){
          const hp = horaProg.toString().match(/\d{1,2}:\d{2}/)?.[0] || horaProg
          const dif = minutos(horaMX()) - minutos(hp)
          if(dif>15) estatus=`RETARDO ${dif}min (Prog ${hp})`
          else estatus=`A TIEMPO Prog ${hp}`
        }
      }catch(e){ console.log('err horario', e) }

      if(!hoy ||!hoy[3]){
        if(dMin>cercana.rEnt){ await sock.sendMessage(jid,{text:`Debes estar a max ${cercana.rEnt}m de ${cercana.nombre}, estas a ${Math.round(dMin)}m`},{quoted:m}); return }
        if(idx===-1) await sClient.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Asistencia!A:H',valueInputOption:'USER_ENTERED',requestBody:{values:[[tel,nombre,fLab,horaMX(),estatus,'','','']]}})
        else await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!D${idx+1}:E${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[[horaMX(),estatus]]}})
        await sock.sendMessage(jid,{text:`✅ ${estatus} - ${nombre} en ${cercana.nombre}`})
      }else{
        if(hoy[5]){ await sock.sendMessage(jid,{text:`Salida ya registrada a las ${hoy[5]}`}); return }
        const h=horaMX(); await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!F${idx+1}:H${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[[h,h,Math.round(dMin).toString()]]}})
        if(dMin>cercana.rSal){
          await sock.sendMessage(GRUPO_REPORTES_ID,{text:`⚠️ Salida de ${nombre} a ${Math.round(dMin)}m de ${cercana.nombre}. Hora ${h}`})
        }
        await sock.sendMessage(jid,{text:`Salida registrada ${nombre} - ${cercana.nombre}`})
      }
    }catch(e){ console.error(e) }
  })
}
start()

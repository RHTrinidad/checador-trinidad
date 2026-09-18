import { default as makeWASocket, useMultiFileAuthState, DisconnectReason } from '@whiskeysockets/baileys'
import qrcode from 'qrcode-terminal'
import { createRequire } from 'module'
const require = createRequire(import.meta.url)

// tus otros requires siguen abajo de esto
const fs = require('fs')
const cron = require('node-cron')

async function startBot() {
    const { state, saveCreds } = await useMultiFileAuthState('auth_info_baileys')
    
    const sock = makeWASocket({
        auth: state
    })

    sock.ev.on('creds.update', saveCreds)

    sock.ev.on('connection.update', async (update) => {
        const { connection, lastDisconnect, qr } = update
        if (qr) {
            console.log('Escanea este QR con el WhatsApp de Trinidad:')
            qrcode.generate(qr, { small: true })
        }
        if (connection === 'close') {
            const shouldReconnect = lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut
            console.log('Desconectado, reconectando:', shouldReconnect)
            if (shouldReconnect) {
                startBot()
            }
        }
        if (connection === 'open') {
            console.log('✅ BOT TRINIDAD CONECTADO')
        }
    })

 const require = createRequire(import.meta.url)
const { google } = require('googleapis')
const cron = require('node-cron')

const SPREADSHEET_ID = process.env.SPREADSHEET_ID
const GRUPO_REPORTES_ID = process.env.GRUPO_REPORTES_ID
const SUCURSALES = [
  { id:"COYOACAN", nombre:"Trinidad Coyoacan", lat:19.352525, lng:-99.161817, rEnt:150, rSal:100 },
  { id:"BUCARELI", nombre:"Trinidad Bucareli", lat:19.426523, lng:-99.153326, rEnt:150, rSal:100 }
]

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets']
})
async function sheetsClient(){ const c=await auth.getClient(); return google.sheets({version:'v4',auth:c}) }
function distM(a,b,c,d){ const R=6371000, toRad=x=>x*Math.PI/180; const dLa=toRad(c-a), dLo=toRad(d-b); const q=Math.sin(dLa/2)**2+Math.cos(toRad(a))*Math.cos(toRad(c))*Math.sin(dLo/2)**2; return R*2*Math.atan2(Math.sqrt(q),Math.sqrt(1-q)) }
function fechaLaboral(d=new Date()){ const x=new Date(d); if(x.getHours()<4) x.setDate(x.getDate()-1); return x.toISOString().split('T')[0] }
function horaMX(d=new Date()){ return d.toLocaleTimeString('es-MX',{hour12:false,timeZone:'America/Mexico_City'}) }
function minutos(h){ const [hh,mm]=h.split(':').map(Number); return hh*60+mm }
async function getRows(range){ const s=await sheetsClient(); const r=await s.spreadsheets.values.get({spreadsheetId:SPREADSHEET_ID,range}); return r.data.values||[] }

async function start(){
  const {state,saveCreds}=await useMultiFileAuthState('/app/baileys_auth')
  const sock=makeWASocket({auth:state,printQRInTerminal:true})
  sock.ev.on('creds.update',saveCreds)
  sock.ev.on('connection.update',({connection})=>{ if(connection==='open') console.log('✅ BOT TRINIDAD LISTO') })

  sock.ev.on('messages.upsert', async ({messages})=>{
    try{
      const m=messages[0]; if(!m||m.key.fromMe) return
      const jid=m.key.remoteJid; if(!jid.endsWith('@g.us')) return
      const tel=(m.key.participant||jid).replace(/\D/g,'')
      const texto=m.message?.conversation||m.message?.extendedTextMessage?.text||m.message?.imageMessage?.caption||''
      const loc=m.message?.locationMessage

      if(texto.trim().toLowerCase()==='id'){ await sock.sendMessage(jid,{text:`ID de este grupo:\n${jid}`}); if(jid!==GRUPO_REPORTES_ID) await sock.sendMessage(GRUPO_REPORTES_ID,{text:`📌 Nuevo grupo detectado ID: ${jid}`}); return }

      if(jid===GRUPO_REPORTES_ID && texto.toLowerCase().startsWith('registrar ')){
        const s=await sheetsClient(); const partes=texto.split(' '); const num=partes[1].replace(/\D/g,''); const resto=texto.substring(texto.indexOf(partes[1])+partes[1].length).trim()||'Sin Nombre'
        await s.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Empleados!A:G',valueInputOption:'USER_ENTERED',requestBody:{values:[[num,resto,'COYOACAN','','','LUNES','']]}})
        await sock.sendMessage(jid,{text:`✅ Empleado ${num} registrado como ${resto}`}); return
      }
      if(jid===GRUPO_REPORTES_ID && texto.toLowerCase()==='reporte hoy'){
        const rows=await getRows('Asistencia!A:H'); const hoy=fechaLaboral(); const deHoy=rows.filter(r=>r[2]===hoy);
        let msg=`📊 Reporte ${hoy}\nTotal fichados: ${deHoy.length}\n`; deHoy.forEach(r=>{ msg+=`- ${r[1]} E:${r[3]||'-'} S:${r[5]||'-'} ${r[4]||''}\n` })
        await sock.sendMessage(jid,{text:msg}); return
      }

      const esEntrada=/entr|ingres|lle?gue|aqui estoy|presente/i.test(texto)
      const esSalida=/salid|me voy|adios|bye/i.test(texto)
      if(!loc){ if(esEntrada) await sock.sendMessage(jid,{text:'Envía tu ubicación para registrar entrada'},{quoted:m}); if(esSalida) await sock.sendMessage(jid,{text:'Envía tu ubicación para registrar salida'},{quoted:m}); return }

      const lat=loc.degreesLatitude, lng=loc.degreesLongitude
      let cercana=null, dMin=Infinity; for(const s of SUCURSALES){ const d=distM(lat,lng,s.lat,s.lng); if(d<dMin){dMin=d; cercana=s} }

      const empRows=await getRows('Empleados!A:G'); const emp=empRows.slice(1).find(r=>r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10))
      const nombre=emp?emp[1]:tel; const fLab=fechaLaboral()
      const asisRows=await getRows('Asistencia!A:H'); const idx=asisRows.findIndex((r,i)=>i>0&&r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10)&&r[2]===fLab)
      const hoy=idx>-1?asisRows[idx]:null; const sClient=await sheetsClient()

      // CALENDARIO PARA RETARDO
      let estatus='A TIEMPO'
      try{
        const cal=await getRows('Calendario_Horarios!A:E'); const prog=cal.slice(1).find(r=>r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10)&&r[2]===fLab)
        if(prog && prog[3]){ const dif=minutos(horaMX())-minutos(prog[3]); if(dif>15) estatus=`RETARDO ${dif}min (Prog ${prog[3]})`; else estatus=`A TIEMPO Prog ${prog[3]}` }
      }catch{}

      if(!hoy ||!hoy[3]){
        if(dMin>cercana.rEnt){ await sock.sendMessage(jid,{text:'la entrada debe registrarse en la unidad'},{quoted:m}); return }
        if(idx===-1) await sClient.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Asistencia!A:H',valueInputOption:'USER_ENTERED',requestBody:{values:[[tel,nombre,fLab,horaMX(),estatus,'','','']]}})
        else await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!D${idx+1}:E${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[[horaMX(),estatus]]}})
        await sock.sendMessage(jid,{text:`Buen turno ${nombre} - ${cercana.nombre}`})
      }else{
        if(hoy[5]){ await sock.sendMessage(jid,{text:`Salida ya registrada a las ${hoy[5]} - ${nombre}`}); return }
        const h=horaMX(); await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!F${idx+1}:H${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[[h,h,Math.round(dMin).toString()]]}})
        if(dMin>cercana.rSal){
          await sock.sendMessage(GRUPO_REPORTES_ID,{text:`⚠️ Salida de ${nombre} (${tel}) a ${Math.round(dMin)}m de ${cercana.nombre}. Hora: ${h}. Fuera de rango 100m. Se conserva hora primer intento.`})
          await sock.sendMessage(jid,{text:`Salida registrada ${nombre}`})
        }else await sock.sendMessage(jid,{text:`Salida registrada ${nombre} - ${cercana.nombre}`})
      }
    }catch(e){ console.error(e) }
  })

  // Chequeo cada 5 min quien no llego +20min
  cron.schedule('*/5 * * * *', async ()=>{
    try{
      const hoy=fechaLaboral(); const cal=await getRows('Calendario_Horarios!A:E'); const asis=await getRows('Asistencia!A:H')
      const ahoraMin=minutos(horaMX()); for(const r of cal.slice(1)){ if(r[2]!==hoy) continue; const tel=r[0]; const prog=r[3]; if(!prog) continue; if(ahoraMin - minutos(prog) === 20){ const tiene=asis.find(a=>a[0]&&a[0].replace(/\D/g,'').slice(-10)===tel.slice(-10)&&a[2]===hoy&&a[3]); if(!tiene){ const sock2=global.sockRef; } } }
    }catch(e){ console.log('cron 20m',e.message) }
  },{timezone:'America/Mexico_City'})

  // Lunes 7am resumen 48h + extras
  cron.schedule('0 7 * * 1', async ()=>{
    try{
      const s=await sheetsClient(); await sock.sendMessage(GRUPO_REPORTES_ID,{text:`📊 Resumen semanal Trinidad\nBase 48h + extras\nRevisa pestaña Asistencia - Columna retardo >15min\nDescanso general LUNES`})
    }catch{}
  },{timezone:'America/Mexico_City'})
}
startBot()
}

require('dotenv').config()
const { default: makeWASocket, useMultiFileAuthState } = require('@whiskeysockets/baileys')
const { google } = require('googleapis')
const cron = require('node-cron')

const SPREADSHEET_ID = process.env.SPREADSHEET_ID
const GRUPO_REPORTES_ID = process.env.GRUPO_REPORTES_ID
const SUCURSALES = [{id:"COYOACAN",nombre:"Trinidad Coyoacan",lat:19.352525,lng:-99.161817},{id:"BUCARELI",nombre:"Trinidad Bucareli",lat:19.426523,lng:-99.153326}]

const auth = new google.auth.GoogleAuth({ credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON), scopes: ['https://www.googleapis.com/auth/spreadsheets'] })
async function sheetsClient(){ const c=await auth.getClient(); return google.sheets({version:'v4',auth:c}) }
function distM(a,b,c,d){ const R=6371000, toRad=x=>x*Math.PI/180; const dLa=toRad(c-a), dLo=toRad(d-b); const q=Math.sin(dLa/2)**2+Math.cos(toRad(a))*Math.cos(toRad(c))*Math.sin(dLo/2)**2; return R*2*Math.atan2(Math.sqrt(q),Math.sqrt(1-q)) }
function fechaLaboral(d=new Date()){ const x=new Date(d); if(x.getHours()<4) x.setDate(x.getDate()-1); return x.toISOString().split('T')[0] }
function horaMX(){ return new Date().toLocaleTimeString('es-MX',{hour12:false,timeZone:'America/Mexico_City'}) }
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
      const texto=m.message?.conversation||m.message?.extendedTextMessage?.text||''
      const loc=m.message?.locationMessage
      if(texto.trim().toLowerCase()==='id'){ await sock.sendMessage(jid,{text:`ID: ${jid}`}); return }
      if(jid===GRUPO_REPORTES_ID && texto.toLowerCase().startsWith('registrar ')){
        const s=await sheetsClient(); const num=texto.split(' ')[1].replace(/\D/g,''); const resto=texto.substring(texto.indexOf(texto.split(' ')[1])+texto.split(' ')[1].length).trim()
        await s.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Empleados!A:F',valueInputOption:'USER_ENTERED',requestBody:{values:[[num,resto,'COYOACAN','','','LUNES']]}})
        await sock.sendMessage(jid,{text:`✅ Registrado ${num} -> ${resto}`}); return
      }
      const esEntrada=/entr|ingres|lle?gue|aqui estoy|presente/i.test(texto); const esSalida=/salid|me voy|adios|bye/i.test(texto)
      if(!loc){ if(esEntrada){ await sock.sendMessage(jid,{text:'Envía tu ubicación para registrar entrada'},{quoted:m}); return } if(esSalida){ await sock.sendMessage(jid,{text:'Envía tu ubicación para registrar salida'},{quoted:m}); return } return }
      const lat=loc.degreesLatitude, lng=loc.degreesLongitude
      let cercana=null, dMin=Infinity; for(const s of SUCURSALES){ const d=distM(lat,lng,s.lat,s.lng); if(d<dMin){dMin=d; cercana=s} }
      const empRows=await getRows('Empleados!A:G'); const emp=empRows.slice(1).find(r=>r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10)); const nombre=emp?emp[1]:tel
      const fLab=fechaLaboral(); const asisRows=await getRows('Asistencia!A:H'); const idx=asisRows.findIndex((r,i)=>i>0&&r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10)&&r[2]===fLab); const hoy=idx>-1?asisRows[idx]:null
      const sClient=await sheetsClient()
      if(!hoy||!hoy[3]){
        if(dMin>150){ await sock.sendMessage(jid,{text:'la entrada debe registrarse en la unidad'},{quoted:m}); return }
        if(idx===-1){ await sClient.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Asistencia!A:H',valueInputOption:'USER_ENTERED',requestBody:{values:[[tel,nombre,fLab,horaMX(),'A TIEMPO','','','']]}})}
        else{ await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!D${idx+1}:E${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[[horaMX(),'A TIEMPO']]}})}
        await sock.sendMessage(jid,{text:`Buen turno ${nombre} - ${cercana.nombre}`})
      }else{
        if(hoy[5]){ await sock.sendMessage(jid,{text:`Salida ya registrada a las ${hoy[5]} - ${nombre}`}); return }
        const h=horaMX(); await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!F${idx+1}:H${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[[h,h,Math.round(dMin).toString()]]}})
        if(dMin>100){ await sock.sendMessage(GRUPO_REPORTES_ID,{text:`⚠️ Salida de ${nombre} a ${Math.round(dMin)}m de ${cercana.nombre}. Hora: ${h}`}); await sock.sendMessage(jid,{text:`Salida registrada ${nombre}`}) }
        else{ await sock.sendMessage(jid,{text:`Salida registrada ${nombre} - ${cercana.nombre}`}) }
      }
    }catch(e){ console.error(e) }
  })
  cron.schedule('0 7 * * 1', async ()=>{ try{ const sc=await sheetsClient(); /* resumen */ }catch{} }, {timezone:'America/Mexico_City'})
}
start()

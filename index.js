import makeWASocket, { useMultiFileAuthState, DisconnectReason, fetchLatestBaileysVersion } from '@whiskeysockets/baileys'
import qrcodeTerminal from 'qrcode-terminal'
import QRCode from 'qrcode'
import express from 'express'
import fs from 'fs'
import path from 'path'
import P from 'pino'
import { GRUPOS, PAQUETES, getTipoGrupo, GRUPO_COYOACAN_ID, GRUPO_BUCARELI_ID, GRUPO_JUAREZ_ID, GRUPO_CHECADOR_COYOACAN_ID, GRUPO_CHECADOR_BUCARELI_ID } from './src/config.js'
import { handleChecador } from './src/checador.js'

const app=express(); let lastQR=null;
app.get('/',(r,s)=>s.send('Bot OK - /qr'));
app.get('/qr',async(r,s)=>{ if(!lastQR) return s.send('No QR'); const d=await QRCode.toDataURL(lastQR); s.send(`<img src="${d}" style="width:400px">`) });
app.get('/backup',(r,s)=>{ try{ const dir='./baileys_auth'; if(!fs.existsSync(dir)) return s.json({error:'no auth'}); const files=fs.readdirSync(dir); let o={}; files.forEach(f=>{ try{o[f]=fs.readFileSync(path.join(dir,f),'utf8')}catch{}}); s.json(o)}catch(e){s.json({error:e.message})} });
app.listen(process.env.PORT||3000,()=>console.log('WEB '+process.env.PORT));

let cache=new Map()
async function getFiltro(jid,sock){
  if(GRUPOS.REPORTES.includes(jid)) return null
  if(jid===GRUPO_COYOACAN_ID||jid===GRUPO_CHECADOR_COYOACAN_ID) return 'coyoacan'
  if(jid===GRUPO_BUCARELI_ID||jid===GRUPO_JUAREZ_ID||jid===GRUPO_CHECADOR_BUCARELI_ID) return 'juarez'
  return null
}

async function start(){
  const {state,saveCreds}=await useMultiFileAuthState('/app/baileys_auth')
  const {version}=await fetchLatestBaileysVersion()
  const sock=makeWASocket({version,auth:state,logger:P({level:'fatal'}),browser:['Trinidad Bot','Chrome','121'],getMessage:async()=>undefined})
  sock.ev.on('creds.update',saveCreds)
  sock.ev.on('connection.update',async({connection,lastDisconnect,qr})=>{ if(qr){lastQR=qr; qrcodeTerminal.generate(qr,{small:false})} if(connection==='close'){ const c=lastDisconnect?.error?.output?.statusCode; if(c!==DisconnectReason.loggedOut) setTimeout(()=>start(),5000)} if(connection==='open'){console.log('CONECTADO MODULAR'); lastQR=null} })
  sock.ev.on('messages.upsert',async({messages})=>{
    try{
      const m=messages[0]; if(!m||m.key.fromMe) return; const jid=m.key.remoteJid; if(!jid.endsWith('@g.us')) return
      const rawLid=m.key.participant||''; const realPn=m.key.participantPn||''; let pn=''; try{pn=await sock.signalRepository?.lidMapping?.getPNForLID(rawLid)||''}catch{}; const rawId=realPn||pn||rawLid||jid
      let tel=(rawId||'').toString().replace(/\D/g,''); let tel10=tel.slice(-10)
      const texto=m.message?.conversation||m.message?.extendedTextMessage?.text||m.message?.imageMessage?.caption||''; const loc=m.message?.locationMessage||m.message?.liveLocationMessage
      const tipo=getTipoGrupo(jid); if(!tipo) return
      if(tipo==='CHECADORES'){
        if(!loc) return
        await handleChecador({sock,jid,m,loc,rawLid,tel10,tel})
      }
    }catch(e){ console.error(e) }
  })
}
start()

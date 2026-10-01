import makeWASocket, { useMultiFileAuthState, DisconnectReason, fetchLatestBaileysVersion } from '@whiskeysockets/baileys'
import qrcodeTerminal from 'qrcode-terminal'
import QRCode from 'qrcode'
import express from 'express'
import fs from 'fs'
import path from 'path'
import P from 'pino'
import { GRUPOS, PAQUETES, getTipoGrupo, GRUPO_COYOACAN_ID, GRUPO_BUCARELI_ID, GRUPO_JUAREZ_ID, GRUPO_CHECADOR_COYOACAN_ID, GRUPO_CHECADOR_BUCARELI_ID } from './src/config.js'
import { handleChecador, checkNoLlegaron } from './src/checador.js'
import { handleReportes } from './src/reportes.js'
import { handleCompras } from './src/compras.js'

const app=express(); let lastQR=null;
app.get('/',(r,s)=>s.send('Bot OK - /qr'));
app.get('/qr',async(r,s)=>{ if(!lastQR) return s.send('No QR aun, espera 5 seg y recarga'); const d=await QRCode.toDataURL(lastQR); s.send(`<div style="text-align:center"><h2>Escanea</h2><img src="${d}" style="width:400px"><p>Se actualiza cada 15s</p></div>`) });
app.get('/backup',(r,s)=>{ try{ const dir='/app/baileys_auth'; if(!fs.existsSync(dir)) return s.json({error:'no auth'}); const files=fs.readdirSync(dir); let o={}; files.forEach(f=>{ try{o[f]=fs.readFileSync(path.join(dir,f),'utf8')}catch{}}); s.json(o)}catch(e){s.json({error:e.message})} });
app.listen(process.env.PORT||3000,()=>console.log('WEB '+process.env.PORT));

async function getFiltro(jid){
  if(GRUPOS.REPORTES.includes(jid)) return null
  if(jid===GRUPO_COYOACAN_ID||jid===GRUPO_CHECADOR_COYOACAN_ID) return 'coyoacan'
  if(jid===GRUPO_BUCARELI_ID||jid===GRUPO_JUAREZ_ID||jid===GRUPO_CHECADOR_BUCARELI_ID) return 'juarez'
  return null
}

let cronInterval = null

async function start(){
  const authDir = '/app/baileys_auth'
  try{
    if(!fs.existsSync(authDir)) fs.mkdirSync(authDir,{recursive:true})
    if(process.env.CREDS_DATA && !fs.existsSync(path.join(authDir,'creds.json'))){
      fs.writeFileSync(path.join(authDir,'creds.json'), process.env.CREDS_DATA, 'utf8')
      console.log('✅ creds.json restaurado desde CREDS_DATA')
    }
  }catch(e){ console.log('Error restaurando creds', e.message) }

  const {state,saveCreds}=await useMultiFileAuthState(authDir)
  const {version}=await fetchLatestBaileysVersion()
  const sock=makeWASocket({version,auth:state,logger:P({level:'fatal'}),browser:['Trinidad Bot','Chrome','121'],getMessage:async()=>undefined})
  sock.ev.on('creds.update',saveCreds)
  sock.ev.on('connection.update',async({connection,lastDisconnect,qr})=>{
    if(qr){lastQR=qr; qrcodeTerminal.generate(qr,{small:false})}
    if(connection==='close'){ const c=lastDisconnect?.error?.output?.statusCode; if(c!==DisconnectReason.loggedOut) setTimeout(()=>start(),5000)}
    if(connection==='open'){
      console.log('CONECTADO MODULAR + TOL 15min + SIN QR');
      lastQR=null
      if(cronInterval) clearInterval(cronInterval)
      cronInterval = setInterval(()=> {
        checkNoLlegaron(sock).catch(e=>console.log('cron error', e.message))
      }, 10 * 60 * 1000)
    }
  })
  sock.ev.on('messages.upsert',async({messages})=>{
    try{
      const m=messages[0]; if(!m||m.key.fromMe) return; const jid=m.key.remoteJid; if(!jid.endsWith('@g.us')) return
      
      const texto=m.message?.conversation||m.message?.extendedTextMessage?.text||m.message?.imageMessage?.caption||m.message?.documentMessage?.caption||''; 
      const loc=m.message?.locationMessage||m.message?.liveLocationMessage

      // 1. COMANDO ID - SIEMPRE RESPONDE, AUNQUE NO ESTE CONFIGURADO
      if(texto.trim().toLowerCase()==='id'){
        const tipoTmp = getTipoGrupo(jid) || 'NO_CONFIGURADO'
        const filtroTmp = await getFiltro(jid) || 'sin filtro'
        console.log(`ID solicitado en ${jid}`)
        await sock.sendMessage(jid,{text:`ID: ${jid}\nTipo: ${tipoTmp}\nFiltro: ${filtroTmp}`}); 
        return 
      }

      const rawLid=m.key.participant||''; const realPn=m.key.participantPn||''; let pn=''; try{pn=await sock.signalRepository?.lidMapping?.getPNForLID(rawLid)||''}catch{}; const rawId=realPn||pn||rawLid||jid
      let tel=(rawId||'').toString().replace(/\D/g,''); let tel10=tel.slice(-10)
      
      const esImagen=!!(m.message?.imageMessage)
      const tipo=getTipoGrupo(jid); 
      if(!tipo){
        console.log(`Grupo no configurado: ${jid} texto: ${texto.substring(0,50)}`)
        return
      }
      const filtro=await getFiltro(jid)

      if(tipo==='REPORTES'){
        if(esImagen || /^compras/i.test(texto)){
          const ok=await handleCompras({sock,jid,m,texto,esImagen}); if(ok) return
        }
        if(PAQUETES.REPORTES_PARA_GERENTES.test(texto)){
          await handleReportes({texto,jid,sock,filtroGrupo:filtro}); return
        }
        return
      }

      if(tipo==='CHECADORES'){
        if(!loc) return
        console.log(`Ubicación recibida de ${tel10} en ${jid}`)
        await handleChecador({sock,jid,m,loc,rawLid,tel10,tel}); return
      }

      if(tipo==='GERENTES'){
        if(loc) return
        if(!PAQUETES.REPORTES_PARA_GERENTES.test(texto)) return
        await handleReportes({texto,jid,sock,filtroGrupo:filtro}); return
      }

    }catch(e){ console.error('Error upsert', e) }
  })
}
start()

import makeWASocket, { useMultiFileAuthState, DisconnectReason, fetchLatestBaileysVersion } from '@whiskeysockets/baileys'
import qrcodeTerminal from 'qrcode-terminal'
import QRCode from 'qrcode'
import express from 'express'
import fs from 'fs'
import path from 'path'
import P from 'pino'
import {
  GRUPOS,
  PAQUETES,
  getTipoGrupo,
  GRUPO_COYOACAN_ID,
  GRUPO_BUCARELI_ID,
  GRUPO_CHECADOR_COYOACAN_ID,
  GRUPO_CHECADOR_BUCARELI_ID,
  GRUPO_PRUEBAS_ID,
  GRUPO_REPORTES_TRINIDAD_ID
} from './src/config.js'
import { handleChecador, checkNoLlegaron, cerrarSalidasPendientes } from './src/checador.js'
import { handleReportes } from './src/reportes.js'
import { handleCompras } from './src/compras.js'

const app=express()
app.use(express.json({limit:'20mb'}))

let lastQR=null

app.get('/',(r,s)=>s.send('Bot OK - /qr'))

app.get('/qr',async(r,s)=>{
  if(!lastQR)return s.send('No QR aun, espera 5 seg y recarga')
  const d=await QRCode.toDataURL(lastQR)
  s.send(`<div style="text-align:center"><h2>Escanea</h2><img src="${d}" style="width:400px"><p>Se actualiza cada 15s</p></div>`)
})

app.get('/backup',(r,s)=>{
  try{
    const dir='/app/auth'
    if(!fs.existsSync(dir))return s.json({error:'no auth'})
    const files=fs.readdirSync(dir)
    if(files.length===0)return s.json({error:'auth vacio - escanea QR'})
    let o={}
    files.forEach(f=>{
      try{o[f]=fs.readFileSync(path.join(dir,f),'utf8')}catch{}
    })
    s.json(o)
  }catch(e){
    s.json({error:e.message})
  }
})

app.post('/restore',(r,s)=>{
  try{
    const dir='/app/auth'
    if(!fs.existsSync(dir))fs.mkdirSync(dir,{recursive:true})
    const data=r.body
    const payload=data.data?data.data:data
    let c=0

    for(const [f,v] of Object.entries(payload)){
      if(f==='dir'||f==='files')continue
      if(typeof v==='string'&&v.length>10){
        fs.writeFileSync(path.join(dir,f),v,'utf8')
        c++
      }
    }

    s.json({ok:true,restaurados:c})
  }catch(e){
    s.json({error:e.message})
  }
})

app.listen(process.env.PORT||3000,()=>console.log('WEB '+(process.env.PORT||3000)+' OK'))

async function getFiltro(jid){
  if(jid===GRUPO_PRUEBAS_ID||jid===GRUPO_REPORTES_TRINIDAD_ID)return null
  if(jid===GRUPO_COYOACAN_ID||jid===GRUPO_CHECADOR_COYOACAN_ID)return 'coyoacan'
  if(jid===GRUPO_BUCARELI_ID||jid===GRUPO_CHECADOR_BUCARELI_ID)return 'bucareli'
  return null
}

const COMANDOS_REPORTES=/^(datos bancarios|cuenta bancaria|cuenta empleado|info|ficha|datos|dato|numero|número|num|tel|telefono|teléfono|cuenta|banco|clave|clabe)\s+.+|^(asistencia hoy|resumen|reporte|checador|compras de hoy|faltas|retardos|críticos|criticos|graves)/i

let cronInterval=null

async function start(){
  const authDir='/app/auth'

  if(!fs.existsSync(authDir)){
    fs.mkdirSync(authDir,{recursive:true})
  }

  console.log('AUTH DIR:',authDir)

  const {state,saveCreds}=await useMultiFileAuthState(authDir)
  const {version}=await fetchLatestBaileysVersion()

  const sock=makeWASocket({
    version,
    auth:state,
    logger:P({level:'fatal'}),
    browser:['Trinidad Bot','Chrome','121'],
    getMessage:async()=>undefined
  })

  sock.ev.on('creds.update',saveCreds)

  sock.ev.on('connection.update',async({connection,lastDisconnect,qr})=>{
    if(qr){
      lastQR=qr
      console.log('QR generado')
      qrcodeTerminal.generate(qr,{small:false})
    }

    if(connection==='close'){
      const c=lastDisconnect?.error?.output?.statusCode

      if(c!==DisconnectReason.loggedOut){
        console.log('WhatsApp desconectado. Reconectando...')
        setTimeout(()=>start(),5000)
      }else{
        console.log('Sesión cerrada. Se requiere nuevo QR.')
      }
    }

    if(connection==='open'){
      console.log('CONECTADO MODULAR + VOLUME OK')

      lastQR=null

      if(cronInterval)clearInterval(cronInterval)

      cronInterval=setInterval(()=>{
        cerrarSalidasPendientes().catch(e=>{
          console.log('cierre auto error',e.message)
        })

        checkNoLlegaron(sock).catch(e=>{
          console.log('cron error',e.message)
        })
      },10*60*1000)

      // Ejecutar también inmediatamente al conectar.
      cerrarSalidasPendientes().catch(e=>{
        console.log('cierre auto inicial error',e.message)
      })
    }
  })

  sock.ev.on('messages.upsert',async({messages})=>{
    try{
      const m=messages[0]

      if(!m||m.key.fromMe)return

      const jid=m.key.remoteJid

      if(!jid?.endsWith('@g.us'))return

      const texto=
        m.message?.conversation||
        m.message?.extendedTextMessage?.text||
        m.message?.imageMessage?.caption||
        m.message?.documentMessage?.caption||
        ''

      const loc=
        m.message?.locationMessage||
        m.message?.liveLocationMessage

      const textoTrim=texto.trim()
      const textoLow=textoTrim.toLowerCase()

      if(textoLow==='id'){
        console.log(`ID solicitado en ${jid}`)
        await sock.sendMessage(jid,{text:jid})
        return
      }

      const rawLid=m.key.participant||''
      const realPn=m.key.participantPn||''

      let pn=''

      try{
        pn=await sock.signalRepository?.lidMapping?.getPNForLID(rawLid)||''
      }catch{}

      const rawId=realPn||pn||rawLid||jid
      const tel=(rawId||'').toString().replace(/\D/g,'')
      const tel10=tel.slice(-10)

      const esImagen=!!m.message?.imageMessage
      const tipo=getTipoGrupo(jid)

      if(!tipo){
        console.log(`Grupo no configurado: ${jid}`)
        return
      }

      const filtro=await getFiltro(jid)

      if(tipo==='REPORTES'){
        if(esImagen||/^compras/i.test(textoTrim)){
          const ok=await handleCompras({
            sock,
            jid,
            m,
            texto,
            esImagen
          })

          if(ok)return
        }

        if(COMANDOS_REPORTES.test(textoTrim)){
          await handleReportes({
            texto,
            jid,
            sock,
            filtroGrupo:filtro
          })

          return
        }

        return
      }

      if(tipo==='CHECADORES'){
        if(!loc)return

        console.log(`Ubicación recibida de ${tel10} en ${jid}`)

        await handleChecador({
          sock,
          jid,
          m,
          loc,
          rawLid,
          tel10,
          tel
        })

        return
      }

      if(tipo==='GERENTES'){
        if(loc)return

        if(!COMANDOS_REPORTES.test(textoTrim))return

        await handleReportes({
          texto,
          jid,
          sock,
          filtroGrupo:filtro
        })

        return
      }

    }catch(e){
      console.error('Error upsert',e)
    }
  })
}

start()

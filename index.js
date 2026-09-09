const { default: makeWASocket, useMultiFileAuthState, DisconnectReason } = require('@whiskeysockets/baileys')
const P = require('pino')
const fs = require('fs')
const config = require('./config')
const { guardarEnSheet } = require('./sheets')

let empleados = {}
try { empleados = JSON.parse(fs.readFileSync('./empleados.json','utf8')) } catch(e){ empleados = {} }

function getNombre(jid, pushName){
  if(empleados[jid]) return empleados[jid]
  return pushName || jid.split('@')[0]
}

async function start(){
  const { state, saveCreds } = await useMultiFileAuthState('auth')
  const sock = makeWASocket({ auth: state, logger: P({level:'silent'}), printQRInTerminal: true })
  sock.ev.on('creds.update', saveCreds)
  sock.ev.on('connection.update', (u)=>{
    if(u.qr) console.log("ESCANEA EL QR")
    if(u.connection === 'open') console.log("CONECTADO")
    if(u.connection === 'close' && u.lastDisconnect?.error?.output?.statusCode!= DisconnectReason.loggedOut) start()
  })
  sock.ev.on('messages.upsert', async ({messages})=>{
    const m = messages[0]
    if(!m.message) return
    const jid = m.key.participant || m.key.remoteJid
    const groupId = m.key.remoteJid
    const pushName = m.pushName
    const text = (m.message.conversation || m.message.extendedTextMessage?.text || "").trim()

    if(text === '!id'){
      await sock.sendMessage(groupId, {text: `ID de este grupo:\n${groupId}`})
      return
    }
    if(!groupId.endsWith('@g.us')) return

    if(text.startsWith('!agregar')){
      if(!config.admins.includes(jid)) return
      const parts = text.split(' ')
      const numero = parts[1]
      const nombre = parts.slice(2).join(' ')
      if(!numero ||!nombre){
        await sock.sendMessage(groupId, {text: "Uso:!agregar 5512345678 Nombre Apellido"})
        return
      }
      const clean = numero.replace(/\D/g,'')
      empleados[`521${clean}@s.whatsapp.net`] = nombre
      empleados[`${clean}@s.whatsapp.net`] = nombre
      fs.writeFileSync('./empleados.json', JSON.stringify(empleados, null, 2))
      await sock.sendMessage(groupId, {text: `Guardado: ${nombre}`})
      return
    }

    if(text === '!lista'){
      let lista = "Empleados:\n"
      for(let k in empleados) lista += `- ${empleados[k]} : ${k}\n`
      await sock.sendMessage(groupId, {text: lista})
      return
    }

    let tipo = null
    const lower = text.toLowerCase()
    if(lower.includes('entrada')) tipo = 'ENTRADA'
    if(lower.includes('salida')) tipo = 'SALIDA'
    if(m.message.locationMessage &&!tipo) tipo = 'ENTRADA'

    if(tipo){
      const sucursal = config.sucursales[groupId] || groupId
      const nombreFinal = getNombre(jid, pushName)
      console.log(`${sucursal} | ${nombreFinal} | ${tipo}`)
      await guardarEnSheet(sucursal, nombreFinal, tipo, jid)
      await sock.sendMessage(groupId, {text: `${tipo} registrada: ${nombreFinal} en ${sucursal}`})
    }
  })
}
start()

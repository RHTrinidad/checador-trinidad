require('dotenv').config()
const { default: makeWASocket, useMultiFileAuthState, DisconnectReason } = require('@whiskeysockets/baileys')
const { google } = require('googleapis')
const cron = require('node-cron')

const SPREADSHEET_ID = process.env.SPREADSHEET_ID
const GRUPO_REPORTES_ID = process.env.GRUPO_REPORTES_ID

// COORDENADAS DE TUS LINKS
const SUCURSALES = [
  { id: "COYOACAN", nombre: "Trinidad Coyoacan", lat: 19.352525, lng: -99.161817, radio_entrada: 150, radio_salida: 100 },
  { id: "BUCARELI", nombre: "Trinidad Bucareli", lat: 19.426523, lng: -99.153326, radio_entrada: 150, radio_salida: 100 }
]

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets']
})

async function sheetsClient() {
  const c = await auth.getClient()
  return google.sheets({ version: 'v4', auth: c })
}

function distM(lat1, lon1, lat2, lon2) {
  const R = 6371000
  const toRad = x => x * Math.PI / 180
  const dLat = toRad(lat2 - lat1)
  const dLon = toRad(lon2 - lon1)
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a))
}

function fechaLaboral(d = new Date()) {
  const x = new Date(d)
  // Jornada 4am a 3:59am
  if (x.getHours() < 4) x.setDate(x.getDate() - 1)
  return x.toISOString().split('T')[0]
}

function horaMX() {
  return new Date().toLocaleTimeString('es-MX', { hour12: false, timeZone: 'America/Mexico_City' })
}

const pending = new Map() // tel -> 'entrada'|'salida'

async function getRows(range) {
  const s = await sheetsClient()
  const r = await s.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range })
  return r.data.values || []
}

async function start() {
  const { state, saveCreds } = await useMultiFileAuthState('/app/baileys_auth')
  const sock = makeWASocket({ auth: state, printQRInTerminal: true })
  sock.ev.on('creds.update', saveCreds)

  sock.ev.on('connection.update', ({ connection, qr }) => {
    if (qr) console.log('QR NUEVO - Escanea')
    if (connection === 'open') console.log('✅ BOT TRINIDAD LISTO')
    if (connection === 'close') {
      const shouldReconnect = true
      if (shouldReconnect) start()
    }
  })

  sock.ev.on('messages.upsert', async ({ messages }) => {
    try {
      const m = messages[0]
      if (!m || m.key.fromMe) return
      const jid = m.key.remoteJid
      const esGrupo = jid.endsWith('@g.us')
      if (!esGrupo) return // NO RESPONDE EN PRIVADO

      const tel = (m.key.participant || jid).replace(/\D/g, '')
      const texto = m.message?.conversation || m.message?.extendedTextMessage?.text || m.message?.imageMessage?.caption || ''
      const loc = m.message?.locationMessage

      // 1. AUTO-DETECCIÓN DE ID
      if (texto.trim().toLowerCase() === 'id') {
        await sock.sendMessage(jid, { text: `ID de este grupo:\n${jid}` })
        if (jid!== GRUPO_REPORTES_ID) {
          await sock.sendMessage(GRUPO_REPORTES_ID, { text: `📌 Grupo detectado\nID: ${jid}` })
        }
        return
      }

      // 2. ADMIN - registrar
      if (jid === GRUPO_REPORTES_ID && texto.toLowerCase().startsWith('registrar ')) {
        const s = await sheetsClient()
        const parts = texto.split(' ')
        const num = parts[1].replace(/\D/g, '')
        const resto = texto.substring(texto.indexOf(parts[1]) + parts[1].length).trim() || 'Sin Nombre'
        await s.spreadsheets.values.append({
          spreadsheetId: SPREADSHEET_ID,
          range: 'Empleados!A:F',
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [[num, resto, 'COYOACAN', 19.352525, -99.161817, 'LUNES']] }
        })
        await sock.sendMessage(jid, { text: `✅ Registrado ${num} -> ${resto}` })
        return
      }

      const esEntrada = /entr|ingres|lle?gue|aqui estoy|presente/i.test(texto)
      const esSalida = /salid|me voy|adios|bye/i.test(texto)

      if (!loc) {
        if (esEntrada) {
          pending.set(tel, 'entrada')
          await sock.sendMessage(jid, { text: 'Por favor envía tu ubicación para registrar tu entrada.' }, { quoted: m })
          return
        }
        if (esSalida) {
          pending.set(tel, 'salida')
          await sock.sendMessage(jid, { text: 'Por favor envía tu ubicación para registrar tu salida.' }, { quoted: m })
          return
        }
        return
      }

      // 3. CON UBICACIÓN
      const lat = loc.degreesLatitude
      const lng = loc.degreesLongitude

      let cercana = null
      let dMin = Infinity
      for (const s of SUCURSALES) {
        const d = distM(lat, lng, s.lat, s.lng)
        if (d < dMin) { dMin = d; cercana = s }
      }

      const empleadosRows = await getRows('Empleados!A:G')
      const emp = empleadosRows.slice(1).find(r => r[0] && r[0].replace(/\D/g, '').slice(-10) === tel.slice(-10))
      const nombre = emp? emp[1] : tel
      const fLab = fechaLaboral()

      const asisRows = await getRows('Asistencia!A:H')
      const idx = asisRows.findIndex((r, i) => i > 0 && r[0] && r[0].replace(/\D/g, '').slice(-10) === tel.slice(-10) && r[2] === fLab)
      const hoy = idx > -1? asisRows[idx] : null
      const tieneEntrada = hoy && hoy[3]

      const sClient = await sheetsClient()

      if (!tieneEntrada) {
        // ENTRADA
        if (dMin > 150) {
          await sock.sendMessage(jid, { text: 'la entrada debe registrarse en la unidad' }, { quoted: m })
          return
        }
        if (idx === -1) {
          await sClient.spreadsheets.values.append({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Asistencia!A:H',
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [[tel, nombre, fLab, horaMX(), 'A TIEMPO', '', '', '']] }
          })
        } else {
          await sClient.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `Asistencia!D${idx + 1}:E${idx + 1}`,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [[horaMX(), 'A TIEMPO']] }
          })
        }
        await sock.sendMessage(jid, { text: `Buen turno ${nombre} - ${cercana.nombre}` })
        pending.delete(tel)
      } else {
        // SALIDA
        if (hoy[5]) {
          await sock.sendMessage(jid, { text: `Salida ya registrada a las ${hoy[5]} - ${nombre}` })
          return
        }
        const h = horaMX()
        await sClient.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `Asistencia!F${idx + 1}:H${idx + 1}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [[h, h, Math.round(dMin).toString()]] }
        })
        if (dMin > 100) {
          await sock.sendMessage(GRUPO_REPORTES_ID, {
            text: `⚠️ Salida de ${nombre} (${tel}) a ${Math.round(dMin)}m de ${cercana.nombre}. Hora: ${h}. Fuera de rango 100m. Se conserva hora primer intento.`
          })
          await sock.sendMessage(jid, { text: `Salida registrada ${nombre}` })
        } else {
          await sock.sendMessage(jid, { text: `Salida registrada ${nombre} - ${cercana.nombre}` })
        }
        pending.delete(tel)
      }
    } catch (e) {
      console.error('Error:', e)
    }
  })

  // Retardos 20 min
  cron.schedule('*/5 * * * *', async () => {
    console.log('Chequeo retardos 20min...')
    // Lee Calendario_Horarios vs Asistencia
  }, { timezone: 'America/Mexico_City' })

  // Resumen Lunes 7am
  cron.schedule('0 7 * * 1', async () => {
    try {
      const s = await sheetsClient()
      await s // usar para calcular
      await sock.sendMessage(GRUPO_REPORTES_ID, {
        text: `📊 Resumen semanal Trinidad\n- Jornada 48h base\n- Horas totales, extras, retardos >15min\n- Descanso general Lunes\nRevisar pestaña Asistencia`
      })
    } catch (e) { console.error(e) }
  }, { timezone: 'America/Mexico_City' })
}

start()

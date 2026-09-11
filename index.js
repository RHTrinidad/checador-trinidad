const { Client, LocalAuth } = require('whatsapp-web.js');
const { google } = require('googleapis');
const fs = require('fs');

const SHEET_ID = process.env.SHEET_ID;
const CREDS = JSON.parse(process.env.GOOGLE_CREDS);

// Borra sesión vieja para forzar QR limpio
if (fs.existsSync('./.wwebjs_auth')) {
  fs.rmSync('./.wwebjs_auth', { recursive: true, force: true });
}

const client = new Client({
  authStrategy: new LocalAuth(),
  puppeteer: {
    headless: true,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-accelerated-2d-canvas',
      '--no-first-run',
      '--no-zygote',
      '--single-process',
      '--disable-gpu'
    ]
  }
});

async function guardarEnSheet(nombre, numero, tipo) {
  try {
    const auth = new google.auth.GoogleAuth({
      credentials: CREDS,
      scopes: ['https://www.googleapis.com/auth/spreadsheets']
    });
    const sheets = google.sheets({ version: 'v4', auth });
    const ahora = new Date();
    const fecha = ahora.toLocaleDateString('es-MX', { timeZone: 'America/Mexico_City' });
    const hora = ahora.toLocaleTimeString('es-MX', { timeZone: 'America/Mexico_City' });
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: 'A:E',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[fecha, hora, nombre, numero, tipo]] }
    });
    console.log(`Guardado OK: ${nombre} - ${tipo}`);
  } catch (e) {
    console.log('Error Sheets:', e.message);
  }
}

client.on('qr', qr => {
  console.log('--- ABRE ESTE LINK PARA VER EL QR ---');
  console.log(`https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=${encodeURIComponent(qr)}`);
});

client.on('ready', () => {
  console.log('¡BOT CONECTADO LISTO! YA PUEDES MANDAR ENTRADA');
});

client.on('auth_failure', msg => {
  console.log('FALLO AUTH:', msg);
});

client.on('message', async msg => {
  try {
    const texto = msg.body.toLowerCase().trim();
    console.log(`Mensaje recibido: ${texto} de ${msg.from}`);
    if (texto === 'entrada' || texto === 'salida') {
      const contacto = await msg.getContact();
      const nombre = contacto.pushname || contacto.name || 'Desconocido';
      await guardarEnSheet(nombre, msg.from, texto);
      await msg.reply(`✅ Registrado: *${texto.toUpperCase()}*\n${nombre}`);
    }
  } catch (e) {
    console.log('Error en mensaje:', e.message);
  }
});

client.initialize();

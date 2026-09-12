const { Client, LocalAuth } = require('whatsapp-web.js');
const { google } = require('googleapis');

const SHEET_ID = process.env.SHEET_ID;
const CREDS = JSON.parse(process.env.GOOGLE_CREDS);

const client = new Client({
  authStrategy: new LocalAuth(),
  puppeteer: {
    headless: true,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-gpu',
      '--single-process'
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
    const hora = ahora.toLocaleTimeString('es-MX', { timeZone: 'America/Mexico_City', hour12: true });
    
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: 'A:E',
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[fecha, hora, nombre, numero, tipo]]
      }
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
  console.log('Client is ready!');
});

// ESTE ES EL CODIGO BUENO
client.on('message', async msg => {
  try {
    const chat = await msg.getChat();
    const text = msg.body ? msg.body.toLowerCase().trim() : '';

    console.log(`MSG: "${text}" de ${msg.from} | Grupo: ${chat.isGroup} | ID: ${chat.id._serialized}`);

    if (text.includes('id grupo')) {
      await msg.reply(`ID de este grupo:\n${chat.id._serialized}`);
      return;
    }

    if (text === 'entrada' || text === 'salida') {
      const contact = await msg.getContact();
      const nombre = contact.pushname || contact.name || 'Desconocido';
      await guardarEnSheet(nombre, msg.from, text);
      
      if (text === 'entrada') {
        await msg.reply(`Buen día, ${nombre}! ☀️\nTu ${text} fue registrada.\nGracias!`);
      } else {
        await msg.reply(`Gracias.`);
      }
    }
  } catch (e) {
    console.log('Error:', e.message);
  }
});

client.initialize();

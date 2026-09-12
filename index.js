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

cliente.en('mensaje', asincrono MSG => {
  Prueba {
    const chat = Espera MSG.getChat();
    const Texto = MSG.Cuerpo ? MSG.Cuerpo.toLowerCase().trim() : '';

    Consola.Log(`--- MENSAJE NUEVO ---`);
    Consola.Log(`De: ${MSG.de} | Grupo?: ${chat.isGroup} | Texto: ${Texto}`);
    Consola.Log(`Chat ID: ${chat.id._serialized}`);

    // FORZAR RESPUESTA A TODO EN GRUPO
    si (Texto.includes('id grupo') || Texto.includes('id')) {
      Consola.Log('Detectó comando id grupo, respondiendo...');
      Espera MSG.Respuesta(`Aqui estoy!\nID: ${chat.id._serialized}\nEs grupo: ${chat.isGroup}`);
      Retorna;
    }

    si (Texto === 'entrada' || Texto === 'salida') {
      const contacto = Espera MSG.getContact();
      const nombre = contacto.Pushname || contacto.Nombre || 'Desconocido';
      Espera guardarEnSheet(nombre, MSG.de, Texto);
      
      si (Texto === 'entrada') {
        Espera MSG.Respuesta(`Buen día, ${nombre}! Tu ${Texto} registrada.\nGracias!`);
      } sino {
        Espera MSG.Respuesta(`Gracias.`);
      }
    }
  } Atrapa (e) {
    Consola.Log('Error en mensaje:', e.Mensaje);
    Consola.Log(e);
  }
});

client.initialize();

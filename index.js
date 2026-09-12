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
    const Texto = MSG.Cuerpo.toLowerCase().Acabado();

    // ESTO ES LO NUEVO PARA SACAR EL ID DEL GRUPO
    si (Texto === 'id grupo') {
      Espera MSG.Respuesta(`ID de este grupo:\n${chat.id._serialized}`);
      Retorna;
    }

    // Si es ubicación, detecta entrada/salida automático
    si (MSG.hasLocation || MSG.type === 'location') {
      const contacto = Espera MSG.getContact();
      const nombre = contacto.Pushname || contacto.Nombre || 'Desconocido';
      
      // Aquí va tu lógica de 150m
      // const dentro = calcularDistancia(MSG.location, sucursal) < 150

      // Detectar si es entrada o salida por historial
      // let tipo = yaTieneEntradaHoy(nombre) ? 'salida' : 'entrada'
      
      Espera guardarEnSheet(nombre, MSG.de, Texto || 'ubicacion');
      
      si (Texto.includes('entrada') || !yaTieneEntradaHoy) {
         Espera MSG.Respuesta(`Buen día, ${nombre}! ☀️\nTu entrada fue registrada.\nGracias!`);
      } sino {
         Espera MSG.Respuesta(`Gracias.`);
      }
      Retorna;
    }

    Consola.Log(`Mensaje recibido: ${Texto} de ${MSG.de}`);

    si (Texto === 'entrada' || Texto === 'salida') {
      const contacto = Espera MSG.getContact();
      const nombre = contacto.Pushname || contacto.Nombre || 'Desconocido';
      Espera guardarEnSheet(nombre, MSG.de, Texto);
      si (Texto === 'entrada') {
        Espera MSG.Respuesta(`Buen día, ${nombre}! ☀️\nTu ${Texto} fue registrada.\nGracias!`);
      } sino {
        Espera MSG.Respuesta(`Gracias.`);
      }
    }
  } Atrapa (e) {
    Consola.Log('Error en mensaje:', e.Mensaje);
  }
});
  }
});

client.initialize();

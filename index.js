const { Client, LocalAuth } = require('whatsapp-web.js');
const { google } = require('googleapis');
const cron = require('node-cron');
const config = require('./config.js');

const SHEET_ID = process.env.SHEET_ID;
const CREDS = JSON.parse(process.env.GOOGLE_CREDS);

const SUCURSALES = config.Sucursales;
const GRUPO_REPORTES = "120363412984528459@g.us"; // tu grupo de reportes

const client = new Client({
  authStrategy: new LocalAuth(),
  puppeteer: {
    headless: true,
    args: ['--no-sandbox','--disable-setuid-sandbox','--disable-dev-shm-usage','--disable-gpu','--single-process']
  }
});

function distanciaMts(lat1, lon1, lat2, lon2) {
  const R = 6371e3;
  const dLat = (lat2-lat1)*Math.PI/180;
  const dLon = (lon2-lon1)*Math.PI/180;
  const a = Math.sin(dLat/2)**2 + Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)**2;
  return R*2*Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}
function getJornadaFecha() {
  let ahora = new Date();
  let h = new Date().toLocaleString("en-US", {timeZone: "America/Mexico_City", hour: 'numeric', hour12: false});
  if (parseInt(h) < 5) ahora.setDate(ahora.getDate() - 1);
  return ahora.toLocaleDateString('es-MX', { timeZone: 'America/Mexico_City' });
}
function getHoraCDMX() {
  return new Date().toLocaleTimeString('es-MX', { timeZone: 'America/Mexico_City', hour12: false });
}

let esperandoUbicacion = {};
let intentosFuera = {};

async function guardarEnSheet(nombre, numero, tipo, sucursal='', nota='') {
  try {
    const auth = new google.auth.GoogleAuth({ credentials: CREDS, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: 'A:G',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[getJornadaFecha(), getHoraCDMX(), nombre, numero, tipo, sucursal, nota]] }
    });
    console.log(`Guardado: ${nombre} ${tipo} ${sucursal}`);
  } catch (e) { console.log('Error Sheets:', e.message); }
}

client.on('qr', qr => console.log(`QR: https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=${encodeURIComponent(qr)}`));
client.on('ready', () => console.log('Client is ready!'));

client.on('message', async msg => {
  try {
    // 1. Si es ubicación
    if (msg.location) {
      const ctx = esperandoUbicacion[msg.from];
      if (!ctx) return;
      const suc = SUCURSALES[ctx.sucursal];
      const dist = distanciaMts(msg.location.latitude, msg.location.longitude, suc.lat, suc.lng);
      console.log(`Distancia: ${dist.toFixed(0)}m a ${suc.nombre}`);

      if (ctx.tipo === 'entrada') {
        if (dist > suc.radioEntrada) {
          await msg.reply(`❌ La entrada se registra en la sucursal, intenta de nuevo en la unidad\nEstás a ${dist.toFixed(0)}m de ${suc.nombre}`);
          return;
        }
        await guardarEnSheet(ctx.nombre, msg.author || msg.from, 'entrada', ctx.sucursal, `${dist.toFixed(0)}m`);
        await msg.reply(`✅ Entrada registrada en ${suc.nombre}\nGracias, buen turno!`);
        delete esperandoUbicacion[msg.from];
      }
      if (ctx.tipo === 'salida') {
        if (dist > suc.radioSalidaLimite) {
          const key = `${msg.from}_${getJornadaFecha()}_salida`;
          if (!intentosFuera[key]) {
            intentosFuera[key] = true;
            await guardarEnSheet(ctx.nombre, msg.author || msg.from, 'intento_salida_fuera', ctx.sucursal, `${dist.toFixed(0)}m ${getHoraCDMX()}`);
          }
          await msg.reply(`❌ La salida se registra en la sucursal, intenta de nuevo en la unidad\nEstás a ${dist.toFixed(0)}m`);
          return;
        }
        if (dist > suc.radioSalidaAviso) {
          await guardarEnSheet(ctx.nombre, msg.author || msg.from, 'salida', ctx.sucursal, `Aviso ${dist.toFixed(0)}m`);
          await msg.reply(`✅ Salida registrada (a ${dist.toFixed(0)}m)\nLa salida es a 100 metros de la sucursal\nGracias!`);
        } else {
          await guardarEnSheet(ctx.nombre, msg.author || msg.from, 'salida', ctx.sucursal, `${dist.toFixed(0)}m`);
          await msg.reply(`✅ Salida registrada en ${suc.nombre}\nGracias!`);
        }
        delete esperandoUbicacion[msg.from];
      }
      return;
    }

    // 2. Si es texto
    const text = msg.body? msg.body.toLowerCase().trim() : '';
    if (!text) return;
    console.log(`NUEVO: "${text}" de ${msg.from}`);

    if (text.includes('id grupo')) {
      await msg.reply(`ID de este grupo:\n${msg.from}`);
      return;
    }

    if (text.includes('entrada') || text.includes('salida')) {
      let tipo = text.includes('entrada')? 'entrada' : 'salida';
      let sucursal = 'bucareli';
      if (text.includes('coyo') || text.includes('mercado') || text.includes('86')) sucursal = 'coyoacan';
      // Si lo escribe desde el grupo de la sucursal, detectar por ID
      if (config.GruposSucursales && config.GruposSucursales[msg.from]) {
        sucursal = config.GruposSucursales[msg.from];
      }
      const nombre = msg._data.notifyName || 'Desconocido';
      esperandoUbicacion[msg.from] = { tipo, sucursal, nombre };
      await msg.reply(`📍 Comparte tu ubicación para registrar tu ${tipo} en ${SUCURSALES[sucursal].nombre} (máx 150m)`);
    }
  } catch (e) { console.log('Error msg:', e.message); }
});

// REPORTE LUNES 8AM
cron.schedule('0 8 * * 1', async () => {
  try {
    const auth = new google.auth.GoogleAuth({ credentials: CREDS, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });
    const res = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'A:G' });
    const rows = res.data.values || [];
    let texto = `📊 REPORTE SEMANAL - Jornada 5am a 4:59am | Lunes descanso\n`;
    texto += `Semana: Mar-Dom | Regular: 48h\n\n`;
    for (let sKey in SUCURSALES) {
      texto += `📍 ${SUCURSALES[sKey].nombre}:\n`;
      // aquí tu lógica de horas la puedes expandir, por ahora conteo
      let conteo = rows.filter(r => r[5] === sKey).length;
      texto += `Registros: ${conteo}\n\n`;
    }
    await client.sendMessage(GRUPO_REPORTES, texto);
  } catch(e){ console.log('Error reporte', e.message); }
}, { timezone: "America/Mexico_City" });

client.initialize();

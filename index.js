const { Client, LocalAuth } = require('whatsapp-web.js');
const { google } = require('googleapis');

const client = new Client({
  authStrategy: new LocalAuth(),
  puppeteer: { args: ['--no-sandbox', '--disable-setuid-sandbox'] }
});

client.on('qr', qr => {
  console.log('--- COPIA ESTE LINK Y ABRELO EN OTRO NAVEGADOR PARA VER EL QR ---');
  console.log(`https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=${encodeURIComponent(qr)}`);
  console.log('--- FIN DEL LINK ---');
});

client.on('ready', () => {
  console.log('¡BOT CONECTADO LISTO!');
});

client.on('message', async msg => {
  const texto = msg.body.toLowerCase();
  if(texto === 'entrada' || texto === 'salida'){
    msg.reply(`Registrado: ${texto} a las ${new Date().toLocaleTimeString('es-MX')}`);
    // aqui luego metemos lo de Google Sheets
  }
});

client.initialize();

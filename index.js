const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');

const client = new Client({
  authStrategy: new LocalAuth(),
  puppeteer: {
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  }
});

client.on('qr', (qr) => {
  console.log('ESCANEA EL QR:');
  qrcode.generate(qr, { small: true });
});

client.on('ready', () => {
  console.log('¡BOT CONECTADO!');
});

client.on('message', async msg => {
  // tu codigo del checador aqui
  console.log(`Mensaje de ${msg.from}: ${msg.body}`);
});

client.initialize();

const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const { google } = require('googleapis');

const client = new Client({
  authStrategy: new LocalAuth(),
  puppeteer: { args: ['--no-sandbox', '--disable-setuid-sandbox'] }
});

client.on('qr', qr => {
  console.log('ESCANEA ESTE QR AHORA:');
  qrcode.generate(qr, {small: true});
});

client.on('ready', () => {
  console.log('BOT CONECTADO LISTO!');
});

client.on('message', async msg => {
  if(msg.body.toLowerCase() === 'entrada' || msg.body.toLowerCase() === 'salida') {
    // Aqui va tu logica de sheets
    msg.reply(`Registrado: ${msg.body} a las ${new Date().toLocaleTimeString()}`);
  }
});

client.initialize();

const { Client, LocalAuth } = require('whatsapp-web.js');
const { google } = require('googleapis');
const configFile = require('./config.js');

const SHEET_ID = process.env.SHEET_ID;
const CREDS = JSON.parse(process.env.GOOGLE_CREDS);

const SUCURSALES = configFile.Sucursales || configFile.sucursales;
const client = new Client({
  authStrategy: new LocalAuth(),
  puppeteer: {
    headless: true,
    executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || '/usr/bin/chromium',
    args: ['--no-sandbox','--disable-setuid-sandbox','--disable-dev-shm-usage','--disable-gpu']
  }
});

function dist(lat1,lon1,lat2,lon2){
  const R=6371000;
  const dLat=(lat2-lat1)*Math.PI/180;
  const dLon=(lon2-lon1)*Math.PI/180;
  const a=Math.sin(dLat/2)**2+Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)**2;
  return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
}

let espera={};
client.on('qr', qr => console.log('QR: https://api.qrserver.com/v1/create-qr-code/?size=400x400&data='+encodeURIComponent(qr)));
client.on('ready', () => console.log('✅ BOT LISTO'));

client.on('message', async msg=>{
  try{
    if(msg.location){
      const ctx=espera[msg.from];
      if(!ctx) return;
      const key = Object.keys(SUCURSALES).find(k=>k.toLowerCase()===ctx.suc.toLowerCase());
      const suc=SUCURSALES[key];
      const sLat=suc.lat||suc.Lat;
      const sLng=suc.lng||suc.LNG||suc.Lng;
      const d=dist(msg.location.latitude, msg.location.longitude, sLat, sLng);
      console.log(`Distancia ${d}m a ${suc.nombre}`);

      if(ctx.tipo==='entrada' && d>150){
        await msg.reply(`❌ La entrada se registra en la sucursal, intenta de nuevo en la unidad\nEstás a ${d.toFixed(0)}m`);
        return;
      }
      if(ctx.tipo==='salida' && d>150){
        await msg.reply(`❌ La salida se registra en la sucursal, intenta de nuevo en la unidad\nEstás a ${d.toFixed(0)}m`);
        return;
      }
      if(ctx.tipo==='salida' && d>100){
        await msg.reply(`✅ Salida registrada a ${d.toFixed(0)}m\nLa salida es a 100 metros de la sucursal\nGracias!`);
      }else{
        await msg.reply(`✅ ${ctx.tipo} registrada en ${suc.nombre}\nGracias, buen turno!`);
      }
      delete espera[msg.from];
      return;
    }

    const text=(msg.body||'').toLowerCase();
    if(text.includes('id grupo')){
      await msg.reply(`ID: ${msg.from}`);
      return;
    }
    if(text.includes('entrada') || text.includes('salida')){
      let tipo=text.includes('entrada')?'entrada':'salida';
      let suc='bucareli';
      if(text.includes('coyo')) suc='coyoacan';
      if(text.includes('bucareli')) suc='bucareli';
      const nombre=msg._data.notifyName||'Desconocido';
      espera[msg.from]={tipo,suc,nombre};
      await msg.reply(`📍 Comparte ubicación para ${tipo} en ${suc.toUpperCase()} (máx 150m)`);
    }
  }catch(e){ console.log('Error:', e.message); }
});
client.initialize();

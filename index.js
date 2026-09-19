import { default as makeWASocket, useMultiFileAuthState, DisconnectReason } from '@whiskeysockets/baileys'
import qrcodeTerminal from 'qrcode-terminal'
import QRCode from 'qrcode'
import express from 'express'
import { createRequire } from 'module'
const require = createRequire(import.meta.url)
const { google } = require('googleapis')

const SPREADSHEET_ID = process.env.SPREADSHEET_ID
const GRUPO_REPORTES_ID = process.env.GRUPO_REPORTES_ID

let SUCURSALES = []
try {
  if (process.env.SUCURSALES_JSON) {
    const parsed = JSON.parse(process.env.SUCURSALES_JSON)
    const arr = Array.isArray(parsed)? parsed : Object.values(parsed)
    SUCURSALES = arr.map(s => ({
      id: s.id || s.nombre,
      nombre: s.nombre || s.id,
      lat: s.lat,
      lng: s.lng || s.lon,
      rEnt: s.radio_entrada || s.rEnt || s.radio || 250,
      rSal: s.radio_salida || s.rSal || s.radio || 250
    }))
  } else throw new Error('no json')
} catch (e) {
  SUCURSALES = [
    { id:"COYOACAN", nombre:"Trinidad Coyoacan", lat:19.352525, lng:-99.161817, rEnt:250, rSal:250 },
    { id:"BUCARELI", nombre:"Trinidad Bucareli", lat:19.426523, lng:-99.153326, rEnt:250, rSal:250 },
    { id:"JUAREZ", nombre:"Trinidad Juarez", lat:19.4314119, lng:-99.1512074, rEnt:250, rSal:250 }
  ]
}
console.log('Sucursales:', SUCURSALES)

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets']
})
async function sheetsClient(){ const c=await auth.getClient(); return google.sheets({version:'v4',auth:c}) }
function distM(a,b,c,d){ const R=6371000, toRad=x=>x*Math.PI/180; const dLa=toRad(c-a), dLo=toRad(d-b); const q=Math.sin(dLa/2)**2+Math.cos(toRad(a))*Math.cos(toRad(c))*Math.sin(dLo/2)**2; return R*2*Math.atan2(Math.sqrt(q),Math.sqrt(1-q)) }
function fechaLaboral(d=new Date()){ const x=new Date(d); if(x.getHours()<4) x.setDate(x.getDate()-1); return x.toISOString().split('T')[0] }
function horaMX(d=new Date()){ return d.toLocaleTimeString('es-MX',{hour12:false,timeZone:'America/Mexico_City'}) }
function minutos(h){ const mm = h?.toString().match(/(\d{1,2}):(\d{2})/); return mm? parseInt(mm[1])*60+parseInt(mm[2]) : 0 }
function parseFechaMX(s){
  if(!s) return null
  if(/^\d{4}-\d{2}-\d{2}$/.test(s)) return new Date(s+'T12:00:00')
  const m=s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/)
  if(m){ return new Date(`${m[3].length==2?'20'+m[3]:m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}T12:00:00`) }
  return new Date(s)
}
async function getRows(range){ const s=await sheetsClient(); const r=await s.spreadsheets.values.get({spreadsheetId:SPREADSHEET_ID,range}); return r.data.values||[] }

async function getDescansosMap(){
  try{
    const rows=await getRows('Horario_Base!A2:J')
    const map={}
    for(const f of rows){
      const nombre=(f[1]||'').toLowerCase().trim()
      if(!nombre) continue
      const descansos=new Set()
      const dias=[f[3],f[4],f[5],f[6],f[7],f[8],f[9]] // Lun a Dom
      const numDia=[1,2,3,4,5,6,0]
      dias.forEach((val,i)=>{
        if((val||'').toString().toLowerCase().includes('descanso')) descansos.add(numDia[i])
      })
      map[nombre]=descansos
    }
    return map
  }catch(e){ console.log('err descansos',e); return {} }
}

async function resumenEmpleado(nombreBuscar, jidRespuesta, sock){
  const sClient=await sheetsClient()
  const [asisRes, descansosMap]=await Promise.all([
    sClient.spreadsheets.values.get({spreadsheetId:SPREADSHEET_ID, range:'Asistencia!A2:H'}),
    getDescansosMap()
  ])
  const filas=asisRes.data.values||[]
  const buscar=nombreBuscar.toLowerCase().trim()

  const ahoraMX=new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'}))
  ahoraMX.setHours(0,0,0,0)
  const hace30=new Date(ahoraMX); hace30.setDate(ahoraMX.getDate()-30)
  const hace14=new Date(ahoraMX); hace14.setDate(ahoraMX.getDate()-14)
  const diaSem=ahoraMX.getDay()
  const lunes=new Date(ahoraMX); lunes.setDate(ahoraMX.getDate() - (diaSem===0?6:diaSem-1))
  const domingo=new Date(lunes); domingo.setDate(lunes.getDate()+6)

  let descEmpleado=new Set()
  let nombreReal=nombreBuscar
  for(const k in descansosMap){ if(k.includes(buscar) || buscar.includes(k)){ descEmpleado=descansosMap[k]; nombreReal=k; break } }

  let dias30=0, diasSem=0, retardosSem=0, minSem=0, retardos14=0, min14=0
  let detalleSem=[], sinSalidaSem=[]

  for(const f of filas){
    const nombre=(f[1]||'').toString().toLowerCase()
    const tel=(f[0]||'').toString()
    if(!nombre.includes(buscar) &&!tel.includes(buscar)) continue
    const fecha=parseFechaMX(f[2]); if(!fecha) continue; fecha.setHours(0,0,0,0)
    const tieneEntrada=!!f[3]
    const tieneSalida=!!f[5]
    const retMatch=(f[4]||'').toString().match(/(\d+)\s*min/)
    const ret=retMatch?parseInt(retMatch[1]):0

    if(fecha>=hace30 && tieneEntrada) dias30++
    if(fecha>=hace14 && ret>0){ retardos14++; min14+=ret }
    if(fecha>=lunes && fecha<=domingo){
      if(tieneEntrada){
        diasSem++
        const flagSalida = tieneSalida? `Sal:${f[5]}` : `⚠️ SIN SALIDA (se cuentan 8h)`
        detalleSem.push(`• ${f[2]}: Ent ${f[3]} | ${flagSalida} ${ret>0?`RET ${ret}m`:''}`)
        if(!tieneSalida) sinSalidaSem.push(f[2])
      }
      if(ret>0){ retardosSem++; minSem+=ret }
    }
  }

  let esperados30=0
  for(let d=new Date(hace30); d<=ahoraMX; d.setDate(d.getDate()+1)){ if(!descEmpleado.has(d.getDay())) esperados30++ }
  let esperadosSem=0
  for(let d=new Date(lunes); d<=ahoraMX; d.setDate(d.getDate()+1)){ if(!descEmpleado.has(d.getDay())) esperadosSem++ }

  const faltas30=Math.max(0, esperados30 - dias30)
  const faltasSem=Math.max(0, esperadosSem - diasSem)
  const diasNom=['Dom','Lun','Mar','Mie','Jue','Vie','Sab']
  const descTxt=[...descEmpleado].map(d=>diasNom[d]).join(', ')||'ninguno'

  const txt=`📊 *${nombreBuscar.toUpperCase()}*
Descansos: ${descTxt}
Semana ${lunes.toLocaleDateString('es-MX')} al ${domingo.toLocaleDateString('es-MX')}

*SEMANA EN CURSO*
- Trabajados: ${diasSem}/${esperadosSem}
- Faltas semana: ${faltasSem} (ya sin contar descansos)
- Retardos semana: ${retardosSem} (${minSem} min)

*ACUMULADO 30 DÍAS NATURALES*
${hace30.toLocaleDateString('es-MX')} al hoy
- Esperados (sin descansos): ${esperados30}
- Trabajados: ${dias30}
- FALTAS 30d: ${faltas30}

*RETARDOS 14 DÍAS*
- ${retardos14} retardos (${min14} min)

*CONTROL DE SALIDAS*
- Días sin checar salida esta semana: ${sinSalidaSem.length>0? sinSalidaSem.join(', ') : 'Ninguno ✅'}
${sinSalidaSem.length>0?'Nota: Esos días se contabilizan como 8h trabajadas pero quedan en registro.':''}

Detalle semana:
${detalleSem.join('\n')||'Sin registros esta semana'}`

  await sock.sendMessage(jidRespuesta,{text:txt})
}

const app = express()
let lastQR = null
app.get('/', (req,res)=> res.send('Bot Trinidad OK - ve a /qr'))
app.get('/qr', async (req,res)=>{
  if(!lastQR) return res.send('<h2>Aun no hay QR, espera 10s</h2>')
  const dataUrl = await QRCode.toDataURL(lastQR)
  res.send(`<div style="text-align:center"><h2>Escanea</h2><img src="${dataUrl}" style="width:350px"><p>Recarga si expira</p></div>`)
})
app.listen(process.env.PORT||3000, ()=> console.log('Web en puerto '+(process.env.PORT||3000)))

async function start(){
  console.log('--- INICIANDO BOT TRINIDAD ---')
  const { state, saveCreds } = await useMultiFileAuthState('/app/auth')
  const sock=makeWASocket({
    auth:state,
    printQRInTerminal:false,
    markOnlineOnConnect:false,
    syncFullHistory:false,
    generateHighQualityLinkPreview:false
  })
  sock.ev.on('creds.update', saveCreds)
  sock.ev.on('connection.update', async ({connection, lastDisconnect, qr}) => {
    if (qr) { lastQR = qr; qrcodeTerminal.generate(qr,{small:false}); console.log('QR nuevo') }
    if (connection === 'close') {
      const shouldReconnect = lastDisconnect?.error?.output?.statusCode!== DisconnectReason.loggedOut
      if (shouldReconnect) start()
    }
    if (connection === 'open') { console.log('✅ BOT CONECTADO'); lastQR=null }
  })

  sock.ev.on('messages.upsert', async ({messages})=>{
    try{
      const m=messages[0]; if(!m||m.key.fromMe) return
      const jid=m.key.remoteJid; if(!jid.endsWith('@g.us')) return

      const rawLid = m.key.participant || ''
      const realPn = m.key.participantPn || m.key.participantAlt || ''
      let pnFromStore = ''
      try{ pnFromStore = await sock.signalRepository?.lidMapping?.getPNForLID(rawLid) || '' }catch{}
      const rawId = realPn || pnFromStore || rawLid || jid
      const tel = rawId.replace(/\D/g,'')

      const texto=m.message?.conversation||m.message?.extendedTextMessage?.text||m.message?.imageMessage?.caption||''
      const loc=m.message?.locationMessage

      console.log(`FROM LID:${rawLid} PN:${realPn||pnFromStore} FINAL_TEL:${tel} TXT:${texto.slice(0,30)}`)

      if(texto.trim().toLowerCase()==='id'){ await sock.sendMessage(jid,{text:`ID: ${jid}`}); return }

      // COMANDO RESUMEN
      if(texto.toLowerCase().startsWith('resumen ')){
        const nombreBuscar=texto.slice(8).trim()
        if(!nombreBuscar){ await sock.sendMessage(jid,{text:'Uso: resumen Daniel R'}); return }
        await resumenEmpleado(nombreBuscar, jid, sock)
        return
      }

      const esEntrada=/entr|ingres|lle?gue|aqui estoy|presente/i.test(texto)
      const esSalida=/salid|me voy|adios|bye/i.test(texto)
      if(!loc){ if(esEntrada) await sock.sendMessage(jid,{text:'Envía tu ubicación para entrada'},{quoted:m}); if(esSalida) await sock.sendMessage(jid,{text:'Envía tu ubicación para salida'},{quoted:m}); return }

      const lat=loc.degreesLatitude, lng=loc.degreesLongitude
      let cercana=null, dMin=Infinity; for(const s of SUCURSALES){ const d=distM(lat,lng,s.lat,s.lng); if(d<dMin){dMin=d; cercana=s} }

      const empRows=await getRows('Empleados!A:G')
      let emp = empRows.slice(1).find(r=>r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10))
      if(!emp && rawLid.includes('@lid')){ emp = empRows.slice(1).find(r=> r[1] && m.pushName && r[1].toLowerCase().includes(m.pushName.toLowerCase().split(' ')[0])) }
      const nombre=emp?emp[1]: (m.pushName || tel)
      const fLab=fechaLaboral()
      const asisRows=await getRows('Asistencia!A:H'); const idx=asisRows.findIndex((r,i)=>i>0&&r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10)&&r[2]===fLab)
      const hoy=idx>-1?asisRows[idx]:null; const sClient=await sheetsClient()

      let estatus='A TIEMPO'
      let horaProg = null
      try{
        const cal=await getRows('Calendario_Horarios!A:E')
        const progExc=cal.slice(1).find(r=>r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10)&&r[2]===fLab)
        if(progExc && progExc[3]) horaProg = progExc[3]
        else {
          const baseRows = await getRows('Horario_Base!A:J')
          const baseRow = baseRows.slice(1).find(r=>r[0]&&r[0].replace(/\D/g,'').slice(-10)===tel.slice(-10))
          if(baseRow){
            const fecha = new Date(fLab+'T12:00:00')
            const mapa = {1:3, 2:4, 3:5, 4:6, 5:7, 6:8, 0:9}
            const valor = baseRow[mapa[fecha.getDay()]] || ''
            if(valor.toLowerCase().includes('descanso')) estatus='DESCANSO'
            else horaProg = valor
          }
        }
        if(horaProg){
          const hp = horaProg.toString().match(/\d{1,2}:\d{2}/)?.[0] || horaProg
          const dif = minutos(horaMX()) - minutos(hp)
          if(dif>15) estatus=`RETARDO ${dif}min (Prog ${hp})`
          else estatus=`A TIEMPO Prog ${hp}`
        }
      }catch(e){ console.log('err horario',e) }

      if(!hoy ||!hoy[3]){
        // ENTRADA
        if(dMin>cercana.rEnt){ await sock.sendMessage(jid,{text:`Debes estar a max ${cercana.rEnt}m de ${cercana.nombre}, estas a ${Math.round(dMin)}m`},{quoted:m}); return }
        if(idx===-1) await sClient.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Asistencia!A:H',valueInputOption:'USER_ENTERED',requestBody:{values:[[tel,nombre,fLab,horaMX(),estatus,'','','']]}})
        else await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!D${idx+1}:E${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[[horaMX(),estatus]]}})
        await sock.sendMessage(jid,{text:`✅ ${estatus} - ${nombre} en ${cercana.nombre}`})
      }else{
        // SALIDA - NUEVA VALIDACIÓN 150m
        if(hoy[5]){ await sock.sendMessage(jid,{text:`Salida ya registrada a las ${hoy[5]}`}); return }
        const LIMITE_SALIDA = 150
        if(dMin>LIMITE_SALIDA){
          await sock.sendMessage(jid,{text:`❌ No puedes checar salida a ${Math.round(dMin)}m de ${cercana.nombre}. Debes estar a max ${LIMITE_SALIDA}m. Acércate a la sucursal.`},{quoted:m})
          return
        }
        const h=horaMX()
        // Si no checó salida antes, se guarda la salida normal. Si nunca checa, el reporte lo cuenta como 8h
        await sClient.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Asistencia!F${idx+1}:H${idx+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[[h,h,Math.round(dMin).toString()]]}})
        await sock.sendMessage(jid,{text:`Salida registrada ${nombre} - ${cercana.nombre} (${Math.round(dMin)}m)`})
      }
    }catch(e){ console.error(e) }
  })
}
start()

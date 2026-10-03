import makeWASocket,{useMultiFileAuthState,DisconnectReason,fetchLatestBaileysVersion} from '@whiskeysockets/baileys'
import pino from 'pino'
import qrcode from 'qrcode-terminal'
import cron from 'node-cron'

import {
  GRUPOS,
  getTipoGrupo,
  GRUPO_REPORTES_TRINIDAD_ID,
  GRUPO_PRUEBAS_ID,
  GRUPO_COYOACAN_ID,
  GRUPO_BUCARELI_ID
} from './src/config.js'

import {
  handleChecador,
  registrarDescansos,
  cerrarSalidasPendientes,
  checkNoLlegaron
} from './src/checador.js'

import {handleReportes} from './src/reportes.js'
import {procesarCompra} from './src/compras.js'

const AUTH_DIR='/app/auth'

const COMANDOS_REPORTES=/^(info|datos|cuenta bancaria|cuenta empleado|datos bancarios|cuenta|banco|clave|clabe|asistencia hoy|resumen|reporte|compras de hoy|faltas|retardos|críticos|criticos|graves|excel|reporte semanal|reporte mensual)\b/i

const sesion={
  sock:null,
  conectado:false
}

let iniciando=false

function textoMensaje(m){
  return (
    m.message?.conversation ||
    m.message?.extendedTextMessage?.text ||
    m.message?.imageMessage?.caption ||
    m.message?.videoMessage?.caption ||
    ''
  ).trim()
}

function obtenerLoc(m){
  const loc=m.message?.locationMessage
  if(!loc)return null

  return {
    degreesLatitude:loc.degreesLatitude,
    degreesLongitude:loc.degreesLongitude
  }
}

function obtenerRawLid(m){
  return (
    m.key?.participant ||
    m.key?.remoteJid ||
    ''
  ).toString()
}

function obtenerTelefono(m){
  const jid=(
    m.key?.participant ||
    m.key?.remoteJid ||
    ''
  ).toString()

  const limpio=jid.split(':')[0].split('@')[0].replace(/\D/g,'')

  return limpio
}

function esGrupo(jid){
  return jid?.endsWith('@g.us')
}

function getFiltro(jid){
  if(jid===GRUPO_COYOACAN_ID)return 'coyoacan'
  if(jid===GRUPO_BUCARELI_ID)return 'juarez'
  return null
}

async function ejecutarProcesosAutomaticos(){
  if(!sesion.sock||!sesion.conectado)return

  try{
    await registrarDescansos()
  }catch(e){
    console.error('Error registrarDescansos:',e)
  }

  try{
    await cerrarSalidasPendientes()
  }catch(e){
    console.error('Error cerrarSalidasPendientes:',e)
  }

  try{
    await checkNoLlegaron(sesion.sock)
  }catch(e){
    console.error('Error checkNoLlegaron:',e)
  }
}

async function conectar(){

  if(iniciando)return
  iniciando=true

  try{

    const {state,saveCreds}=await useMultiFileAuthState(AUTH_DIR)

    console.log('AUTH DIR:',AUTH_DIR)

    let version

    try{
      const latest=await fetchLatestBaileysVersion()
      version=latest.version
      console.log('Baileys version:',version)
    }catch{
      console.log('No se pudo obtener version de Baileys, usando default')
    }

    const sock=makeWASocket({
      auth:state,
      version,
      logger:pino({level:'silent'}),
      printQRInTerminal:false,
      browser:['RH Trinidad','Chrome','1.0.0'],
      markOnlineOnConnect:false,
      syncFullHistory:false,
      generateHighQualityLinkPreview:false
    })

    sesion.sock=sock

    sock.ev.on('creds.update',saveCreds)

    sock.ev.on('connection.update',async(update)=>{

      const {connection,lastDisconnect,qr}=update

      if(qr){
        console.log('ESCANEA ESTE QR:')
        qrcode.generate(qr,{small:true})
      }

      if(connection==='open'){

        console.log('CONECTADO MODULAR + VOLUME OK')

        sesion.conectado=true
        iniciando=false

        try{
          await ejecutarProcesosAutomaticos()
        }catch(e){
          console.error('Error procesos iniciales:',e)
        }

        return
      }

      if(connection==='close'){

        sesion.conectado=false
        sesion.sock=null

        const code=lastDisconnect?.error?.output?.statusCode

        console.log('Conexion cerrada. Codigo:',code)

        if(code===DisconnectReason.loggedOut){

          console.error('Sesion cerrada / logout. Se requiere volver a vincular.')

          iniciando=false

        }else{

          iniciando=false

          setTimeout(()=>{
            conectar().catch(e=>{
              console.error('Error reconectando:',e)
            })
          },5000)
        }
      }
    })

    sock.ev.on('messages.upsert',async({messages,type})=>{

      if(type!=='notify')return

      for(const m of messages){

        try{

          if(!m.message)continue
          if(m.key?.fromMe)continue

          const jid=m.key?.remoteJid

          if(!jid)continue

          const tipoGrupo=getTipoGrupo(jid)

          const texto=textoMensaje(m)
          const textoLow=texto.toLowerCase().trim()

          const loc=obtenerLoc(m)

          const rawLid=obtenerRawLid(m)
          const tel=obtenerTelefono(m)
          const tel10=tel.slice(-10)

          // ==========================================
          // GRUPOS CHECADOR
          // ==========================================

          if(tipoGrupo==='CHECADORES'){

            // Los grupos checadores solo procesan ubicaciones.
            if(!loc)continue

            await handleChecador({
              sock,
              jid,
              m,
              loc,
              rawLid,
              tel10,
              tel
            })

            continue
          }

          // ==========================================
          // COMPRAS / IMAGENES
          // ==========================================

          if(
            tipoGrupo==='REPORTES' ||
            tipoGrupo==='GERENTES'
          ){

            const esImagen=!!m.message?.imageMessage

            if(esImagen){

              try{
                await procesarCompra({
                  sock,
                  jid,
                  m
                })
              }catch(e){
                console.error('Error procesando compra:',e)
              }

              continue
            }
          }

          // ==========================================
          // REPORTES / GERENTES / PRUEBAS
          // ==========================================

          if(
            tipoGrupo==='REPORTES' ||
            tipoGrupo==='GERENTES'
          ){

            if(COMANDOS_REPORTES.test(textoLow)){

              try{

                await handleReportes({
                  sock,
                  jid,
                  m,
                  texto,
                  filtroGrupo:getFiltro(jid)
                })

              }catch(e){

                console.error('Error handleReportes:',e)

                try{
                  await sock.sendMessage(
                    jid,
                    {text:'⚠️ Ocurrió un error al generar el reporte.'},
                    {quoted:m}
                  )
                }catch{}
              }

              continue
            }
          }

        }catch(e){

          console.error('Error procesando mensaje:',e)

        }
      }
    })

  }catch(e){

    console.error('Error iniciando WhatsApp:',e)

    sesion.conectado=false
    sesion.sock=null
    iniciando=false

    setTimeout(()=>{
      conectar().catch(err=>{
        console.error('Error en reconexión:',err)
      })
    },5000)
  }
}


// =====================================================
// PROCESOS AUTOMÁTICOS CADA 10 MINUTOS
// =====================================================

cron.schedule('*/10 * * * *',async()=>{

  if(!sesion.conectado||!sesion.sock)return

  console.log('⏱️ Ejecutando procesos automáticos...')

  await ejecutarProcesosAutomaticos()

})


// =====================================================
// INICIAR BOT
// =====================================================

console.log('🚀 Iniciando RH Trinidad...')

conectar().catch(e=>{
  console.error('Error inicial:',e)
})

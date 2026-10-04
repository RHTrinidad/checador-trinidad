import makeWASocket,{useMultiFileAuthState,DisconnectReason,fetchLatestBaileysVersion}from '@whiskeysockets/baileys'
import pino from 'pino'
import qrcode from 'qrcode-terminal'
import cron from 'node-cron'

import {getTipoGrupo,GRUPO_COYOACAN_ID,GRUPO_BUCARELI_ID} from './src/config.js'
import {handleChecador,registrarDescansos,cerrarSalidasPendientes,checkNoLlegaron} from './src/checador.js'
import {handleReportes} from './src/reportes.js'
import {handleCompras} from './src/compras/compras.js'
import {handlePagos} from './src/compras/pagos.js'

const AUTH_DIR='/app/auth'

const COMANDOS_REPORTES=/^(reportes|info|datos|cuenta bancaria|cuenta empleado|datos bancarios|cuenta|banco|clave|clabe|asistencia hoy|resumen|reporte|compras\b|faltas|retardos|críticos|criticos|graves|excel|reporte semanal|reporte mensual)\b/i

const sesion={sock:null,conectado:false}
let iniciando=false

const textoMensaje=m=>(
  m.message?.conversation||
  m.message?.extendedTextMessage?.text||
  m.message?.imageMessage?.caption||
  m.message?.videoMessage?.caption||
  m.message?.documentMessage?.caption||
  ''
).trim()

const obtenerLoc=m=>{
  const x=m.message?.locationMessage

  return x?{
    degreesLatitude:x.degreesLatitude,
    degreesLongitude:x.degreesLongitude
  }:null
}

const obtenerIdentidad=m=>{
  const participante=(m.key?.participant||'').toString()
  const remoto=(m.key?.remoteJid||'').toString()

  const candidatos=[
    m.key?.participantPn,
    m.key?.senderPn,
    m.key?.participantAlt,
    participante,
    remoto
  ]
  .filter(Boolean)
  .map(x=>x.toString())

  const lid=
    candidatos.find(
      x=>x.includes('@lid')
    )||''

  const fuenteTelefono=
    candidatos.find(
      x=>
        x &&
        !x.includes('@lid') &&
        !x.includes('@g.us')
    )||''

  const tel=
    fuenteTelefono
      .split(':')[0]
      .split('@')[0]
      .replace(/\D/g,'')

  return{
    rawLid:lid,
    tel,
    tel10:tel.slice(-10)
  }
}

const getFiltro=jid=>
  jid===GRUPO_COYOACAN_ID
    ?'coyoacan'
    :jid===GRUPO_BUCARELI_ID
      ?'juarez'
      :null

async function procesos(){

  if(
    !sesion.sock||
    !sesion.conectado
  )return

  try{
    await registrarDescansos()
  }catch(e){
    console.error(
      'Descansos:',
      e
    )
  }

  try{
    await cerrarSalidasPendientes()
  }catch(e){
    console.error(
      'Cierres:',
      e
    )
  }

  try{
    await checkNoLlegaron(
      sesion.sock
    )
  }catch(e){
    console.error(
      'No llegados:',
      e
    )
  }
}

async function conectar(){

  if(iniciando)return

  iniciando=true

  try{

    const {
      state,
      saveCreds
    }=
      await useMultiFileAuthState(
        AUTH_DIR
      )

    console.log(
      'AUTH DIR:',
      AUTH_DIR
    )

    let version

    try{

      version=
        (
          await fetchLatestBaileysVersion()
        ).version

      console.log(
        'Baileys version:',
        version
      )

    }catch{

      console.log(
        'No se pudo obtener version de Baileys'
      )
    }

    const sock=
      makeWASocket({

        auth:state,

        version,

        logger:
          pino({
            level:'silent'
          }),

        printQRInTerminal:false,

        browser:[
          'RH Trinidad',
          'Chrome',
          '1.0.0'
        ],

        markOnlineOnConnect:false,

        syncFullHistory:false,

        generateHighQualityLinkPreview:false

      })

    sesion.sock=sock

    sock.ev.on(
      'creds.update',
      saveCreds
    )

    sock.ev.on(
      'connection.update',
      async({
        connection,
        lastDisconnect,
        qr
      })=>{

        if(qr){

          console.log(
            'ESCANEA ESTE QR:'
          )

          qrcode.generate(
            qr,
            {small:true}
          )
        }

        if(connection==='open'){

          console.log(
            'CONECTADO MODULAR + VOLUME OK'
          )

          sesion.conectado=true
          iniciando=false

          await procesos()
        }

        if(connection==='close'){

          sesion.conectado=false
          sesion.sock=null

          const code=
            lastDisconnect
              ?.error
              ?.output
              ?.statusCode

          console.log(
            'Conexion cerrada. Codigo:',
            code
          )

          iniciando=false

          if(
            code!==DisconnectReason.loggedOut
          ){

            setTimeout(
              ()=>conectar().catch(
                console.error
              ),
              5000
            )

          }else{

            console.error(
              'Sesion cerrada / logout. Se requiere volver a vincular.'
            )
          }
        }
      }
    )

    sock.ev.on(
      'messages.upsert',
      async({
        messages,
        type
      })=>{

        if(type!=='notify')
          return

        for(const m of messages){

          try{

            if(
              !m.message||
              m.key?.fromMe
            )continue

            const jid=
              m.key?.remoteJid

            if(!jid)continue

            const tipoGrupo=
              getTipoGrupo(jid)

            const texto=
              textoMensaje(m)

            const loc=
              obtenerLoc(m)

            const identidad=
              obtenerIdentidad(m)

            /*
              FOTO
            */

            const esImagen=
              !!m.message?.imageMessage

            /*
              PDF
              Se identifica desde ahora para
              el siguiente paso de Compras.
            */

            const esDocumento=
              !!m.message?.documentMessage

            const esPDF=
              esDocumento&&
              (
                m.message
                  ?.documentMessage
                  ?.mimetype||''
              )
                .toLowerCase()===
                'application/pdf'

            if(
              tipoGrupo===
              'CHECADORES'
            ){

              if(loc){

                await handleChecador({

                  sock,
                  jid,
                  m,
                  loc,

                  rawLid:
                    identidad.rawLid,

                  tel10:
                    identidad.tel10,

                  tel:
                    identidad.tel

                })
              }

              continue
            }

            /*
              =================================================
              PAGOS
              =================================================

              Se mantiene ANTES de Compras para que un
              comprobante de pago no termine como compra.
            */

            if(
              (
                tipoGrupo==='REPORTES'||
                tipoGrupo==='GERENTES'
              )&&
              esImagen
            ){

              const esPago=
                await handlePagos({

                  sock,
                  jid,
                  m,
                  texto,
                  esImagen

                })

              if(esPago)
                continue
            }

            /*
              =================================================
              COMPRAS
              =================================================

              Las fotografías siguen exactamente la misma
              entrada.

              compras.js ahora se encarga de agrupar varias
              fotografías consecutivas.
            */

            if(
              (
                tipoGrupo==='REPORTES'||
                tipoGrupo==='GERENTES'
              )&&
              (
                esImagen||
                /^compras\b/i.test(texto)
              )
            ){

              await handleCompras({

                sock,
                jid,
                m,
                texto,
                esImagen

              })

              continue
            }

            /*
              =================================================
              REPORTES
              =================================================
            */

            if(
              (
                tipoGrupo==='REPORTES'||
                tipoGrupo==='GERENTES'
              )&&
              COMANDOS_REPORTES.test(texto)
            ){

              await handleReportes({

                sock,
                jid,
                m,
                texto,
                filtroGrupo:
                  getFiltro(jid)

              })

              continue
            }

          }catch(e){

            console.error(
              'Error procesando mensaje:',
              e
            )
          }
        }
      }
    )

  }catch(e){

    console.error(
      'Error iniciando WhatsApp:',
      e
    )

    sesion.conectado=false
    sesion.sock=null
    iniciando=false

    setTimeout(
      ()=>conectar().catch(
        console.error
      ),
      5000
    )
  }
}

cron.schedule(
  '*/10 * * * *',
  async()=>{

    console.log(
      '⏱️ Procesos automáticos...'
    )

    await procesos()

  },
  {
    timezone:
      'America/Mexico_City'
  }
)

console.log(
  '🚀 Iniciando RH Trinidad...'
)

conectar().catch(
  console.error
)

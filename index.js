// ============================================================
// INDEX.JS
// ============================================================
// Archivo principal del bot.
//
// Aquí se encarga de:
//
// 1. Iniciar el servidor WEB.
// 2. Manejar el QR de WhatsApp.
// 3. Mantener la sesión de Baileys.
// 4. Conectar con los grupos de WhatsApp.
// 5. Identificar el tipo de grupo.
// 6. Recibir ubicaciones para el checador.
// 7. Procesar reportes.
// 8. Procesar compras.
// 9. Ejecutar periódicamente el aviso de empleados
//    que no han llegado.
// ===========================================================


import makeWASocket, {
  useMultiFileAuthState,
  DisconnectReason,
  fetchLatestBaileysVersion
} from '@whiskeysockets/baileys'

import qrcodeTerminal from 'qrcode-terminal'
import QRCode from 'qrcode'
import express from 'express'
import fs from 'fs'
import path from 'path'
import P from 'pino'


// ============================================================
// IMPORTACIONES DEL BOT
// ============================================================

import {
  GRUPOS,
  PAQUETES,
  getTipoGrupo,
  GRUPO_COYOACAN_ID,
  GRUPO_BUCARELI_ID,
  GRUPO_CHECADOR_COYOACAN_ID,
  GRUPO_CHECADOR_BUCARELI_ID,
  GRUPO_PRUEBAS_ID,
  GRUPO_REPORTES_TRINIDAD_ID
} from './src/config.js'


import {
  handleChecador,
  checkNoLlegaron
} from './src/checador.js'


import {
  handleReportes
} from './src/reportes.js'


import {
  handleCompras
} from './src/compras.js'


// ============================================================
// 1. SERVIDOR WEB
// ============================================================

const app = express()

app.use(
  express.json({
    limit: '20mb'
  })
)


// Guarda el último QR generado por WhatsApp.
let lastQR = null


// ------------------------------------------------------------
// Página principal.
// ------------------------------------------------------------

app.get(
  '/',
  (r, s) =>
    s.send('Bot OK - /qr')
)


// ------------------------------------------------------------
// Página para visualizar el QR.
// ------------------------------------------------------------

app.get(
  '/qr',
  async (r, s) => {

    if (!lastQR) {
      return s.send(
        'No QR aun, espera 5 seg y recarga'
      )
    }

    const d =
      await QRCode.toDataURL(lastQR)

    s.send(`
      <div style="text-align:center">
        <h2>Escanea</h2>
        <img
          src="${d}"
          style="width:400px"
        >
        <p>Se actualiza cada 15s</p>
      </div>
    `)
  }
)


// ============================================================
// 2. BACKUP DE AUTENTICACIÓN
// ============================================================
//
// Estos endpoints trabajan con:
// /app/auth
//
// Se mantienen para poder respaldar/restaurar la sesión
// de WhatsApp si fuera necesario.
// ============================================================


// ------------------------------------------------------------
// BACKUP
// ------------------------------------------------------------

app.get(
  '/backup',
  (r, s) => {

    try {

      const dir =
        '/app/auth'


      if (!fs.existsSync(dir)) {
        return s.json({
          error: 'no auth'
        })
      }


      const files =
        fs.readdirSync(dir)


      if (files.length === 0) {

        return s.json({
          error: 'auth vacio - escanea QR'
        })
      }


      let o = {}


      files.forEach(
        f => {

          try {

            o[f] =
              fs.readFileSync(
                path.join(dir, f),
                'utf8'
              )

          } catch {}
        }
      )


      s.json(o)

    } catch (e) {

      s.json({
        error: e.message
      })
    }
  }
)


// ------------------------------------------------------------
// RESTORE
// ------------------------------------------------------------

app.post(
  '/restore',
  (r, s) => {

    try {

      const dir =
        '/app/auth'


      if (!fs.existsSync(dir)) {

        fs.mkdirSync(
          dir,
          {
            recursive: true
          }
        )
      }


      const data =
        r.body


      const payload =
        data.data
          ? data.data
          : data


      let c = 0


      for (
        const [f, v]
        of Object.entries(payload)
      ) {

        if (
          f === 'dir' ||
          f === 'files'
        ) {
          continue
        }


        if (
          typeof v === 'string' &&
          v.length > 10
        ) {

          fs.writeFileSync(
            path.join(dir, f),
            v,
            'utf8'
          )

          c++
        }
      }


      s.json({
        ok: true,
        restaurados: c
      })

    } catch (e) {

      s.json({
        error: e.message
      })
    }
  }
)


// ============================================================
// 3. INICIAR SERVIDOR WEB
// ============================================================

app.listen(
  process.env.PORT || 3000,
  () =>
    console.log(
      'WEB ' +
      (process.env.PORT || 3000) +
      ' OK'
    )
)


// ============================================================
// 4. DETERMINAR FILTRO DE SUCURSAL
// ============================================================
//
// Reportes y Pruebas:
//   Sin filtro.
//
// Grupo Coyoacán:
//   Filtro Coyoacán.
//
// Grupo Checador Coyoacán:
//   Filtro Coyoacán.
//
// Grupo Bucareli:
//   Filtro Bucareli.
//
// Grupo Checador Bucareli:
//   Filtro Bucareli.
// ============================================================

async function getFiltro(jid) {

  if (
    jid === GRUPO_PRUEBAS_ID ||
    jid === GRUPO_REPORTES_TRINIDAD_ID
  ) {

    return null
  }


  if (
    jid === GRUPO_COYOACAN_ID ||
    jid === GRUPO_CHECADOR_COYOACAN_ID
  ) {

    return 'coyoacan'
  }


  if (
    jid === GRUPO_BUCARELI_ID ||
    jid === GRUPO_CHECADOR_BUCARELI_ID
  ) {

    return 'bucareli'
  }


  return null
}


// ============================================================
// 5. CONTROL DEL CRON DEL CHECADOR
// ============================================================

let cronInterval = null


// ============================================================
// 6. INICIAR WHATSAPP / auth
// ============================================================

async function start() {

  // ----------------------------------------------------------
  // CARPETA DE AUTENTICACIÓN
  // ----------------------------------------------------------
  //
  // IMPORTANTE:
  // Esta es la carpeta donde Baileys guarda la sesión.
  //
  // Railway debe conservar esta carpeta mediante el Volume.
  // ----------------------------------------------------------

  const authDir =
    '/app/auth'


  // Si no existe, crearla.
  if (!fs.existsSync(authDir)) {

    fs.mkdirSync(
      authDir,
      {
        recursive: true
      }
    )
  }


  // ==========================================================
  // DIAGNÓSTICO DEL VOLUME
  // ==========================================================
  //
  // Estas líneas son nuevas.
  //
  // Nos permiten saber si Railway realmente está entregando
  // los archivos de autenticación que Baileys guardó
  // anteriormente.
  // ==========================================================

  console.log(
    'AUTH DIR:',
    authDir
  )


  try {

    const archivosAuth =
      fs.readdirSync(authDir)


    console.log(
      'ARCHIVOS AUTH:',
      archivosAuth
    )


    console.log(
      'TOTAL ARCHIVOS AUTH:',
      archivosAuth.length
    )

  } catch (e) {

    console.log(
      'ERROR LEYENDO AUTH:',
      e.message
    )
  }


  // ==========================================================
  // 7. CARGAR SESIÓN DE BAILEYS
  // ==========================================================

  const {
    state,
    saveCreds
  } =
    await useMultiFileAuthState(
      authDir
    )


  // Obtener versión actual de WhatsApp/Baileys.
  const {
    version
  } =
    await fetchLatestBaileysVersion()


  // ==========================================================
  // 8. CREAR SOCKET DE WHATSAPP
  // ==========================================================

  const sock =
    makeWASocket({

      version,

      auth: state,

      logger:
        P({
          level: 'fatal'
        }),

      browser: [
        'Trinidad Bot',
        'Chrome',
        '121'
      ],

      getMessage:
        async () => undefined
    })


  // ==========================================================
  // 9. GUARDAR CAMBIOS DE CREDENCIALES
  // ==========================================================
  //
  // Cada vez que Baileys modifica las credenciales,
  // saveCreds las guarda dentro de /app/baileys_auth.
  // ==========================================================

  sock.ev.on(
    'creds.update',
    saveCreds
  )


  // ==========================================================
  // 10. ESTADO DE CONEXIÓN
  // ==========================================================

  sock.ev.on(
    'connection.update',
    async ({
      connection,
      lastDisconnect,
      qr
    }) => {

      // ------------------------------------------------------
      // QR GENERADO
      // ------------------------------------------------------

      if (qr) {

        lastQR = qr

        console.log(
          'QR generado'
        )

        qrcodeTerminal.generate(
          qr,
          {
            small: false
          }
        )
      }


      // ------------------------------------------------------
      // CONEXIÓN CERRADA
      // ------------------------------------------------------

      if (connection === 'close') {

        const c =
          lastDisconnect
            ?.error
            ?.output
            ?.statusCode


        console.log(
          'CONEXION CERRADA - CODIGO:',
          c
        )


        // Si no fue logout, intentar reconectar.
        if (
          c !== DisconnectReason.loggedOut
        ) {

          console.log(
            'Reintentando conexión en 5 segundos...'
          )


          setTimeout(
            () => start(),
            5000
          )
        }
      }


      // ------------------------------------------------------
      // CONEXIÓN ABIERTA
      // ------------------------------------------------------

      if (connection === 'open') {

        console.log(
          'CONECTADO MODULAR + VOLUME OK'
        )


        // Si ya conectó, eliminar QR de memoria.
        lastQR = null


        // Evitar tener varios intervalos simultáneos.
        if (cronInterval) {

          clearInterval(
            cronInterval
          )
        }


        // ====================================================
        // 11. REVISIÓN AUTOMÁTICA DE NO LLEGADAS
        // ====================================================
        //
        // Cada 10 minutos revisa:
        //
        // - empleados que deberían haber llegado
        // - si tienen entrada
        // - si ya pasaron 20 minutos
        //
        // Los avisos son enviados por checkNoLlegaron().
        // ====================================================

        cronInterval =
          setInterval(
            () => {

              checkNoLlegaron(sock)
                .catch(
                  e =>
                    console.log(
                      'cron error',
                      e.message
                    )
                )

            },
            10 * 60 * 1000
          )
      }
    }
  )


  // ==========================================================
  // 12. RECEPCIÓN DE MENSAJES
  // ==========================================================

  sock.ev.on(
    'messages.upsert',
    async ({
      messages
    }) => {

      try {

        const m =
          messages[0]


        // Ignorar mensajes inexistentes.
        if (!m) return


        // Ignorar mensajes enviados por el propio bot.
        if (m.key.fromMe) return


        // Grupo donde se recibió el mensaje.
        const jid =
          m.key.remoteJid


        // El bot solamente trabaja dentro de grupos.
        if (
          !jid ||
          !jid.endsWith('@g.us')
        ) {
          return
        }


        // ====================================================
        // 13. OBTENER TEXTO DEL MENSAJE
        // ====================================================

        const texto =
          m.message?.conversation ||
          m.message?.extendedTextMessage?.text ||
          m.message?.imageMessage?.caption ||
          m.message?.documentMessage?.caption ||
          ''


        // ====================================================
        // 14. OBTENER UBICACIÓN
        // ====================================================

        const loc =
          m.message?.locationMessage ||
          m.message?.liveLocationMessage


        // ====================================================
        // 15. COMANDO "ID"
        // ====================================================
        //
        // Permite conocer el ID del grupo.
        // ====================================================

        if (
          texto
            .trim()
            .toLowerCase() === 'id'
        ) {

          console.log(
            `ID solicitado en ${jid}`
          )


          await sock.sendMessage(
            jid,
            {
              text: jid
            }
          )


          return
        }


        // ====================================================
        // 16. IDENTIFICAR TELÉFONO / LID
        // ====================================================

        const rawLid =
          m.key.participant || ''


        const realPn =
          m.key.participantPn || ''


        let pn = ''


        try {

          pn =
            await sock
              .signalRepository
              ?.lidMapping
              ?.getPNForLID(rawLid) ||
            ''

        } catch {}


        const rawId =
          realPn ||
          pn ||
          rawLid ||
          jid


        let tel =
          (rawId || '')
            .toString()
            .replace(/\D/g, '')


        let tel10 =
          tel.slice(-10)


        // ====================================================
        // 17. IDENTIFICAR SI ES IMAGEN
        // ====================================================

        const esImagen =
          !!(
            m.message?.imageMessage
          )


        // ====================================================
        // 18. IDENTIFICAR TIPO DE GRUPO
        // ====================================================

        const tipo =
          getTipoGrupo(jid)


        // Si el grupo no está configurado, ignorar.
        if (!tipo) {

          console.log(
            `Grupo no configurado: ${jid}`
          )

          return
        }


        // ====================================================
        // 19. OBTENER FILTRO DE SUCURSAL
        // ====================================================

        const filtro =
          await getFiltro(jid)


        // ====================================================
        // 20. GRUPOS DE REPORTES
        // ====================================================

        if (
          tipo === 'REPORTES'
        ) {

          // --------------------------------------------------
          // Compras:
          // - imágenes
          // - mensajes que comienzan con "compras"
          // --------------------------------------------------

          if (
            esImagen ||
            /^compras/i.test(texto)
          ) {

            const ok =
              await handleCompras({
                sock,
                jid,
                m,
                texto,
                esImagen
              })


            if (ok) {
              return
            }
          }


          // --------------------------------------------------
          // Reportes y consultas.
          // --------------------------------------------------

          if (
            PAQUETES
              .REPORTES_PARA_GERENTES
              .test(texto)
          ) {

            await handleReportes({
              texto,
              jid,
              sock,
              filtroGrupo: filtro
            })


            return
          }


          return
        }


        // ====================================================
        // 21. GRUPOS DE CHECADORES
        // ====================================================

        if (
          tipo === 'CHECADORES'
        ) {

          // Los grupos de checador únicamente procesan
          // mensajes que contienen ubicación.
          if (!loc) return


          console.log(
            `Ubicación recibida de ${tel10} en ${jid}`
          )


          // --------------------------------------------------
          // Procesar entrada/salida.
          //
          // IMPORTANTE:
          // jid es el grupo donde se mandó la ubicación.
          //
          // Por eso la confirmación se envía a ese mismo grupo.
          // --------------------------------------------------

          await handleChecador({
            sock,
            jid,
            m,
            loc,
            rawLid,
            tel10,
            tel
          })


          return
        }


        // ====================================================
        // 22. GRUPOS DE GERENTES
        // ====================================================

        if (
          tipo === 'GERENTES'
        ) {

          // Los grupos de gerentes no procesan ubicaciones
          // como checador.
          if (loc) return


          // Si el mensaje no corresponde a un comando
          // permitido para reportes, ignorarlo.
          if (
            !PAQUETES
              .REPORTES_PARA_GERENTES
              .test(texto)
          ) {
            return
          }


          // Procesar reporte.
          await handleReportes({
            texto,
            jid,
            sock,
            filtroGrupo: filtro
          })


          return
        }

      } catch (e) {

        // ====================================================
        // 23. ERROR GENERAL DE MENSAJES
        // ====================================================

        console.error(
          'Error upsert',
          e
        )
      }
    }
  )
}


// ============================================================
// 24. INICIAR BOT
// ============================================================

start()

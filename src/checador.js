// ============================================================
// CHECADOR
// ============================================================
// Este archivo se encarga de:
//
// 1. Recibir la ubicación del empleado.
// 2. Determinar la sucursal según GPS.
// 3. Identificar al empleado.
// 4. Consultar su horario.
// 5. Registrar ENTRADA.
// 6. Registrar SALIDA.
// 7. Calcular horas trabajadas y horas extra.
// 8. Avisar cuando un empleado lleva 20 minutos sin llegar.
//
// IMPORTANTE:
// - La sucursal registrada en Sheets se determina por GPS.
// - La confirmación de entrada/salida se manda al grupo
//   donde el empleado envió su ubicación.
// - Los avisos de 20 minutos se mandan al grupo de GERENTES.
// ============================================================


// ============================================================
// 1. IMPORTACIONES
// ============================================================

// Sucursales y grupos de WhatsApp.
// Los grupos CHECADOR se utilizan para las confirmaciones.
// Los grupos GERENTES se utilizan para avisos de faltas/retardos.
import {
  SUCURSALES,
  GRUPO_COYOACAN_ID,
  GRUPO_BUCARELI_ID,
  SPREADSHEET_ID
} from './config.js'


// Funciones auxiliares de fechas, horarios, distancia y horas trabajadas.
import {
  distM,
  fechaLaboral,
  horaMX,
  minutos,
  normaliza,
  calcularHorasTrabajadas,
  calcularExtra,
  parseHorarioRango
} from './utils.js'


// Funciones para consultar y modificar Google Sheets.
import {
  getRows,
  sheetsClient,
  getHorarioBaseMap
} from './sheets.js'


// ============================================================
// 2. CONFIGURACIÓN DEL CHECADOR
// ============================================================

// Minutos de tolerancia para considerar una entrada "A TIEMPO".
//
// Ejemplo:
// Horario: 08:00
// Entrada: 08:15 -> A TIEMPO
// Entrada: 08:16 -> RETARDO 16min
const TOLERANCIA_MIN = 15


// Después de estos minutos sin registrar entrada,
// se manda un aviso al grupo de GERENTES.
//
// Ejemplo:
// Horario: 08:00
// 08:20 -> aviso de NO HA LLEGADO
const AVISO_FALTA_MIN = 20


// ============================================================
// 3. CONTROL DE AVISOS DE FALTA
// ============================================================

// Guarda los empleados a los que ya se les mandó el aviso
// de 20 minutos durante el día.
//
// Esto evita que el bot mande el mismo aviso cada 10 minutos.
let avisosHoy = new Set()


// Guarda la fecha correspondiente al contenido de avisosHoy.
let fechaAvisos = ''


// ============================================================
// 4. REGISTRO DE ENTRADA / SALIDA
// ============================================================

export async function handleChecador({
  sock,
  jid,
  m,
  loc,
  rawLid,
  tel10,
  tel
}) {

  // ==========================================================
  // 4.1 OBTENER UBICACIÓN GPS
  // ==========================================================

  const lat = loc.degreesLatitude
  const lng = loc.degreesLongitude


  // Aquí buscamos cuál sucursal está más cerca de la ubicación.
  let cercana = null
  let dMin = Infinity


  for (const s of SUCURSALES) {
    const d = distM(lat, lng, s.lat, s.lng)

    if (d < dMin) {
      dMin = d
      cercana = s
    }
  }


  // ==========================================================
  // 4.2 BUSCAR EMPLEADO
  // ==========================================================

  // Se consulta la hoja de empleados.
  const empRows = await getRows('Empleados!A:K')


  // Función para obtener el LID de WhatsApp de una fila.
  const getLid = r =>
    (r.find(x => String(x).includes('@lid')) || '').trim()


  let emp = null


  // ----------------------------------------------------------
  // Primero intentamos localizar al empleado por teléfono.
  // ----------------------------------------------------------

  if (tel10.length >= 10) {

    const i = empRows.findIndex(
      (r, idx) =>
        idx > 0 &&
        r[0] &&
        r[0].replace(/\D/g, '').slice(-10) === tel10
    )


    if (i > -1) {
      emp = empRows[i]
    }
  }


  // ----------------------------------------------------------
  // Si no apareció por teléfono, intentamos por LID.
  // ----------------------------------------------------------

  if (!emp && rawLid.includes('@lid')) {

    const i = empRows.findIndex(
      (r, idx) =>
        idx > 0 &&
        getLid(r) === rawLid
    )


    if (i > -1) {

      emp = empRows[i]

      tel10 =
        (emp[0] || '')
          .replace(/\D/g, '')
          .slice(-10)
    }
  }


  // ==========================================================
  // 4.3 DATOS BÁSICOS DEL EMPLEADO
  // ==========================================================

  // Nombre completo.
  //
  // Columna D = Nombre Completo
  // Columna B = Nombre Corto
  const nombre =
    emp
      ? (emp[3] || emp[1])
      : (m.pushName || tel10 || 'Desconocido')


  // Teléfono del empleado.
  const telF =
    emp
      ? (emp[0] || '').replace(/\D/g, '')
      : tel


  const tel10F =
    telF.slice(-10) || tel10


  // ==========================================================
  // 4.4 FECHA Y REGISTRO DEL DÍA
  // ==========================================================

  const fLab = fechaLaboral()


  // Consultamos las asistencias actuales.
  const asis = await getRows('Asistencia!A:M')


  // Buscamos si este empleado ya tiene registro de hoy.
  const idx = asis.findIndex(
    (r, i) =>
      i > 0 &&
      r[0] &&
      r[0].replace(/\D/g, '').slice(-10) === tel10F &&
      r[2] === fLab
  )


  // Si existe, guardamos el registro.
  const hoy =
    idx > -1
      ? asis[idx]
      : null


  // ==========================================================
  // 4.5 OBTENER HORARIO DEL EMPLEADO
  // ==========================================================

  const sClient = await sheetsClient()


  // Obtiene el mapa de horarios.
  const baseMap = await getHorarioBaseMap()


  // Primero intenta encontrar el horario por teléfono.
  // Si no, intenta por nombre corto.
  const base =
    baseMap[tel10F] ||
    baseMap[normaliza(nombre).split(' ')[0]] ||
    null


  // ==========================================================
  // 4.6 DETERMINAR ESTATUS DE ENTRADA
  // ==========================================================

  let estatus = 'A TIEMPO'

  let hObj = null

  let esRet = false

  let minRet = 0

  let hProg = ''


  if (base) {

    // Fecha utilizada para determinar el día de la semana.
    const fe = new Date(fLab + 'T12:00:00')

    const dn = fe.getDay()


    // --------------------------------------------------------
    // Si el día es descanso.
    // --------------------------------------------------------

    if (base.descansos.has(dn)) {

      estatus = 'DESCANSO'


    // --------------------------------------------------------
    // Si existe horario para ese día.
    // --------------------------------------------------------

    } else if (base.horas[dn]) {

      hObj = base.horas[dn]

      hProg = hObj.entrada || ''


      // ------------------------------------------------------
      // Horario normal, no LIBRE.
      // ------------------------------------------------------

      if (
        hObj.entrada !== 'LIBRE' &&
        hObj.entrada
      ) {

        // Diferencia entre hora actual y hora programada.
        const dif =
          minutos(horaMX()) -
          minutos(hObj.entrada)


        // ----------------------------------------------------
        // Más de 15 minutos = RETARDO.
        // ----------------------------------------------------

        if (dif > TOLERANCIA_MIN) {

          estatus =
            `RETARDO ${dif}min (Prog ${hObj.entrada})`

          esRet = true

          minRet = dif


        // ----------------------------------------------------
        // Dentro de los primeros 15 minutos = A TIEMPO.
        // ----------------------------------------------------

        } else {

          estatus =
            `A TIEMPO (Prog ${hObj.entrada})`
        }
      }
    }
  }


  // ==========================================================
  // 5. REGISTRO DE ENTRADA / SALIDA
  // ==========================================================

  try {

    // ========================================================
    // 5.1 ENTRADA
    // ========================================================

    if (!hoy || !hoy[3]) {

      // ------------------------------------------------------
      // Validar distancia para poder registrar entrada.
      // ------------------------------------------------------

      if (dMin > cercana.rEnt) {

        await sock.sendMessage(
          jid,
          {
            text:
              `Debes estar a max ${cercana.rEnt}m de ${cercana.nombre}`
          },
          { quoted: m }
        )

        return
      }


      // Hora actual.
      const h = horaMX()


      // Texto del horario que se guarda en columna M.
      const jTxt =
        hObj
          ? `${hObj.entrada}${hObj.salida ? ` - ${hObj.salida}` : ''}`
          : '8h'


      // ------------------------------------------------------
      // Crear fila de asistencia.
      //
      // A = teléfono
      // B = nombre
      // C = fecha
      // D = entrada
      // E = estatus
      // F = salida
      // G = sucursal
      // H = distancia
      // I/J = campos existentes
      // K = horas trabajadas
      // L = horas extra
      // M = horario
      // ------------------------------------------------------

      const row = [
        tel10F,
        nombre,
        fLab,
        h,
        estatus,
        '',
        cercana.nombre,
        Math.round(dMin).toString(),
        '',
        '',
        calcularHorasTrabajadas(h, ''),
        '0',
        jTxt
      ]


      // ------------------------------------------------------
      // Si no existe registro, crear una nueva fila.
      // ------------------------------------------------------

      if (idx === -1) {

        await sClient.spreadsheets.values.append({
          spreadsheetId: SPREADSHEET_ID,
          range: 'Asistencia!A:M',
          valueInputOption: 'USER_ENTERED',
          requestBody: {
            values: [row]
          }
        })


      // ------------------------------------------------------
      // Si ya existe una fila pero no tenía entrada,
      // actualizar esa fila.
      // ------------------------------------------------------

      } else {

        await sClient.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `Asistencia!A${idx + 1}:M${idx + 1}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: {
            values: [row]
          }
        })
      }


      // ------------------------------------------------------
      // CONFIRMACIÓN DE ENTRADA
      //
      // IMPORTANTE:
      // "jid" es el grupo donde el empleado mandó
      // la ubicación.
      //
      // Por eso la confirmación siempre vuelve al grupo
      // donde mandó la ubicación.
      // ------------------------------------------------------

      await sock.sendMessage(
        jid,
        {
          text:
            `✅ ${estatus} - ${nombre} en ${cercana.nombre} - ${h}`
        }
      )


    // ========================================================
    // 5.2 SALIDA
    // ========================================================

    } else {

      // ------------------------------------------------------
      // Ya existe entrada y salida.
      // ------------------------------------------------------

      if (hoy[5]) {

        await sock.sendMessage(
          jid,
          {
            text: `Salida ya registrada`
          }
        )

        return
      }


      // ------------------------------------------------------
      // Validar distancia para poder registrar salida.
      // ------------------------------------------------------

      if (dMin > cercana.rSal) {

        await sock.sendMessage(
          jid,
          {
            text:
              `No puedes checar salida a ${Math.round(dMin)}m`
          },
          { quoted: m }
        )

        return
      }


      // Hora de salida.
      const h = horaMX()


      // Recuperar horario guardado.
      const jTxt =
        hoy[12] || '8h'


      // ------------------------------------------------------
      // Calcular horas trabajadas y horas extra.
      // Estos valores se guardan en Sheets.
      // ------------------------------------------------------

      const {
        trabajadas,
        extra
      } =
        calcularExtra(
          hoy[3],
          h,
          hObj?.entrada || null,
          hObj?.salida || null
        )


      // ------------------------------------------------------
      // Actualizar salida y resultados en Sheets.
      // ------------------------------------------------------

      await sClient.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Asistencia!F${idx + 1}:M${idx + 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [[
            h,
            hoy[6] || '',
            hoy[7] || '',
            cercana.nombre,
            Math.round(dMin).toString(),
            trabajadas,
            extra,
            jTxt
          ]]
        }
      })


      // ------------------------------------------------------
      // CONFIRMACIÓN DE SALIDA
      // ------------------------------------------------------
      //
      // Actualmente el código también muestra Trab y Extra
      // en WhatsApp.
      //
      // Esto lo podemos cambiar en el siguiente ajuste para
      // que WhatsApp solamente diga que la salida fue registrada.
      //
      // Las horas trabajadas y extras seguirán quedando
      // guardadas en Sheets.
      // ------------------------------------------------------

      await sock.sendMessage(
        jid,
        {
          text:
            `✅ Salida - ${nombre} en ${cercana.nombre} - ${h} - Trab ${trabajadas} Extra ${extra}`
        }
      )
    }


  } catch (e) {

    // ========================================================
    // 6. MANEJO DE ERRORES
    // ========================================================

    console.error(e)
  }
}


// ============================================================
// 7. AVISOS DE EMPLEADOS QUE NO HAN LLEGADO
// ============================================================
//
// Esta función se ejecuta periódicamente desde index.js.
//
// Revisa:
// - Horario programado
// - Hora actual
// - Si existe entrada registrada
//
// Si ya pasaron 20 minutos y no existe entrada,
// manda aviso al grupo de GERENTES correspondiente.
//
// NO manda estos avisos a los grupos de CHECADORES.
// ============================================================

export async function checkNoLlegaron(sock) {

  // Fecha laboral actual.
  const hoy = fechaLaboral()


  // ----------------------------------------------------------
  // Si cambió el día, limpiar el control de avisos.
  // ----------------------------------------------------------

  if (fechaAvisos !== hoy) {

    avisosHoy.clear()

    fechaAvisos = hoy
  }


  // Hora actual convertida a minutos.
  const ahoraMin = minutos(horaMX())


  // ----------------------------------------------------------
  // Obtener horarios base.
  // ----------------------------------------------------------

  const baseRows =
    await getRows('Horario_Base!A2:K')


  // ----------------------------------------------------------
  // Obtener asistencias actuales.
  // ----------------------------------------------------------

  const asisRows =
    await getRows('Asistencia!A2:M')


  // ----------------------------------------------------------
  // Determinar el día de la semana.
  // ----------------------------------------------------------

  const diaNum =
    new Date(
      new Date().toLocaleString(
        'en-US',
        {
          timeZone: 'America/Mexico_City'
        }
      )
    ).getDay()


  // ==========================================================
  // 8. REVISAR CADA EMPLEADO
  // ==========================================================

  for (const r of baseRows) {

    // --------------------------------------------------------
    // Teléfono del empleado.
    // --------------------------------------------------------

    const tel =
      (r[0] || '')
        .replace(/\D/g, '')
        .slice(-10)


    if (!tel) continue


    // --------------------------------------------------------
    // Llave única para evitar avisos repetidos.
    // --------------------------------------------------------

    const key =
      `${hoy}_${tel}`


    if (avisosHoy.has(key)) continue


    // --------------------------------------------------------
    // Obtener el horario correspondiente al día.
    //
    // Lunes = r[3]
    // Martes = r[4]
    // Miércoles = r[5]
    // Jueves = r[6]
    // Viernes = r[7]
    // Sábado = r[8]
    // Domingo = r[9]
    // --------------------------------------------------------

    const mapa = {
      1: r[3],
      2: r[4],
      3: r[5],
      4: r[6],
      5: r[7],
      6: r[8],
      0: r[9]
    }


    const v =
      (mapa[diaNum] || '')
        .toString()
        .trim()


    // Interpretar el rango de horario.
    const parsed =
      parseHorarioRango(v)


    // --------------------------------------------------------
    // Si no hay horario o es LIBRE, no generar aviso.
    // --------------------------------------------------------

    if (
      !parsed ||
      parsed.entrada === 'LIBRE' ||
      !parsed.entrada
    ) {
      continue
    }


    // --------------------------------------------------------
    // Calcular cuántos minutos han pasado desde la entrada
    // programada.
    // --------------------------------------------------------

    const dif =
      ahoraMin -
      minutos(parsed.entrada)


    // --------------------------------------------------------
    // Todavía no han pasado 20 minutos.
    // --------------------------------------------------------

    if (dif < AVISO_FALTA_MIN) {
      continue
    }


    // ========================================================
    // 9. COMPROBAR SI YA REGISTRÓ ENTRADA
    // ========================================================

    const registro =
      asisRows.find(
        a =>
          (a[0] || '')
            .replace(/\D/g, '')
            .slice(-10) === tel &&
          a[2] === hoy &&
          a[3]
      )


    // Si ya registró entrada, no mandar aviso.
    if (registro) {
      continue
    }


    // ========================================================
    // 10. PREPARAR AVISO
    // ========================================================

    // Marcar que ya se mandó el aviso de este empleado.
    avisosHoy.add(key)


    // Nombre del empleado.
    const nombre =
      r[1] || tel


    // Sucursal asignada en Horario_Base.
    const sucId =
      (r[2] || '')
        .toLowerCase()


    // ========================================================
    // 11. DETERMINAR GRUPO DE GERENTES
    // ========================================================
    //
    // Coyoacán / Hotel -> Grupo Gerentes Coyoacán
    //
    // Bucareli / Juárez -> Grupo Gerentes Bucareli
    //
    // IMPORTANTE:
    // Ya NO se utilizan los grupos de CHECADORES para
    // estos avisos.
    //
    // Tampoco se utiliza "trinidad" para enviarlo
    // automáticamente a Coyoacán.
    // ========================================================

    let grupoAviso = null


    if (
      sucId.includes('coyo') ||
      sucId.includes('hotel')
    ) {

      grupoAviso =
        GRUPO_COYOACAN_ID


    } else if (
      sucId.includes('bucareli') ||
      sucId.includes('juarez')
    ) {

      grupoAviso =
        GRUPO_BUCARELI_ID
    }


    // ========================================================
    // 12. ENVIAR AVISO AL GRUPO DE GERENTES
    // ========================================================

    if (grupoAviso) {

      try {

        await sock.sendMessage(
          grupoAviso,
          {
            text:
              `⚠️ NO HA LLEGADO - ${nombre} - Prog ${parsed.entrada} - ${dif}min tarde - [${r[2]}]`
          }
        )

      } catch {}
    }
  }
}

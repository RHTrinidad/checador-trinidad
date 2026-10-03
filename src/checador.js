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
// 9. Avisar cuando un empleado llega con 60 minutos o más.
// 10. Controlar 3 retardos en 30 días hábiles.
// 11. Permitir múltiples entradas/salidas en días LIBRE.
//
// IMPORTANTE:
// - La sucursal registrada en Sheets se determina por GPS.
// - La confirmación de entrada/salida se manda al grupo
//   donde el empleado envió su ubicación.
// - Los avisos se mandan al grupo correspondiente de GERENTES.
// - En días LIBRE no existe cálculo de retardo.
// - En LIBRE cada intervalo ENTRADA -> SALIDA se guarda
//   en una fila independiente de Asistencia.
// ============================================================


// ============================================================
// 1. IMPORTACIONES
// ============================================================

import {
  SUCURSALES,
  GRUPO_COYOACAN_ID,
  GRUPO_BUCARELI_ID,
  SPREADSHEET_ID
} from './config.js'


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


import {
  getRows,
  sheetsClient,
  getHorarioBaseMap
} from './sheets.js'


// ============================================================
// 2. CONFIGURACIÓN
// ============================================================

// Tolerancia para entrada.
const TOLERANCIA_MIN = 15


// Después de 20 minutos sin entrada,
// se manda aviso a GERENTES.
const AVISO_FALTA_MIN = 20


// Después de 60 minutos de retraso,
// además de la confirmación normal,
// se manda aviso a GERENTES.
const AVISO_RETARDO_60_MIN = 60


// A partir de esta fecha se consideran los retardos
// para el control de 3 retardos.
const FECHA_CORTE_RETARDOS = '2026-10-02'


// Nombre de la hoja donde se guarda el estado
// de los avisos de retardos acumulados.
const HOJA_ALERTAS_RETARDOS = 'Alertas_Retardos'


// ============================================================
// 3. CONTROL DE AVISOS DEL DÍA
// ============================================================

// Avisos de empleados que no han llegado.
let avisosHoy = new Set()


// Fecha correspondiente a avisosHoy.
let fechaAvisos = ''


// Avisos de retardos de 60 minutos.
let avisos60Hoy = new Set()


// Fecha correspondiente a avisos60Hoy.
let fechaAvisos60 = ''


// ============================================================
// 4. FUNCIONES AUXILIARES
// ============================================================


// ------------------------------------------------------------
// Determinar grupo de GERENTES según sucursal asignada.
// ------------------------------------------------------------

function grupoGerentesPorSucursal(sucursal) {

  const s =
    normaliza(sucursal || '')
      .toLowerCase()


  // Coyoacán / Hotel
  if (
    s.includes('coyo') ||
    s.includes('hotel')
  ) {
    return GRUPO_COYOACAN_ID
  }


  // Bucareli / Juárez
  if (
    s.includes('bucareli') ||
    s.includes('juarez')
  ) {
    return GRUPO_BUCARELI_ID
  }


  return null
}


// ------------------------------------------------------------
// Determinar si una fecha es día hábil.
//
// Para este control consideramos:
// Lunes a viernes = hábil
// Sábado y domingo = no hábil
// ------------------------------------------------------------

function esDiaHabil(fecha) {

  const d =
    new Date(`${fecha}T12:00:00`)


  const dia = d.getDay()


  return dia !== 0 && dia !== 6
}


// ------------------------------------------------------------
// Obtener la fecha inicial de una ventana de 30 días
// hábiles hacia atrás.
//
// El día actual cuenta si es lunes-viernes.
// ------------------------------------------------------------

function inicioVentana30DiasHabiles(fechaFin) {

  const d =
    new Date(`${fechaFin}T12:00:00`)


  let contados = 0


  while (contados < 30) {

    const fecha =
      d.toISOString().slice(0, 10)


    if (esDiaHabil(fecha)) {
      contados++
    }


    if (contados >= 30) {
      return fecha
    }


    d.setDate(d.getDate() - 1)
  }


  return fechaFin
}


// ------------------------------------------------------------
// Asegurar que exista la hoja Alertas_Retardos.
// ------------------------------------------------------------
//
// Columnas:
//
// A = Teléfono
// B = Nombre
// C = Último aviso
// D = Retardos incluidos en último aviso
// ------------------------------------------------------------

async function asegurarHojaAlertasRetardos(sClient) {

  const meta =
    await sClient.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID
    })


  let sheet =
    meta.data.sheets?.find(
      s =>
        s.properties?.title === HOJA_ALERTAS_RETARDOS
    )


  // ----------------------------------------------------------
  // Si no existe, crearla.
  // ----------------------------------------------------------

  if (!sheet) {

    await sClient.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: HOJA_ALERTAS_RETARDOS
              }
            }
          }
        ]
      }
    })


    // Volvemos a obtener la información.
    const meta2 =
      await sClient.spreadsheets.get({
        spreadsheetId: SPREADSHEET_ID
      })


    sheet =
      meta2.data.sheets?.find(
        s =>
          s.properties?.title === HOJA_ALERTAS_RETARDOS
      )
  }


  // ----------------------------------------------------------
  // Si la hoja está vacía, crear encabezados.
  // ----------------------------------------------------------

  const rows =
    await getRows(
      `${HOJA_ALERTAS_RETARDOS}!A:D`
    )


  if (!rows || rows.length === 0) {

    await sClient.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${HOJA_ALERTAS_RETARDOS}!A1:D1`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[
          'Teléfono',
          'Nombre',
          'Último aviso',
          'Retardos incluidos'
        ]]
      }
    })
  }
}


// ------------------------------------------------------------
// Registrar/actualizar la fecha del último aviso de retardos.
// ------------------------------------------------------------

async function guardarUltimoAvisoRetardo({
  sClient,
  tel,
  nombre,
  fecha,
  detalles
}) {

  await asegurarHojaAlertasRetardos(sClient)


  const rows =
    await getRows(
      `${HOJA_ALERTAS_RETARDOS}!A:D`
    )


  const idx =
    rows.findIndex(
      (r, i) =>
        i > 0 &&
        (r[0] || '')
          .replace(/\D/g, '')
          .slice(-10) === tel
    )


  // ----------------------------------------------------------
  // Si ya existe registro, actualizar.
  // ----------------------------------------------------------

  if (idx > -1) {

    await sClient.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range:
        `${HOJA_ALERTAS_RETARDOS}!A${idx + 1}:D${idx + 1}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[
          tel,
          nombre,
          fecha,
          detalles
        ]]
      }
    })


    return
  }


  // ----------------------------------------------------------
  // Si no existe, crear.
  // ----------------------------------------------------------

  await sClient.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${HOJA_ALERTAS_RETARDOS}!A:D`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[
        tel,
        nombre,
        fecha,
        detalles
      ]]
    }
  })
}


// ------------------------------------------------------------
// Revisar si el empleado llegó a 3 retardos.
//
// Solo considera:
// - Retardos > 15 minutos.
// - Desde 02/10/2026.
// - Dentro de los últimos 30 días hábiles.
// - Después del último aviso, para reiniciar el conteo.
// ------------------------------------------------------------

async function revisarTresRetardos({
  sock,
  sClient,
  tel,
  nombre,
  grupoAviso,
  fechaActual
}) {

  try {

    // --------------------------------------------------------
    // Si no tenemos grupo de gerentes, no podemos avisar.
    // --------------------------------------------------------

    if (!grupoAviso) {
      return
    }


    // --------------------------------------------------------
    // Asegurar hoja de control.
    // --------------------------------------------------------

    await asegurarHojaAlertasRetardos(sClient)


    // --------------------------------------------------------
    // Obtener último aviso de este empleado.
    // --------------------------------------------------------

    const alertasRows =
      await getRows(
        `${HOJA_ALERTAS_RETARDOS}!A:D`
      )


    const alertaEmpleado =
      alertasRows.find(
        (r, i) =>
          i > 0 &&
          (r[0] || '')
            .replace(/\D/g, '')
            .slice(-10) === tel
      )


    const ultimoAviso =
      alertaEmpleado?.[2] || ''


    // --------------------------------------------------------
    // Determinar ventana de 30 días hábiles.
    // --------------------------------------------------------

    let inicio =
      inicioVentana30DiasHabiles(
        fechaActual
      )


    // Nunca revisar antes del corte establecido.
    if (inicio < FECHA_CORTE_RETARDOS) {
      inicio = FECHA_CORTE_RETARDOS
    }


    // --------------------------------------------------------
    // Obtener asistencias.
    // --------------------------------------------------------

    const asisRows =
      await getRows('Asistencia!A:M')


    const retardos = []


    // --------------------------------------------------------
    // Buscar retardos.
    // --------------------------------------------------------

    for (const r of asisRows) {

      const fecha =
        String(r[2] || '').trim()


      if (!fecha) continue


      // Fecha posterior al último aviso.
      if (
        ultimoAviso &&
        fecha <= ultimoAviso
      ) {
        continue
      }


      // No considerar antes del corte.
      if (fecha < FECHA_CORTE_RETARDOS) {
        continue
      }


      // No considerar fuera de la ventana.
      if (fecha < inicio) {
        continue
      }


      // No considerar fechas futuras.
      if (fecha > fechaActual) {
        continue
      }


      const telFila =
        (r[0] || '')
          .replace(/\D/g, '')
          .slice(-10)


      if (telFila !== tel) {
        continue
      }


      // ------------------------------------------------------
      // E = estatus.
      //
      // Buscamos:
      // RETARDO 16min
      // RETARDO 25min
      // etc.
      // ------------------------------------------------------

      const estatus =
        String(r[4] || '')


      const match =
        estatus.match(
          /RETARDO\s+(\d+)/i
        )


      if (!match) {
        continue
      }


      const minutosRetardo =
        Number(match[1])


      if (
        !Number.isFinite(minutosRetardo) ||
        minutosRetardo <= TOLERANCIA_MIN
      ) {
        continue
      }


      retardos.push({
        fecha,
        minutos: minutosRetardo
      })
    }


    // --------------------------------------------------------
    // Ordenar por fecha.
    // --------------------------------------------------------

    retardos.sort(
      (a, b) =>
        a.fecha.localeCompare(b.fecha)
    )


    // --------------------------------------------------------
    // Todavía no llega a 3.
    // --------------------------------------------------------

    if (retardos.length < 3) {
      return
    }


    // --------------------------------------------------------
    // Tomar los 3 retardos que disparan el aviso.
    // --------------------------------------------------------

    const tres =
      retardos.slice(0, 3)


    const detalles =
      tres
        .map(
          r =>
            `${r.fecha}: ${r.minutos}min`
        )
        .join(' | ')


    // --------------------------------------------------------
    // Enviar aviso a GERENTES.
    // --------------------------------------------------------

    await sock.sendMessage(
      grupoAviso,
      {
        text:
          `⚠️ RETARDOS ACUMULADOS\n\n` +
          `${nombre}\n` +
          `Se registraron 3 retardos de más de ${TOLERANCIA_MIN} minutos ` +
          `en los últimos 30 días hábiles.\n\n` +
          `${detalles}\n\n` +
          `Favor de revisar y dar seguimiento.`
      }
    )


    // --------------------------------------------------------
    // GUARDAR EL AVISO.
    //
    // Esto reinicia el conteo.
    //
    // Los siguientes retardos deberán ser posteriores
    // a esta fecha para volver a llegar a 3.
    // --------------------------------------------------------

    await guardarUltimoAvisoRetardo({
      sClient,
      tel,
      nombre,
      fecha: fechaActual,
      detalles
    })

  } catch (e) {

    console.error(
      'Error revisando 3 retardos:',
      e.message
    )
  }
}


// ============================================================
// 5. REGISTRO DE ENTRADA / SALIDA
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
  // 5.1 OBTENER UBICACIÓN
  // ==========================================================

  const lat =
    loc.degreesLatitude

  const lng =
    loc.degreesLongitude


  let cercana = null
  let dMin = Infinity


  for (const s of SUCURSALES) {

    const d =
      distM(
        lat,
        lng,
        s.lat,
        s.lng
      )


    if (d < dMin) {

      dMin = d
      cercana = s
    }
  }


  // Si por alguna razón no se encontró sucursal.
  if (!cercana) {

    await sock.sendMessage(
      jid,
      {
        text:
          'No se pudo determinar la sucursal por GPS.'
      },
      { quoted: m }
    )

    return
  }


  // ==========================================================
  // 5.2 BUSCAR EMPLEADO
  // ==========================================================

  const empRows =
    await getRows('Empleados!A:K')


  const getLid = r =>
    (
      r.find(
        x =>
          String(x).includes('@lid')
      ) || ''
    ).trim()


  let emp = null


  // ----------------------------------------------------------
  // Buscar por teléfono.
  // ----------------------------------------------------------

  if (tel10.length >= 10) {

    const i =
      empRows.findIndex(
        (r, idx) =>
          idx > 0 &&
          r[0] &&
          r[0]
            .replace(/\D/g, '')
            .slice(-10) === tel10
      )


    if (i > -1) {
      emp = empRows[i]
    }
  }


  // ----------------------------------------------------------
  // Si no apareció por teléfono, buscar por LID.
  // ----------------------------------------------------------

  if (
    !emp &&
    rawLid.includes('@lid')
  ) {

    const i =
      empRows.findIndex(
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
  // 5.3 DATOS DEL EMPLEADO
  // ==========================================================

  const nombre =
    emp
      ? (emp[3] || emp[1])
      : (m.pushName || tel10 || 'Desconocido')


  const telF =
    emp
      ? (emp[0] || '').replace(/\D/g, '')
      : tel


  const tel10F =
    telF.slice(-10) || tel10


  // Sucursal asignada del empleado.
  //
  // Empleados:
  // C = Sucursal
  const sucursalAsignada =
    emp?.[2] || ''


  // ==========================================================
  // 5.4 FECHA
  // ==========================================================

  const fLab =
    fechaLaboral()


  // ----------------------------------------------------------
  // Reiniciar avisos diarios si cambió el día.
  // ----------------------------------------------------------

  if (fechaAvisos !== fLab) {

    avisosHoy.clear()
    fechaAvisos = fLab
  }


  if (fechaAvisos60 !== fLab) {

    avisos60Hoy.clear()
    fechaAvisos60 = fLab
  }


  // ==========================================================
  // 5.5 OBTENER ASISTENCIAS
  // ==========================================================

  const asis =
    await getRows('Asistencia!A:M')


  // ----------------------------------------------------------
  // Obtener todas las filas del empleado de hoy.
  // ----------------------------------------------------------

  const registrosHoy = []


  for (let i = 1; i < asis.length; i++) {

    const r = asis[i]


    const telFila =
      (r[0] || '')
        .replace(/\D/g, '')
        .slice(-10)


    if (
      telFila === tel10F &&
      r[2] === fLab
    ) {

      registrosHoy.push({
        idx: i,
        row: r
      })
    }
  }


  // ==========================================================
  // 5.6 OBTENER HORARIO
  // ==========================================================

  const sClient =
    await sheetsClient()


  const baseMap =
    await getHorarioBaseMap()


  const base =
    baseMap[tel10F] ||
    baseMap[
      normaliza(nombre)
        .split(' ')[0]
    ] ||
    null


  let estatus = 'A TIEMPO'

  let hObj = null

  let esRet = false

  let minRet = 0

  let hProg = ''


  // ==========================================================
  // 5.7 DETERMINAR HORARIO DEL DÍA
  // ==========================================================

  if (base) {

    const fe =
      new Date(
        `${fLab}T12:00:00`
      )


    const dn =
      fe.getDay()


    // --------------------------------------------------------
    // DESCANSO
    // --------------------------------------------------------

    if (base.descansos.has(dn)) {

      estatus = 'DESCANSO'


    // --------------------------------------------------------
    // Existe horario.
    // --------------------------------------------------------

    } else if (base.horas[dn]) {

      hObj =
        base.horas[dn]


      hProg =
        hObj.entrada || ''


      // ------------------------------------------------------
      // Horario normal.
      // ------------------------------------------------------

      if (
        hObj.entrada !== 'LIBRE' &&
        hObj.entrada
      ) {

        const dif =
          minutos(horaMX()) -
          minutos(hObj.entrada)


        if (dif > TOLERANCIA_MIN) {

          estatus =
            `RETARDO ${dif}min (Prog ${hObj.entrada})`

          esRet = true

          minRet = dif

        } else {

          estatus =
            `A TIEMPO (Prog ${hObj.entrada})`
        }
      }


      // ------------------------------------------------------
      // LIBRE.
      // ------------------------------------------------------

      else if (
        hObj.entrada === 'LIBRE'
      ) {

        estatus = 'LIBRE'
      }
    }
  }


  // ==========================================================
  // 5.8 DETERMINAR SI ES DÍA LIBRE
  // ==========================================================

  const esLibre =
    hObj?.entrada === 'LIBRE'


  // ==========================================================
  // 5.9 BUSCAR REGISTRO ABIERTO
  // ==========================================================
  //
  // Un registro abierto significa:
  //
  // Entrada existe
  // Salida todavía no existe
  //
  // En LIBRE esto permite:
  //
  // 09:00 -> entrada
  // 12:00 -> salida
  // 14:00 -> entrada
  // 18:00 -> salida
  //
  // Cada intervalo queda en su propia fila.
  // ==========================================================

  const abierto =
    registrosHoy.find(
      x =>
        x.row[3] &&
        !x.row[5]
    )


  const idxAbierto =
    abierto
      ? abierto.idx
      : -1


  // ----------------------------------------------------------
  // Para LIBRE:
  //
  // - Si existe una entrada abierta -> salida.
  // - Si no existe entrada abierta -> nueva entrada.
  //
  // Para horario normal:
  //
  // - Se utiliza el registro existente del día.
  // ----------------------------------------------------------

  let idx = -1
  let hoy = null


  if (esLibre) {

    if (idxAbierto > -1) {

      idx = idxAbierto
      hoy = asis[idx]

    } else {

      idx = -1
      hoy = null
    }

  } else {

    if (registrosHoy.length > 0) {

      idx =
        registrosHoy[0].idx

      hoy =
        asis[idx]
    }
  }


  // ==========================================================
  // 6. ENTRADA / SALIDA
  // ==========================================================

  try {

    // ========================================================
    // 6.1 ENTRADA
    // ========================================================

    if (!hoy || !hoy[3]) {

      // ------------------------------------------------------
      // Validar distancia.
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


      // ------------------------------------------------------
      // Hora actual.
      // ------------------------------------------------------

      const h =
        horaMX()


      // ------------------------------------------------------
      // Horario que se guarda en M.
      // ------------------------------------------------------

      let jTxt = '8h'


      if (esLibre) {

        jTxt = 'LIBRE'

      } else if (hObj) {

        jTxt =
          `${hObj.entrada}` +
          `${hObj.salida ? ` - ${hObj.salida}` : ''}`
      }


      // ------------------------------------------------------
      // En entrada todavía no hay horas trabajadas.
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
        '0',
        '0',
        jTxt
      ]


      // ------------------------------------------------------
      // Si no existe registro, crear fila.
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
      // Si existe registro pero no tenía entrada,
      // actualizar.
      // ------------------------------------------------------

      } else {

        await sClient.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range:
            `Asistencia!A${idx + 1}:M${idx + 1}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: {
            values: [row]
          }
        })
      }


      // ======================================================
      // CONFIRMACIÓN DE ENTRADA
      // ======================================================
      //
      // jid = grupo donde mandó ubicación.
      //
      // La sucursal mostrada es la determinada por GPS.
      // ======================================================

      await sock.sendMessage(
        jid,
        {
          text:
            `✅ ${estatus} - ${nombre} en ${cercana.nombre} - ${h}`
        }
      )


      // ======================================================
      // AVISO DE 60 MINUTOS O MÁS
      // ======================================================

      if (
        esRet &&
        minRet >= AVISO_RETARDO_60_MIN
      ) {

        const key60 =
          `${fLab}_${tel10F}`


        if (!avisos60Hoy.has(key60)) {

          avisos60Hoy.add(key60)


          const grupoAviso =
            grupoGerentesPorSucursal(
              sucursalAsignada
            )


          if (grupoAviso) {

            try {

              await sock.sendMessage(
                grupoAviso,
                {
                  text:
                    `⚠️ RETARDO DE 60 MINUTOS O MÁS\n\n` +
                    `${nombre}\n` +
                    `Entrada: ${h}\n` +
                    `Horario programado: ${hProg || 'N/D'}\n` +
                    `Retardo: ${minRet} minutos\n` +
                    `Sucursal asignada: ${sucursalAsignada || 'N/D'}\n\n` +
                    `Favor de confirmar si el empleado se queda a trabajar el turno.`
                }
              )

            } catch (e) {

              console.error(
                'Error enviando aviso de 60 min:',
                e.message
              )
            }
          }
        }
      }


      // ======================================================
      // REVISAR 3 RETARDOS EN 30 DÍAS HÁBILES
      // ======================================================

      if (esRet) {

        const grupoRetardos =
          grupoGerentesPorSucursal(
            sucursalAsignada
          )


        await revisarTresRetardos({
          sock,
          sClient,
          tel: tel10F,
          nombre,
          grupoAviso: grupoRetardos,
          fechaActual: fLab
        })
      }


    // ========================================================
    // 6.2 SALIDA
    // ========================================================

    } else {

      // ------------------------------------------------------
      // Si por alguna razón no tenemos registro abierto,
      // no permitir salida.
      // ------------------------------------------------------

      if (!hoy[3]) {

        await sock.sendMessage(
          jid,
          {
            text:
              'Primero debes registrar la entrada.'
          }
        )

        return
      }


      // ------------------------------------------------------
      // En horario normal:
      // si ya tiene salida, no permitir otra.
      //
      // En LIBRE no debería entrar aquí con salida porque
      // buscamos únicamente registros abiertos.
      // ------------------------------------------------------

      if (hoy[5]) {

        await sock.sendMessage(
          jid,
          {
            text:
              `Salida ya registrada`
          }
        )

        return
      }


      // ------------------------------------------------------
      // Validar distancia de salida.
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


      // ------------------------------------------------------
      // Hora de salida.
      // ------------------------------------------------------

      const h =
        horaMX()


      // ------------------------------------------------------
      // Recuperar horario guardado.
      // ------------------------------------------------------

      const jTxt =
        hoy[12] || (esLibre ? 'LIBRE' : '8h')


      // ======================================================
      // CALCULAR HORAS
      // ======================================================

      let trabajadas = '0'
      let extra = '0'


      // ------------------------------------------------------
      // LIBRE:
      //
      // Solo calculamos horas trabajadas.
      // No hay horas extra ni retardo.
      // ------------------------------------------------------

      if (esLibre) {

        trabajadas =
          calcularHorasTrabajadas(
            hoy[3],
            h
          )

        extra = '0'


      // ------------------------------------------------------
      // HORARIO NORMAL:
      //
      // calcularExtra conserva el cálculo de:
      // - horas trabajadas
      // - horas extra
      // ------------------------------------------------------

      } else {

        const resultado =
          calcularExtra(
            hoy[3],
            h,
            hObj?.entrada || null,
            hObj?.salida || null
          )


        trabajadas =
          resultado.trabajadas


        extra =
          resultado.extra
      }


      // ======================================================
      // ACTUALIZAR SHEETS
      // ======================================================

      await sClient.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range:
          `Asistencia!F${idx + 1}:M${idx + 1}`,
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


      // ======================================================
      // CONFIRMACIÓN DE SALIDA
      // ======================================================
      //
      // IMPORTANTE:
      // WhatsApp YA NO muestra Trab ni Extra.
      //
      // Esas cantidades permanecen en Sheets.
      // ======================================================

      await sock.sendMessage(
        jid,
        {
          text:
            `✅ Salida registrada - ${nombre} en ${cercana.nombre} - ${h}`
        }
      )
    }


  } catch (e) {

    console.error(
      'Error en handleChecador:',
      e
    )
  }
}


// ============================================================
// 7. AVISOS DE EMPLEADOS QUE NO HAN LLEGADO
// ============================================================
//
// Esta función se ejecuta periódicamente desde index.js.
//
// Revisa:
//
// - Horario programado.
// - Hora actual.
// - Si existe entrada.
//
// Si ya pasaron 20 minutos y no existe entrada,
// manda aviso al grupo de GERENTES.
//
// NO manda estos avisos a los grupos de CHECADORES.
// ============================================================

export async function checkNoLlegaron(sock) {

  try {

    // ========================================================
    // FECHA
    // ========================================================

    const hoy =
      fechaLaboral()


    // --------------------------------------------------------
    // Reiniciar avisos cuando cambia el día.
    // --------------------------------------------------------

    if (fechaAvisos !== hoy) {

      avisosHoy.clear()

      fechaAvisos = hoy
    }


    // ========================================================
    // HORA
    // ========================================================

    const ahoraMin =
      minutos(horaMX())


    // ========================================================
    // OBTENER HORARIOS
    // ========================================================

    const baseRows =
      await getRows('Horario_Base!A2:K')


    // ========================================================
    // OBTENER ASISTENCIAS
    // ========================================================

    const asisRows =
      await getRows('Asistencia!A2:M')


    // ========================================================
    // DÍA DE LA SEMANA
    // ========================================================

    const diaNum =
      new Date(
        new Date().toLocaleString(
          'en-US',
          {
            timeZone:
              'America/Mexico_City'
          }
        )
      ).getDay()


    // ========================================================
    // REVISAR CADA EMPLEADO
    // ========================================================

    for (const r of baseRows) {

      // ------------------------------------------------------
      // Teléfono.
      // ------------------------------------------------------

      const tel =
        (r[0] || '')
          .replace(/\D/g, '')
          .slice(-10)


      if (!tel) {
        continue
      }


      // ------------------------------------------------------
      // Llave para no repetir aviso.
      // ------------------------------------------------------

      const key =
        `${hoy}_${tel}`


      if (avisosHoy.has(key)) {
        continue
      }


      // ------------------------------------------------------
      // Obtener horario del día.
      //
      // Lunes = r[3]
      // Martes = r[4]
      // Miércoles = r[5]
      // Jueves = r[6]
      // Viernes = r[7]
      // Sábado = r[8]
      // Domingo = r[9]
      // ------------------------------------------------------

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


      const parsed =
        parseHorarioRango(v)


      // ------------------------------------------------------
      // Sin horario o LIBRE:
      // no generar aviso de falta.
      // ------------------------------------------------------

      if (
        !parsed ||
        parsed.entrada === 'LIBRE' ||
        !parsed.entrada
      ) {
        continue
      }


      // ======================================================
      // MINUTOS DE RETRASO
      // ======================================================

      const dif =
        ahoraMin -
        minutos(parsed.entrada)


      // ------------------------------------------------------
      // Todavía no han pasado 20 minutos.
      // ------------------------------------------------------

      if (dif < AVISO_FALTA_MIN) {
        continue
      }


      // ======================================================
      // COMPROBAR SI YA REGISTRÓ ENTRADA
      // ======================================================

      const registro =
        asisRows.find(
          a =>
            (a[0] || '')
              .replace(/\D/g, '')
              .slice(-10) === tel &&
            a[2] === hoy &&
            a[3]
        )


      // ------------------------------------------------------
      // Si ya registró entrada, no avisar.
      // ------------------------------------------------------

      if (registro) {
        continue
      }


      // ======================================================
      // PREPARAR AVISO
      // ======================================================

      avisosHoy.add(key)


      const nombre =
        r[1] || tel


      const sucursal =
        r[2] || ''


      // ======================================================
      // GRUPO DE GERENTES
      // ======================================================

      const grupoAviso =
        grupoGerentesPorSucursal(
          sucursal
        )


      // ======================================================
      // ENVIAR AVISO
      // ======================================================

      if (grupoAviso) {

        try {

          await sock.sendMessage(
            grupoAviso,
            {
              text:
                `⚠️ NO HA LLEGADO\n\n` +
                `${nombre}\n` +
                `Horario: ${parsed.entrada}\n` +
                `Retardo: ${dif} minutos\n` +
                `Sucursal: ${sucursal}`
            }
          )

        } catch (e) {

          console.error(
            'Error enviando aviso de no llegada:',
            e.message
          )
        }
      }
    }

  } catch (e) {

    console.error(
      'Error en checkNoLlegaron:',
      e
    )
  }
}

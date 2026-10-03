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
// CONFIGURACIÓN
// ============================================================

const TOLERANCIA_MIN = 15
const AVISO_FALTA_MIN = 20
const AVISO_RETARDO_60_MIN = 60

// A partir de esta fecha comienzan a contar los retardos
// para la alerta de 3 retardos.
const FECHA_CORTE_RETARDOS = '2026-10-02'

const HOJA_ALERTAS_RETARDOS = 'Alertas_Retardos'

// Control temporal de avisos de 60 minutos.
// Los avisos de NO HA LLEGADO usan la hoja Avisos,
// por lo que sobreviven a reinicios/deploys.
let avisos60Hoy = new Set()
let fechaAvisos60 = ''


// ============================================================
// FUNCIONES AUXILIARES
// ============================================================

function esDiaHabil(fecha) {
  const d = new Date(`${fecha}T12:00:00`).getDay()
  return d >= 1 && d <= 5
}


function inicioVentana30DiasHabiles(fechaFinal) {
  let actual = new Date(`${fechaFinal}T12:00:00`)
  let contados = 1

  while (contados < 30) {
    actual.setDate(actual.getDate() - 1)

    const f =
      `${actual.getFullYear()}-` +
      `${String(actual.getMonth() + 1).padStart(2, '0')}-` +
      `${String(actual.getDate()).padStart(2, '0')}`

    if (esDiaHabil(f)) {
      contados++
    }
  }

  return (
    `${actual.getFullYear()}-` +
    `${String(actual.getMonth() + 1).padStart(2, '0')}-` +
    `${String(actual.getDate()).padStart(2, '0')}`
  )
}


function normalizaFechaCelda(v) {
  const s = (v || '').toString().trim()

  if (!s) return ''

  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    return s.slice(0, 10)
  }

  const m = s.match(
    /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/
  )

  if (m) {
    const dd = m[1].padStart(2, '0')
    const mm = m[2].padStart(2, '0')
    const yyyy =
      m[3].length === 2
        ? `20${m[3]}`
        : m[3]

    return `${yyyy}-${mm}-${dd}`
  }

  return s.slice(0, 10)
}


function tel10De(v) {
  return (v || '')
    .toString()
    .replace(/\D/g, '')
    .slice(-10)
}


// ============================================================
// GRUPO DE GERENTES SEGÚN SUCURSAL ASIGNADA
// ============================================================

function grupoGerentesPorSucursal(sucursal) {
  const s = (sucursal || '').toLowerCase()

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
    s.includes('juarez') ||
    s.includes('júarez')
  ) {
    return GRUPO_BUCARELI_ID
  }

  return null
}


// ============================================================
// CREAR AUTOMÁTICAMENTE Alertas_Retardos
// ============================================================

async function asegurarHojaAlertasRetardos(sClient) {
  const meta =
    await sClient.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
      fields: 'sheets.properties.title'
    })

  const existe =
    (meta.data.sheets || []).some(
      s =>
        s.properties?.title ===
        HOJA_ALERTAS_RETARDOS
    )

  if (existe) {
    return
  }

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

  await sClient.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range:
      `${HOJA_ALERTAS_RETARDOS}!A1:G1`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[
        'Tel',
        'Nombre',
        'Sucursal',
        'FechaAviso',
        'HoraAviso',
        'RetardosConsiderados',
        'DetalleRetardos'
      ]]
    }
  })
}


// ============================================================
// OBTENER ÚLTIMO AVISO DE 3 RETARDOS
// ============================================================

async function getUltimoAvisoRetardo(
  sClient,
  tel10
) {
  const rows =
    await sClient.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range:
        `${HOJA_ALERTAS_RETARDOS}!A2:G`
    }).then(
      r => r.data.values || []
    )

  let ultimo = null

  for (const r of rows) {
    if (tel10De(r[0]) !== tel10) {
      continue
    }

    const fecha =
      normalizaFechaCelda(r[3])

    if (!fecha) {
      continue
    }

    if (
      !ultimo ||
      fecha > ultimo.fecha
    ) {
      ultimo = {
        fecha,
        hora:
          (r[4] || '')
            .toString()
            .trim()
      }
    }
  }

  return ultimo
}


// ============================================================
// GUARDAR AVISO DE 3 RETARDOS
// ============================================================

async function guardarAvisoRetardo(
  sClient,
  {
    tel,
    nombre,
    sucursal,
    fecha,
    hora,
    detalles
  }
) {
  await sClient.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range:
      `${HOJA_ALERTAS_RETARDOS}!A:G`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[
        tel,
        nombre,
        sucursal,
        fecha,
        hora,
        '3',
        detalles.join(' | ')
      ]]
    }
  })
}


// ============================================================
// REVISAR 3 RETARDOS EN 30 DÍAS HÁBILES
// ============================================================

async function revisarTresRetardos({
  sock,
  sClient,
  tel10,
  nombre,
  sucursal
}) {
  if (!tel10) {
    return
  }

  const grupoAviso =
    grupoGerentesPorSucursal(
      sucursal
    )

  if (!grupoAviso) {
    return
  }

  await asegurarHojaAlertasRetardos(
    sClient
  )

  const ultimoAviso =
    await getUltimoAvisoRetardo(
      sClient,
      tel10
    )

  const fechaHoy =
    fechaLaboral()

  // Antes del corte no cuenta.
  if (
    fechaHoy <
    FECHA_CORTE_RETARDOS
  ) {
    return
  }

  const inicioVentana =
    inicioVentana30DiasHabiles(
      fechaHoy
    )

  const asisRows =
    await getRows(
      'Asistencia!A2:M'
    )

  const retardos =
    asisRows
      .map(r => {
        const telFila =
          tel10De(r[0])

        const fecha =
          normalizaFechaCelda(r[2])

        const match =
          (r[4] || '')
            .toString()
            .match(
              /RETARDO\s+(\d+)\s*min/i
            )

        if (
          !telFila ||
          telFila !== tel10
        ) {
          return null
        }

        if (
          !fecha ||
          !match
        ) {
          return null
        }

        if (
          fecha <
          FECHA_CORTE_RETARDOS
        ) {
          return null
        }

        if (
          fecha <
            inicioVentana ||
          fecha >
            fechaHoy
        ) {
          return null
        }

        if (
          ultimoAviso &&
          fecha <=
            ultimoAviso.fecha
        ) {
          return null
        }

        if (!esDiaHabil(fecha)) {
          return null
        }

        return {
          fecha,
          minutos:
            parseInt(
              match[1],
              10
            )
        }
      })
      .filter(Boolean)
      .sort(
        (a, b) =>
          a.fecha.localeCompare(
            b.fecha
          )
      )

  if (retardos.length < 3) {
    return
  }

  const primerosTres =
    retardos.slice(0, 3)

  const detalles =
    primerosTres.map(
      r =>
        `${r.fecha}: ${r.minutos} min tarde`
    )

  const horaAviso =
    horaMX()

  try {
    await sock.sendMessage(
      grupoAviso,
      {
        text:
          `🚨 *3 RETARDOS*\n` +
          `${nombre}\n` +
          `Sucursal: ${sucursal}\n` +
          `Se acumularon 3 retardos dentro de los últimos 30 días hábiles.\n` +
          `• ${detalles.join('\n• ')}\n\n` +
          `Favor de dar seguimiento.`
      }
    )

    await guardarAvisoRetardo(
      sClient,
      {
        tel: tel10,
        nombre,
        sucursal,
        fecha: fechaHoy,
        hora: horaAviso,
        detalles
      }
    )
  } catch (e) {
    console.error(
      'ERROR AVISO 3 RETARDOS:',
      e.message
    )
  }
}


// ============================================================
// REGISTRAR AVISO DE NO LLEGADA EN HOJA Avisos
//
// Avisos actual:
//
// A = Tel
// B = Fecha
// C = HoraAviso
// ============================================================

async function registrarAvisoNoLlegada(
  sClient,
  tel10,
  fechaHoy,
  horaAviso
) {
  try {
    await sClient.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Avisos!A:C',
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[
          tel10,
          fechaHoy,
          horaAviso
        ]]
      }
    })

    return true
  } catch (e) {
    console.error(
      'ERROR GUARDANDO Avisos:',
      e.message
    )

    return false
  }
}


// ============================================================
// CHECADOR
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
  const lat =
    loc.degreesLatitude

  const lng =
    loc.degreesLongitude

  let cercana = null
  let dMin = Infinity

  // ==========================================================
  // SUCURSAL MÁS CERCANA
  // ==========================================================

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

  if (!cercana) {
    await sock.sendMessage(
      jid,
      {
        text:
          'No se pudo determinar la sucursal más cercana.'
      }
    )

    return
  }


  // ==========================================================
  // BUSCAR EMPLEADO
  // ==========================================================

  const empRows =
    await getRows(
      'Empleados!A:K'
    )

  const getLid = r =>
    (
      r.find(
        x =>
          String(x)
            .includes('@lid')
      ) || ''
    ).trim()

  let emp = null


  // Buscar por teléfono.
  if (tel10.length >= 10) {
    const i =
      empRows.findIndex(
        (r, idx) =>
          idx > 0 &&
          r[0] &&
          tel10De(r[0]) ===
            tel10
      )

    if (i > -1) {
      emp = empRows[i]
    }
  }


  // Buscar por LID.
  if (
    !emp &&
    rawLid.includes('@lid')
  ) {
    const i =
      empRows.findIndex(
        (r, idx) =>
          idx > 0 &&
          getLid(r) ===
            rawLid
      )

    if (i > -1) {
      emp = empRows[i]

      tel10 =
        tel10De(
          emp[0]
        )
    }
  }


  const nombre =
    emp
      ? (
          emp[3] ||
          emp[1]
        )
      : (
          m.pushName ||
          tel10 ||
          'Desconocido'
        )

  const telF =
    emp
      ? (
          emp[0] || ''
        ).replace(
          /\D/g,
          ''
        )
      : tel

  const tel10F =
    telF.slice(-10) ||
    tel10

  // Sucursal asignada en Empleados.
  const sucursalEmpleado =
    emp?.[2] || ''

  const fLab =
    fechaLaboral()


  // ==========================================================
  // LEER ASISTENCIA
  // ==========================================================

  const asis =
    await getRows(
      'Asistencia!A:M'
    )


  // LIBRE:
  // buscamos un intervalo abierto.
  const idxAbiertoLibre =
    asis.findIndex(
      (r, i) =>
        i > 0 &&
        tel10De(r[0]) ===
          tel10F &&
        r[2] === fLab &&
        r[3] &&
        !r[5]
    )


  // Jornada definida:
  // buscamos la fila del día.
  const idxDia =
    asis.findIndex(
      (r, i) =>
        i > 0 &&
        tel10De(r[0]) ===
          tel10F &&
        r[2] === fLab
    )


  const sClient =
    await sheetsClient()

  const baseMap =
    await getHorarioBaseMap()

  const base =
    baseMap[tel10F] ||
    baseMap[
      normaliza(
        nombre
      ).split(' ')[0]
    ] ||
    null


  // ==========================================================
  // DETERMINAR HORARIO
  // ==========================================================

  let estatus =
    'A TIEMPO'

  let hObj = null

  let esRet = false

  let minRet = 0


  if (base) {
    const fe =
      new Date(
        `${fLab}T12:00:00`
      )

    const dn =
      fe.getDay()

    if (
      base.descansos.has(dn)
    ) {
      estatus =
        'DESCANSO'
    }
    else if (
      base.horas[dn]
    ) {
      hObj =
        base.horas[dn]

      if (
        hObj.entrada ===
        'LIBRE'
      ) {
        estatus =
          'LIBRE'
      }
      else if (
        hObj.entrada
      ) {
        const dif =
          minutos(
            horaMX()
          ) -
          minutos(
            hObj.entrada
          )

        if (
          dif >
          TOLERANCIA_MIN
        ) {
          estatus =
            `RETARDO ${dif}min (Prog ${hObj.entrada})`

          esRet = true
          minRet = dif
        }
        else {
          estatus =
            `A TIEMPO (Prog ${hObj.entrada})`
        }
      }
    }
  }


  try {

    // ========================================================
    // ¿LIBRE?
    // ========================================================

    const esLibre =
      hObj?.entrada ===
      'LIBRE'

    let esEntrada = false

    let idxTrabajo = -1


    if (esLibre) {

      // Sin intervalo abierto:
      // Entrada.
      if (
        idxAbiertoLibre === -1
      ) {
        esEntrada = true
      }
      else {
        // Intervalo abierto:
        // Salida.
        idxTrabajo =
          idxAbiertoLibre
      }

    }
    else {

      // ======================================================
      // JORNADA DEFINIDA
      // ======================================================

      if (
        idxDia === -1 ||
        !asis[idxDia]?.[3]
      ) {
        esEntrada = true
      }

      else if (
        asis[idxDia]?.[5]
      ) {
        await sock.sendMessage(
          jid,
          {
            text:
              'Salida ya registrada'
          }
        )

        return
      }

      else {
        idxTrabajo =
          idxDia
      }
    }


    // ========================================================
    // ENTRADA
    // ========================================================

    if (esEntrada) {

      // Validación GPS de entrada.
      if (
        dMin >
        cercana.rEnt
      ) {
        await sock.sendMessage(
          jid,
          {
            text:
              `Debes estar a max ${cercana.rEnt}m de ${cercana.nombre}`
          },
          {
            quoted: m
          }
        )

        return
      }

      const h =
        horaMX()

      const jTxt =
        hObj
          ? (
              hObj.entrada ===
              'LIBRE'
                ? 'LIBRE'
                : (
                    `${hObj.entrada}` +
                    `${
                      hObj.salida
                        ? ` - ${hObj.salida}`
                        : ''
                    }`
                  )
            )
          : '8h'

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
        calcularHorasTrabajadas(
          h,
          ''
        ),
        '0',
        jTxt
      ]


      // ======================================================
      // GUARDAR ENTRADA
      // ======================================================

      // LIBRE:
      // siempre crea una nueva fila.
      //
      // Jornada definida:
      // si no hay fila, crea.
      // si existe fila sin entrada, actualiza.
      if (
        esLibre ||
        idxDia === -1
      ) {
        await sClient.spreadsheets.values.append({
          spreadsheetId:
            SPREADSHEET_ID,
          range:
            'Asistencia!A:M',
          valueInputOption:
            'USER_ENTERED',
          requestBody: {
            values: [row]
          }
        })
      }
      else {
        await sClient.spreadsheets.values.update({
          spreadsheetId:
            SPREADSHEET_ID,
          range:
            `Asistencia!A${idxDia + 1}:M${idxDia + 1}`,
          valueInputOption:
            'USER_ENTERED',
          requestBody: {
            values: [row]
          }
        })
      }


      // ======================================================
      // CONFIRMACIÓN DE ENTRADA
      // ======================================================

      await sock.sendMessage(
        jid,
        {
          text:
            `✅ ${estatus} - ${nombre} en ${cercana.nombre} - ${h}`
        }
      )


      // ======================================================
      // ALERTA DE 60+ MINUTOS
      // ======================================================

      if (
        esRet &&
        minRet >=
          AVISO_RETARDO_60_MIN
      ) {

        if (
          fechaAvisos60 !==
          fLab
        ) {
          avisos60Hoy.clear()
          fechaAvisos60 =
            fLab
        }

        const key60 =
          `${fLab}_${tel10F}`

        if (
          !avisos60Hoy.has(
            key60
          )
        ) {

          const grupoGerentes =
            grupoGerentesPorSucursal(
              sucursalEmpleado
            )

          if (grupoGerentes) {

            avisos60Hoy.add(
              key60
            )

            try {
              await sock.sendMessage(
                grupoGerentes,
                {
                  text:
                    `⚠️ *RETARDO DE ${minRet} MIN*\n` +
                    `${nombre}\n` +
                    `Sucursal: ${sucursalEmpleado || 'No indicada'}\n` +
                    `Entrada programada: ${hObj?.entrada || 'N/D'}\n` +
                    `Entrada registrada: ${h}\n\n` +
                    `Favor de confirmar si el empleado se queda / trabaja su turno.`
                }
              )
            }
            catch (e) {
              console.error(
                'ERROR AVISO 60 MIN:',
                e.message
              )
            }
          }
        }
      }


      // ======================================================
      // REVISAR 3 RETARDOS
      // ======================================================

      if (esRet) {
        await revisarTresRetardos({
          sock,
          sClient,
          tel10: tel10F,
          nombre,
          sucursal:
            sucursalEmpleado
        })
      }

      return
    }


    // ========================================================
    // SALIDA
    // ========================================================

    const idxSalida =
      idxTrabajo

    const hoy =
      idxSalida > -1
        ? asis[idxSalida]
        : null

    if (
      !hoy ||
      !hoy[3]
    ) {
      await sock.sendMessage(
        jid,
        {
          text:
            'No hay una entrada registrada para esta salida.'
        }
      )

      return
    }


    if (hoy[5]) {
      await sock.sendMessage(
        jid,
        {
          text:
            'Salida ya registrada'
        }
      )

      return
    }


    // Validación GPS de salida.
    if (
      dMin >
      cercana.rSal
    ) {
      await sock.sendMessage(
        jid,
        {
          text:
            `No puedes checar salida a ${Math.round(dMin)}m`
        },
        {
          quoted: m
        }
      )

      return
    }


    const h =
      horaMX()

    const jTxt =
      hoy[12] ||
      (
        esLibre
          ? 'LIBRE'
          : '8h'
      )


    let trabajadas =
      '0'

    let extra =
      '0'


    // ======================================================
    // LIBRE
    // ======================================================

    if (esLibre) {

      trabajadas =
        calcularHorasTrabajadas(
          hoy[3],
          h
        )

      extra =
        '0'
    }

    // ======================================================
    // JORNADA DEFINIDA
    // ======================================================

    else {

      const calc =
        calcularExtra(
          hoy[3],
          h,
          hObj?.entrada ||
            null,
          hObj?.salida ||
            null
        )

      trabajadas =
        calc.trabajadas

      extra =
        calc.extra
    }


    // ======================================================
    // GUARDAR SALIDA
    // ======================================================

    await sClient.spreadsheets.values.update({
      spreadsheetId:
        SPREADSHEET_ID,
      range:
        `Asistencia!F${idxSalida + 1}:M${idxSalida + 1}`,
      valueInputOption:
        'USER_ENTERED',
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
    // WHATSAPP
    // ======================================================

    await sock.sendMessage(
      jid,
      {
        text:
          `✅ Salida - ${nombre} en ${cercana.nombre} - ${h}`
      }
    )

  }
  catch (e) {
    console.error(
      'ERROR handleChecador:',
      e
    )
  }
}


// ============================================================
// REVISAR QUIÉN NO HA LLEGADO
//
// Avisos:
// A = Tel
// B = Fecha
// C = HoraAviso
//
// Esto hace que un deploy/reinicio NO repita la alerta
// del mismo empleado el mismo día.
// ============================================================

export async function checkNoLlegaron(sock) {

  const hoy =
    fechaLaboral()

  const ahoraMin =
    minutos(horaMX())


  // Reset del control de avisos de 60 minutos.
  if (
    fechaAvisos60 !==
    hoy
  ) {
    avisos60Hoy.clear()
    fechaAvisos60 =
      hoy
  }


  const sClient =
    await sheetsClient()


  // ==========================================================
  // CREAR Alertas_Retardos AUTOMÁTICAMENTE
  // ==========================================================

  try {
    await asegurarHojaAlertasRetardos(
      sClient
    )
  }
  catch (e) {
    console.error(
      'ERROR CREANDO Alertas_Retardos:',
      e.message
    )
  }


  // ==========================================================
  // LEER DATOS
  // ==========================================================

  const [
    baseRows,
    asisRows,
    avisosRows
  ] =
    await Promise.all([
      getRows(
        'Horario_Base!A2:K'
      ),

      getRows(
        'Asistencia!A2:M'
      ),

      getRows(
        'Avisos!A:C'
      )
    ])


  // ==========================================================
  // AVISOS YA ENVIADOS HOY
  // ==========================================================

  const avisosNoLlegadaHoy =
    new Set(
      avisosRows
        .filter(
          r =>
            normalizaFechaCelda(
              r[1]
            ) === hoy
        )
        .map(
          r =>
            tel10De(r[0])
        )
        .filter(Boolean)
    )


  // ==========================================================
  // DÍA DE LA SEMANA
  // ==========================================================

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


  // ==========================================================
  // REVISAR CADA EMPLEADO
  // ==========================================================

  for (const r of baseRows) {

    const tel =
      tel10De(r[0])

    if (!tel) {
      continue
    }


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
      (
        mapa[diaNum] ||
        ''
      )
        .toString()
        .trim()


    const parsed =
      parseHorarioRango(v)


    // Sin horario / LIBRE / DESCANSO
    if (
      !parsed ||
      parsed.entrada ===
        'LIBRE' ||
      !parsed.entrada
    ) {
      continue
    }


    const dif =
      ahoraMin -
      minutos(
        parsed.entrada
      )


    // Todavía no pasan 20 min.
    if (
      dif <
      AVISO_FALTA_MIN
    ) {
      continue
    }


    // ========================================================
    // ¿YA REGISTRÓ ENTRADA?
    // ========================================================

    const registro =
      asisRows.find(
        a =>
          tel10De(a[0]) ===
            tel &&
          a[2] === hoy &&
          a[3]
      )


    if (registro) {
      continue
    }


    // ========================================================
    // ¿YA TIENE AVISO HOY?
    // ========================================================

    if (
      avisosNoLlegadaHoy.has(
        tel
      )
    ) {
      continue
    }


    const nombre =
      r[1] || tel

    const sucursalAsignada =
      r[2] || ''


    // ========================================================
    // GRUPO DE GERENTES
    // ========================================================

    const grupoAviso =
      grupoGerentesPorSucursal(
        sucursalAsignada
      )


    if (!grupoAviso) {
      console.log(
        `SIN GRUPO DE GERENTES PARA NO LLEGADA: ${nombre} - ${sucursalAsignada}`
      )

      continue
    }


    const horaAviso =
      horaMX()


    try {

      // ======================================================
      // MENSAJE
      // ======================================================

      await sock.sendMessage(
        grupoAviso,
        {
          text:
            `⚠️ *NO HA LLEGADO* - ${nombre}\n` +
            `Prog: ${parsed.entrada}\n` +
            `Más de ${AVISO_FALTA_MIN} min sin registrar entrada\n` +
            `[${sucursalAsignada}]`
        }
      )


      // ======================================================
      // GUARDAR EN Avisos
      // ======================================================

      const guardado =
        await registrarAvisoNoLlegada(
          sClient,
          tel,
          hoy,
          horaAviso
        )


      if (guardado) {
        avisosNoLlegadaHoy.add(
          tel
        )
      }

    }
    catch (e) {
      console.error(
        'ERROR AVISO NO LLEGADA:',
        e.message
      )
    }
  }
}

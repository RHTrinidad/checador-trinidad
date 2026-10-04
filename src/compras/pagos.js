import { downloadMediaMessage } from '@whiskeysockets/baileys'
import P from 'pino'
import {
  SPREADSHEET_COMPRAS_ID,
  GRUPO_PRUEBAS_ID
} from '../config.js'
import { fechaLaboral, normaliza } from '../utils.js'
import { sheetsClient, getRows } from '../sheets.js'

const OPENAI_API_KEY = process.env.OPENAI_API_KEY
const SHEET_PAGOS = 'Pagos'

function limpiarTexto(valor){
  return (valor || '').toString().trim()
}

function limpiarMonto(valor){
  const limpio = (valor || '')
    .toString()
    .replace(/[$,\s]/g,'')
    .trim()

  if(!limpio) return ''

  const numero = parseFloat(limpio)

  return Number.isNaN(numero)
    ? ''
    : numero.toFixed(2)
}

async function leerPagoConOpenAI(bufferImagen){

  try{

    const base64 =
      bufferImagen.toString('base64')

    const res = await fetch(
      'https://api.openai.com/v1/chat/completions',
      {
        method:'POST',
        headers:{
          'Authorization':`Bearer ${OPENAI_API_KEY}`,
          'Content-Type':'application/json'
        },
        body:JSON.stringify({
          model:'gpt-4o',
          messages:[{
            role:'user',
            content:[
              {
                type:'text',
                text:`Eres lector de comprobantes de pago para restaurante Trinidad.

Analiza esta imagen.

Devuelve SOLO JSON válido:

{
  "tipo_documento": "",
  "fecha_pago": "",
  "destinatario": "",
  "monto": "",
  "banco": "",
  "cuenta_destino": "",
  "referencia": "",
  "concepto": ""
}

REGLAS:

- tipo_documento: "PAGO" únicamente si es claramente un comprobante de pago. Si no, "NO_PAGO".
- fecha_pago: fecha visible del pago en formato YYYY-MM-DD. Si no existe o no es legible, "".
- destinatario: nombre visible de la persona o empresa que recibe el pago.
- monto: importe realmente pagado, solo número.
- banco: banco visible. Si no es claro, "".
- cuenta_destino: cuenta, CLABE o tarjeta visible. Si aparece parcialmente, conserva únicamente lo visible.
- referencia: referencia, folio, operación o clave de rastreo visible.
- concepto: concepto visible relacionado con el pago.
- NO inventes información.
- NO completes números faltantes.
- Si un dato no aparece o no es legible, devuelve "".
`
              },
              {
                type:'image_url',
                image_url:{
                  url:`data:image/jpeg;base64,${base64}`
                }
              }
            ]
          }],
          max_tokens:1200,
          temperature:0.1
        })
      }
    )

    const data = await res.json()

    const content =
      data.choices?.[0]?.message?.content || ''

    console.log('OpenAI PAGOS RAW:',content)

    const match =
      content.match(/\{[\s\S]*\}/)

    if(!match) return null

    const json = JSON.parse(match[0])

    return {
      tipo_documento:
        limpiarTexto(json.tipo_documento).toUpperCase(),

      fecha_pago:
        limpiarTexto(json.fecha_pago),

      destinatario:
        limpiarTexto(json.destinatario),

      monto:
        limpiarMonto(json.monto),

      banco:
        limpiarTexto(json.banco),

      cuenta_destino:
        limpiarTexto(json.cuenta_destino),

      referencia:
        limpiarTexto(json.referencia),

      concepto:
        limpiarTexto(json.concepto)
    }

  }catch(e){

    console.log(
      'OpenAI PAGOS:',
      e.message
    )

    return null
  }
}

async function buscarDuplicado(datos){

  try{

    const rows =
      await getRows(
        `${SHEET_PAGOS}!A2:H`,
        SPREADSHEET_COMPRAS_ID
      )

    const fechaPago =
      normaliza(
        datos.fechaPago || ''
      ).toUpperCase()

    const destinatario =
      normaliza(
        datos.destinatario || ''
      ).toUpperCase()

    const monto =
      parseFloat(
        datos.monto || '0'
      ) || 0

    const referencia =
      normaliza(
        datos.referencia || ''
      ).toUpperCase()

    const banco =
      normaliza(
        datos.banco || ''
      ).toUpperCase()

    if(
      !fechaPago ||
      !destinatario ||
      !monto
    ){
      return false
    }

    return rows.some(r=>{

      const rFechaPago =
        normaliza(r[1] || '').toUpperCase()

      const rDestinatario =
        normaliza(r[2] || '').toUpperCase()

      const rMonto =
        parseFloat(
          (r[3] || '0')
            .toString()
            .replace(/[$,]/g,'')
        ) || 0

      const rBanco =
        normaliza(r[4] || '').toUpperCase()

      const rReferencia =
        normaliza(r[6] || '').toUpperCase()

      if(rFechaPago !== fechaPago)
        return false

      if(rDestinatario !== destinatario)
        return false

      if(
        Math.abs(rMonto - monto) >= 0.01
      ){
        return false
      }

      if(
        referencia &&
        rReferencia
      ){
        return referencia === rReferencia
      }

      return (
        banco &&
        rBanco &&
        banco === rBanco
      )
    })

  }catch(e){

    console.log(
      'Error buscando duplicado de pago:',
      e.message
    )

    return false
  }
}

async function obtenerSheetId(
  sClient,
  nombreHoja
){

  const meta =
    await sClient.spreadsheets.get({
      spreadsheetId:
        SPREADSHEET_COMPRAS_ID,
      fields:'sheets.properties'
    })

  const hoja =
    meta.data.sheets?.find(
      s =>
        s.properties?.title === nombreHoja
    )

  return hoja?.properties?.sheetId ?? 0
}

async function marcarFilaDuplicada(rowNumber){

  try{

    const sClient =
      await sheetsClient()

    await sClient.spreadsheets.batchUpdate({

      spreadsheetId:
        SPREADSHEET_COMPRAS_ID,

      requestBody:{
        requests:[{
          repeatCell:{
            range:{
              sheetId:
                await obtenerSheetId(
                  sClient,
                  SHEET_PAGOS
                ),
              startRowIndex:rowNumber - 1,
              endRowIndex:rowNumber,
              startColumnIndex:0,
              endColumnIndex:8
            },
            cell:{
              userEnteredFormat:{
                backgroundColor:{
                  red:1,
                  green:1,
                  blue:0.6
                }
              }
            },
            fields:
              'userEnteredFormat.backgroundColor'
          }
        }]
      }
    })

  }catch(e){

    console.log(
      'Error marcando pago duplicado:',
      e.message
    )
  }
}

async function registrarPago(datos){

  const sClient =
    await sheetsClient()

  const duplicado =
    await buscarDuplicado(datos)

  await sClient.spreadsheets.values.append({

    spreadsheetId:
      SPREADSHEET_COMPRAS_ID,

    range:
      `${SHEET_PAGOS}!A:H`,

    valueInputOption:
      'USER_ENTERED',

    requestBody:{
      values:[[
        datos.fechaCaptura,
        datos.fechaPago,
        datos.destinatario,
        datos.monto,
        datos.banco,
        datos.cuentaDestino,
        datos.referencia,
        datos.concepto
      ]]
    }
  })

  if(duplicado){

    try{

      const rowsActuales =
        await getRows(
          `${SHEET_PAGOS}!A2:H`,
          SPREADSHEET_COMPRAS_ID
        )

      await marcarFilaDuplicada(
        rowsActuales.length + 1
      )

    }catch(e){

      console.log(
        'Error marcando fila duplicada:',
        e.message
      )
    }
  }

  return {
    ok:true,
    duplicado,
    msg:
      `✅ Pago registrado · ` +
      `Fecha ${datos.fechaPago || 'SIN FECHA'} · ` +
      `${datos.destinatario || 'SIN DESTINATARIO'} · ` +
      `$${datos.monto || '0.00'}`
  }
}

export async function handlePagos({
  sock,
  jid,
  m,
  texto,
  esImagen
}){

  if(
    esImagen &&
    jid !== GRUPO_PRUEBAS_ID
  ){
    return false
  }

  if(!esImagen)
    return false

  try{

    if(!OPENAI_API_KEY){

      await sock.sendMessage(
        jid,
        {
          text:
            '❌ Falta OPENAI_API_KEY en Railway'
        }
      )

      return true
    }

    const buffer =
      await downloadMediaMessage(
        m,
        'buffer',
        {},
        {
          logger:P({
            level:'fatal'
          }),
          reuploadRequest:
            sock.updateMediaMessage
        }
      )

    const datosIA =
      await leerPagoConOpenAI(buffer)

    if(!datosIA){

      await sock.sendMessage(
        jid,
        {
          text:
            '❌ No pude analizar el comprobante de pago. Manda una foto más clara.'
        }
      )

      return true
    }

    if(
      datosIA.tipo_documento !== 'PAGO'
    ){
      return false
    }

    if(!datosIA.monto){

      await sock.sendMessage(
        jid,
        {
          text:
            '❌ No pude leer el monto del pago. Manda una foto más clara.'
        }
      )

      return true
    }

    const datosFinal = {

      fechaCaptura:
        fechaLaboral(),

      fechaPago:
        datosIA.fecha_pago || '',

      destinatario:
        datosIA.destinatario || '',

      monto:
        datosIA.monto || '',

      banco:
        datosIA.banco || '',

      cuentaDestino:
        datosIA.cuenta_destino || '',

      referencia:
        datosIA.referencia || '',

      concepto:
        datosIA.concepto || ''
    }

    const res =
      await registrarPago(
        datosFinal
      )

    await sock.sendMessage(
      jid,
      {
        text:
          res.msg +
          (
            res.duplicado
              ? '\n⚠️ POSIBLE DUPLICADO — fila marcada en amarillo.'
              : ''
          )
      }
    )

    return true

  }catch(e){

    console.error(
      'Error procesando pago:',
      e
    )

    await sock.sendMessage(
      jid,
      {
        text:
          '❌ Error al registrar pago: ' +
          e.message
      }
    )

    return true
  }
}

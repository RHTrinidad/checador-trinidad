import { downloadMediaMessage } from '@whiskeysockets/baileys'
import P from 'pino'
import { SPREADSHEET_COMPRAS_ID, SHEET_RESUMEN, SHEET_INSUMOS, GRUPO_PRUEBAS_ID } from '../config.js'
import { fechaLaboral, normaliza } from '../utils.js'
import { sheetsClient, getRows } from '../sheets.js'
import ExcelJS from 'exceljs'
import fs from 'fs'
import path from 'path'
import os from 'os'

const OPENAI_API_KEY = process.env.OPENAI_API_KEY

const CONCEPTOS_VALIDOS = [
  'ABARROTES',
  'BAÑOS',
  'BASURA',
  'BEBIDA SIN ALCOHOL',
  'CARNE DE CERDO',
  'CARNE DE RES',
  'CERVEZAS Y AGUAMIEL',
  'CHAPULINES',
  'CHICHARRON DURO',
  'COCAS',
  'COMIDA PERSONAL',
  'DESECHABLES',
  'DESTILADOS',
  'ESTACIONAMIENTO',
  'FLORES',
  'FRUTAS Y VERDURAS',
  'FUMIGACION',
  'GARRAFONES',
  'GAS LP',
  'GASTO',
  'HIELO',
  'INSECTOS',
  'JUGOS',
  'LACTEOS',
  'LIMPIEZA',
  'MASA Y TORTILLA',
  'MEDICAMENTOS',
  'PAPELERIA Y COPIAS',
  'PESCADOS Y MARISCOS',
  'POLLO Y PATO',
  'PRENSADO',
  'RENTA',
  'REPARACIONES',
  'UNIFORMES',
  'VIATICOS',
  'VIGILACIA'
]

function obtenerConceptoValido(valor){
  const x = normaliza(valor || '').toUpperCase().trim()

  if(!x) return ''

  const encontrado = CONCEPTOS_VALIDOS.find(c =>
    normaliza(c).toUpperCase() === x
  )

  return encontrado || ''
}

function obtenerSemanaCompras(fecha){
  const d = new Date(`${fecha}T12:00:00`)
  if(Number.isNaN(d.getTime())) return ''

  /*
    La semana empieza en martes y termina en lunes.

    La primera semana del año es la que contiene
    el 1 de enero.

    Para 2026:
    01-26 = martes 30/12/2025 a lunes 05/01/2026
    40-26 = martes 29/09/2026 a lunes 05/10/2026
  */

  const anio = d.getFullYear()

  const inicioAnio = new Date(anio, 0, 1, 12)

  const diaSemana = inicioAnio.getDay()

  const diasDesdeMartes = (diaSemana - 2 + 7) % 7

  const inicioSemana01 = new Date(inicioAnio)
  inicioSemana01.setDate(inicioSemana01.getDate() - diasDesdeMartes)

  const diferenciaDias = Math.floor(
    (d - inicioSemana01) / 86400000
  )

  const numeroSemana = Math.floor(diferenciaDias / 7) + 1

  return `${String(numeroSemana).padStart(2,'0')}-${String(anio).slice(-2)}`
}

async function buscarProveedorCatalogo(nombre){
  try{
    const rows = await getRows(
      'lista Proveedores!A2:A',
      SPREADSHEET_COMPRAS_ID
    )

    const buscado = normaliza(nombre || '').toUpperCase().trim()

    if(!buscado) return 'PROVEEDOR OCASIONAL'

    for(const row of rows){
      const oficial = (row[0] || '').toString().trim()

      if(!oficial) continue

      const normalizado = normaliza(oficial).toUpperCase()

      if(
        normalizado === buscado ||
        normalizado.includes(buscado) ||
        buscado.includes(normalizado)
      ){
        return oficial
      }
    }

    return 'PROVEEDOR OCASIONAL'

  }catch(e){
    console.log('Error lista Proveedores:', e.message)
    return 'PROVEEDOR OCASIONAL'
  }
}

async function leerTicketConOpenAI(bufferImagen){
  try{
    const base64 = bufferImagen.toString('base64')

    const res = await fetch('https://api.openai.com/v1/chat/completions', {
      method:'POST',
      headers:{
        'Authorization':`Bearer ${OPENAI_API_KEY}`,
        'Content-Type':'application/json'
      },
      body:JSON.stringify({
        model:'gpt-4o',
        messages:[
          {
            role:'user',
            content:[
              {
                type:'text',
                text:`Eres lector de tickets de compras para restaurante Trinidad.

Analiza esta imagen. Puede ser FACTURA impresa, TICKET o NOTA MANUSCRITA.

Devuelve SOLO JSON válido con esta estructura:

{
  "proveedor": "",
  "folio": "",
  "importe_total": "",
  "fecha_ticket": "YYYY-MM-DD",
  "concepto": "",
  "area": "",
  "pagado": false,
  "productos": [
    {
      "nombre": "",
      "cantidad": "",
      "unidad": "",
      "precio": "",
      "total": ""
    }
  ]
}

CONCEPTOS PERMITIDOS:
${CONCEPTOS_VALIDOS.join(', ')}

REGLAS:

- proveedor: escribe exactamente lo que aparezca en el comprobante.
- No inventes proveedores.
- folio: usa el número real del comprobante si existe.
- Si no existe folio, devuelve "".
- importe_total: busca TOTAL y devuelve solo número.
- fecha_ticket: usa la fecha visible. Si no es legible devuelve "".
- concepto: selecciona SOLO uno de los conceptos permitidos.
- Si no hay coincidencia clara, devuelve "".
- area: solo puede ser COCINA, BARRA o GASTOS.
- Si no hay evidencia clara del área, devuelve "".
- pagado: true únicamente si se observa claramente PAGADO, sello de pagado o evidencia equivalente.
- No marques pagado por suposición.
- productos: extrae los productos visibles.
- cantidad y precio deben ser números cuando sea posible.
- unidad puede ser KG, PIEZA, CAJA, LITRO, etc.
- No inventes productos que no aparezcan.
- Si algo no es legible, deja el campo vacío.`
              },
              {
                type:'image_url',
                image_url:{
                  url:`data:image/jpeg;base64,${base64}`
                }
              }
            ]
          }
        ],
        max_tokens:1500,
        temperature:0.1
      })
    })

    const data = await res.json()

    const content =
      data.choices?.[0]?.message?.content || ''

    console.log('OpenAI RAW:', content)

    const match = content.match(/\{[\s\S]*\}/)

    if(!match) return null

    const json = JSON.parse(match[0])

    return {
      proveedor:json.proveedor || '',
      folio:json.folio || '',
      importe_total:(json.importe_total || '0')
        .toString()
        .replace(/,/g,''),
      fecha_ticket:json.fecha_ticket || fechaLaboral(),
      concepto_sugerido:obtenerConceptoValido(json.concepto),
      area:['COCINA','BARRA','GASTOS'].includes(
        (json.area || '').toString().toUpperCase()
      )
        ? json.area.toString().toUpperCase()
        : '',
      pagado:json.pagado === true,
      productos:Array.isArray(json.productos)
        ? json.productos
        : []
    }

  }catch(e){
    console.log('OpenAI fail', e.message)
    return null
  }
}

async function buscarDuplicado(datos, proveedorFinal){
  try{
    const rows = await getRows(
      `${SHEET_RESUMEN}!A2:I`,
      SPREADSHEET_COMPRAS_ID
    )

    const folio = normaliza(datos.folio || '').toUpperCase()
    const fecha = datos.fecha
    const monto = parseFloat(
      (datos.monto || '0').toString().replace(/,/g,'')
    ) || 0

    if(!folio) return false

    return rows.some(r=>{
      const rProveedor = normaliza(r[3] || '').toUpperCase()
      const rFolio = normaliza(r[4] || '').toUpperCase()
      const rFecha = r[1] || ''
      const rMonto = parseFloat(
        (r[5] || '0').toString().replace(/,/g,'')
      ) || 0

      return (
        rProveedor === normaliza(proveedorFinal).toUpperCase() &&
        rFolio === folio &&
        rFecha === fecha &&
        Math.abs(rMonto - monto) < 0.01
      )
    })

  }catch(e){
    console.log('Error buscando duplicado:', e.message)
    return false
  }
}

async function marcarFilaDuplicada(rowNumber){
  try{
    const sClient = await sheetsClient()

    await sClient.spreadsheets.batchUpdate({
      spreadsheetId:SPREADSHEET_COMPRAS_ID,
      requestBody:{
        requests:[
          {
            repeatCell:{
              range:{
                sheetId:await obtenerSheetId(
                  sClient,
                  SHEET_RESUMEN
                ),
                startRowIndex:rowNumber - 1,
                endRowIndex:rowNumber,
                startColumnIndex:0,
                endColumnIndex:9
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
              fields:'userEnteredFormat.backgroundColor'
            }
          }
        ]
      }
    })

  }catch(e){
    console.log('Error marcando duplicado:', e.message)
  }
}

async function obtenerSheetId(sClient, nombreHoja){
  const meta = await sClient.spreadsheets.get({
    spreadsheetId:SPREADSHEET_COMPRAS_ID,
    fields:'sheets.properties'
  })

  const hoja = meta.data.sheets?.find(
    s => s.properties?.title === nombreHoja
  )

  return hoja?.properties?.sheetId ?? 0
}

async function registrarCompra(datos){
  const sClient = await sheetsClient()

  const proveedorFinal =
    await buscarProveedorCatalogo(datos.proveedor)

  const duplicado =
    await buscarDuplicado(datos, proveedorFinal)

  const formaPago =
    datos.pagado ? 'PAGADO' : ''

  await sClient.spreadsheets.values.append({
    spreadsheetId:SPREADSHEET_COMPRAS_ID,
    range:`${SHEET_RESUMEN}!A:I`,
    valueInputOption:'USER_ENTERED',
    requestBody:{
      values:[[
        datos.semana,
        datos.fecha,
        datos.sucursal,
        proveedorFinal,
        datos.folio,
        datos.monto,
        datos.concepto,
        datos.area,
        formaPago
      ]]
    }
  })

  try{
    if(datos.productos && datos.productos.length > 0){

      const rows = datos.productos.map(p=>[
        datos.folio,
        p.nombre || datos.concepto,
        p.cantidad || '1',
        p.unidad || '',
        p.precio || '',
        p.total || datos.monto,
        datos.sucursal,
        datos.fecha
      ])

      await sClient.spreadsheets.values.append({
        spreadsheetId:SPREADSHEET_COMPRAS_ID,
        range:`${SHEET_INSUMOS}!A:H`,
        valueInputOption:'USER_ENTERED',
        requestBody:{
          values:rows
        }
      })

    }else{

      await sClient.spreadsheets.values.append({
        spreadsheetId:SPREADSHEET_COMPRAS_ID,
        range:`${SHEET_INSUMOS}!A:H`,
        valueInputOption:'USER_ENTERED',
        requestBody:{
          values:[[
            datos.folio,
            datos.concepto,
            '1',
            '',
            datos.monto,
            datos.monto,
            datos.sucursal,
            datos.fecha
          ]]
        }
      })
    }

  }catch(e){
    console.log('Error INSUMOS:', e.message)
  }

  /*
    Después del append obtenemos la última fila
    y, si era duplicado, la marcamos en amarillo.
  */

  if(duplicado){
    try{
      const rowsActuales = await getRows(
        `${SHEET_RESUMEN}!A2:I`,
        SPREADSHEET_COMPRAS_ID
      )

      await marcarFilaDuplicada(rowsActuales.length + 1)

    }catch(e){
      console.log('Error marcando fila duplicada:', e.message)
    }
  }

  return {
    ok:true,
    duplicado,
    proveedor:proveedorFinal,
    msg:
      `✅ Compra registrada · ${datos.fecha} · `+
      `${proveedorFinal} · ${datos.folio || 'SIN FOLIO'} · `+
      `$${datos.monto} · ${datos.sucursal}`
  }
}

export async function generarExcelCompras(filtro, jid, sock){
  const rows = await getRows(
    `${SHEET_RESUMEN}!A2:I`,
    SPREADSHEET_COMPRAS_ID
  )

  let filtradas = rows.filter(r=>{

    if(
      filtro.sucursal &&
      !(r[2] || '')
        .toUpperCase()
        .includes(filtro.sucursal.toUpperCase())
    ) return false

    if(filtro.fecha && r[1] !== filtro.fecha) return false

    return true
  })

  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Compras')

  ws.addRow([
    'Semana',
    'Fecha',
    'Sucursal',
    'Proveedor',
    '#Comprobante',
    'Importe',
    'Concepto',
    'Area',
    'FormaPago'
  ]).font={bold:true}

  filtradas.forEach(r=>ws.addRow(r))

  const total = filtradas.reduce(
    (a,r)=>
      a +
      (
        parseFloat(
          (r[5] || '0')
            .toString()
            .replace(/,/g,'')
        ) || 0
      ),
    0
  )

  ws.addRow([])
  ws.addRow(['TOTAL','','','','',total])

  ws.columns.forEach(c=>c.width=18)

  const fileName =
    `Compras_${filtro.sucursal || 'TODAS'}_${fechaLaboral()}.xlsx`

  const fp = path.join(os.tmpdir(), fileName)

  await wb.xlsx.writeFile(fp)

  await sock.sendMessage(jid,{
    document:fs.readFileSync(fp),
    mimetype:
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    fileName,
    caption:
      `📊 Compras ${filtro.sucursal || ''} - `+
      `${filtradas.length} regs - Total $${total}`
  })

  fs.unlinkSync(fp)
}

function parseFiltroCompras(texto){
  const low = normaliza(texto)

  let suc = null

  if(low.includes('bucareli')) suc='BUCARELI'
  else if(low.includes('coyo')) suc='COYOACAN'
  else if(low.includes('juarez')) suc='JUAREZ'

  let fecha = null

  if(low.includes('hoy')){
    fecha = fechaLaboral()
  }

  return {
    sucursal:suc,
    fecha
  }
}

export async function handleCompras({
  sock,
  jid,
  m,
  texto,
  esImagen
}){

  if(/^compras/i.test(texto)){
    const f = parseFiltroCompras(texto)

    await generarExcelCompras(
      f,
      jid,
      sock
    )

    return true
  }

  /*
    Por ahora SOLO el grupo de pruebas puede
    registrar compras mediante fotografías.
  */

  if(esImagen && jid !== GRUPO_PRUEBAS_ID){
    return false
  }

  if(esImagen){
    try{

      if(!OPENAI_API_KEY){
        await sock.sendMessage(jid,{
          text:'❌ Falta OPENAI_API_KEY en Railway'
        })
        return true
      }

      const buffer = await downloadMediaMessage(
        m,
        'buffer',
        {},
        {
          logger:P({level:'fatal'}),
          reuploadRequest:sock.updateMediaMessage
        }
      )

      /*
        No se manda mensaje de procesamiento.
        La lectura se hace internamente.
      */

      const datosIA =
        await leerTicketConOpenAI(buffer)

      if(
        !datosIA ||
        datosIA.importe_total === '0' ||
        !datosIA.importe_total
      ){
        await sock.sendMessage(jid,{
          text:'❌ No pude leer el total. Manda una foto más clara.'
        })
        return true
      }

      const fecha =
        datosIA.fecha_ticket || fechaLaboral()

      const datosFinal = {
        folio:datosIA.folio,
        concepto:datosIA.concepto_sugerido,
        proveedor:datosIA.proveedor,
        monto:datosIA.importe_total,
        fecha,
        sucursal:'BUCARELI',
        semana:obtenerSemanaCompras(fecha),
        area:datosIA.area,
        pagado:datosIA.pagado,
        productos:datosIA.productos
      }

      const res =
        await registrarCompra(datosFinal)

      await sock.sendMessage(jid,{
        text:res.msg
      })

      return true

    }catch(e){
      console.error(e)

      await sock.sendMessage(jid,{
        text:'❌ Error al registrar compra: '+e.message
      })

      return true
    }
  }

  return false
}

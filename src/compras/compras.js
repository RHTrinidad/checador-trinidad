import{downloadMediaMessage}from'@whiskeysockets/baileys'
import P from'pino'
import{SPREADSHEET_COMPRAS_ID,SHEET_RESUMEN,SHEET_INSUMOS,GRUPO_PRUEBAS_ID}from'../config.js'
import{fechaLaboral,normaliza}from'../utils.js'
import{sheetsClient,getRows}from'../sheets.js'
import ExcelJS from'exceljs'
import fs from'fs'
import path from'path'
import os from'os'
import crypto from'crypto'

const OPENAI_API_KEY=process.env.OPENAI_API_KEY
const colasFotos=new Map()
const hashesRecientes=new Map()
const TIEMPO_AGRUPACION=5000
const TIEMPO_HASH_RECIENTE=15*60*1000

function hashBuffer(buffer){
  return crypto.createHash('sha256').update(buffer).digest('hex')
}

function limpiarHashesRecientes(){
  const ahora=Date.now()
  for(const[hash,fecha]of hashesRecientes){
    if(ahora-fecha>TIEMPO_HASH_RECIENTE)hashesRecientes.delete(hash)
  }
}

async function agregarFotoACola({sock,jid,m,texto}){
  try{
    const buffer=await downloadMediaMessage(m,'buffer',{},{
      logger:P({level:'fatal'}),
      reuploadRequest:sock.updateMediaMessage
    })
    limpiarHashesRecientes()
    const hash=hashBuffer(buffer)

    if(hashesRecientes.has(hash)){
      console.log('Foto duplicada ignorada:',jid)
      return false
    }

    hashesRecientes.set(hash,Date.now())

    let cola=colasFotos.get(jid)

    if(!cola){
      cola={items:[],timer:null}
      colasFotos.set(jid,cola)
    }

    if(cola.items.some(x=>x.hash===hash))return false

    cola.items.push({buffer,hash,texto})

    if(cola.timer)clearTimeout(cola.timer)

    cola.timer=setTimeout(
      ()=>procesarColaFotos(jid,sock),
      TIEMPO_AGRUPACION
    )

    return true
  }catch(e){
    console.error('Error agregando foto a cola:',e)
    await sock.sendMessage(jid,{
      text:'❌ No pude recibir correctamente la fotografía.'
    })
    return true
  }
}

async function procesarColaFotos(jid,sock){
  const cola=colasFotos.get(jid)

  if(!cola||!cola.items.length)return

  colasFotos.delete(jid)

  try{
    const buffers=cola.items.map(x=>x.buffer)
    const datosIA=await leerTicketsConOpenAI(buffers)

    await registrarDocumentosIA(
      jid,
      sock,
      datosIA,
      'fotografías'
    )
  }catch(e){
    console.error('Error procesando lote de compras:',e)

    await sock.sendMessage(jid,{
      text:'❌ Error al procesar las fotografías: '+e.message
    })
  }
}

async function registrarDocumentosIA(jid,sock,datosIA,tipoEntrada){
  try{
    if(!datosIA||!Array.isArray(datosIA)||!datosIA.length){
      await sock.sendMessage(jid,{
        text:`❌ No pude identificar la factura. Manda ${tipoEntrada} nuevamente con mayor claridad.`
      })
      return
    }

    const fechaCaptura=fechaLaboral()

    let registros=0
    let totalInsumos=0
    const detalleCompras=[]

    for(const documento of datosIA){
      if(
        !documento||
        documento.importe_total==='0'||
        !documento.importe_total
      )continue

      const datosFinal={
        folio:documento.folio||'',
        concepto:documento.concepto_sugerido||'',
        proveedor:documento.proveedor||'',
        monto:documento.importe_total,
        fechaCaptura,
        fechaDocumento:documento.fecha_ticket||'',
        sucursal:documento.sucursal||'',
        semana:obtenerSemanaCompras(fechaCaptura),
        area:documento.area||'',
        pagado:documento.pagado===true,
        productos:documento.productos
      }

      const res=await registrarCompra(datosFinal)

      registros++

      detalleCompras.push({
        proveedor:res.proveedor||'PROVEEDOR OCASIONAL',
        monto:datosFinal.monto||'0.00'
      })

      totalInsumos+=
        Array.isArray(datosFinal.productos)
          ?datosFinal.productos.length
          :1
    }

    if(!registros){
      await sock.sendMessage(jid,{
        text:`❌ No pude leer el total de ninguna factura. Manda ${tipoEntrada} nuevamente con mayor claridad.`
      })
      return
    }

    let mensaje=
      registros===1
        ?'✅ Compra registrada correctamente.'
        :`✅ ${registros} compras registradas correctamente.`

    for(const compra of detalleCompras){
      const monto=parseFloat(
        (compra.monto||'0')
          .toString()
          .replace(/,/g,'')
      )||0

      mensaje+=
        `\n${registros>1?'• ':''}${compra.proveedor} · $${monto.toFixed(2)}`
    }

    if(totalInsumos){
      mensaje+=`\n📦 ${totalInsumos} insumos registrados.`
    }

    await sock.sendMessage(jid,{text:mensaje})
  }catch(e){
    console.error('Error registrando documentos IA:',e)

    await sock.sendMessage(jid,{
      text:'❌ Error al registrar la compra: '+e.message
    })
  }
}

async function obtenerConceptosCatalogo(){
  try{
    const rows=await getRows(
      'lista Proveedores!J2:J',
      SPREADSHEET_COMPRAS_ID
    )

    const conceptos=[]

    for(const row of rows){
      const concepto=(row[0]||'').toString().trim()
      if(concepto)conceptos.push(concepto)
    }

    return conceptos
  }catch(e){
    console.log(
      'Error leyendo conceptos de lista Proveedores:',
      e.message
    )
    return[]
  }
}

function obtenerConceptoValido(valor,conceptosValidos){
  const x=normaliza(valor||'').toUpperCase().trim()

  if(!x)return''

  const encontrado=conceptosValidos.find(
    c=>normaliza(c).toUpperCase()===x
  )

  return encontrado||''
}

function obtenerSucursalValida(valor){
  const x=normaliza(valor||'').toUpperCase().trim()

  if(!x)return''

  if(
    x==='BUCARELI'||
    x.includes('BUCARELI')
  )return'BUCARELI'

  if(
    x==='COYOACAN'||
    x.includes('COYOACAN')||
    x.includes('COYOACÁN')
  )return'COYOACAN'

  return''
}

function obtenerSemanaCompras(fecha){
  const d=new Date(`${fecha}T12:00:00`)

  if(Number.isNaN(d.getTime()))return''

  const anio=d.getFullYear()

  const inicioAnio=
    new Date(anio,0,1,12)

  const diaSemana=
    inicioAnio.getDay()

  const diasDesdeMartes=
    (diaSemana-2+7)%7

  const inicioSemana01=
    new Date(inicioAnio)

  inicioSemana01.setDate(
    inicioSemana01.getDate()-diasDesdeMartes
  )

  const diferenciaDias=
    Math.floor(
      (d-inicioSemana01)/86400000
    )

  const numeroSemana=
    Math.floor(diferenciaDias/7)+1

  return`${String(numeroSemana).padStart(2,'0')}-${String(anio).slice(-2)}`
}

async function buscarProveedorCatalogo(nombre){
  try{
    const rows=await getRows(
      'lista Proveedores!A2:A',
      SPREADSHEET_COMPRAS_ID
    )

    const buscado=
      normaliza(nombre||'')
        .toUpperCase()
        .trim()

    if(!buscado)return'PROVEEDOR OCASIONAL'

    for(const row of rows){
      const oficial=
        (row[0]||'').toString().trim()

      if(!oficial)continue

      const normalizado=
        normaliza(oficial)
          .toUpperCase()

      if(
        normalizado===buscado||
        normalizado.includes(buscado)||
        buscado.includes(normalizado)
      ){
        return oficial
      }
    }

    return'PROVEEDOR OCASIONAL'
  }catch(e){
    console.log(
      'Error lista Proveedores:',
      e.message
    )

    return'PROVEEDOR OCASIONAL'
  }
}

async function leerTicketsConOpenAI(buffers){
  try{
    const conceptosValidos=
      await obtenerConceptosCatalogo()

    const listaConceptos=
      conceptosValidos.length>0
        ?conceptosValidos.join(', ')
        :'(
NO HAY CONCEPTOS CONFIGURADOS EN LA COLUMNA J
)'

    const contenido=[{
      type:'text',
      text:`Eres lector de comprobantes de compras para restaurante Trinidad.

Recibirás una o varias fotografías.

IMPORTANTE:
Las fotografías pueden ser páginas diferentes de UNA MISMA FACTURA.
También pueden existir VARIAS FACTURAS diferentes dentro del conjunto.
Debes identificar correctamente cada documento.

Devuelve SOLO JSON válido:
{
  "documentos":[
    {
      "proveedor":"",
      "folio":"",
      "importe_total":"",
      "fecha_ticket":"",
      "concepto":"",
      "area":"",
      "sucursal":"",
      "pagado":false,
      "productos":[
        {
          "nombre":"",
          "cantidad":"",
          "unidad":"",
          "precio":"",
          "total":""
        }
      ]
    }
  ]
}

CONCEPTOS DISPONIBLES EN GOOGLE SHEETS:
${listaConceptos}

REGLAS PARA AGRUPAR:
- Si varias fotografías muestran distintas partes o páginas de LA MISMA FACTURA, deben formar UN SOLO documento.
- NO crees una factura nueva solamente porque aparezca otra página.
- Une TODOS los productos visibles de las diferentes fotografías de la misma factura.
- Si una misma línea aparece repetida porque dos fotografías muestran la misma zona, NO la dupliques.
- Si una fotografía es exactamente la misma página repetida, no dupliques sus productos.
- Si existen facturas diferentes, crea un documento independiente para cada una.
- Un cambio claro de folio, total, proveedor o documento significa que probablemente es otra factura.
- NO inventes información.

REGLAS DEL PROVEEDOR:
- proveedor: escribe exactamente lo que aparezca.
- No inventes proveedores.

REGLAS DEL FOLIO:
- folio: usa el número real del comprobante si existe.
- Si no existe folio, devuelve "".

REGLAS DEL IMPORTE:
- importe_total: busca TOTAL.
- Devuelve solo número.
- No inventes el importe.

REGLAS DE FECHA:
- fecha_ticket: fecha visible del comprobante.
- Formato YYYY-MM-DD.
- Si no existe o no es legible, devuelve "".
- NO inventes una fecha.

REGLAS DE CONCEPTO:
- selecciona SOLO uno de los conceptos disponibles.
- No inventes conceptos.
- Si no hay coincidencia clara, devuelve "".

REGLAS DE AREA:
- solo puede ser COCINA, BARRA o GASTOS.
- Si no hay evidencia clara, devuelve "".

REGLAS DE SUCURSAL:
La sucursal se determina UNICAMENTE mediante SELLO DE SUCURSAL visible.
- BUCARELI → "BUCARELI"
- COYOACAN o COYOACÁN → "COYOACAN"
- Sin sello claramente visible → ""
NO deduzcas sucursal por proveedor, dirección, RFC, teléfono, encabezado, domicilio, texto de factura o grupo de WhatsApp.

REGLAS DE PAGADO:
- true únicamente si se observa claramente PAGADO, sello de pagado o evidencia equivalente.
- No marques pagado por suposición.
- Sello de sucursal y sello de PAGADO son independientes.

REGLAS DE PRODUCTOS:
- Extrae TODOS los productos visibles de TODAS las páginas correspondientes a esa factura.
- cantidad y precio deben ser números cuando sea posible.
- unidad puede ser KG, PIEZA, CAJA, LITRO, etc.
- No inventes productos.
- Si algo no es legible, deja el campo vacío.
- No dupliques productos que aparecen repetidos únicamente por solapamiento entre fotografías.

Si una factura ocupa 4 fotografías:
→ 1 solo documento
→ 1 solo registro en Resumen Compras
→ TODOS sus productos en Insumos.`
    }]

    for(const buffer of buffers){
      contenido.push({
        type:'image_url',
        image_url:{
          url:`data:image/jpeg;base64,${buffer.toString('base64')}`
        }
      })
    }

    const res=await fetch(
      'https://api.openai.com/v1/chat/completions',
      {
        method:'POST',
        headers:{
          Authorization:`Bearer ${OPENAI_API_KEY}`,
          'Content-Type':'application/json'
        },
        body:JSON.stringify({
          model:'gpt-4o',
          messages:[{
            role:'user',
            content:contenido
          }],
          max_tokens:3000,
          temperature:0.1
        })
      }
    )

    const data=await res.json()

    if(!res.ok){
      console.log(
        'OpenAI COMPRAS FOTO ERROR:',
        data
      )
      return null
    }

    const content=
      data.choices?.[0]?.message?.content||''

    console.log(
      'OpenAI COMPRAS RAW:',
      content
    )

    const match=
      content.match(/\{[\s\S]*\}/)

    if(!match)return null

    const json=
      JSON.parse(match[0])

    const documentos=
      Array.isArray(json.documentos)
        ?json.documentos
        :[]

    return documentos.map(documento=>({
      proveedor:documento.proveedor||'',
      folio:documento.folio||'',
      importe_total:
        (documento.importe_total||'0')
          .toString()
          .replace(/,/g,''),
      fecha_ticket:
        documento.fecha_ticket||'',
      concepto_sugerido:
        obtenerConceptoValido(
          documento.concepto,
          conceptosValidos
        ),
      area:
        ['COCINA','BARRA','GASTOS'].includes(
          (documento.area||'')
            .toString()
            .toUpperCase()
        )
          ?documento.area.toString().toUpperCase()
          :'',
      sucursal:
        obtenerSucursalValida(
          documento.sucursal
        ),
      pagado:
        documento.pagado===true,
      productos:
        Array.isArray(documento.productos)
          ?documento.productos
          :[]
    }))
  }catch(e){
    console.log(
      'OpenAI FOTO fail:',
      e.message
    )
    return null
  }
}

async function leerPDFConOpenAI(
  buffer,
  nombreArchivo='comprobante.pdf'
){
  try{
    const conceptosValidos=
      await obtenerConceptosCatalogo()

    const listaConceptos=
      conceptosValidos.length>0
        ?conceptosValidos.join(', ')
        :'(
NO HAY CONCEPTOS CONFIGURADOS EN LA COLUMNA J
)'

    const prompt=`Eres lector de comprobantes de compras para restaurante Trinidad.

Recibirás un archivo PDF.

El PDF puede tener UNA O VARIAS páginas.

IMPORTANTE:
- Varias páginas pueden pertenecer a UNA MISMA FACTURA.
- Un PDF también puede contener VARIAS FACTURAS diferentes.
- Debes identificar correctamente cada documento.
- No crees una factura nueva solamente porque cambie de página.
- Si varias páginas pertenecen al mismo comprobante, únelas en UN SOLO documento.
- Extrae TODOS los productos de todas las páginas de ese documento.
- Si existen varias facturas realmente diferentes, crea un documento independiente para cada una.
- No inventes información.

Devuelve SOLO JSON válido, sin markdown ni explicaciones:

{
  "documentos":[
    {
      "proveedor":"",
      "folio":"",
      "importe_total":"",
      "fecha_ticket":"",
      "concepto":"",
      "area":"",
      "sucursal":"",
      "pagado":false,
      "productos":[
        {
          "nombre":"",
          "cantidad":"",
          "unidad":"",
          "precio":"",
          "total":""
        }
      ]
    }
  ]
}

CONCEPTOS DISPONIBLES EN GOOGLE SHEETS:
${listaConceptos}

REGLAS DEL PROVEEDOR:
- proveedor: escribe exactamente lo visible.
- No inventes proveedores.
- Después el sistema validará el proveedor contra Google Sheets.

REGLAS DEL FOLIO:
- Usa el número real del comprobante si existe.
- Si no existe o no es legible, devuelve "".

REGLAS DEL IMPORTE:
- Busca TOTAL, IMPORTE TOTAL o equivalente.
- Devuelve solamente el número.
- No inventes el importe.

REGLAS DE FECHA:
- Usa la fecha visible del documento.
- Formato YYYY-MM-DD.
- Si no existe o no es legible, devuelve "".
- NO uses la fecha de captura como fecha del documento.

REGLAS DE CONCEPTO:
- Selecciona SOLO uno de los conceptos disponibles.
- No inventes conceptos.
- Si no hay coincidencia clara, devuelve "".

REGLAS DE AREA:
- Solo puede ser COCINA, BARRA o GASTOS.
- Si no hay evidencia clara, devuelve "".

REGLAS DE SUCURSAL:
La sucursal se determina UNICAMENTE mediante SELLO DE SUCURSAL visible.
- BUCARELI → "BUCARELI"
- COYOACAN o COYOACÁN → "COYOACAN"
- Sin sello claramente visible → ""

NO deduzcas sucursal por:
- proveedor
- dirección
- RFC
- teléfono
- encabezado
- domicilio
- texto del documento
- nombre del archivo

REGLAS DE PAGADO:
- true únicamente si aparece claramente PAGADO, sello de pagado o evidencia equivalente.
- No marques PAGADO por suposición.
- El sello de sucursal y el sello de PAGADO son independientes.

REGLAS DE PRODUCTOS:
- Extrae TODOS los productos visibles.
- Incluye productos de TODAS las páginas pertenecientes a la misma factura.
- No dupliques productos.
- Si una página repite información de otra, no vuelvas a registrar las mismas líneas.
- cantidad y precio deben ser números cuando sea posible.
- unidad puede ser KG, PIEZA, CAJA, LITRO, etc.
- Si un dato no es legible, déjalo vacío.
- No inventes productos.

Si el PDF tiene 4 páginas de UNA factura:
→ 1 solo documento
→ 1 solo registro en Resumen Compras
→ TODOS sus productos en Insumos.`

    const base64=
      buffer.toString('base64')

    const res=await fetch(
      'https://api.openai.com/v1/responses',
      {
        method:'POST',
        headers:{
          Authorization:`Bearer ${OPENAI_API_KEY}`,
          'Content-Type':'application/json'
        },
        body:JSON.stringify({
          model:'gpt-4o',
          input:[{
            role:'user',
            content:[
              {
                type:'input_text',
                text:prompt
              },
              {
                type:'input_file',
                filename:
                  nombreArchivo||'comprobante.pdf',
                file_data:
                  `data:application/pdf;base64,${base64}`
              }
            ]
          }],
          temperature:0.1
        })
      }
    )

    const data=await res.json()

    if(!res.ok){
      console.log(
        'OpenAI PDF ERROR:',
        JSON.stringify(data,null,2)
      )
      return null
    }

    let content=
      data.output_text||''

    if(!content){
      const textos=[]

      for(const item of(data.output||[])){
        for(const c of(item.content||[])){
          if(
            c.type==='output_text'&&
            c.text
          ){
            textos.push(c.text)
          }
        }
      }

      content=textos.join('\n')
    }

    console.log(
      'OpenAI COMPRAS PDF RAW:',
      content
    )

    if(!content){
      console.log(
        'OpenAI COMPRAS PDF RESPONSE:',
        JSON.stringify(data,null,2)
      )
      return null
    }

    const match=
      content.match(/\{[\s\S]*\}/)

    if(!match){
      console.log(
        'OpenAI PDF SIN JSON:',
        content
      )
      return null
    }

    let json

    try{
      json=JSON.parse(match[0])
    }catch(e){
      console.log(
        'OpenAI PDF JSON ERROR:',
        e.message,
        content
      )
      return null
    }

    const documentos=
      Array.isArray(json.documentos)
        ?json.documentos
        :[]

    return documentos.map(documento=>({
      proveedor:documento.proveedor||'',
      folio:documento.folio||'',
      importe_total:
        (documento.importe_total||'0')
          .toString()
          .replace(/,/g,''),
      fecha_ticket:
        documento.fecha_ticket||'',
      concepto_sugerido:
        obtenerConceptoValido(
          documento.concepto,
          conceptosValidos
        ),
      area:
        ['COCINA','BARRA','GASTOS'].includes(
          (documento.area||'')
            .toString()
            .toUpperCase()
        )
          ?documento.area.toString().toUpperCase()
          :'',
      sucursal:
        obtenerSucursalValida(
          documento.sucursal
        ),
      pagado:
        documento.pagado===true,
      productos:
        Array.isArray(documento.productos)
          ?documento.productos
          :[]
    }))
  }catch(e){
    console.log(
      'OpenAI PDF fail:',
      e.message
    )
    return null
  }
}

async function procesarPDF({sock,jid,m}){
  try{
    if(!OPENAI_API_KEY){
      await sock.sendMessage(jid,{
        text:'❌ Falta OPENAI_API_KEY en Railway'
      })
      return true
    }

    const buffer=
      await downloadMediaMessage(
        m,
        'buffer',
        {},
        {
          logger:P({level:'fatal'}),
          reuploadRequest:
            sock.updateMediaMessage
        }
      )

    limpiarHashesRecientes()

    const hash=
      hashBuffer(buffer)

    if(hashesRecientes.has(hash)){
      console.log(
        'PDF duplicado ignorado:',
        jid
      )

      await sock.sendMessage(jid,{
        text:'⚠️ Este PDF ya fue recibido recientemente y no se volverá a registrar.'
      })

      return true
    }

    hashesRecientes.set(
      hash,
      Date.now()
    )

    const nombreArchivo=
      m.message?.documentMessage?.fileName||
      m.message?.documentWithCaptionMessage?.message?.documentMessage?.fileName||
      'comprobante.pdf'

    const datosIA=
      await leerPDFConOpenAI(
        buffer,
        nombreArchivo
      )

    await registrarDocumentosIA(
      jid,
      sock,
      datosIA,
      'el PDF'
    )

    return true
  }catch(e){
    console.error(
      'Error procesando PDF:',
      e
    )

    await sock.sendMessage(jid,{
      text:'❌ Error al procesar el PDF: '+e.message
    })

    return true
  }
}

async function buscarDuplicado(
  datos,
  proveedorFinal
){
  try{
    const rows=
      await getRows(
        `${SHEET_RESUMEN}!A2:J`,
        SPREADSHEET_COMPRAS_ID
      )

    const folio=
      normaliza(datos.folio||'')
        .toUpperCase()

    const fechaDocumento=
      datos.fechaDocumento||''

    const fechaCaptura=
      datos.fechaCaptura||''

    const monto=
      parseFloat(
        (datos.monto||'0')
          .toString()
          .replace(/,/g,'')
      )||0

    if(!folio)return false

    return rows.some(r=>{
      const rProveedor=
        normaliza(r[4]||'')
          .toUpperCase()

      const rFolio=
        normaliza(r[5]||'')
          .toUpperCase()

      const rFechaDocumento=
        r[2]||''

      const rFechaCaptura=
        r[1]||''

      const rMonto=
        parseFloat(
          (r[6]||'0')
            .toString()
            .replace(/,/g,'')
        )||0

      const mismaFecha=
        fechaDocumento&&rFechaDocumento
          ?rFechaDocumento===fechaDocumento
          :rFechaCaptura===fechaCaptura

      return(
        rProveedor===
          normaliza(proveedorFinal)
            .toUpperCase()&&
        rFolio===folio&&
        mismaFecha&&
        Math.abs(rMonto-monto)<0.01
      )
    })
  }catch(e){
    console.log(
      'Error buscando duplicado:',
      e.message
    )
    return false
  }
}

async function marcarFilaDuplicada(
  rowNumber
){
  try{
    const sClient=
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
                  SHEET_RESUMEN
                ),
              startRowIndex:
                rowNumber-1,
              endRowIndex:
                rowNumber,
              startColumnIndex:0,
              endColumnIndex:10
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
      'Error marcando duplicado:',
      e.message
    )
  }
}

async function obtenerSheetId(
  sClient,
  nombreHoja
){
  const meta=
    await sClient.spreadsheets.get({
      spreadsheetId:
        SPREADSHEET_COMPRAS_ID,
      fields:'sheets.properties'
    })

  const hoja=
    meta.data.sheets?.find(
      s=>s.properties?.title===nombreHoja
    )

  return hoja?.properties?.sheetId??0
}

async function registrarCompra(datos){
  const sClient=
    await sheetsClient()

  const proveedorFinal=
    await buscarProveedorCatalogo(
      datos.proveedor
    )

  const duplicado=
    await buscarDuplicado(
      datos,
      proveedorFinal
    )

  const formaPago=
    datos.pagado?'PAGADO':''

  await sClient.spreadsheets.values.append({
    spreadsheetId:
      SPREADSHEET_COMPRAS_ID,

    range:
      `${SHEET_RESUMEN}!A:J`,

    valueInputOption:
      'USER_ENTERED',

    requestBody:{
      values:[[
        datos.semana,
        datos.fechaCaptura,
        datos.fechaDocumento,
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
    if(
      datos.productos&&
      datos.productos.length>0
    ){
      const rows=
        datos.productos.map(p=>[
          datos.folio,
          p.nombre||datos.concepto,
          p.cantidad||'1',
          p.unidad||'',
          p.precio||'',
          p.total||datos.monto,
          proveedorFinal,
          datos.sucursal,
          datos.fechaCaptura
        ])

      await sClient.spreadsheets.values.append({
        spreadsheetId:
          SPREADSHEET_COMPRAS_ID,

        range:
          `${SHEET_INSUMOS}!A:I`,

        valueInputOption:
          'USER_ENTERED',

        requestBody:{
          values:rows
        }
      })
    }else{
      await sClient.spreadsheets.values.append({
        spreadsheetId:
          SPREADSHEET_COMPRAS_ID,

        range:
          `${SHEET_INSUMOS}!A:I`,

        valueInputOption:
          'USER_ENTERED',

        requestBody:{
          values:[[
            datos.folio,
            datos.concepto,
            '1',
            '',
            datos.monto,
            datos.monto,
            proveedorFinal,
            datos.sucursal,
            datos.fechaCaptura
          ]]
        }
      })
    }
  }catch(e){
    console.log(
      'Error INSUMOS:',
      e.message
    )
  }

  if(duplicado){
    try{
      const rowsActuales=
        await getRows(
          `${SHEET_RESUMEN}!A2:J`,
          SPREADSHEET_COMPRAS_ID
        )

      await marcarFilaDuplicada(
        rowsActuales.length+1
      )
    }catch(e){
      console.log(
        'Error marcando fila duplicada:',
        e.message
      )
    }
  }

  return{
    ok:true,
    duplicado,
    proveedor:proveedorFinal,
    msg:
      `✅ Compra registrada · `+
      `Captura ${datos.fechaCaptura} · `+
      `Documento ${datos.fechaDocumento||'SIN FECHA'} · `+
      `${proveedorFinal} · `+
      `${datos.folio||'SIN FOLIO'} · `+
      `$${datos.monto}`+
      `${datos.sucursal?` · ${datos.sucursal}`:''}`
  }
}

export async function generarExcelCompras(
  filtro,
  jid,
  sock
){
  const rows=
    await getRows(
      `${SHEET_RESUMEN}!A2:J`,
      SPREADSHEET_COMPRAS_ID
    )

  let filtradas=rows.filter(r=>{
    if(
      filtro.sucursal&&
      !(r[3]||'')
        .toUpperCase()
        .includes(
          filtro.sucursal.toUpperCase()
        )
    )return false

    if(
      filtro.fecha&&
      r[1]!==filtro.fecha
    )return false

    return true
  })

  const wb=
    new ExcelJS.Workbook()

  const ws=
    wb.addWorksheet('Compras')

  ws.addRow([
    'Semana',
    'Fecha de Captura',
    'Fecha Documento',
    'Sucursal',
    'Proveedor',
    '#Comprobante',
    'Importe',
    'Concepto',
    'Area',
    'FormaPago'
  ]).font={bold:true}

  filtradas.forEach(
    r=>ws.addRow(r)
  )

  const total=
    filtradas.reduce(
      (a,r)=>
        a+
        (
          parseFloat(
            (r[6]||'0')
              .toString()
              .replace(/,/g,'')
          )||0
        ),
      0
    )

  ws.addRow([])
  ws.addRow([
    'TOTAL',
    '',
    '',
    '',
    '',
    '',
    total
  ])

  ws.columns.forEach(
    c=>{c.width=18}
  )

  const fileName=
    `Compras_${filtro.sucursal||'TODAS'}_${fechaLaboral()}.xlsx`

  const fp=
    path.join(
      os.tmpdir(),
      fileName
    )

  await wb.xlsx.writeFile(fp)

  await sock.sendMessage(jid,{
    document:
      fs.readFileSync(fp),

    mimetype:
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',

    fileName,

    caption:
      `📊 Compras ${filtro.sucursal||''} - ${filtradas.length} regs - Total $${total}`
  })

  fs.unlinkSync(fp)
}

function parseFiltroCompras(texto){
  const low=
    normaliza(texto)

  let suc=null

  if(low.includes('bucareli'))
    suc='BUCARELI'
  else if(low.includes('coyo'))
    suc='COYOACAN'
  else if(low.includes('juarez'))
    suc='JUAREZ'

  let fecha=null

  if(low.includes('hoy'))
    fecha=fechaLaboral()

  return{
    sucursal:suc,
    fecha
  }
}

export async function handleCompras({
  sock,
  jid,
  m,
  texto,
  esImagen,
  esPDF
}){
  if(/^compras/i.test(texto)){
    const f=
      parseFiltroCompras(texto)

    await generarExcelCompras(
      f,
      jid,
      sock
    )

    return true
  }

  if(
    (esImagen||esPDF)&&
    jid!==GRUPO_PRUEBAS_ID
  )return false

  if(esPDF){
    try{
      return await procesarPDF({
        sock,
        jid,
        m
      })
    }catch(e){
      console.error(
        'Error en PDF Compras:',
        e
      )

      await sock.sendMessage(jid,{
        text:
          '❌ Error al recibir PDF: '+
          e.message
      })

      return true
    }
  }

  if(esImagen){
    try{
      if(!OPENAI_API_KEY){
        await sock.sendMessage(jid,{
          text:
            '❌ Falta OPENAI_API_KEY en Railway'
        })
        return true
      }

      await agregarFotoACola({
        sock,
        jid,
        m,
        texto
      })

      return true
    }catch(e){
      console.error(e)

      await sock.sendMessage(jid,{
        text:
          '❌ Error al recibir compra: '+
          e.message
      })

      return true
    }
  }

  return false
}

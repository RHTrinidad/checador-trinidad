import { downloadMediaMessage } from '@whiskeysockets/baileys'
import P from 'pino'
import { SPREADSHEET_COMPRAS_ID, SHEET_RESUMEN, SHEET_INSUMOS } from './config.js'
import { fechaLaboral, getSemanaActual, normalizarProveedor, normaliza } from './utils.js'
import { sheetsClient, getRows } from './sheets.js'
import ExcelJS from 'exceljs'
import fs from 'fs'
import path from 'path'
import os from 'os'

const OPENAI_API_KEY = process.env.OPENAI_API_KEY

async function leerTicketConOpenAI(bufferImagen){
  try{
    const base64 = bufferImagen.toString('base64')
    const res = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${OPENAI_API_KEY}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: 'gpt-4o',
        messages: [
          {
            role: 'user',
            content: [
              { type: 'text', text: `Eres lector de tickets de compras para restaurante Trinidad. Analiza esta imagen que puede ser FACTURA impresa o NOTA MANUSCRITA.

Devuelve SOLO JSON válido así:
{
  "proveedor": "nombre proveedor en MAYUSCULAS",
  "folio": "AC-12345 o folio real",
  "importe_total": "1234.56 solo numero",
  "fecha_ticket": "YYYY-MM-DD",
  "concepto": "CARNE, POLLO, VERDURAS, ABARROTES, LACTEOS, LIMPIEZA, DESECHABLES, OTROS",
  "productos": [{"nombre": "...", "cantidad": "1", "precio": "100", "total": "100"}]
}

Reglas:
- Si es manuscrita, intenta leer lo mejor posible.
- Si no hay folio, inventa AC- + 6 numeros.
- Proveedor: si dice Gastro, pon GASTROSOPHIA, si dice 3B pon TIENDAS TRES B, si dice Florentina pon QUESOS FLORENTINA, si dice Marco Antonio pon MARCO ANTONIO.
- Importe: busca TOTAL.
- Si no ves nada, igual intenta.` },
              { type: 'image_url', image_url: { url: `data:image/jpeg;base64,${base64}` } }
            ]
          }
        ],
        max_tokens: 1000,
        temperature: 0.1
      })
    })
    const data = await res.json()
    const content = data.choices?.[0]?.message?.content || ''
    console.log('OpenAI RAW:', content)
    // Extrae JSON
    const match = content.match(/\{[\s\S]*\}/)
    if(!match) return null
    const json = JSON.parse(match[0])
    return {
      proveedor: json.proveedor || 'PROVEEDOR OCASIONAL',
      folio: json.folio || `AC-${Date.now().toString().slice(-6)}`,
      importe_total: (json.importe_total||'0').toString().replace(/,/g,''),
      fecha_ticket: json.fecha_ticket || fechaLaboral(),
      concepto_sugerido: json.concepto || 'OTROS',
      productos: json.productos || null
    }
  }catch(e){
    console.log('OpenAI fail', e.message)
    return null
  }
}

async function registrarCompra(datos){
  const sClient = await sheetsClient()
  const provNorm = normalizarProveedor(datos.proveedor)
  const proveedorFinal = provNorm || "PROVEEDOR OCASIONAL"
  await sClient.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_COMPRAS_ID,
    range: `${SHEET_RESUMEN}!A:I`,
    valueInputOption:'USER_ENTERED',
    requestBody:{ values:[[ datos.semana, datos.fecha, datos.sucursal, proveedorFinal, datos.folio, datos.monto, datos.concepto, "", "" ]] }
  })
  try{
    if(datos.productos && datos.productos.length>0){
      let rows = datos.productos.map(p=> [ datos.folio, p.nombre||datos.concepto, p.cantidad||"1", p.precio||"", p.total||datos.monto, datos.sucursal, datos.fecha ])
      await sClient.spreadsheets.values.append({ spreadsheetId: SPREADSHEET_COMPRAS_ID, range: `${SHEET_INSUMOS}!A:G`, valueInputOption:'USER_ENTERED', requestBody:{ values: rows } })
    } else {
      await sClient.spreadsheets.values.append({ spreadsheetId: SPREADSHEET_COMPRAS_ID, range: `${SHEET_INSUMOS}!A:G`, valueInputOption:'USER_ENTERED', requestBody:{ values:[[ datos.folio, datos.concepto, "1", datos.monto, datos.monto, datos.sucursal, datos.fecha ]] } })
    }
  }catch(e){ console.log(e) }
  return { ok:true, msg: `✅ ${datos.folio} $${datos.monto} - ${datos.concepto} - ${proveedorFinal} guardado con OpenAI.` }
}

export async function generarExcelCompras(filtro, jid, sock){
  const rows = await getRows(`${SHEET_RESUMEN}!A2:I`, SPREADSHEET_COMPRAS_ID)
  let filtradas = rows.filter(r=>{ if(filtro.sucursal &&!(r[2]||'').toUpperCase().includes(filtro.sucursal.toUpperCase())) return false; if(filtro.fecha && r[1]!==filtro.fecha) return false; return true })
  const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('Compras'); ws.addRow(['Semana','Fecha','Sucursal','Proveedor','#Comprobante','Importe','Concepto','Area','FormaPago']).font={bold:true}; filtradas.forEach(r=>ws.addRow(r)); const total = filtradas.reduce((a,r)=> a + (parseFloat((r[5]||'0').toString().replace(/,/g,''))||0), 0); ws.addRow([]); ws.addRow(['TOTAL','','','','',total]); ws.columns.forEach(c=>c.width=18); const fileName = `Compras_${filtro.sucursal||'TODAS'}_${fechaLaboral()}.xlsx`; const fp = path.join(os.tmpdir(), fileName); await wb.xlsx.writeFile(fp); await sock.sendMessage(jid,{document:fs.readFileSync(fp),mimetype:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',fileName, caption: `📊 Compras ${filtro.sucursal||''} - ${filtradas.length} regs - Total $${total}`}); fs.unlinkSync(fp)
}

function parseFiltroCompras(texto){ const low = normaliza(texto); let suc = null; if(low.includes('bucareli')) suc='BUCARELI'; else if(low.includes('coyo')) suc='COYOACAN'; else if(low.includes('juarez')) suc='JUAREZ'; let fecha = null; if(low.includes('hoy')) fecha = fechaLaboral(); return { sucursal:suc, fecha } }

export async function handleCompras({ sock, jid, m, texto, esImagen }){
  if(/^compras/i.test(texto)){ const f = parseFiltroCompras(texto); await generarExcelCompras(f, jid, sock); return true }
  if(esImagen){
    try{
      if(!OPENAI_API_KEY){ await sock.sendMessage(jid,{text:'❌ Falta OPENAI_API_KEY en Railway'}); return true }
      const buffer = await downloadMediaMessage(m, 'buffer', {}, { logger: P({level:'fatal'}), reuploadRequest: sock.updateMediaMessage })
      await sock.sendMessage(jid,{text:`🤖 Leyendo con OpenAI (manuscrita + factura)...`})
      let datosIA = await leerTicketConOpenAI(buffer)
      if(!datosIA || datosIA.importe_total=="0"){ await sock.sendMessage(jid,{text:`No pude leer el total con OpenAI. Manda foto más clara`}); return true }
      let datosFinal = {
        folio: datosIA.folio,
        concepto: datosIA.concepto_sugerido,
        proveedor: datosIA.proveedor,
        monto: datosIA.importe_total,
        fecha: datosIA.fecha_ticket || fechaLaboral(),
        sucursal: "BUCARELI",
        semana: getSemanaActual(),
        productos: datosIA.productos
      }
      const res = await registrarCompra(datosFinal)
      await sock.sendMessage(jid,{text: res.msg})
      return true
    }catch(e){ console.error(e); await sock.sendMessage(jid,{text:'Error OpenAI: '+e.message}); return true }
  }
  return false
}

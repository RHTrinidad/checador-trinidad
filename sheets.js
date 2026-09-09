const { google } = require('googleapis')
async function guardarEnSheet(sucursal, nombre, tipo, jid){
  try{
    const creds = JSON.parse(process.env.GOOGLE_CREDS)
    const auth = new google.auth.GoogleAuth({ credentials: creds, scopes: ['https://www.googleapis.com/auth/spreadsheets'] })
    const sheets = google.sheets({version:'v4', auth})
    const fecha = new Date().toLocaleString("es-MX",{timeZone:"America/Mexico_City"})
    await sheets.spreadsheets.values.append({
      spreadsheetId: process.env.SHEET_ID,
      range: 'A:E',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[fecha, sucursal, nombre, tipo, jid]] }
    })
  }catch(e){ console.log("Error Sheet:", e.message) }
}
module.exports = { guardarEnSheet }

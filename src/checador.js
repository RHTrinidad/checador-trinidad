```js
import { SUCURSALES, GRUPO_COYOACAN_ID, GRUPO_BUCARELI_ID, SPREADSHEET_ID } from './config.js'
import { distM, fechaLaboral, horaMX, minutos, calcularHorasTrabajadas, calcularExtra, parseHorarioRango } from './utils.js'
import { getRows, sheetsClient } from './sheets.js'

const TOLERANCIA_MIN=15
const AVISO_FALTA_MIN=20
const CIERRE_AUTO_HORAS=16

let ultimoRegistroDescansos=''

function esDescanso(v){
  return (v||'').toString().trim().toLowerCase().includes('descanso')
}

function diaMexico(){
  return new Date(
    new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'})
  ).getDay()
}

function horarioDelDia(r,dia){
  const mapa={
    1:r[3],
    2:r[4],
    3:r[5],
    4:r[6],
    5:r[7],
    6:r[8],
    0:r[9]
  }
  return (mapa[dia]||'').toString().trim()
}

async function avisoYaRegistrado(tel,fecha,tipo='NO HA LLEGADO'){
  try{
    const rows=await getRows('Avisos!A:H')
    const tel10=(tel||'').replace(/\D/g,'').slice(-10)

    return rows.some((r,i)=>{
      if(i===0)return false

      const fechaRow=(r[0]||'').toString().trim()
      const telRow=(r[1]||'').toString().replace(/\D/g,'').slice(-10)
      const tipoRow=(r[7]||'').toString().trim().toUpperCase()

      return fechaRow===fecha &&
        telRow===tel10 &&
        tipoRow===tipo.toUpperCase()
    })
  }catch(e){
    console.error('Error consultando Avisos:',e)
    return false
  }
}

async function registrarAviso({
  fecha,
  tel,
  nombre,
  programada,
  minutosTarde,
  sucursal,
  tipo='NO HA LLEGADO'
}){
  try{
    const client=await sheetsClient()
    const ahora=horaMX()
    const tel10=(tel||'').replace(/\D/g,'').slice(-10)

    await client.spreadsheets.values.append({
      spreadsheetId:SPREADSHEET_ID,
      range:'Avisos!A:H',
      valueInputOption:'USER_ENTERED',
      requestBody:{
        values:[[
          fecha,
          tel10,
          nombre,
          programada,
          minutosTarde,
          sucursal,
          ahora,
          tipo
        ]]
      }
    })

    return true
  }catch(e){
    console.error('Error registrando Aviso:',e)
    return false
  }
}

export async function handleChecador({sock,jid,m,loc,rawLid,tel10,tel}){
  const lat=loc.degreesLatitude
  const lng=loc.degreesLongitude

  let cercana=null
  let dMin=Infinity

  for(const s of SUCURSALES){
    const d=distM(lat,lng,s.lat,s.lng)
    if(d<dMin){
      dMin=d
      cercana=s
    }
  }

  const empRows=await getRows('Empleados!A:K')

  const getLid=r=>(r.find(x=>String(x).includes('@lid'))||'').trim()

  let emp=null

  if(tel10.length>=10){
    const i=empRows.findIndex((r,idx)=>
      idx>0 &&
      r[0] &&
      r[0].replace(/\D/g,'').slice(-10)===tel10
    )

    if(i>-1)emp=empRows[i]
  }

  if(!e
```

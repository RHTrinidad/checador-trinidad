import { createRequire } from 'module'
const require=createRequire(import.meta.url)
const {google}=require('googleapis')

import {parseHorarioRango} from './utils.js'
import {SPREADSHEET_ID} from './config.js'

const auth=new google.auth.GoogleAuth({
  credentials:JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
  scopes:['https://www.googleapis.com/auth/spreadsheets']
})

export async function sheetsClient(){
  const c=await auth.getClient()
  return google.sheets({version:'v4',auth:c})
}

export async function getRows(range,sid){
  const s=await sheetsClient()
  const r=await s.spreadsheets.values.get({
    spreadsheetId:sid||SPREADSHEET_ID,
    range
  })
  return r.data.values||[]
}

export async function buscarEmpleadoPorTelefono(tel,sid){
  try{
    const tel10=(tel||'').toString().replace(/\D/g,'').slice(-10)
    if(!tel10)return null

    const rows=await getRows('Empleados!A:T',sid)

    for(let i=1;i<rows.length;i++){
      const row=rows[i]
      const telRow=(row[0]||'').toString().replace(/\D/g,'').slice(-10)

      if(telRow===tel10){
        return {
          rowIndex:i+1,
          row
        }
      }
    }

    return null
  }catch(e){
    console.error('Error buscando empleado por teléfono:',e)
    return null
  }
}

export async function buscarEmpleadoPorLid(lid,sid){
  try{
    const valor=(lid||'').toString().trim()
    if(!valor||!valor.includes('@lid'))return null

    const rows=await getRows('Empleados!A:T',sid)

    for(let i=1;i<rows.length;i++){
      const row=rows[i]
      const lidRow=(row[10]||'').toString().trim()

      if(lidRow===valor){
        return {
          rowIndex:i+1,
          row
        }
      }
    }

    return null
  }catch(e){
    console.error('Error buscando empleado por LID:',e)
    return null
  }
}

export async function guardarLidEmpleado(rowIndex,lid,sid){
  try{
    const valor=(lid||'').toString().trim()

    if(!rowIndex||!valor||!valor.includes('@lid'))return false

    const client=await sheetsClient()

    await client.spreadsheets.values.update({
      spreadsheetId:sid||SPREADSHEET_ID,
      range:`Empleados!K${rowIndex}`,
      valueInputOption:'USER_ENTERED',
      requestBody:{
        values:[[valor]]
      }
    })

    return true
  }catch(e){
    console.error('Error guardando LID del empleado:',e)
    return false
  }
}

export async function getHorarioBaseMap(){
  try{
    const rows=await getRows('Horario_Base!A2:K')
    const map={}

    for(const f of rows){
      const tel=(f[0]||'').replace(/\D/g,'').slice(-10)
      const nombre=(f[1]||'').toLowerCase().trim()

      if(!nombre)continue

      const dias={
        1:f[3],
        2:f[4],
        3:f[5],
        4:f[6],
        5:f[7],
        6:f[8],
        0:f[9]
      }

      const descansos=new Set()
      const horas={}

      for(const [k,v] of Object.entries(dias)){
        const val=(v||'').toString().trim()

        if(!val){
          descansos.add(parseInt(k))
          continue
        }

        const p=parseHorarioRango(val)

        if(!p)descansos.add(parseInt(k))
        else horas[k]=p
      }

      const obj={
        descansos,
        horas,
        nombreOriginal:f[1],
        tel,
        sucursal:f[2]||''
      }

      if(tel)map[tel]=obj
      map[nombre]=obj

      const pri=nombre.split(' ')[0]

      if(pri&&!map[pri])map[pri]=obj
    }

    return map
  }catch(e){
    console.error('Error obteniendo Horario_Base:',e)
    return {}
  }
}

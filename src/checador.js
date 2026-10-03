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
  calcularHorasTrabajadas,
  calcularExtra,
  parseHorarioRango
} from './utils.js'

import {
  getRows,
  sheetsClient,
  buscarEmpleadoPorTelefono,
  buscarEmpleadoPorLid,
  guardarLidEmpleado
} from './sheets.js'

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

async function buscarEmpleado({tel10,rawLid}){
  let resultado=null

  if(rawLid){
    resultado=await buscarEmpleadoPorLid(rawLid)
    if(resultado)return resultado
  }

  if(tel10){
    resultado=await buscarEmpleadoPorTelefono(tel10)
    if(resultado)return resultado
  }

  return null
}

function datosEmpleado(resultado){
  if(!resultado?.row)return null

  const r=resultado.row

  return {
    rowIndex:resultado.rowIndex,
    telefono:(r[0]||'').toString(),
    nombreCorto:(r[1]||'').toString().trim(),
    sucursal:(r[2]||'').toString().trim(),
    nombre:(r[3]||r[1]||'').toString().trim(),
    puesto:(r[4]||'').toString().trim(),
    lid:(r[10]||'').toString().trim()
  }
}

async function identificarEmpleado({tel10,rawLid,m}){
  const resultado=await buscarEmpleado({tel10,rawLid})
  const emp=datosEmpleado(resultado)

  if(emp){
    if(rawLid && emp.lid!==rawLid){
      await guardarLidEmpleado(emp.rowIndex,rawLid)
      emp.lid=rawLid
    }

    return emp
  }

  return null
}

export async function handleChecador({
  sock,
  jid,
  m,
  loc,
  rawLid,
  tel10,
  tel
}){
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

  const fecha=fechaLaboral()
  const ahora=horaMX()
  const emp=await identificarEmpleado({tel10,rawLid,m})

  const nombre=emp?.nombre||m.pushName||tel10||'Desconocido'
  const telefono=emp?.telefono||tel||tel10||''
  const telefono10=telefono.replace(/\D/g,'').slice(-10)

  const horarioRows=await getRows('Horario_Base!A2:K')
  const dia=diaMexico()

  let horarioEmp=null

  for(const r of horarioRows){
    const telRow=(r[0]||'').toString().replace(/\D/g,'').slice(-10)
    const nombreRow=(r[1]||'').toString().trim().toLowerCase()

    if(
      (telefono10 && telRow===telefono10) ||
      (nombreRow && nombreRow===nombre.toLowerCase())
    ){
      horarioEmp=r
      break
    }
  }

  const horario=horarioEmp?horarioDelDia(horarioEmp,dia):''
  const descanso=esDescanso(horario)

  const partes=horario.match(/(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})/)

  let horaEntrada=''
  let horaSalida=''

  if(partes){
    horaEntrada=partes[1]
    horaSalida=partes[2]
  }

  const tipoRegistro=await determinarTipoRegistro(
    telefono10,
    fecha
  )

  if(descanso){
    await responder(sock,jid,
      `🏖️ *${nombre}*\n\nHoy tienes descanso según tu horario.`
    )
    return
  }

  if(tipoRegistro==='ENTRADA'){
    await registrarEntrada({
      fecha,
      hora:ahora,
      telefono:telefono10,
      nombre,
      sucursal:cercana?.nombre||'',
      distancia:dMin,
      programada:horaEntrada,
      emp,
      sock,
      jid
    })
    return
  }

  await registrarSalida({
    fecha,
    hora:ahora,
    telefono:telefono10,
    nombre,
    sucursal:cercana?.nombre||'',
    distancia:dMin,
    programada:horaSalida,
    emp,
    sock,
    jid
  })
}

async function determinarTipoRegistro(tel,fecha){
  try{
    const rows=await getRows('Asistencia!A:K')
    let entrada=false
    let salida=false

    for(let i=1;i<rows.length;i++){
      const r=rows[i]

      const fechaRow=(r[0]||'').toString().trim()
      const telRow=(r[1]||'').toString().replace(/\D/g,'').slice(-10)
      const tipo=(r[4]||'').toString().trim().toUpperCase()

      if(fechaRow!==fecha || telRow!==tel)continue

      if(tipo.includes('ENTRADA')||tipo.includes('A TIEMPO')||tipo.includes('RETARDO')){
        entrada=true
      }

      if(tipo.includes('SALIDA')){
        salida=true
      }
    }

    if(!entrada)return 'ENTRADA'
    if(!salida)return 'SALIDA'

    return 'ENTRADA'
  }catch(e){
    console.error('Error determinando tipo de registro:',e)
    return 'ENTRADA'
  }
}

async function registrarEntrada({
  fecha,
  hora,
  telefono,
  nombre,
  sucursal,
  distancia,
  programada,
  emp,
  sock,
  jid
}){
  try{
    const client=await sheetsClient()

    let estado='A TIEMPO'
    let minutosTarde=0

    if(programada){
      minutosTarde=minutos(hora,programada)

      if(minutosTarde>0)estado='RETARDO'
    }

    await client.spreadsheets.values.append({
      spreadsheetId:SPREADSHEET_ID,
      range:'Asistencia!A:K',
      valueInputOption:'USER_ENTERED',
      requestBody:{
        values:[[
          telefono,
          nombre,
          fecha,
          hora,
          estado,
          minutosTarde,
          sucursal,
          distancia,
          '',
          emp?.puesto||'',
          programada||''
        ]]
      }
    })

    await responder(
      sock,
      jid,
      `✅ *ENTRADA REGISTRADA*\n\n`+
      `👤 ${nombre}\n`+
      `📅 ${fecha}\n`+
      `🕐 ${hora}\n`+
      `📍 ${sucursal}\n`+
      `📏 ${Math.round(distancia)} m\n`+
      `📋 ${estado}`+
      (programada?`\n⏰ Prog: ${programada}`:'')
    )

    return true
  }catch(e){
    console.error('Error registrando entrada:',e)

    await responder(
      sock,
      jid,
      '❌ No se pudo registrar la entrada.'
    )

    return false
  }
}

async function registrarSalida({
  fecha,
  hora,
  telefono,
  nombre,
  sucursal,
  distancia,
  programada,
  emp,
  sock,
  jid
}){
  try{
    const rows=await getRows('Asistencia!A:K')
    let entrada=null
    let rowIndex=-1

    for(let i=rows.length-1;i>=1;i--){
      const r=rows[i]

      const fechaRow=(r[0]||'').toString().trim()
      const telRow=(r[1]||'').toString().replace(/\D/g,'').slice(-10)
      const tipo=(r[4]||'').toString().trim().toUpperCase()

      if(
        fechaRow===fecha &&
        telRow===telefono &&
        !tipo.includes('SALIDA')
      ){
        entrada=r
        rowIndex=i+1
        break
      }
    }

    if(!entrada){
      await responder(
        sock,
        jid,
        `⚠️ *No encontré una entrada registrada hoy para ${nombre}.*`
      )
      return false
    }

    const horaEntrada=(entrada[3]||'').toString().trim()
    const horas=calcularHorasTrabajadas(horaEntrada,hora)
    const extra=calcularExtra(horas)

    const client=await sheetsClient()

    await client.spreadsheets.values.append({
      spreadsheetId:SPREADSHEET_ID,
      range:'Asistencia!A:K',
      valueInputOption:'USER_ENTERED',
      requestBody:{
        values:[[
          telefono,
          nombre,
          fecha,
          hora,
          'SALIDA',
          '',
          sucursal,
          distancia,
          horas,
          extra,
          programada||''
        ]]
      }
    })

    await responder(
      sock,
      jid,
      `✅ *SALIDA REGISTRADA*\n\n`+
      `👤 ${nombre}\n`+
      `🕐 Entrada: ${horaEntrada}\n`+
      `🕐 Salida: ${hora}\n`+
      `⏱️ Horas: ${horas}`+
      (extra?`\n➕ Extra: ${extra}`:'')
    )

    return true
  }catch(e){
    console.error('Error registrando salida:',e)

    await responder(
      sock,
      jid,
      '❌ No se pudo registrar la salida.'
    )

    return false
  }
}

export async function registrarDescansos(){
  const fecha=fechaLaboral()

  if(ultimoRegistroDescansos===fecha)return

  try{
    const rows=await getRows('Horario_Base!A2:K')
    const dia=diaMexico()

    const client=await sheetsClient()
    const valores=[]

    for(const r of rows){
      const telefono=(r[0]||'').toString().replace(/\D/g,'').slice(-10)
      const nombre=(r[1]||'').toString().trim()
      const sucursal=(r[2]||'').toString().trim()
      const horario=horarioDelDia(r,dia)

      if(!telefono||!nombre||!esDescanso(horario))continue

      valores.push([
        telefono,
        nombre,
        fecha,
        '',
        'DESCANSO',
        '',
        sucursal,
        '',
        '',
        '',
        ''
      ])
    }

    if(valores.length){
      await client.spreadsheets.values.append({
        spreadsheetId:SPREADSHEET_ID,
        range:'Asistencia!A:K',
        valueInputOption:'USER_ENTERED',
        requestBody:{values:valores}
      })
    }

    ultimoRegistroDescansos=fecha
  }catch(e){
    console.error('Error registrando descansos:',e)
  }
}

export async function cerrarSalidasPendientes(){
  try{
    const fecha=fechaLaboral()
    const rows=await getRows('Asistencia!A:K')

    const pendientes=[]

    for(let i=1;i<rows.length;i++){
      const r=rows[i]

      const fechaRow=(r[2]||'').toString().trim()
      const tipo=(r[4]||'').toString().trim().toUpperCase()
      const horaEntrada=(r[3]||'').toString().trim()

      if(fechaRow!==fecha||!horaEntrada)continue
      if(tipo.includes('SALIDA')||tipo==='DESCANSO')continue

      const telefono=(r[0]||'').toString().replace(/\D/g,'').slice(-10)

      if(!telefono)continue

      const yaSalida=rows.some(x=>
        (x[2]||'').toString().trim()===fecha &&
        (x[0]||'').toString().replace(/\D/g,'').slice(-10)===telefono &&
        (x[4]||'').toString().toUpperCase().includes('SALIDA')
      )

      if(!yaSalida)pendientes.push({r,telefono})
    }

    if(!pendientes.length)return

    const ahora=horaMX()
    const horaActual=parseHorarioRango(`00:00-${ahora}`)

    if(!horaActual)return

    for(const p of pendientes){
      const horaEntrada=(p.r[3]||'').toString().trim()
      const horas=calcularHorasTrabajadas(horaEntrada,ahora)

      if(!horas)continue

      const num=parseFloat(horas)||0
      if(num<CIERRE_AUTO_HORAS)continue

      const client=await sheetsClient()

      await client.spreadsheets.values.append({
        spreadsheetId:SPREADSHEET_ID,
        range:'Asistencia!A:K',
        valueInputOption:'USER_ENTERED',
        requestBody:{
          values:[[
            p.telefono,
            p.r[1]||'',
            fecha,
            ahora,
            'SALIDA AUTOMATICA',
            '',
            p.r[6]||'',
            p.r[7]||'',
            horas,
            calcularExtra(horas),
            p.r[10]||''
          ]]
        }
      })
    }
  }catch(e){
    console.error('Error cerrando salidas pendientes:',e)
  }
}

export async function checkNoLlegaron(sock){
  try{
    const fecha=fechaLaboral()
    const ahora=horaMX()
    const horaAhora=minutos(ahora,'00:00')

    const horarioRows=await getRows('Horario_Base!A2:K')
    const asistenciaRows=await getRows('Asistencia!A:K')

    const dia=diaMexico()

    for(const r of horarioRows){
      const telefono=(r[0]||'').toString().replace(/\D/g,'').slice(-10)
      const nombre=(r[1]||'').toString().trim()
      const sucursal=(r[2]||'').toString().trim()
      const horario=horarioDelDia(r,dia)

      if(!telefono||!nombre||!horario||esDescanso(horario))continue

      const partes=horario.match(/(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})/)
      if(!partes)continue

      const programada=partes[1]
      const minProgramada=minutos(programada,'00:00')

      if(horaAhora-minProgramada<AVISO_FALTA_MIN)continue

      const tieneEntrada=asistenciaRows.some(x=>{
        const fechaRow=(x[2]||'').toString().trim()
        const telRow=(x[0]||'').toString().replace(/\D/g,'').slice(-10)
        const tipo=(x[4]||'').toString().trim().toUpperCase()

        return fechaRow===fecha &&
          telRow===telefono &&
          (
            tipo.includes('ENTRADA')||
            tipo.includes('A TIEMPO')||
            tipo.includes('RETARDO')
          )
      })

      if(tieneEntrada)continue

      const yaAviso=await avisoYaRegistrado(
        telefono,
        fecha,
        'NO HA LLEGADO'
      )

      if(yaAviso)continue

      const minutosTarde=horaAhora-minProgramada

      let grupoAviso=null

      const sucId=sucursal
        .toString()
        .trim()
        .toLowerCase()

      if(
        sucId.includes('coyo')||
        sucId.includes('hotel')
      ){
        grupoAviso=GRUPO_COYOACAN_ID
      }else if(
        sucId.includes('bucareli')
      ){
        grupoAviso=GRUPO_BUCARELI_ID
      }

      if(!grupoAviso){
        console.error(
          `⚠️ Sucursal sin grupo configurado: "${sucursal}" - ${nombre}`
        )
        continue
      }

      const mensaje=
        `⚠️ *NO HA LLEGADO* - ${nombre}\n`+
        `Prog: ${programada}\n`+
        `Más de ${AVISO_FALTA_MIN} min sin registrar entrada\n`+
        `[${sucursal}]`

      await sock.sendMessage(
        grupoAviso,
        {text:mensaje}
      )

      await registrarAviso({
        fecha,
        tel:telefono,
        nombre,
        programada,
        minutosTarde,
        sucursal,
        tipo:'NO HA LLEGADO'
      })
    }
  }catch(e){
    console.error('Error revisando no llegados:',e)
  }
}

async function responder(sock,jid,texto){
  try{
    await sock.sendMessage(jid,{text:texto})
  }catch(e){
    console.error('Error enviando respuesta:',e)
  }
}

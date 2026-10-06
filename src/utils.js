export function distM(a,b,c,d){
  const R=6371000
  const toRad=x=>x*Math.PI/180

  const dLa=toRad(c-a)
  const dLo=toRad(d-b)

  const q=
    Math.sin(dLa/2)**2+
    Math.cos(toRad(a))*
    Math.cos(toRad(c))*
    Math.sin(dLo/2)**2

  return R*2*Math.atan2(Math.sqrt(q),Math.sqrt(1-q))
}

export function fechaLaboral(d=new Date()){
  const mx=new Date(
    d.toLocaleString('en-US',{
      timeZone:'America/Mexico_City'
    })
  )

  export function diaLaboral(d=new Date()){
  const mx=new Date(
    d.toLocaleString('en-US',{
      timeZone:'America/Mexico_City'
    })
  )

  /*
    El día laboral cambia a las 05:00 AM.

    00:00 - 04:59 todavía pertenece
    al día laboral anterior.

    Ejemplo:
    Martes 01:30 → lunes laboral
    Martes 05:00 → martes laboral
  */
  if(mx.getHours()<5){
    mx.setDate(mx.getDate()-1)
  }

  return mx.getDay()
}
  /*
    El día laboral cambia a las 05:00 AM.
    Esto permite que una salida, por ejemplo a la 01:00 AM,
    siga perteneciendo al día laboral anterior.
  */
  if(mx.getHours()<5){
    mx.setDate(mx.getDate()-1)
  }

  return `${mx.getFullYear()}-${String(mx.getMonth()+1).padStart(2,'0')}-${String(mx.getDate()).padStart(2,'0')}`
}

export function getSemanaActual(){
  const hoy=new Date()
  const jan1=new Date(hoy.getFullYear(),0,1)

  return Math.ceil(
    (((hoy-jan1)/86400000)+jan1.getDay()+1)/7
  ).toString()
}

export function horaMX(d=new Date()){
  const partes=new Intl.DateTimeFormat('en-GB',{
    timeZone:'America/Mexico_City',
    hour:'2-digit',
    minute:'2-digit',
    second:'2-digit',
    hourCycle:'h23'
  }).formatToParts(d)

  const datos={}

  for(const p of partes){
    if(p.type!=='literal'){
      datos[p.type]=p.value
    }
  }

  return `${datos.hour}:${datos.minute}:${datos.second}`
}

export function minutos(h){
  const mm=h?.toString().match(/(\d{1,2}):(\d{2})/)

  return mm
    ?parseInt(mm[1])*60+parseInt(mm[2])
    :0
}

export function parseHorarioRango(v){
  const s=(v||'').toString().trim()

  if(!s)return null

  const low=s.toLowerCase()

  if(
    low.includes('libre')||
    low.includes('flex')
  ){
    return{
      entrada:'LIBRE',
      salida:null,
      raw:s
    }
  }

  if(low.includes('descanso'))return null

  const m=s.match(
    /(\d{1,2}:\d{2})\s*(?:-|a)?\s*(\d{1,2}:\d{2})?/i
  )

  if(!m)return null

  return{
    entrada:m[1],
    salida:m[2]||null,
    raw:s
  }
}

/*
  Convierte una hora HH:MM o HH:MM:SS a minutos.
*/
function minutosHora(h){
  if(!h)return null

  const partes=h.toString().trim().split(':').map(Number)

  if(!Number.isFinite(partes[0]))return null

  return(
    (partes[0]||0)*60+
    (partes[1]||0)
  )
}

/*
  Calcula minutos trabajados entre entrada y salida.

  Si la salida pertenece al día siguiente,
  se corrige automáticamente.

  IMPORTANTE:
  Ya no se fuerza "8 horas" para cualquier registro.
  La jornada real se conserva.
*/
export function minutosTrabajados(hE,hS){
  const entrada=minutosHora(hE)
  const salida=minutosHora(hS)

  if(
    entrada===null||
    salida===null
  ){
    return 0
  }

  let diff=salida-entrada

  if(diff<0){
    diff+=24*60
  }

  return diff
}

/*
  Convierte minutos a HH:MM:SS.

  Para mantener compatibilidad con el formato
  que ya usa Asistencia.
*/
export function formatoHoras(min){
  const minutosTotal=Math.max(0,Math.round(min||0))

  const h=Math.floor(minutosTotal/60)
  const m=minutosTotal%60

  return `${h}:${String(m).padStart(2,'0')}:00`
}

/*
  Calcula las horas realmente trabajadas.

  Jornada libre:
    - Si no hay horario de salida, se conserva
      la lógica de 8 horas cuando corresponda.

  Jornada programada:
    - Se calcula exactamente el tiempo entre entrada y salida.
*/
export function calcularHorasTrabajadas(hE,hS){
  if(!hE)return''

  /*
    Si todavía no existe salida, no inventamos
    horas trabajadas.
  */
  if(!hS)return''

  const diff=minutosTrabajados(hE,hS)

  return formatoHoras(diff)
}

/*
  Redondea horas extra HACIA ABAJO a bloques de 30 minutos.

  Ejemplos:
    15  -> 0
    29  -> 0
    30  -> 30
    44  -> 30
    59  -> 30
    60  -> 60
    75  -> 60
    89  -> 60
    90  -> 90
*/
export function redondearExtra30(extraMin){
  const minutosExtra=Math.max(
    0,
    Math.floor(Number(extraMin)||0)
  )

  return Math.floor(
    minutosExtra/30
  )*30
}

/*
  Calcula las horas extra.

  REGLA:
  - La jornada programada se obtiene de pe -> ps.
  - La jornada libre conserva 8 horas como base.
  - El extra REAL se conserva sin redondear.
  - El extraRedondeado se utiliza posteriormente
    en reportes para bloques de 30 minutos.

  Esto permite que Sheets conserve el tiempo real,
  mientras que los reportes puedan aplicar la regla:
    1:15 -> 1:00
    1:29 -> 1:00
    1:30 -> 1:30
    1:45 -> 1:30
    2:00 -> 2:00
*/
export function calcularExtra(hE,hS,pe,ps){
  const trab=calcularHorasTrabajadas(hE,hS)

  /*
    Sin salida todavía.
  */
  if(!hS){
    return{
      trabajadas:trab,
      extra:'0',
      extraMin:0,
      extraRedondeada:'0',
      extraMinRedondeada:0
    }
  }

  /*
    Jornada libre:
    base = 8 horas.
  */
  if(
    !pe||
    !ps||
    pe==='LIBRE'
  ){
    const minutosTrab=minutosTrabajados(hE,hS)

    const ex=Math.max(
      0,
      minutosTrab-(8*60)
    )

    const exRed=redondearExtra30(ex)

    return{
      trabajadas:trab,
      extra:ex>0?formatoHoras(ex):'0',
      extraMin:ex,
      extraRedondeada:exRed>0?formatoHoras(exRed):'0',
      extraMinRedondeada:exRed
    }
  }

  const entradaProgramada=minutosHora(pe)
  const salidaProgramada=minutosHora(ps)

  if(
    entradaProgramada===null||
    salidaProgramada===null
  ){
    return{
      trabajadas:trab,
      extra:'0',
      extraMin:0,
      extraRedondeada:'0',
      extraMinRedondeada:0
    }
  }

  let jornadaProgramada=
    salidaProgramada-entradaProgramada

  if(jornadaProgramada<0){
    jornadaProgramada+=24*60
  }

  const minutosTrab=minutosTrabajados(
    hE,
    hS
  )

  /*
    El extra se determina contra la jornada
    programada completa.

    Ejemplos:

    09:14 - 19:00
    jornada 09:00 - 19:00
    trabajado = 9:46
    jornada = 10:00
    extra = 0

    09:44 - 19:00
    jornada 09:00 - 18:00
    trabajado = 9:16
    jornada = 9:00
    extra = 0:16

    09:44 - 18:44
    jornada 09:00 - 18:00
    trabajado = 9:00
    jornada = 9:00
    extra = 0
  */
  const ex=Math.max(
    0,
    minutosTrab-jornadaProgramada
  )

  const exRed=redondearExtra30(ex)

  return{
    trabajadas:trab,
    extra:ex>0?formatoHoras(ex):'0',
    extraMin:ex,
    extraRedondeada:exRed>0?formatoHoras(exRed):'0',
    extraMinRedondeada:exRed
  }
}

export function parseFechaMX(s){
  if(!s)return null

  if(/^\d{4}-\d{2}-\d{2}$/.test(s)){
    return new Date(
      s+'T12:00:00'
    )
  }

  const m=s.match(/(\d{1,2})\/(\d{2,4})/)

  if(m){
    return new Date(
      `${m[3].length==2?'20'+m[3]:m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}T12:00:00`
    )
  }

  return new Date(s)
}

export function normaliza(s){
  return(s||'')
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'')
    .toLowerCase()
    .trim()
}

export function getRangoSemana(t){
  const hoy=new Date(
    new Date().toLocaleString(
      'en-US',
      {
        timeZone:'America/Mexico_City'
      }
    )
  )

  const d=hoy.getDay()

  const lunes=new Date(hoy)

  lunes.setDate(
    hoy.getDate()-(d===0?6:d-1)
  )

  let l
  let dom

  if(t==='pasada'){
    l=new Date(lunes)
    l.setDate(
      l.getDate()-7
    )

    dom=new Date(l)
    dom.setDate(
      dom.getDate()+6
    )
  }else{
    l=lunes
    dom=hoy
  }

  l.setHours(0,0,0,0)
  dom.setHours(23,59,59,999)

  return{
    lunes:l,
    domingo:dom,
    rangoTxt:
      `${l.toLocaleDateString('es-MX')} al ${dom.toLocaleDateString('es-MX')} (${t})`
  }
}

export function sucursalCoincideConFiltro(s,f){
  if(!f)return true

  const x=(s||'').toLowerCase()

  if(f==='coyoacan'){
    return x.includes('coyo')||x.includes('hotel')
  }

  if(f==='bucareli'){
    return x.includes('bucareli')
  }

  return x.includes(f)
}

export function scoreEmpleado(c,d,b){
  const cc=normaliza(c)
  const dd=normaliza(d)

  if(cc===b)return 100
  if(cc.startsWith(b))return 90
  if(cc.includes(b)||dd.includes(b))return 10

  return-1
}

export function normalizarProveedor(n){
  let x=(n||"")
    .toUpperCase()
    .trim()

  if(
    x.includes("VICTOR HUGO")||
    x.includes("TRINIDAD")
  ){
    return null
  }

  if(x.includes("GASTRO"))
    return"GASTROSOPHIA"

  if(
    x.includes("TRES B")||
    x.includes("3B")
  ){
    return"TIENDAS TRES B"
  }

  if(x.includes("FLORENTINA"))
    return"QUESOS FLORENTINA"

  if(!x||x.length<3)
    return"PROVEEDOR OCASIONAL"

  return x
}

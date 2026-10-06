import{getRows,sheetsClient}from'../sheets.js'
import{SPREADSHEET_RECETAS_ID,GRUPO_COYOACAN_ID,GRUPO_BUCARELI_ID,GRUPO_REPORTES_TRINIDAD_ID,GRUPO_PRUEBAS_ID}from'../config.js'

const CONTEXTO_MENU=new Map()

const limpiarTexto=x=>
  (x||'').toString().trim().replace(/\s+/g,' ')

const normaliza=x=>
  limpiarTexto(x)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'')
    .toLowerCase()

const dinero=x=>{

  const n=parseFloat(
    (x||'')
      .toString()
      .replace(/[$,%\s,]/g,'')
  )

  return Number.isFinite(n)?n:0
}

const numeroSolicitud=x=>{

  let s=
    limpiarTexto(x)
      .replace(/\s/g,'')

  if(s.includes(',')&&s.includes('.')){

    s=s.replace(/,/g,'')

  }else if(s.includes(',')){

    s=s.replace(',','.')
  }

  const n=Number(s)

  return Number.isFinite(n)?n:0
}

const moneda=x=>
  `$${Number(x||0).toLocaleString('es-MX',{
    minimumFractionDigits:2,
    maximumFractionDigits:2
  })}`

const guardarContexto=(jid,datos)=>
  CONTEXTO_MENU.set(jid,{
    ...datos,
    expira:Date.now()+60000
  })

const obtenerContexto=jid=>{

  const x=CONTEXTO_MENU.get(jid)

  if(!x)
    return null

  if(Date.now()>x.expira){

    CONTEXTO_MENU.delete(jid)

    return null
  }

  return x
}

const borrarContexto=jid=>
  CONTEXTO_MENU.delete(jid)

const esReporteGeneral=jid=>
  jid===GRUPO_REPORTES_TRINIDAD_ID||
  jid===GRUPO_PRUEBAS_ID

const sucursalGrupo=jid=>{

  if(jid===GRUPO_BUCARELI_ID)
    return'bucareli'

  if(jid===GRUPO_COYOACAN_ID)
    return'coyoacan'

  return''
}

const nombreSucursal=x=>{

  const n=normaliza(x)

  if(n.includes('coyoacan'))
    return'coyoacan'

  if(n.includes('bucareli'))
    return'bucareli'

  return''
}

const nombreArea=x=>{

  const n=normaliza(x)

  if(n.includes('cocina'))
    return'cocina'

  if(n.includes('barra'))
    return'barra'

  return''
}

const etiquetaSucursal=x=>
  normaliza(x)==='coyoacan'
    ?'Coyoacán'
    :'Bucareli'

const etiquetaArea=x=>
  normaliza(x)==='cocina'
    ?'COCINA'
    :'BARRA'

const esComandoMenu=t=>{

  const n=normaliza(t)

  return /^(receta|recetas|recetario|menu|menú|cotizar|cotizacion|cotización)\b/i.test(n)
}

/* =========================================================
   TIPO DE REPORTE
========================================================= */

const tipoReporte=texto=>{

  const n=normaliza(texto)

  if(
    n==='1'||
    /^costos?$/i.test(n)
  ){

    return'costos'
  }

  if(
    n==='2'||
    /^(descriptivo|descritivo|descripcion)$/i.test(n)
  ){

    return'descriptivo'
  }

  if(
    n==='3'||
    /^procesos?$/i.test(n)
  ){

    return'procesos'
  }

  return''
}

const tipoReporteSubreceta=texto=>{

  const n=normaliza(texto)

  if(
    n==='1'||
    /^costos?$/i.test(n)
  ){

    return'costos'
  }

  if(
    n==='2'||
    /^procesos?$/i.test(n)
  ){

    return'procesos'
  }

  return''
}

/* =========================================================
   UNIDADES
========================================================= */

const convertirUnidadBase=(cantidad,unidad)=>{

  const n=normaliza(unidad)

  if([
    'l',
    'lt',
    'lts',
    'lto',
    'litro',
    'litros',
    'lit',
    'ltrs'
  ].includes(n)){

    return{
      cantidad:cantidad*1000,
      unidad:'ml'
    }
  }

  if([
    'ml',
    'mililitro',
    'mililitros',
    'mililitro'
  ].includes(n)){

    return{
      cantidad,
      unidad:'ml'
    }
  }

  if([
    'kg',
    'kgs',
    'kilo',
    'kilos',
    'kilogramo',
    'kilogramos'
  ].includes(n)){

    return{
      cantidad:cantidad*1000,
      unidad:'g'
    }
  }

  /*
   * GR / GRL / G / GRAMO
   *
   * Para el sistema:
   *
   * 1 kg = 1000 gr
   * 1 kg = 1000 grl
   *
   * Por lo tanto GR y GRL pertenecen
   * exactamente a la misma unidad base.
   */

  if([
    'g',
    'gr',
    'grl',
    'grs',
    'grls',
    'gramo',
    'gramos',
    'graml',
    'gramls'
  ].includes(n)){

    return{
      cantidad,
      unidad:'g'
    }
  }

  if([
    'pz',
    'pza',
    'pzas',
    'pieza',
    'piezas'
  ].includes(n)){

    return{
      cantidad,
      unidad:'pz'
    }
  }

  return{
    cantidad,
    unidad:n
  }
}

const formatearCantidad=x=>{

  const n=Number(x||0)

  if(!Number.isFinite(n))
    return'0'

  if(Number.isInteger(n))
    return n.toLocaleString('es-MX')

  return n.toLocaleString('es-MX',{
    maximumFractionDigits:3
  })
}

/* =========================================================
   DETECTAR SOLICITUD DE SUBRECETA
========================================================= */

function extraerTipoFinalSubreceta(nombre){

  let texto=limpiarTexto(nombre)

  let tipo=''

  const m=texto.match(
    /\s+(costos?|procesos?)$/i
  )

  if(m){

    const palabra=normaliza(m[1])

    if(palabra.startsWith('costo'))
      tipo='costos'

    if(palabra.startsWith('proceso'))
      tipo='procesos'

    texto=limpiarTexto(
      texto.slice(
        0,
        texto.length-m[0].length
      )
    )
  }

  return{
    nombre:texto,
    tipo
  }
}

function parseSolicitudSubreceta(texto){

  const t=limpiarTexto(texto)

  const m=t.match(
    /^receta(?:s)?\s+(?:para\s+([\d.,]+)\s*([a-zA-ZáéíóúñÁÉÍÓÚÑ]+)\s+de\s+)?(sub\b.+)$/i
  )

  if(!m)
    return null

  const cantidad=
    m[1]
      ?numeroSolicitud(m[1])
      :null

  const unidad=
    m[2]||''

  const extraido=
    extraerTipoFinalSubreceta(m[3])

  return{

    cantidad:
      cantidad&&cantidad>0
        ?cantidad
        :null,

    unidad,

    nombre:limpiarTexto(
      extraido.nombre
    ),

    tipoSolicitado:
      extraido.tipo||''
  }
}

/* =========================================================
   LECTURA ROBUSTA DE HOJAS
========================================================= */

async function leerHojaPorNombre(sheets,titulo){

  const r=await sheets.spreadsheets.get({
    spreadsheetId:SPREADSHEET_RECETAS_ID,
    includeGridData:true
  })

  const hoja=r.data.sheets?.find(
    s=>s.properties?.title===titulo
  )

  if(!hoja){

    throw new Error(
      `No existe la hoja "${titulo}" en el archivo de recetas.`
    )
  }

  const bloques=hoja.data||[]

  const filas=[]

  for(const bloque of bloques){

    for(const row of (bloque.rowData||[])){

      const valores=(row.values||[]).map(c=>{

        if(c?.formattedValue!==undefined)
          return c.formattedValue

        if(c?.effectiveValue?.stringValue!==undefined)
          return c.effectiveValue.stringValue

        if(c?.effectiveValue?.numberValue!==undefined)
          return c.effectiveValue.numberValue

        if(c?.effectiveValue?.boolValue!==undefined)
          return c.effectiveValue.boolValue

        return''
      })

      filas.push(valores)
    }
  }

  return filas
}

async function cargarRecetas(){

  const sheets=await sheetsClient()

  const[
    catalogo,
    subrecetas,
    rendimientos,
    platillos,
    menu
  ]=await Promise.all([

    leerHojaPorNombre(
      sheets,
      '1_CATALOGO'
    ),

    leerHojaPorNombre(
      sheets,
      '2_SUBRECETAS'
    ),

    leerHojaPorNombre(
      sheets,
      '3_RENDIMIENTOS'
    ),

    leerHojaPorNombre(
      sheets,
      '4_PLATILLOS'
    ),

    leerHojaPorNombre(
      sheets,
      '5_MENU_COSTOS'
    )

  ])

  return{
    catalogo,
    subrecetas,
    rendimientos,
    platillos,
    menu
  }
}

/* =========================================================
   MENÚ / PLATILLOS
========================================================= */

function activo(x){

  const n=normaliza(x)

  return !x||
    [
      'si',
      'sí',
      'activo',
      'activa',
      '1',
      'true',
      'verdadero',
      'yes'
    ].includes(n)
}

function filaMenuCoincideSucursal(r,sucursal){

  if(!sucursal)
    return true

  const n=normaliza(r[1])

  return n.includes(sucursal)||
    n.includes('ambos')||
    n.includes('coyoacan y bucareli')||
    n.includes('bucareli y coyoacan')
}

function filasMenu(menu,sucursal='',area=''){

  return menu.slice(1).filter(r=>{

    if(!normaliza(r[0]))
      return false

    if(!activo(r[4]))
      return false

    if(!filaMenuCoincideSucursal(r,sucursal))
      return false

    if(area&&normaliza(r[2])!==area)
      return false

    return true
  })
}

function buscarPlatillos(
  nombre,
  menu,
  platillos,
  sucursal='',
  area=''
){

  const n=normaliza(nombre)

  const lista=filasMenu(
    menu,
    sucursal,
    area
  )

  const nombres=new Map()

  for(const r of lista){

    const p=normaliza(r[0])

    if(p)
      nombres.set(p,r[0])
  }

  if(!sucursal&&!area){

    for(const r of platillos.slice(1)){

      const p=normaliza(r[0])

      if(p&&!nombres.has(p))
        nombres.set(p,r[0])
    }
  }

  const exactos=[]
  const similares=[]

  for(const[nombreNormal,nombreReal]of nombres){

    if(nombreNormal===n){

      exactos.push(nombreReal)

    }else if(
      nombreNormal.replace(/\s+/g,'')===
      n.replace(/\s+/g,'')
    ){

      exactos.push(nombreReal)

    }else if(
      nombreNormal.includes(n)||
      n.includes(nombreNormal)
    ){

      similares.push(nombreReal)

    }else{

      const palabras=
        n.split(' ').filter(Boolean)

      if(
        palabras.length>1&&
        palabras.every(
          w=>nombreNormal.includes(w)
        )
      ){

        similares.push(nombreReal)
      }
    }
  }

  return[
    ...new Set([
      ...exactos,
      ...similares
    ])
  ]
}

function buscarPlatillo(nombre,rows){

  const n=normaliza(nombre)

  const datos=rows
    .slice(1)
    .filter(r=>normaliza(r[0]))

  return datos.find(
    r=>normaliza(r[0])===n
  )||

  datos.find(
    r=>
      normaliza(r[0]).replace(/\s+/g,'')===
      n.replace(/\s+/g,'')
  )||

  datos.find(
    r=>normaliza(r[0]).includes(n)
  )||

  datos.find(
    r=>n.includes(normaliza(r[0]))
  )||

  null
}

/* =========================================================
   CATÁLOGO / RENDIMIENTOS
========================================================= */

function mapaCatalogo(rows){

  const map=new Map()

  for(const r of rows.slice(1)){

    const nombre=normaliza(r[0])

    if(!nombre)
      continue

    map.set(nombre,{

      nombre:r[0]||'',

      unidad:r[1]||'',

      costo:dinero(r[5]),

      alergico:r[6]||''
    })
  }

  return map
}

function mapaRendimientos(rows){

  const map=new Map()

  for(const r of rows.slice(1)){

    const sub=normaliza(r[0])

    if(!sub)
      continue

    if(!map.has(sub))
      map.set(sub,[])

    map.get(sub).push({

      subreceta:r[0]||'',

      rendimiento:dinero(r[1]),

      unidad:r[2]||'',

      sucursal:r[3]||'',

      tipo:r[4]||'',

      costoTotal:dinero(r[5]),

      costoUnidad:dinero(r[6]),

      /*
       * Columna H de 3_RENDIMIENTOS
       */
      proceso:r[7]||''
    })
  }

  return map
}

function obtenerRendimientoSubreceta(
  nombre,
  map,
  origen
){

  const lista=
    map.get(
      normaliza(nombre)
    )||[]

  if(!lista.length)
    return null

  const o=normaliza(origen)

  let x=lista.find(
    r=>o&&normaliza(r.sucursal)===o
  )

  if(x)
    return x

  x=lista.find(
    r=>o&&(
      normaliza(r.sucursal).includes(o)||
      o.includes(normaliza(r.sucursal))
    )
  )

  if(x)
    return x

  const suc=
    o.includes('bucareli')
      ?'bucareli'
      :o.includes('coyoacan')
        ?'coyoacan'
        :''

  if(suc){

    x=lista.find(
      r=>normaliza(r.sucursal).includes(suc)
    )

    if(x)
      return x
  }

  return lista[0]
}

function obtenerIngredientesPlatillo(nombre,rows){

  const n=normaliza(nombre)

  return rows
    .slice(1)
    .filter(
      r=>normaliza(r[0])===n
    )
}

/* =========================================================
   SUBRECETAS
========================================================= */

function listaNombresSubrecetas(rows){

  const nombres=new Map()

  for(const r of rows.slice(1)){

    const nombre=r[0]||''
    const n=normaliza(nombre)

    if(!n)
      continue

    /*
     * Todas las subrecetas comienzan con SUB.
     */
    if(!/^sub\b/.test(n))
      continue

    if(!nombres.has(n))
      nombres.set(n,nombre)
  }

  return[
    ...nombres.values()
  ]
}

function buscarSubrecetas(nombre,rows){

  const n=normaliza(nombre)

  const lista=
    listaNombresSubrecetas(rows)

  const exactas=[]
  const similares=[]

  for(const nombreReal of lista){

    const x=normaliza(nombreReal)

    if(x===n){

      exactas.push(nombreReal)

      continue
    }

    if(
      x.replace(/\s+/g,'')===
      n.replace(/\s+/g,'')
    ){

      exactas.push(nombreReal)

      continue
    }

    if(
      x.includes(n)||
      n.includes(x)
    ){

      similares.push(nombreReal)

      continue
    }

    const palabras=
      n.split(' ').filter(Boolean)

    if(
      palabras.length>1&&
      palabras.every(
        w=>x.includes(w)
      )
    ){

      similares.push(nombreReal)
    }
  }

  return[
    ...new Set([
      ...exactas,
      ...similares
    ])
  ]
}

function esSubreceta(nombre,subrows){

  const n=normaliza(nombre)

  return subrows
    .slice(1)
    .some(
      r=>normaliza(r[0])===n
    )
}

function obtenerFilasSubreceta(nombre,rows){

  const n=normaliza(nombre)

  return rows
    .slice(1)
    .filter(
      r=>normaliza(r[0])===n
    )
}

function buscarIngredienteCatalogo(nombre,map){

  const n=normaliza(nombre)

  if(map.has(n))
    return map.get(n)

  for(const[k,v]of map){

    if(
      k.replace(/\s+/g,'')===
      n.replace(/\s+/g,'')
    ){

      return v
    }
  }

  return null
}

function agregarAlergeno(
  lista,
  nombre,
  alergico
){

  if(!alergico)
    return

  const clave=normaliza(nombre)

  if(
    !lista.some(
      x=>normaliza(x.nombre)===clave
    )
  ){

    lista.push({
      nombre,
      alergico
    })
  }
}

function limpiarNombreSubreceta(nombre){

  return(
    nombre||''
  )
    .toString()
    .replace(/^sub\s+/i,'')
    .trim()
}

function obtenerAlergenosSubreceta(
  nombre,
  subrows,
  catalogoMap,
  visitados=new Set()
){

  const clave=normaliza(nombre)

  if(visitados.has(clave))
    return[]

  visitados.add(clave)

  const salida=[]

  const filas=subrows
    .slice(1)
    .filter(
      r=>normaliza(r[0])===clave
    )

  for(const r of filas){

    const ingrediente=r[1]||''

    if(!ingrediente)
      continue

    const cat=
      buscarIngredienteCatalogo(
        ingrediente,
        catalogoMap
      )

    if(cat?.alergico){

      agregarAlergeno(
        salida,
        limpiarNombreSubreceta(
          ingrediente
        ),
        cat.alergico
      )

      continue
    }

    if(esSubreceta(ingrediente,subrows)){

      const sub=
        obtenerAlergenosSubreceta(
          ingrediente,
          subrows,
          catalogoMap,
          new Set(visitados)
        )

      for(const a of sub){

        agregarAlergeno(
          salida,
          a.nombre,
          a.alergico
        )
      }
    }
  }

  return salida
}

/* =========================================================
   REPORTE COSTOS SUBRECETA
========================================================= */

async function reporteCostosSubreceta({
  sock,
  jid,
  subreceta,
  sucursal='',
  cantidad=null,
  unidadSolicitud=''
}){

  try{

    const{
      catalogo,
      subrecetas,
      rendimientos
    }=await cargarRecetas()

    const filas=
      obtenerFilasSubreceta(
        subreceta,
        subrecetas
      )

    if(!filas.length){

      await sock.sendMessage(jid,{
        text:
          `❌ No encontré "${subreceta}" `+
          `en 2_SUBRECETAS.`
      })

      return true
    }

    const rendMap=
      mapaRendimientos(
        rendimientos
      )

    const rend=
      obtenerRendimientoSubreceta(
        subreceta,
        rendMap,
        sucursal
      )

    if(!rend){

      await sock.sendMessage(jid,{
        text:
          `❌ Encontré "${subreceta}" `+
          `pero no tiene rendimiento registrado `+
          `en 3_RENDIMIENTOS.`
      })

      return true
    }

    const rendimientoBase=
      Number(rend.rendimiento||0)

    const unidadBase=
      convertirUnidadBase(
        rendimientoBase,
        rend.unidad
      )

    let factor=1

    let solicitudBase=null

    if(cantidad!==null){

      solicitudBase=
        convertirUnidadBase(
          cantidad,
          unidadSolicitud
        )

      if(
        solicitudBase.unidad!==
        unidadBase.unidad
      ){

        await sock.sendMessage(jid,{
          text:
            `❌ No puedo convertir `+
            `${unidadSolicitud} a ${rend.unidad} `+
            `para "${subreceta}".\n\n`+
            `El rendimiento de esta subreceta `+
            `está registrado en ${rend.unidad}.`
        })

        return true
      }

      if(
        !unidadBase.cantidad||
        unidadBase.cantidad<=0
      ){

        await sock.sendMessage(jid,{
          text:
            `❌ El rendimiento de "${subreceta}" `+
            `no es válido.`
        })

        return true
      }

      factor=
        solicitudBase.cantidad/
        unidadBase.cantidad
    }

    const catalogoMap=
      mapaCatalogo(catalogo)

    const ingredientes=[]

    let totalCalculado=0

    for(const r of filas){

      const ingrediente=r[1]||''

      if(!ingrediente)
        continue

      const cantidadBase=
        dinero(r[2])

      const unidad=
        r[3]||''

      const costoTotalBase=
        dinero(r[5])

      const cantidadFinal=
        cantidadBase*factor

      const costoFinal=
        costoTotalBase*factor

      totalCalculado+=costoFinal

      ingredientes.push({

        nombre:
          esSubreceta(
            ingrediente,
            subrecetas
          )
            ?limpiarNombreSubreceta(
              ingrediente
            )
            :ingrediente,

        cantidad:cantidadFinal,

        unidad,

        costo:costoFinal
      })
    }

    let total=
      rend.costoTotal
        ?rend.costoTotal*factor
        :totalCalculado

    if(!total)
      total=totalCalculado

    const alergenos=
      obtenerAlergenosSubreceta(
        subreceta,
        subrecetas,
        catalogoMap
      )

    let texto=
      `🍲 ${subreceta.toUpperCase()}\n`

    if(sucursal){

      texto+=
        `📍 ${etiquetaSucursal(sucursal)}\n`
    }

    texto+=
      `📏 RENDIMIENTO BASE: `+
      `${formatearCantidad(unidadBase.cantidad)} `+
      `${unidadBase.unidad.toUpperCase()}`

    if(cantidad!==null){

      texto+=
        `\n📦 SOLICITADO: `+
        `${formatearCantidad(solicitudBase.cantidad)} `+
        `${solicitudBase.unidad.toUpperCase()}`

      texto+=
        `\n🔢 FACTOR DE CONVERSIÓN: `+
        `${factor.toLocaleString('es-MX',{
          maximumFractionDigits:4
        })}`
    }

    texto+=
      '\n\nIngredientes:\n'

    for(const x of ingredientes){

      texto+=
        `• ${x.nombre} ........ `+
        `${formatearCantidad(x.cantidad)} `+
        `${x.unidad} ........ `+
        `${moneda(x.costo)}\n`
    }

    texto+=
      `\n💰 COSTO TOTAL: ${moneda(total)}`

    if(alergenos.length){

      texto+=
        '\n\n────────────────────\n'+
        '⚠️ ALÉRGENOS:\n'

      for(const a of alergenos){

        texto+=
          `• ${a.nombre} — `+
          `posible alérgeno: ${a.alergico}\n`
      }
    }

    await sock.sendMessage(jid,{
      text:texto
    })

    return true

  }catch(e){

    console.error(
      '❌ ERROR COSTOS SUBRECETA:',
      e
    )

    await sock.sendMessage(jid,{
      text:
        '❌ Ocurrió un error consultando el costo de la subreceta.'
    })

    return true
  }
}

/* =========================================================
   REPORTE PROCESO SUBRECETA
========================================================= */

async function reporteProcesosSubreceta({
  sock,
  jid,
  subreceta,
  sucursal='',
  cantidad=null,
  unidadSolicitud=''
}){

  try{

    const{
      rendimientos
    }=await cargarRecetas()

    const rendMap=
      mapaRendimientos(
        rendimientos
      )

    const rend=
      obtenerRendimientoSubreceta(
        subreceta,
        rendMap,
        sucursal
      )

    if(!rend){

      await sock.sendMessage(jid,{
        text:
          `❌ No encontré rendimiento para `+
          `"${subreceta}" en 3_RENDIMIENTOS.`
      })

      return true
    }

    const unidadBase=
      convertirUnidadBase(
        Number(rend.rendimiento||0),
        rend.unidad
      )

    let solicitud=null

    if(cantidad!==null){

      solicitud=
        convertirUnidadBase(
          cantidad,
          unidadSolicitud
        )

      if(
        solicitud.unidad!==
        unidadBase.unidad
      ){

        await sock.sendMessage(jid,{
          text:
            `❌ No puedo convertir `+
            `${unidadSolicitud} a ${rend.unidad} `+
            `para esta subreceta.`
        })

        return true
      }
    }

    let texto=
      `👨‍🍳 ${subreceta.toUpperCase()}\n`

    if(sucursal){

      texto+=
        `📍 ${etiquetaSucursal(sucursal)}\n`
    }

    texto+=
      `📏 RENDIMIENTO: `+
      `${formatearCantidad(unidadBase.cantidad)} `+
      `${unidadBase.unidad.toUpperCase()}`

    if(cantidad!==null){

      texto+=
        `\n📦 SOLICITADO: `+
        `${formatearCantidad(solicitud.cantidad)} `+
        `${solicitud.unidad.toUpperCase()}`
    }

    /*
     * Columna H de 3_RENDIMIENTOS.
     *
     * Si existe proceso, se muestra.
     * Si está vacía, no agregamos ningún
     * mensaje adicional.
     */

    if(limpiarTexto(rend.proceso)){

      texto+=
        `\n\n👨‍🍳 PROCEDIMIENTO:\n`+
        `${rend.proceso}`
    }

    await sock.sendMessage(jid,{
      text:texto
    })

    return true

  }catch(e){

    console.error(
      '❌ ERROR PROCESOS SUBRECETA:',
      e
    )

    await sock.sendMessage(jid,{
      text:
        '❌ Ocurrió un error consultando el proceso de la subreceta.'
    })

    return true
  }
}

/* =========================================================
   MOSTRAR OPCIONES SUBRECETA
========================================================= */

async function mostrarCoincidenciasSubreceta({
  sock,
  jid,
  coincidencias,
  sucursal='',
  cantidad=null,
  unidadSolicitud='',
  tipoSolicitado=''
}){

  const lista=
    coincidencias.slice(0,15)

  if(lista.length===1){

    const elegida=
      lista[0]

    if(tipoSolicitado){

      borrarContexto(jid)

      if(tipoSolicitado==='costos'){

        return await reporteCostosSubreceta({
          sock,
          jid,
          subreceta:elegida,
          sucursal,
          cantidad,
          unidadSolicitud
        })
      }

      if(tipoSolicitado==='procesos'){

        return await reporteProcesosSubreceta({
          sock,
          jid,
          subreceta:elegida,
          sucursal,
          cantidad,
          unidadSolicitud
        })
      }
    }

    guardarContexto(jid,{

      subreceta:elegida,

      sucursal,

      cantidad,

      unidadSolicitud,

      esperandoTipoSubreceta:true
    })

    await sock.sendMessage(jid,{
      text:
        `🍲 ${elegida}\n\n`+
        `¿Qué reporte quieres?\n\n`+
        `1. Costos\n`+
        `2. Procesos`
    })

    return true
  }

  guardarContexto(jid,{

    coincidencias:lista,

    sucursal,

    cantidad,

    unidadSolicitud,

    tipoSolicitado,

    esperandoSubreceta:true
  })

  let texto=
    `🍲 Encontré ${lista.length} subrecetas que coinciden:\n\n`

  lista.forEach((x,i)=>{

    texto+=
      `${i+1}. ${x}\n`
  })

  texto+=
    '\nEscribe el número o el nombre de la subreceta.'

  await sock.sendMessage(jid,{
    text:texto
  })

  return true
}

/* =========================================================
   MOSTRAR OPCIONES PLATILLO
========================================================= */

async function mostrarCoincidencias({
  sock,
  jid,
  coincidencias,
  sucursal='',
  area=''
}){

  const lista=
    coincidencias.slice(0,15)

  if(lista.length===1){

    guardarContexto(jid,{

      platillo:lista[0],

      sucursal,

      area,

      esperandoTipo:true
    })

    await sock.sendMessage(jid,{
      text:
        `🍽️ ${lista[0]}\n\n`+
        `¿Qué reporte quieres?\n\n`+
        `1. Costos\n`+
        `2. Descriptivo\n`+
        `3. Procesos`
    })

    return true
  }

  guardarContexto(jid,{

    coincidencias:lista,

    sucursal,

    area,

    esperandoPlatillo:true
  })

  let texto=
    '🍽️ Encontré estos platillos:\n\n'

  lista.forEach((x,i)=>{

    texto+=
      `${i+1}. ${x}\n`
  })

  texto+=
    '\nEscribe el nombre del platillo o su número.'

  await sock.sendMessage(jid,{
    text:texto
  })

  return true
}

/* =========================================================
   REPORTES DE PLATILLOS
========================================================= */

function buscarFilaMenu(
  nombre,
  menu,
  sucursal='',
  area=''
){

  const lista=
    filasMenu(
      menu,
      sucursal,
      area
    )

  const n=normaliza(nombre)

  return lista.find(
    r=>normaliza(r[0])===n
  )||

  lista.find(
    r=>
      normaliza(r[0]).replace(/\s+/g,'')===
      n.replace(/\s+/g,'')
  )||

  lista.find(
    r=>normaliza(r[0]).includes(n)
  )||

  lista.find(
    r=>n.includes(normaliza(r[0]))
  )||

  null
}

async function reporteListaMenu({
  sock,
  jid,
  sucursal='',
  area=''
}){

  try{

    const{menu}=await cargarRecetas()

    const filas=
      filasMenu(
        menu,
        sucursal,
        area
      )

    if(!filas.length){

      await sock.sendMessage(jid,{
        text:
          `❌ No encontré productos activos para `+
          `${sucursal?etiquetaSucursal(sucursal):'el menú'}`+
          `${area?` — ${etiquetaArea(area)}`:''}.`
      })

      return true
    }

    const grupos={
      cocina:[],
      barra:[]
    }

    for(const r of filas){

      const a=
        normaliza(r[2])==='barra'
          ?'barra'
          :'cocina'

      grupos[a].push(r)
    }

    let texto=
      `🍽️ MENÚ ${sucursal?etiquetaSucursal(sucursal).toUpperCase():'GENERAL'}\n`

    if(area)
      texto+=`\n${etiquetaArea(area)}\n`

    if(grupos.cocina.length){

      texto+='\n👨‍🍳 COCINA\n'

      for(const r of grupos.cocina){

        texto+=
          `• ${r[0]} — `+
          `costo ${moneda(dinero(r[6]))} — `+
          `venta ${moneda(dinero(r[5]))}\n`
      }
    }

    if(grupos.barra.length){

      texto+='\n🍸 BARRA\n'

      for(const r of grupos.barra){

        texto+=
          `• ${r[0]} — `+
          `costo ${moneda(dinero(r[6]))} — `+
          `venta ${moneda(dinero(r[5]))}\n`
      }
    }

    await sock.sendMessage(jid,{
      text:texto.trim()
    })

    return true

  }catch(e){

    console.error(
      '❌ ERROR LISTA MENU:',
      e
    )

    await sock.sendMessage(jid,{
      text:
        '❌ Ocurrió un error consultando el menú.'
    })

    return true
  }
}

async function reporteCostos({
  sock,
  jid,
  platillo,
  sucursal='',
  area=''
}){

  try{

    const{
      catalogo,
      subrecetas,
      rendimientos,
      platillos,
      menu
    }=await cargarRecetas()

    const filaMenu=
      buscarFilaMenu(
        platillo,
        menu,
        sucursal,
        area
      )

    let nombrePlatillo=
      filaMenu?.[0]||platillo

    const filaPlatillo=
      buscarPlatillo(
        nombrePlatillo,
        platillos
      )

    if(!filaPlatillo){

      await sock.sendMessage(jid,{
        text:
          `❌ No encontré el platillo `+
          `"${platillo}" en 4_PLATILLOS.`
      })

      return true
    }

    nombrePlatillo=
      filaPlatillo[0]

    const filas=
      obtenerIngredientesPlatillo(
        nombrePlatillo,
        platillos
      )

    if(!filas.length){

      await sock.sendMessage(jid,{
        text:
          `❌ Encontré "${nombrePlatillo}", `+
          `pero no tiene ingredientes registrados.`
      })

      return true
    }

    const catalogoMap=
      mapaCatalogo(catalogo)

    const rendMap=
      mapaRendimientos(
        rendimientos
      )

    let total=0

    const ingredientes=[]
    const alergenos=[]

    for(const r of filas){

      const ingrediente=
        r[1]||''

      const cantidad=
        dinero(r[2])

      const unidad=
        r[3]||''

      const origen=
        r[4]||''

      let costoUnidad=
        dinero(r[5])

      let costoTotal=
        dinero(r[6])

      const sub=
        esSubreceta(
          ingrediente,
          subrecetas
        )

      if(sub){

        const rend=
          obtenerRendimientoSubreceta(
            ingrediente,
            rendMap,
            origen
          )

        if(rend){

          costoUnidad=
            rend.costoUnidad

          costoTotal=
            cantidad*costoUnidad
        }

        const subAlergenos=
          obtenerAlergenosSubreceta(
            ingrediente,
            subrecetas,
            catalogoMap
          )

        for(const a of subAlergenos){

          agregarAlergeno(
            alergenos,
            limpiarNombreSubreceta(
              ingrediente
            ),
            a.alergico
          )
        }

      }else{

        const cat=
          buscarIngredienteCatalogo(
            ingrediente,
            catalogoMap
          )

        if(cat?.alergico){

          agregarAlergeno(
            alergenos,
            ingrediente,
            cat.alergico
          )
        }

        if(!costoUnidad&&cat){

          costoUnidad=
            cat.costo

          costoTotal=
            cantidad*costoUnidad
        }
      }

      if(
        !costoTotal&&
        cantidad&&
        costoUnidad
      ){

        costoTotal=
          cantidad*costoUnidad
      }

      total+=costoTotal

      ingredientes.push({

        nombre:
          sub
            ?limpiarNombreSubreceta(
              ingrediente
            )
            :ingrediente,

        cantidad,

        unidad,

        costo:costoTotal
      })
    }

    let texto=
      `🍽️ ${nombrePlatillo.toUpperCase()}\n`

    if(sucursal){

      texto+=
        `📍 ${etiquetaSucursal(sucursal)}`
    }

    if(area){

      texto+=
        `${sucursal?'\n':''}`+
        `📂 ${etiquetaArea(area)}`
    }

    texto+=
      '\n\nIngredientes:\n'

    for(const x of ingredientes){

      texto+=
        `• ${x.nombre} ........ `+
        `${x.cantidad} ${x.unidad} ........ `+
        `${moneda(x.costo)}\n`
    }

    texto+=
      `\n💰 COSTO DEL PLATILLO: ${moneda(total)}`

    if(alergenos.length){

      texto+=
        '\n\n────────────────────\n'+
        '⚠️ ALÉRGENOS:\n'

      for(const a of alergenos){

        texto+=
          `• ${a.nombre} — `+
          `posible alérgeno: ${a.alergico}\n`
      }
    }

    await sock.sendMessage(jid,{
      text:texto
    })

    return true

  }catch(e){

    console.error(
      '❌ ERROR REPORTE COSTOS MENU:',
      e
    )

    await sock.sendMessage(jid,{
      text:
        '❌ Ocurrió un error consultando la receta.'
    })

    return true
  }
}

async function reporteDescriptivo({
  sock,
  jid,
  platillo,
  sucursal='',
  area=''
}){

  try{

    const{menu}=await cargarRecetas()

    const fila=
      buscarFilaMenu(
        platillo,
        menu,
        sucursal,
        area
      )

    if(!fila){

      await sock.sendMessage(jid,{
        text:
          `❌ No encontré "${platillo}" `+
          `en el menú solicitado.`
      })

      return true
    }

    let texto=
      `🍽️ ${fila[0].toUpperCase()}\n`

    if(sucursal){

      texto+=
        `📍 ${etiquetaSucursal(sucursal)}\n`
    }

    if(area){

      texto+=
        `📂 ${etiquetaArea(area)}\n`
    }

    /*
     * Columna K = DESCRIPTIVO
     */

    if(fila[10]){

      texto+=
        `\n${fila[10]}`
    }

    /*
     * Columna M = FOTO_EMPLATADO
     */

    if(fila[12]){

      texto+=
        `\n\n📷 ${fila[12]}`
    }

    /*
     * Si ambos están vacíos:
     *
     * NO agregamos mensaje.
     */

    await sock.sendMessage(jid,{
      text:texto
    })

    return true

  }catch(e){

    console.error(
      '❌ ERROR DESCRIPTIVO MENU:',
      e
    )

    await sock.sendMessage(jid,{
      text:
        '❌ Ocurrió un error consultando el descriptivo.'
    })

    return true
  }
}

async function reporteProcesos({
  sock,
  jid,
  platillo,
  sucursal='',
  area=''
}){

  try{

    const{
      menu,
      platillos
    }=await cargarRecetas()

    const fila=
      buscarFilaMenu(
        platillo,
        menu,
        sucursal,
        area
      )

    if(!fila){

      await sock.sendMessage(jid,{
        text:
          `❌ No encontré "${platillo}" `+
          `en el menú solicitado.`
      })

      return true
    }

    const filas=
      obtenerIngredientesPlatillo(
        fila[0],
        platillos
      )

    let texto=
      `👨‍🍳 ${fila[0].toUpperCase()}\n`

    if(sucursal){

      texto+=
        `📍 ${etiquetaSucursal(sucursal)}\n`
    }

    if(area){

      texto+=
        `📂 ${etiquetaArea(area)}\n`
    }

    texto+=
      '\nIngredientes:\n'

    if(filas.length){

      for(const r of filas){

        texto+=
          `• ${r[1]||''} — `+
          `${r[2]||''} ${r[3]||''}\n`
      }

    }else{

      texto+=
        '• Sin ingredientes registrados.\n'
    }

    /*
     * Columna L = PROCEDIMIENTO
     */

    if(fila[11]){

      texto+=
        `\n👨‍🍳 PROCEDIMIENTO:\n${fila[11]}`
    }

    /*
     * Columna M = FOTO_EMPLATADO
     */

    if(fila[12]){

      texto+=
        `\n\n📷 ${fila[12]}`
    }

    await sock.sendMessage(jid,{
      text:texto
    })

    return true

  }catch(e){

    console.error(
      '❌ ERROR PROCESOS MENU:',
      e
    )

    await sock.sendMessage(jid,{
      text:
        '❌ Ocurrió un error consultando el proceso.'
    })

    return true
  }
}

/* =========================================================
   PROCESAR CONTEXTO
========================================================= */

async function procesarRespuestaContexto({
  sock,
  jid,
  texto,
  contexto
}){

  const n=normaliza(texto)

  /* =====================================================
     SELECCIÓN DE SUBRECETA
  ===================================================== */

  if(contexto.esperandoSubreceta){

    let subreceta=''

    if(/^\d+$/.test(n)){

      const indice=
        parseInt(n,10)-1

      if(
        indice>=0&&
        indice<contexto.coincidencias.length
      ){

        subreceta=
          contexto.coincidencias[indice]
      }

    }else{

      subreceta=
        contexto.coincidencias?.find(
          x=>normaliza(x)===n
        )||

        contexto.coincidencias?.find(
          x=>normaliza(x).includes(n)
        )||

        ''
    }

    if(!subreceta){

      await sock.sendMessage(jid,{
        text:
          `❌ No reconocí esa subreceta.\n\n`+
          `Escribe un número del 1 al `+
          `${contexto.coincidencias?.length||0}.`
      })

      return true
    }

    if(contexto.tipoSolicitado){

      const tipo=
        contexto.tipoSolicitado

      borrarContexto(jid)

      if(tipo==='costos'){

        return await reporteCostosSubreceta({
          sock,
          jid,
          subreceta,
          sucursal:contexto.sucursal||'',
          cantidad:contexto.cantidad,
          unidadSolicitud:
            contexto.unidadSolicitud||''
        })
      }

      return await reporteProcesosSubreceta({
        sock,
        jid,
        subreceta,
        sucursal:contexto.sucursal||'',
        cantidad:contexto.cantidad,
        unidadSolicitud:
          contexto.unidadSolicitud||''
      })
    }

    borrarContexto(jid)

    guardarContexto(jid,{

      subreceta,

      sucursal:contexto.sucursal||'',

      cantidad:contexto.cantidad,

      unidadSolicitud:
        contexto.unidadSolicitud||'',

      esperandoTipoSubreceta:true
    })

    await sock.sendMessage(jid,{
      text:
        `🍲 ${subreceta}\n\n`+
        `¿Qué reporte quieres?\n\n`+
        `1. Costos\n`+
        `2. Procesos`
    })

    return true
  }

  /* =====================================================
     TIPO DE REPORTE SUBRECETA
  ===================================================== */

  if(contexto.esperandoTipoSubreceta){

    const tipo=
      tipoReporteSubreceta(n)

    if(!tipo){

      await sock.sendMessage(jid,{
        text:
          '¿Qué reporte quieres?\n\n'+
          '1. Costos\n'+
          '2. Procesos'
      })

      return true
    }

    if(
      n==='3'||
      /^descript/i.test(n)
    ){

      await sock.sendMessage(jid,{
        text:
          '❌ Las subrecetas solo tienen:\n\n'+
          '1. Costos\n'+
          '2. Procesos'
      })

      return true
    }

    const subreceta=
      contexto.subreceta

    const sucursal=
      contexto.sucursal||''

    const cantidad=
      contexto.cantidad

    const unidadSolicitud=
      contexto.unidadSolicitud||''

    borrarContexto(jid)

    if(tipo==='costos'){

      return await reporteCostosSubreceta({
        sock,
        jid,
        subreceta,
        sucursal,
        cantidad,
        unidadSolicitud
      })
    }

    return await reporteProcesosSubreceta({
      sock,
      jid,
      subreceta,
      sucursal,
      cantidad,
      unidadSolicitud
    })
  }

  /* =====================================================
     ÁREA
  ===================================================== */

  if(contexto.esperandoArea){

    const area=
      nombreArea(n)

    if(!area){

      await sock.sendMessage(jid,{
        text:
          '📂 Indica el área:\n\n'+
          '• Cocina\n'+
          '• Barra'
      })

      return true
    }

    const suc=
      contexto.sucursal||''

    if(suc){

      borrarContexto(jid)

      guardarContexto(jid,{
        accion:'buscar',
        sucursal:suc,
        area
      })

      await sock.sendMessage(jid,{
        text:
          `📋 RECETARIO ${etiquetaArea(area)} — `+
          `${etiquetaSucursal(suc)}\n\n`+
          `Escribe el nombre del platillo que buscas.`
      })

      return true
    }

    borrarContexto(jid)

    guardarContexto(jid,{
      accion:'recetario',
      area,
      esperandoSucursal:true
    })

    await sock.sendMessage(jid,{
      text:
        '📍 ¿De qué sucursal?\n\n'+
        '• Coyoacán\n'+
        '• Bucareli'
    })

    return true
  }

  /* =====================================================
     SUCURSAL
  ===================================================== */

  if(contexto.esperandoSucursal){

    const suc=
      nombreSucursal(n)

    if(!suc){

      await sock.sendMessage(jid,{
        text:
          '📍 Indica la sucursal:\n\n'+
          '• Coyoacán\n'+
          '• Bucareli'
      })

      return true
    }

    contexto.sucursal=suc
    contexto.esperandoSucursal=false

    if(contexto.accion==='lista'){

      borrarContexto(jid)

      return await reporteListaMenu({
        sock,
        jid,
        sucursal:suc,
        area:contexto.area||''
      })
    }

    if(contexto.accion==='recetario'){

      if(contexto.platillo){

        const{
          menu,
          platillos
        }=await cargarRecetas()

        const coincidencias=
          buscarPlatillos(
            contexto.platillo,
            menu,
            platillos,
            suc,
            contexto.area||''
          )

        borrarContexto(jid)

        if(!coincidencias.length){

          await sock.sendMessage(jid,{
            text:
              `❌ No encontré "${contexto.platillo}" `+
              `en el recetario solicitado.`
          })

          return true
        }

        return await mostrarCoincidencias({
          sock,
          jid,
          coincidencias,
          sucursal:suc,
          area:contexto.area||''
        })
      }

      borrarContexto(jid)

      guardarContexto(jid,{
        accion:'buscar',
        sucursal:suc,
        area:contexto.area||''
      })

      await sock.sendMessage(jid,{
        text:
          `📋 Recetario ${etiquetaArea(contexto.area)} `+
          `${etiquetaSucursal(suc)}.\n\n`+
          `Escribe el nombre del platillo que buscas.`
      })

      return true
    }
  }

  /* =====================================================
     TIPO DE REPORTE PLATILLO
  ===================================================== */

  if(contexto.esperandoTipo){

    const tipo=
      tipoReporte(n)

    if(!tipo){

      await sock.sendMessage(jid,{
        text:
          '¿Qué reporte quieres?\n\n'+
          '1. Costos\n'+
          '2. Descriptivo\n'+
          '3. Procesos'
      })

      return true
    }

    const platillo=
      contexto.platillo

    const sucursal=
      contexto.sucursal||''

    const area=
      contexto.area||''

    borrarContexto(jid)

    if(tipo==='costos'){

      return await reporteCostos({
        sock,
        jid,
        platillo,
        sucursal,
        area
      })
    }

    if(tipo==='descriptivo'){

      return await reporteDescriptivo({
        sock,
        jid,
        platillo,
        sucursal,
        area
      })
    }

    return await reporteProcesos({
      sock,
      jid,
      platillo,
      sucursal,
      area
    })
  }

  /* =====================================================
     SELECCIÓN DE PLATILLO
  ===================================================== */

  if(contexto.esperandoPlatillo){

    let platillo=''

    if(/^\d+$/.test(n)){

      const indice=
        parseInt(n,10)-1

      if(
        indice>=0&&
        indice<contexto.coincidencias.length
      ){

        platillo=
          contexto.coincidencias[indice]
      }

    }else{

      platillo=
        contexto.coincidencias?.find(
          x=>normaliza(x)===n
        )||

        contexto.coincidencias?.find(
          x=>normaliza(x).includes(n)
        )||

        ''
    }

    if(!platillo){

      await sock.sendMessage(jid,{
        text:
          `❌ No reconocí esa opción.\n\n`+
          `Escribe un número del 1 al `+
          `${contexto.coincidencias?.length||0}.`
      })

      return true
    }

    borrarContexto(jid)

    guardarContexto(jid,{

      platillo,

      sucursal:
        contexto.sucursal||'',

      area:
        contexto.area||'',

      esperandoTipo:true
    })

    await sock.sendMessage(jid,{
      text:
        `🍽️ ${platillo}\n\n`+
        `¿Qué reporte quieres?\n\n`+
        `1. Costos\n`+
        `2. Descriptivo\n`+
        `3. Procesos`
    })

    return true
  }

  /* =====================================================
     BÚSQUEDA DE RECETARIO
  ===================================================== */

  if(contexto.accion==='buscar'){

    const{
      menu,
      platillos
    }=await cargarRecetas()

    const coincidencias=
      buscarPlatillos(
        texto,
        menu,
        platillos,
        contexto.sucursal||'',
        contexto.area||''
      )

    if(!coincidencias.length){

      await sock.sendMessage(jid,{
        text:
          `❌ No encontré "${texto}" `+
          `en el recetario solicitado.`
      })

      return true
    }

    return await mostrarCoincidencias({
      sock,
      jid,
      coincidencias,
      sucursal:contexto.sucursal||'',
      area:contexto.area||''
    })
  }

  return false
}

/* =========================================================
   INICIAR SOLICITUD SUBRECETA
========================================================= */

async function iniciarSolicitudSubreceta({
  sock,
  jid,
  solicitud,
  tipoGrupo
}){

  try{

    const{
      subrecetas
    }=await cargarRecetas()

    const suc=
      sucursalGrupo(jid)

    const coincidencias=
      buscarSubrecetas(
        solicitud.nombre,
        subrecetas
      )

    if(!coincidencias.length){

      await sock.sendMessage(jid,{
        text:
          `❌ No encontré "${solicitud.nombre}" `+
          `en 2_SUBRECETAS.`
      })

      return true
    }

    return await mostrarCoincidenciasSubreceta({
      sock,
      jid,
      coincidencias,
      sucursal:suc,
      cantidad:solicitud.cantidad,
      unidadSolicitud:solicitud.unidad,
      tipoSolicitado:solicitud.tipoSolicitado
    })

  }catch(e){

    console.error(
      '❌ ERROR BUSCANDO SUBRECETA:',
      e
    )

    await sock.sendMessage(jid,{
      text:
        '❌ Ocurrió un error buscando la subreceta.'
    })

    return true
  }
}

/* =========================================================
   HANDLE PRINCIPAL
========================================================= */

export async function handleReportesMenu({
  sock,
  jid,
  m,
  texto,
  tipoGrupo
}){

  const t=
    limpiarTexto(texto)

  const n=
    normaliza(t)

  /*
   * ======================================================
   * SUBRECETA TIENE PRIORIDAD
   * ======================================================
   */

  const solicitudSub=
    parseSolicitudSubreceta(t)

  if(solicitudSub){

    borrarContexto(jid)

    return await iniciarSolicitudSubreceta({
      sock,
      jid,
      solicitud:solicitudSub,
      tipoGrupo
    })
  }

  /*
   * ======================================================
   * CONTEXTO
   *
   * IMPORTANTE:
   * Aquí entran también:
   *
   * 9
   * 1
   * 2
   * 3
   *
   * siempre que exista contexto pendiente.
   * ======================================================
   */

  const contexto=
    obtenerContexto(jid)

  if(contexto){

    const atendido=
      await procesarRespuestaContexto({
        sock,
        jid,
        texto:t,
        contexto
      })

    if(atendido)
      return true
  }

  /*
   * ======================================================
   * SI NO ES COMANDO, SALIMOS
   * ======================================================
   */

  if(!esComandoMenu(t))
    return false

  const sucGrupo=
    sucursalGrupo(jid)

  const esReportes=
    esReporteGeneral(jid)||
    tipoGrupo==='REPORTES'

  /* =====================================================
     MENU
  ===================================================== */

  if(/^menu\b/i.test(n)){

    const suc=
      nombreSucursal(n)

    const area=
      nombreArea(n)

    if(suc){

      return await reporteListaMenu({
        sock,
        jid,
        sucursal:suc,
        area
      })
    }

    if(sucGrupo){

      return await reporteListaMenu({
        sock,
        jid,
        sucursal:sucGrupo,
        area
      })
    }

    guardarContexto(jid,{
      accion:'lista',
      area,
      esperandoSucursal:true
    })

    await sock.sendMessage(jid,{
      text:
        '📍 ¿De qué sucursal?\n\n'+
        '• Coyoacán\n'+
        '• Bucareli'
    })

    return true
  }

  /* =====================================================
     RECETARIO
  ===================================================== */

  if(/^recetario\b/i.test(n)){

    const area=
      nombreArea(n)

    const suc=
      sucGrupo||
      nombreSucursal(n)

    if(area&&suc){

      guardarContexto(jid,{
        accion:'buscar',
        sucursal:suc,
        area
      })

      await sock.sendMessage(jid,{
        text:
          `📋 RECETARIO ${etiquetaArea(area)} — `+
          `${etiquetaSucursal(suc)}\n\n`+
          `Escribe el nombre del platillo que buscas.`
      })

      return true
    }

    if(area&&!suc&&!esReportes){

      guardarContexto(jid,{
        accion:'buscar',
        sucursal:'',
        area
      })

      await sock.sendMessage(jid,{
        text:
          `📋 RECETARIO ${etiquetaArea(area)}\n\n`+
          `Escribe el nombre del platillo que buscas.`
      })

      return true
    }

    if(area&&!suc&&esReportes){

      guardarContexto(jid,{
        accion:'recetario',
        area,
        esperandoSucursal:true
      })

      await sock.sendMessage(jid,{
        text:
          `📋 RECETARIO ${etiquetaArea(area)}\n\n`+
          `¿De qué sucursal?\n\n`+
          `• Coyoacán\n`+
          `• Bucareli`
      })

      return true
    }

    guardarContexto(jid,{
      accion:'recetario',
      sucursal:suc,
      esperandoArea:true
    })

    await sock.sendMessage(jid,{
      text:
        '📂 ¿Qué recetario?\n\n'+
        '• Cocina\n'+
        '• Barra'
    })

    return true
  }

  /* =====================================================
     RECETA NORMAL / PLATILLO
  ===================================================== */

  const matchReceta=
    t.match(
      /^receta(?:s)?\s+(?:de\s+)?(.+?)(?:\s+(costos?|procesos?|descriptivo|descritivo|descripcion|descripción))?$/i
    )

  if(matchReceta){

    const platillo=
      limpiarTexto(matchReceta[1])

    const tipoTexto=
      matchReceta[2]||''

    const suc=
      sucGrupo

    const area=''

    const{
      menu,
      platillos
    }=await cargarRecetas()

    const coincidencias=
      buscarPlatillos(
        platillo,
        menu,
        platillos,
        suc,
        area
      )

    if(!coincidencias.length){

      await sock.sendMessage(jid,{
        text:
          `❌ No encontré "${platillo}" `+
          `en las recetas`+
          `${suc?` de ${etiquetaSucursal(suc)}`:''}.`
      })

      return true
    }

    /*
     * Si hay varias coincidencias y no se indicó
     * tipo de reporte, mostramos la lista.
     */

    if(
      coincidencias.length>1&&
      !tipoTexto
    ){

      return await mostrarCoincidencias({
        sock,
        jid,
        coincidencias,
        sucursal:suc,
        area
      })
    }

    /*
     * Si hay varias y ya se indicó tipo,
     * primero hay que elegir platillo.
     */

    if(
      coincidencias.length>1&&
      tipoTexto
    ){

      const tipo=
        tipoReporte(tipoTexto)

      await sock.sendMessage(jid,{
        text:
          'Encontré varios platillos:\n\n'+
          coincidencias
            .slice(0,15)
            .map(
              (x,i)=>`${i+1}. ${x}`
            )
            .join('\n')+
          '\n\nEscribe el número o el nombre del que buscas.'
      })

      guardarContexto(jid,{

        coincidencias:
          coincidencias.slice(0,15),

        sucursal:suc,

        area,

        tipoSolicitado:
          tipo||'',

        esperandoPlatillo:true
      })

      return true
    }

    const elegido=
      coincidencias[0]

    if(tipoTexto){

      const tipo=
        tipoReporte(tipoTexto)

      if(tipo==='costos'){

        return await reporteCostos({
          sock,
          jid,
          platillo:elegido,
          sucursal:suc,
          area
        })
      }

      if(tipo==='procesos'){

        return await reporteProcesos({
          sock,
          jid,
          platillo:elegido,
          sucursal:suc,
          area
        })
      }

      return await reporteDescriptivo({
        sock,
        jid,
        platillo:elegido,
        sucursal:suc,
        area
      })
    }

    guardarContexto(jid,{

      platillo:elegido,

      sucursal:suc,

      area,

      esperandoTipo:true
    })

    await sock.sendMessage(jid,{
      text:
        `🍽️ ${elegido}\n\n`+
        `¿Qué reporte quieres?\n\n`+
        `1. Costos\n`+
        `2. Descriptivo\n`+
        `3. Procesos`
    })

    return true
  }

  /* =====================================================
     COTIZAR
  ===================================================== */

  if(
    /^(cotizar|cotizacion|cotización)\b/i.test(t)
  ){

    await sock.sendMessage(jid,{
      text:
        '💰 Cotización de menú\n\n'+
        '⏳ La cotización la conectamos después de validar costos.'
    })

    return true
  }

  return false
}

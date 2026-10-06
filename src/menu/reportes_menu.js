import {getRows}from'../sheets.js'
import {SPREADSHEET_RECETAS_ID}from'../config.js'

const CONTEXTO_MENU=new Map()

const limpiarTexto=x=>(x||'').toString().trim().replace(/\s+/g,' ')

const normaliza=x=>limpiarTexto(x).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase()

const dinero=x=>{
  const n=parseFloat((x||'').toString().replace(/[$,%\s,]/g,''))
  return Number.isFinite(n)?n:0
}

const moneda=x=>`$${Number(x||0).toLocaleString('es-MX',{minimumFractionDigits:2,maximumFractionDigits:2})}`

const esComandoMenu=t=>{
  const x=normaliza(t)
  return /^(receta|recetas|cotizar|cotizacion)\b/i.test(x)
}

const guardarContexto=(jid,datos)=>{
  CONTEXTO_MENU.set(jid,{...datos,expira:Date.now()+60000})
}

const obtenerContexto=jid=>{
  const x=CONTEXTO_MENU.get(jid)
  if(!x)return null
  if(Date.now()>x.expira){
    CONTEXTO_MENU.delete(jid)
    return null
  }
  return x
}

const borrarContexto=jid=>CONTEXTO_MENU.delete(jid)

async function cargarRecetas(){
  const [catalogo,subrecetas,rendimientos,platillos,menu]=await Promise.all([
    getRows('CATALOGO!A:G',SPREADSHEET_RECETAS_ID),
    getRows('2_SUBRECETAS!A:F',SPREADSHEET_RECETAS_ID),
    getRows('3_RENDIMIENTOS!A:G',SPREADSHEET_RECETAS_ID),
    getRows('4_PLATILLOS!A:G',SPREADSHEET_RECETAS_ID),
    getRows('5_MENU_COSTOS!A:P',SPREADSHEET_RECETAS_ID)
  ])

  return{catalogo,subrecetas,rendimientos,platillos,menu}
}

function buscarPlatillo(nombre,rows){
  const n=normaliza(nombre)

  return rows.slice(1).find(r=>normaliza(r[0])===n)||
         rows.slice(1).find(r=>normaliza(r[0]).includes(n))||
         rows.slice(1).find(r=>n.includes(normaliza(r[0])))
}

function mapaCatalogo(rows){
  const map=new Map()

  for(const r of rows.slice(1)){
    const nombre=normaliza(r[0])
    if(!nombre)continue

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
    const suc=normaliza(r[3])

    if(!sub)continue

    if(!map.has(sub))map.set(sub,[])
    map.get(sub).push({
      subreceta:r[0]||'',
      rendimiento:parseFloat(r[1])||0,
      unidad:r[2]||'',
      sucursal:r[3]||'',
      tipo:r[4]||'',
      costoTotal:dinero(r[5]),
      costoUnidad:dinero(r[6]),
      suc
    })
  }

  return map
}

function obtenerRendimientoSubreceta(nombre,map,origen){
  const lista=map.get(normaliza(nombre))||[]
  if(!lista.length)return null

  const o=normaliza(origen)

  return lista.find(x=>o&&normaliza(x.sucursal)===o)||
         lista.find(x=>o&&normaliza(x.sucursal).includes(o))||
         lista[0]
}

function obtenerIngredientesPlatillo(nombre,rows){
  const n=normaliza(nombre)

  return rows.slice(1).filter(r=>normaliza(r[0])===n)
}

function esSubreceta(nombre,subrows){
  const n=normaliza(nombre)
  return subrows.slice(1).some(r=>normaliza(r[0])===n)
}

function buscarIngredienteCatalogo(nombre,map){
  return map.get(normaliza(nombre))||null
}

function agregarAlergeno(lista,nombre,alergico){
  if(!alergico)return

  const clave=normaliza(nombre)

  if(!lista.some(x=>normaliza(x.nombre)===clave)){
    lista.push({nombre,alergico})
  }
}

function limpiarNombreSubreceta(nombre){
  return (nombre||'').toString().replace(/^sub\s+/i,'').trim()
}

function obtenerAlergenosSubreceta(nombre,subrows,catalogoMap,visitados=new Set()){
  const clave=normaliza(nombre)

  if(visitados.has(clave))return[]
  visitados.add(clave)

  const salida=[]
  const filas=subrows.slice(1).filter(r=>normaliza(r[0])===clave)

  for(const r of filas){

    const ingrediente=r[1]||''
    if(!ingrediente)continue

    const cat=buscarIngredienteCatalogo(ingrediente,catalogoMap)

    if(cat?.alergico){
      agregarAlergeno(salida,limpiarNombreSubreceta(ingrediente),cat.alergico)
      continue
    }

    if(esSubreceta(ingrediente,subrows)){
      const sub=obtenerAlergenosSubreceta(ingrediente,subrows,catalogoMap,visitados)
      for(const a of sub){
        agregarAlergeno(salida,a.nombre,a.alergico)
      }
    }
  }

  return salida
}

async function reporteCostos({sock,jid,platillo}){
  try{

    const{
      catalogo,
      subrecetas,
      rendimientos,
      platillos
    }=await cargarRecetas()

    const filaPlatillo=buscarPlatillo(platillo,platillos)

    if(!filaPlatillo){
      await sock.sendMessage(jid,{
        text:`❌ No encontré el platillo "${platillo}" en el catálogo de recetas.`
      })
      return true
    }

    const nombrePlatillo=filaPlatillo[0]
    const filas=obtenerIngredientesPlatillo(nombrePlatillo,platillos)

    if(!filas.length){
      await sock.sendMessage(jid,{
        text:`❌ El platillo "${nombrePlatillo}" no tiene ingredientes registrados.`
      })
      return true
    }

    const catalogoMap=mapaCatalogo(catalogo)
    const rendMap=mapaRendimientos(rendimientos)

    let total=0
    const ingredientes=[]
    const alergenos=[]

    for(const r of filas){

      const ingrediente=r[1]||''
      const cantidad=parseFloat(r[2])||0
      const unidad=r[3]||''
      const origen=r[4]||''
      let costoUnidad=dinero(r[5])
      let costoTotal=dinero(r[6])

      const sub=esSubreceta(ingrediente,subrecetas)

      if(sub){

        const rend=obtenerRendimientoSubreceta(ingrediente,rendMap,origen)

        if(rend){
          costoUnidad=rend.costoUnidad
          costoTotal=cantidad*costoUnidad
        }

        const subAlergenos=obtenerAlergenosSubreceta(
          ingrediente,
          subrecetas,
          catalogoMap
        )

        for(const a of subAlergenos){
          agregarAlergeno(alergenos,limpiarNombreSubreceta(ingrediente),a.alergico)
        }

      }else{

        const cat=buscarIngredienteCatalogo(ingrediente,catalogoMap)

        if(cat?.alergico){
          agregarAlergeno(alergenos,ingrediente,cat.alergico)
        }

        if(!costoUnidad&&cat){
          costoUnidad=cat.costo
          costoTotal=cantidad*costoUnidad
        }
      }

      if(!costoTotal&&cantidad&&costoUnidad){
        costoTotal=cantidad*costoUnidad
      }

      total+=costoTotal

      ingredientes.push({
        nombre:sub?limpiarNombreSubreceta(ingrediente):ingrediente,
        cantidad,
        unidad,
        costo:costoTotal
      })
    }

    let texto=`🍽️ ${nombrePlatillo.toUpperCase()}\n\nIngredientes:\n`

    for(const x of ingredientes){
      texto+=`• ${x.nombre} ........ ${x.cantidad} ${x.unidad} ........ ${moneda(x.costo)}\n`
    }

    texto+=`\n💰 COSTO DEL PLATILLO: ${moneda(total)}`

    if(alergenos.length){
      texto+=`\n\n────────────────────\n⚠️ ALÉRGENOS:\n`

      for(const a of alergenos){
        texto+=`• ${a.nombre} — posible alérgeno: ${a.alergico}\n`
      }
    }

    await sock.sendMessage(jid,{text})

    return true

  }catch(e){
    console.error('Error reporte costos menú:',e)

    await sock.sendMessage(jid,{
      text:'❌ Ocurrió un error consultando el costo de la receta.'
    })

    return true
  }
}

export async function handleReportesMenu({sock,jid,m,texto,tipoGrupo}){

  const t=limpiarTexto(texto)
  const n=normaliza(t)

  const contexto=obtenerContexto(jid)

  if(contexto&&/^(costos|costo|procesos|proceso|descriptivo|descritivo|descripcion)$/i.test(n)){

    const tipo=
      /^costo/i.test(n)?'costos':
      /^proceso/i.test(n)?'procesos':
      'descriptivo'

    borrarContexto(jid)

    if(tipo==='costos'){
      return await reporteCostos({
        sock,
        jid,
        platillo:contexto.platillo
      })
    }

    await sock.sendMessage(jid,{
      text:`📋 Receta: ${contexto.platillo}\nTipo solicitado: ${tipo}\n\n⏳ Este reporte lo conectamos en el siguiente paso.`
    })

    return true
  }

  if(!esComandoMenu(t))return false

  const matchReceta=t.match(/^receta(?:s)?\s+de\s+(.+?)(?:\s+(costos?|procesos?|descriptivo|descritivo|descripcion))?$/i)

  if(matchReceta){

    const platillo=limpiarTexto(matchReceta[1])
    const tipoTexto=matchReceta[2]||''

    if(tipoTexto){

      const tipo=
        /^costo/i.test(tipoTexto)?'costos':
        /^proceso/i.test(tipoTexto)?'procesos':
        'descriptivo'

      if(tipo==='costos'){
        return await reporteCostos({
          sock,
          jid,
          platillo
        })
      }

      await sock.sendMessage(jid,{
        text:`📋 Receta: ${platillo}\nTipo solicitado: ${tipo}\n\n⏳ Este reporte lo conectamos en el siguiente paso.`
      })

      return true
    }

    guardarContexto(jid,{platillo})

    await sock.sendMessage(jid,{
      text:'¿Qué receta buscas?\n\n• Costos\n• Descriptivo\n• Procesos'
    })

    return true
  }

  if(/^(cotizar|cotizacion)\b/i.test(t)){

    await sock.sendMessage(jid,{
      text:'💰 Cotización de menú\n\n⏳ La cotización la conectamos después de validar costos.'
    })

    return true
  }

  return false
}

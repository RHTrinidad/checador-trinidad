/*
  ============================================================
  FLUJO DE LISTA NUMÉRICA
  ============================================================
  Sistema general para procesos que necesitan:
  1. Mostrar una lista numerada.
  2. Esperar una respuesta numérica.
  3. Recuperar el contexto de la solicitud.
  4. Interpretar la opción elegida.
  5. Continuar automáticamente el proceso.
  El contexto se guarda por grupo + usuario.
  ============================================================
*/
const listasPendientes=new Map()
function obtenerClave(jid,usuario){
  return `${jid}|${usuario||'desconocido'}`
}
export function iniciarLista({
  jid,
  usuario,
  contexto={},
  opciones={},
  mensaje=''
}){
  const clave=obtenerClave(jid,usuario)
  const opcionesNormalizadas={}
  for(const [numero,opcion] of Object.entries(opciones)){
    if(opcion===null||opcion===undefined)continue
    if(typeof opcion==='object'&&!Array.isArray(opcion)){
      opcionesNormalizadas[String(numero)]={
        texto:opcion.texto||String(numero),
        valor:opcion.valor!==undefined?opcion.valor:opcion.texto
      }
    }else{
      opcionesNormalizadas[String(numero)]={
        texto:String(opcion),
        valor:opcion
      }
    }
  }
  listasPendientes.set(clave,{
    contexto,
    opciones:opcionesNormalizadas,
    creado:Date.now()
  })
  return construirMensajeLista(opcionesNormalizadas,mensaje)
}
function construirMensajeLista(opciones,mensaje=''){
  const lineas=[]
  if(mensaje){
    lineas.push(mensaje)
    lineas.push('')
  }
  for(const numero of Object.keys(opciones)){
    lineas.push(`${numero}️⃣ ${opciones[numero].texto}`)
  }
  lineas.push('')
  lineas.push('Responde con el número de la opción.')
  return lineas.join('\n')
}
export function resolverLista({
  jid,
  usuario,
  texto
}){
  const clave=obtenerClave(jid,usuario)
  const pendiente=listasPendientes.get(clave)
  if(!pendiente)return null
  const respuesta=(texto||'').toString().trim()
  if(!/^\d+$/.test(respuesta)){
    return{
      estado:'INVALIDA',
      mensaje:'⚠️ Responde únicamente con el número de la opción.'
    }
  }
  const opcion=pendiente.opciones[respuesta]
  if(!opcion){
    return{
      estado:'INVALIDA',
      mensaje:'⚠️ Esa opción no existe. Responde con uno de los números de la lista.'
    }
  }
  listasPendientes.delete(clave)
  return{
    estado:'RESUELTA',
    numero:respuesta,
    valor:opcion.valor,
    texto:opcion.texto,
    contexto:pendiente.contexto
  }
}
export function hayListaPendiente({
  jid,
  usuario
}){
  const clave=obtenerClave(jid,usuario)
  return listasPendientes.has(clave)
}
export function obtenerListaPendiente({
  jid,
  usuario
}){
  const clave=obtenerClave(jid,usuario)
  return listasPendientes.get(clave)||null
}
export function cancelarLista({
  jid,
  usuario
}){
  const clave=obtenerClave(jid,usuario)
  return listasPendientes.delete(clave)
}
const TIEMPO_MAXIMO=15*60*1000
setInterval(()=>{
  const ahora=Date.now()
  for(const [clave,lista] of listasPendientes){
    if(ahora-lista.creado>TIEMPO_MAXIMO){
      listasPendientes.delete(clave)
    }
  }
},60*1000)

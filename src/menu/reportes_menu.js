const CONTEXTO_MENU=new Map()

const limpiarTexto=x=>(x||'').toString().trim().replace(/\s+/g,' ')

const normaliza=x=>limpiarTexto(x).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase()

const esComandoMenu=t=>{
  const x=normaliza(t)
  return /^(receta|recetas|cotizar|cotizacion|cotización)\b/i.test(x)
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

export async function handleReportesMenu({sock,jid,m,texto,tipoGrupo}){

  const t=limpiarTexto(texto)
  const n=normaliza(t)

  /*
    RESPUESTA PENDIENTE DE TIPO DE RECETA

    Ejemplo:
    Usuario: receta de chilaquiles verdes
    Bot: ¿Qué receta buscas?
    Usuario: costos

    Se conserva el contexto durante 1 minuto.
  */

  const contexto=obtenerContexto(jid)

  if(contexto&&/^(costos|costo|procesos|proceso|descriptivo|descripcion|descripción)$/i.test(n)){

    const tipo=
      /^costo/i.test(n)?'costos':
      /^proceso/i.test(n)?'procesos':
      'descriptivo'

    borrarContexto(jid)

    await sock.sendMessage(jid,{
      text:`📋 Receta: ${contexto.platillo}\nTipo solicitado: ${tipo}\n\n⏳ Módulo de reportes de menú en preparación.`
    })

    return true
  }

  if(!esComandoMenu(t))return false

  /*
    RECETA
  */

  const matchReceta=t.match(/^receta(?:s)?\s+de\s+(.+?)(?:\s+(costos?|procesos?|descriptivo|descripcion|descripción))?$/i)

  if(matchReceta){

    const platillo=limpiarTexto(matchReceta[1])
    const tipoTexto=matchReceta[2]||''

    if(tipoTexto){

      const tipo=
        /^costo/i.test(tipoTexto)?'costos':
        /^proceso/i.test(tipoTexto)?'procesos':
        'descriptivo'

      await sock.sendMessage(jid,{
        text:`📋 Receta: ${platillo}\nTipo solicitado: ${tipo}\n\n⏳ Módulo de reportes de menú en preparación.`
      })

      return true
    }

    guardarContexto(jid,{platillo})

    await sock.sendMessage(jid,{
      text:`¿Qué receta buscas?\n\n• Costos\n• Descriptivo\n• Procesos`
    })

    return true
  }

  /*
    COTIZACIONES
  */

  if(/^(cotizar|cotizacion|cotización)\b/i.test(t)){

    await sock.sendMessage(jid,{
      text:'💰 Cotización de menú\n\n⏳ Módulo de cotizaciones en preparación.'
    })

    return true
  }

  return false
}

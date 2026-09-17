module.exports = {
    // Coordenadas geográficas oficiales de las unidades
    SUCURSALES: {
        "Trinidad Coyoacan": { lat: 19.3503, lng: -99.1623 },
        "Trinidad Bucareli": { lat: 19.4304, lng: -99.1501 }
    },
    
    // Reglas de negocio del Checador
    REGLAS: {
        DISTANCIA_MAX_ENTRADA_M: 150,     // Margen de aceptación entrada
        DISTANCIA_MAX_SALIDA_M: 100,      // Margen ideal para salida sin alerta
        TOLERANCIA_RETARDO_MIN: 15,       // Minutos de tolerancia para retardos
        TIEMPO_ALERTA_FALTA_MIN: 20,      // Minutos para avisar que no ha llegado
        JORNADA_REGULAR_SEMANAL_HRS: 48,  // Jornada base semanal
        JORNADA_STANDALONE_DIARIA_HRS: 8  // Si no hay calendario asignado
    }
};

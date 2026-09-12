module.exports = {
  sucursales: {
    bucareli: { 
      lat: 19.4314119, 
      lng: -99.1512074, 
      nombre: 'BUCARELI 66',
      radioEntrada: 150,
      radioSalidaAviso: 100,
      radioSalidaLimite: 150
    },
    coyoacan: { 
      lat: 19.3526359, 
      lng: -99.1808071, 
      nombre: 'COYOACAN MERCADO 86',
      radioEntrada: 150,
      radioSalidaAviso: 100,
      radioSalidaLimite: 150
    }
  },
  administradores: ['120363412984528459@g.us'],
  jornada: {
    inicio: 5, // 5am
    descanso: 1 // 1 = Lunes
  },
  mensajes: {
    fueraRangoEntrada: '❌ La entrada se registra en la sucursal, intenta de nuevo en la unidad',
    fueraRangoSalida: '❌ La salida se registra en la sucursal, intenta de nuevo en la unidad',
    avisoSalida: 'La salida es a 100 metros de la sucursal',
    graciasEntrada: '✅ Gracias, buen turno!',
    graciasSalida: '✅ Gracias!'
  }
};

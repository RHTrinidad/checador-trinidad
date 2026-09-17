const { createBot, createProvider, createFlow, addKeyword } = require('@bot-whatsapp/bot');
const BaileysProvider = require('@bot-whatsapp/provider/baileys');
const cron = require('node-cron');

const { SUCURSALES, REGLAS } = require('./config');
const { guardarAsistencia, obtenerHorarioEmpleado, generarSiguienteSemana, obtenerResumenSemanal } = require('./sheets');

// --- FÓRMULA DE DISTANCIA (HAVERSINE) ---
function calcularDistancia(lat1, lon1, lat2, lon2) {
    const R = 6371000; // Radio de la Tierra en metros
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLon = (lon2 - lon1) * Math.PI / 180;
    const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
              Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * 
              Math.sin(dLon/2) * Math.sin(dLon/2);
    return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// --- DETERMINAR FECHA LABORAL CON CORTE 4:00 AM ---
function obtenerFechaLaboralActual() {
    const ahora = new Date();
    if (ahora.getHours() < 4) {
        ahora.setDate(ahora.getDate() - 1); // Pertenece a la jornada del día de ayer
    }
    return ahora.toISOString().split('T')[0];
}

// --- FLUJO ENTRADA (Soporta variantes como: entrada, entre, entrar, etrada, endrada) ---
const flujoEntrada = addKeyword([/^(entrada|entre|entrar|etrada|endrada)\$/i], { regex: true })
    .addAnswer('📍 Por favor, comparte tu **Ubicación en tiempo real** o actual para registrar tu entrada.')
    .addAction({ capture: true }, async (ctx, { flowDynamic, fallBack }) => {
        if (!ctx.message?.locationMessage) {
            return fallBack('⚠️ Envío inválido. Debes presionar el clip en WhatsApp y seleccionar "Ubicación".');
        }

        const lat = ctx.message.locationMessage.degreesLatitude;
        const lng = ctx.message.locationMessage.degreesLongitude;
        const telefono = ctx.from;
        const fechaLaboral = obtenerFechaLaboralActual();

        // Consultar los datos del empleado en la base de datos (sheets.js)
        const empleado = await obtenerHorarioEmpleado(telefono, fechaLaboral);
        if (!empleado) return await flowDynamic('❌ No estás registrado en el sistema. Solicita tu alta con el administrador.');

        const sucursal = SUCURSALES[empleado.sucursalAsignada];
        if (!sucursal) return await flowDynamic('❌ Sucursal no identificada en el sistema.');

        const distancia = calcularDistancia(lat, lng, sucursal.lat, sucursal.lng);

        if (distancia > REGLAS.DISTANCIA_MAX_ENTRADA_M) {
            return await flowDynamic('❌ *La entrada debe registrarse en la unidad.* Vuelve a intentarlo dentro de la sucursal.');
        }

        // Validar Retardos si existe horario planificado
        let estatus = "A TIEMPO";
        if (empleado.horaEsperada && empleado.horaEsperada !== 'DESCANSO') {
            const [espHoras, espMinutos] = empleado.horaEsperada.split(':').map(Number);
            const ahora = new Date();
            const limite = new Date();
            limite.setHours(espHoras, espMinutos + REGLAS.TOLERANCIA_RETARDO_MIN, 0);
            if (ahora > limite) estatus = "RETARDO";
        }

        await guardarAsistencia(telefono, fechaLaboral, { horaEntrada: new Date().toLocaleTimeString(), estatusEntrada: estatus });
        await flowDynamic(`✅ ¡Buen turno, *${empleado.nombre}*! Entrada registrada (${estatus}).`);
    });

// --- FLUJO SALIDA (Soporta variantes como: salida, salir, me voy, salda, salid) ---
const flujoSalida = addKeyword([/^(salida|salir|me voy|salda|salid)\$/i], { regex: true })
    .addAnswer('📍 Por favor, comparte tu **Ubicación** para validar tu registro de salida.')
    .addAction({ capture: true }, async (ctx, { flowDynamic, fallBack, provider }) => {
        if (!ctx.message?.locationMessage) {
            return fallBack('⚠️ Por favor comparte tu ubicación de WhatsApp.');
        }

        const lat = ctx.message.locationMessage.degreesLatitude;
        const lng = ctx.message.locationMessage.degreesLongitude;
        const telefono = ctx.from;
        const fechaLaboral = obtenerFechaLaboralActual();

        const empleado = await obtenerHorarioEmpleado(telefono, fechaLaboral);
        const sucursal = SUCURSALES[empleado.sucursalAsignada];
        const distancia = calcularDistancia(lat, lng, sucursal.lat, sucursal.lng);
        const horaActual = new Date().toLocaleTimeString();

        if (distancia > REGLAS.DISTANCIA_MAX_SALIDA_M) {
            // Se registra el incidente pero SÍ se guarda la marca temporal
            await guardarAsistencia(telefono, fechaLaboral, { horaPrimerIntentoSalida: horaActual, fueraRango: true, distancia: distancia });
            await flowDynamic('⚠️ *La salida debe de registrarse a menos de 100 metros.* Tu marca de salida fue pre-registrada, informando la incidencia al grupo de reportes.');
            
            // Envío automático al grupo de reportes
            const grupoReportesId = process.env.GRUPO_REPORTES_ID;
            await provider.sendMessage(grupoReportesId, `⚠️ *ALERTA DE SALIDA FUERA DE RANGO*\n👤 Empleado: ${empleado.nombre}\n🏢 Sucursal: ${empleado.sucursalAsignada}\n📏 Distancia: ${Math.round(distancia)} metros\n⏰ Hora: ${horaActual}`);
        } else {
            // Confirmación limpia o confirmación tras haber regresado a la sucursal
            await guardarAsistencia(telefono, fechaLaboral, { horaSalidaDefinitiva: horaActual, fueraRango: false });
            await flowDynamic(`👋 Hasta luego, *${empleado.nombre}*. Salida registrada correctamente.`);
        }
    });

// --- CRON JOBS (PLANIFICACIONES EN RAILWAY) ---

// 1. Automatización Opción A: Generar semanas de manera automática (Sábados 11:00 PM)
cron.schedule('0 23 * * 6', async () => {
    console.log('📅 Generando calendario automático para la siguiente semana...');
    await generarSiguienteSemana();
});

// 2. Reporte consolidado de Incidencias al grupo de reportes (Lunes 7:00 AM)
cron.schedule('0 7 * * 1', async (provider) => {
    console.log('📊 Enviando reporte consolidado semanal...');
    const reporte = await obtenerResumenSemanal();
    const grupoReportesId = process.env.GRUPO_REPORTES_ID;
    await provider.sendMessage(grupoReportesId, reporte);
});

// 📌 PEGA ESTO JUSTO AQUÍ (ANTES DEL MAIN):
const flujoRegistrar = addKeyword(['/registrar'])
    .addAction(async (ctx, { flowDynamic }) => {
        const listaAdministradores = await obtenerAdministradores(); 

        if (!listaAdministradores.includes(ctx.from)) {
            return await flowDynamic('❌ No tienes permisos de administrador en el sistema para realizar registros.');
        }

        const partes = ctx.body.split(' ');
        if (partes.length < 5) {
            return await flowDynamic('⚠️ Formato incorrecto. Usa: `/registrar [Teléfono] [Nombre] [Sucursal (Coyoacan o Bucareli)]`');
        }

        const telefonoEmpleado = partes[1];
        const sucursalTexto = ctx.body.toLowerCase().includes('bucareli') ? 'Trinidad Bucareli' : 'Trinidad Coyoacan';
        
        const nombreEmpleado = ctx.body
            .replace(`/registrar ${telefonoEmpleado} `, '')
            .replace(/trinidad coyoacan/i, '')
            .replace(/trinidad bucareli/i, '')
            .replace(/coyoacan/i, '')
            .replace(/bucareli/i, '')
            .trim();

        const coords = SUCURSALES[sucursalTexto];

        const { google } = require('googleapis');
        const auth = new google.auth.GoogleAuth({
            credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON || '{}'),
            scopes: ['https://googleapis.com'],
        });
        const sheets = google.sheets({ version: 'v4', auth });

        await sheets.spreadsheets.values.append({
            spreadsheetId: process.env.SPREADSHEET_ID,
            range: 'Empleados!A2',
            valueInputOption: 'USER_ENTERED',
            resource: { 
                values: [[telefonoEmpleado, nombreEmpleado, sucursalTexto, coords.lat, coords.lng, 'Lunes']] 
            },
        });

        await flowDynamic(`👤 *¡Empleado Registrado!*\n• Nombre: ${nombreEmpleado}\n• Tel: ${telefonoEmpleado}\n• Unidad: ${sucursalTexto}\n\nPermisos validados correctamente desde la nube.`);
    });


// 🚀 AHORA SÍ, TU FUNCIÓN MAIN CON EL SERVIDOR WEB QR:
const main = async () => {
    // 📌 Configuración limpia del proveedor Baileys
    const adapterProvider = createProvider(BaileysProvider);

    createBot({
        flow: createFlow([flujoEntrada, flujoSalida, flujoRegistrar]),
        provider: adapterProvider,
        database: null,
    });

    // 🔥 SOLUCIÓN DIRECTA: Escuchar el evento del QR y dibujarlo en la pantalla de Railway
    adapterProvider.on('qr', (qr) => {
        console.log('📢 ¡NUEVO CÓDIGO QR GENERADO! ESCANEA AQUÍ ABAJO:');
        require('qrcode-terminal').generate(qr, { small: true });
    });
};

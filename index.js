const { createBot, createProvider, createFlow, addKeyword } = require('@bot-whatsapp/bot');
const BaileysProvider = require('@bot-whatsapp/provider/baileys');
const MockAdapter = require('@bot-whatsapp/database/mock');
const cron = require('node-cron');

const { SUCURSALES, REGLAS } = require('./config');
const { guardarAsistencia, obtenerHorarioEmpleado, generarSiguienteSemana, obtenerResumenSemanal } = require('./sheets');

// --- FÓRMULA DE DISTANCIA (HAVERSINE) ---
function calcularDistancia(lat1, lon1, lat2, lon2) {
    const R = 6371000;
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLon = (lon2 - lon1) * Math.PI / 180;
    const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
              Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
              Math.sin(dLon/2) * Math.sin(dLon/2);
    return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function obtenerFechaLaboralActual() {
    const ahora = new Date();
    if (ahora.getHours() < 4) {
        ahora.setDate(ahora.getDate() - 1);
    }
    return ahora.toISOString().split('T')[0];
}

const flujoEntrada = addKeyword([/^(entrada|entre|entrar|etrada|endrada)$/i], { regex: true })
   .addAnswer('📍 Por favor, comparte tu **Ubicación en tiempo real** o actual para registrar tu entrada.')
   .addAction({ capture: true }, async (ctx, { flowDynamic, fallBack }) => {
        if (!ctx.message?.locationMessage) {
            return fallBack('⚠️ Envío inválido. Debes presionar el clip en WhatsApp y seleccionar "Ubicación".');
        }
        const lat = ctx.message.locationMessage.degreesLatitude;
        const lng = ctx.message.locationMessage.degreesLongitude;
        const telefono = ctx.from;
        const fechaLaboral = obtenerFechaLaboralActual();
        const empleado = await obtenerHorarioEmpleado(telefono, fechaLaboral);
        if (!empleado) return await flowDynamic('❌ No estás registrado en el sistema.');
        const sucursal = SUCURSALES[empleado.sucursalAsignada];
        if (!sucursal) return await flowDynamic('❌ Sucursal no identificada.');
        const distancia = calcularDistancia(lat, lng, sucursal.lat, sucursal.lng);
        if (distancia > REGLAS.DISTANCIA_MAX_ENTRADA_M) {
            return await flowDynamic('❌ *La entrada debe registrarse en la unidad.* Vuelve a intentarlo dentro de la sucursal.');
        }
        let estatus = "A TIEMPO";
        if (empleado.horaEsperada && empleado.horaEsperada!== 'DESCANSO') {
            const [espHoras, espMinutos] = empleado.horaEsperada.split(':').map(Number);
            const ahora = new Date();
            const limite = new Date();
            limite.setHours(espHoras, espMinutos + REGLAS.TOLERANCIA_RETARDO_MIN, 0);
            if (ahora > limite) estatus = "RETARDO";
        }
        await guardarAsistencia(telefono, fechaLaboral, { horaEntrada: new Date().toLocaleTimeString(), estatusEntrada: estatus });
        await flowDynamic(`✅ ¡Buen turno, *${empleado.nombre}*! Entrada registrada (${estatus}).`);
    });

const flujoSalida = addKeyword([/^(salida|salir|me voy|salda|salid)$/i], { regex: true })
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
        if(!empleado) return await flowDynamic('❌ No estás registrado.');
        const sucursal = SUCURSALES[empleado.sucursalAsignada];
        const distancia = calcularDistancia(lat, lng, sucursal.lat, sucursal.lng);
        const horaActual = new Date().toLocaleTimeString();
        if (distancia > REGLAS.DISTANCIA_MAX_SALIDA_M) {
            await guardarAsistencia(telefono, fechaLaboral, { horaPrimerIntentoSalida: horaActual, fueraRango: true, distancia: distancia });
            await flowDynamic('⚠️ *La salida debe de registrarse a menos de 100 metros.* Tu marca fue pre-registrada.');
            try {
                const grupoReportesId = process.env.GRUPO_REPORTES_ID;
                if(grupoReportesId) await provider.sendMessage(grupoReportesId, `⚠️ *ALERTA FUERA DE RANGO*\n👤 ${empleado.nombre}\n🏢 ${empleado.sucursalAsignada}\n📏 ${Math.round(distancia)} m\n⏰ ${horaActual}`);
            } catch(e){ console.log('Error enviando alerta', e.message) }
        } else {
            await guardarAsistencia(telefono, fechaLaboral, { horaSalidaDefinitiva: horaActual, fueraRango: false });
            await flowDynamic(`👋 Hasta luego, *${empleado.nombre}*. Salida registrada correctamente.`);
        }
    });

const flujoRegistrar = addKeyword(['/registrar'])
   .addAction(async (ctx, { flowDynamic }) => {
        // Si no tienes obtenerAdministradores, usa ADMIN_NUMBER en env
        const admins = (process.env.ADMIN_NUMBERS || '').split(',').map(s=>s.trim());
        if (admins.length > 0 &&!admins.includes(ctx.from)) {
             return await flowDynamic('❌ No tienes permisos de administrador.');
        }
        const partes = ctx.body.split(' ');
        if (partes.length < 3) {
            return await flowDynamic('⚠️ Formato: `/registrar 5255... Nombre Apellido Coyoacan` o `Bucareli`');
        }
        const telefonoEmpleado = partes[1];
        const sucursalTexto = ctx.body.toLowerCase().includes('bucareli')? 'Trinidad Bucareli' : 'Trinidad Coyoacan';
        const nombreEmpleado = ctx.body.replace(`/registrar ${telefonoEmpleado} `, '').replace(/trinidad coyoacan/i, '').replace(/trinidad bucareli/i, '').replace(/coyoacan/i, '').replace(/bucareli/i, '').trim();

        try {
            const { google } = require('googleapis');
            let creds = process.env.GOOGLE_CREDS;
            try { creds = JSON.parse(creds); } catch(e){}
            if(typeof creds === 'string'){ try{ creds = JSON.parse(creds); }catch(e){} }
            const auth = new google.auth.GoogleAuth({
                credentials: creds,
                scopes: ['https://www.googleapis.com/auth/spreadsheets'],
            });
            const sheets = google.sheets({ version: 'v4', auth });
            await sheets.spreadsheets.values.append({
                spreadsheetId: process.env.SHEET_ID || process.env.SPREADSHEET_ID,
                range: 'Empleados!A2',
                valueInputOption: 'USER_ENTERED',
                resource: { values: [[telefonoEmpleado, nombreEmpleado, sucursalTexto]] },
            });
            await flowDynamic(`👤 *¡Empleado Registrado!*\n• Nombre: ${nombreEmpleado}\n• Tel: ${telefonoEmpleado}\n• Unidad: ${sucursalTexto}`);
        } catch(e){
            console.log(e);
            await flowDynamic(`❌ Error registrando: ${e.message}`);
        }
    });

cron.schedule('0 23 * * 6', async () => {
    console.log('📅 Generando calendario automático...');
    await generarSiguienteSemana();
});

// --- MAIN ---
const main = async () => {
    const adapterDB = new MockAdapter();
    const adapterFlow = createFlow([flujoEntrada, flujoSalida, flujoRegistrar]);
    const adapterProvider = createProvider(BaileysProvider);

        try {
        const QRPortalWeb = require('@bot-whatsapp/portal');
        QRPortalWeb({ port: process.env.PORT || 8080 });
    } catch(e){}
    
    createBot({
        flow: adapterFlow,
        provider: adapterProvider,
        database: adapterDB,
    });

    console.log('BOT LISTO - Esperando QR en logs...');
}

main();

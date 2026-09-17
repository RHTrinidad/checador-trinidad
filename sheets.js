const { google } = require('googleapis');

// Configuración de autenticación con Google API a través de variables de entorno de Railway
// Configuración de autenticación mejorada con manejo de errores integrados
const jsonCredenciales = process.env.GOOGLE_SERVICE_ACCOUNT_JSON || process.env.GOOGLE_CREDS || "{}";

const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(jsonCredenciales.trim()),
    scopes: ['https://googleapis.com'],
});
});

const sheets = google.sheets({ version: 'v4', auth });
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

/**
 * Busca a un empleado por teléfono y obtiene su horario programado para una fecha específica.
 */
async function obtenerHorarioEmpleado(telefono, fechaLaboral) {
    try {
        // 1. Primero buscamos los datos generales del empleado (Nombre y Sucursal)
        const vistaEmpleados = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Empleados!A2:C', // Columnas: Teléfono, Nombre, Sucursal Asignada
        });
        
        const filasEmpleados = vistaEmpleados.data.values || [];
        const infoEmpleado = filasEmpleados.find(fila => fila[0] === telefono);
        
        if (!infoEmpleado) return null; // No está registrado

        const empleado = {
            telefono: infoEmpleado[0],
            nombre: infoEmpleado[1],
            sucursalAsignada: infoEmpleado[2],
            horaEsperada: null
        };

        // 2. Buscamos si tiene un horario específico programado para el día de hoy
        const vistaCalendario = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Calendario_Horarios!A2:D', // Columnas: Teléfono, Nombre, Fecha, Hora Entrada Esperada
        });

        const filasCalendario = vistaCalendario.data.values || [];
        // Filtrar por teléfono y fecha exacta (AAAA-MM-DD)
        const registroHorario = filasCalendario.find(fila => fila[0] === telefono && fila[2] === fechaLaboral);

        if (registroHorario) {
            empleado.horaEsperada = registroHorario[3]; // Guarda "08:00" o "DESCANSO"
        }

        return empleado;
    } catch (error) {
        console.error('Error en obtenerHorarioEmpleado:', error);
        return null;
    }
}

/**
 * Registra o actualiza la asistencia (Entradas, Primeros Intentos y Salidas Definitivas).
 */
async function guardarAsistencia(telefono, fechaLaboral, datos) {
    try {
        const vistaAsistencia = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Asistencia!A2:H',
        });

        const filas = vistaAsistencia.data.values || [];
        // Intentar encontrar si el empleado ya abrió turno en esta fecha laboral
        const indiceFila = filas.findIndex(fila => fila[0] === telefono && fila[2] === fechaLaboral);

        // Traer datos del empleado para rellenar campos obligatorios
        const infoEmpleado = await obtenerHorarioEmpleado(telefono, fechaLaboral);

        if (indiceFila === -1) {
            // --- REGISTRO NUEVO (Normalmente la Entrada) ---
            const nuevaFila = [
                telefono,
                infoEmpleado ? infoEmpleado.nombre : 'Desconocido',
                fechaLaboral,
                datos.horaEntrada || '',       // Hora Entrada
                datos.estatusEntrada || '',    // Estatus Entrada OK?
                datos.horaPrimerIntentoSalida || '', // Hora Primer Intento Salida
                datos.horaSalidaDefinitiva || '',    // Hora Salida Definitiva
                datos.distancia ? Math.round(datos.distancia) : '' // Distancia Salida (m)
            ];

            await sheets.spreadsheets.values.append({
                spreadsheetId: SPREADSHEET_ID,
                range: 'Asistencia!A2',
                valueInputOption: 'USER_ENTERED',
                resource: { values: [nuevaFila] },
            });
        } else {
            // --- ACTUALIZACIÓN DE FILA EXISTENTE (Salidas o Reintentos) ---
            const numeroFilaReal = indiceFila + 2; // Compensar índice base 0 y encabezado
            
            // Si mandan salida definitiva limpia, actualizamos solo esa columna
            if (datos.horaSalidaDefinitiva) {
                await sheets.spreadsheets.values.update({
                    spreadsheetId: SPREADSHEET_ID,
                    range: `Asistencia!G${numeroFilaReal}`, // Columna G: Hora Salida Definitiva
                    valueInputOption: 'USER_ENTERED',
                    resource: { values: [[datos.horaSalidaDefinitiva]] },
                });
            }
            
            // Si mandan un intento fuera de rango, registramos el incidente sin pisar lo demás
            if (datos.horaPrimerIntentoSalida) {
                // Actualiza Columna F (Primer intento) y Columna H (Distancia)
                await sheets.spreadsheets.values.update({
                    spreadsheetId: SPREADSHEET_ID,
                    range: `Asistencia!F${numeroFilaReal}:H${numeroFilaReal}`,
                    valueInputOption: 'USER_ENTERED',
                    resource: { values: [[datos.horaPrimerIntentoSalida, '', Math.round(datos.distancia)]] },
                });
            }
        }
    } catch (error) {
        console.error('Error en guardarAsistencia:', error);
    }
}

/**
 * Automatización Opción A: Lee 'Horario_Base' y expande las fechas de la Próxima Semana.
 */
async function generarSiguienteSemana() {
    try {
        const vistaBase = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Horario_Base!A2:J', // Teléfono, Nombre, Sucursal, Lun, Mar, Mie, Jue, Vie, Sab, Dom
        });

        const filasBase = vistaBase.data.values || [];
        if (filasBase.length === 0) return;

        // Calcular fechas del próximo lunes al próximo domingo
        const proximasFilas = [];
        const hoy = new Date();
        const diasHastaLunes = (1 + 7 - hoy.getDay()) % 7 || 7;
        
        const lunesProximo = new Date(hoy);
        lunesProximo.setDate(hoy.getDate() + diasHastaLunes);

        for (const empleado of filasBase) {
            const [telefono, nombre, sucursal, ...diasSemana] = empleado;

            for (let i = 0; i < 7; i++) {
                const fechaDia = new Date(lunesProximo);
                fechaDia.setDate(lunesProximo.getDate() + i);
                const fechaFormateada = fechaDia.toISOString().split('T')[0];
                const horaEsperada = diasSemana[i] || 'DESCANSO';

                proximasFilas.push([telefono, nombre, fechaFormateada, horaEsperada]);
            }
        }

        // Insertar en bloque para no saturar la API de Google
        await sheets.spreadsheets.values.append({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Calendario_Horarios!A2',
            valueInputOption: 'USER_ENTERED',
            resource: { values: proximasFilas },
        });

        console.log('✅ Calendario de la próxima semana generado con éxito.');
    } catch (error) {
        console.error('Error generando la siguiente semana:', error);
    }
}

/**
 * Calcula horas totales, horas extras de 48h semanales y compila incidencias para el lunes 7 AM.
 */
async function obtenerResumenSemanal() {
    try {
        const vistaAsistencia = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Asistencia!A2:H',
        });

        const filas = vistaAsistencia.data.values || [];
        if (filas.length === 0) return '📊 *Resumen Semanal de Incidencias:*\nNo hay registros esta semana.';

        const resumen = {};

        // Procesar fila por fila mapeando por empleado
        filas.forEach(fila => {
            const [telefono, nombre, , entrada, estatus, , salida, distancia, fuera] = fila;
            if (!resumen[telefono]) {
                resumen[telefono] = { nombre, horasTrabajadas: 0, retardos: 0, fueraRango: 0 };
            }

            if (estatus === 'RETARDO') resumen[telefono].retardos++;
            if (fuera) resumen[telefono].fueraRango++;

            if (entrada && salida) {
                // Cálculo simple de horas de diferencia por día
                const [hE, mE] = entrada.split(':').map(Number);
                const [hS, mS] = salida.split(':').map(Number);
                let diffHrs = (hS + mS/60) - (hE + mE/60);
                if (diffHrs < 0) diffHrs += 24; // Resuelve turnos que pasan de la medianoche
                resumen[telefono].horasTrabajadas += diffHrs;
            }
        });

        // Construir mensaje de texto estructurado para WhatsApp
        let mensaje = `📊 *RESUMEN SEMANAL DE INCIDENCIAS*\n🗓️ Horario de corte: Lun 7:00 AM\n\n`;
        
        for (const id in resumen) {
            const emp = resumen[id];
            const hrsRegulares = 48;
            const extras = emp.horasTrabajadas > hrsRegulares ? (emp.horasTrabajadas - hrsRegulares).toFixed(1) : 0;
            
            mensaje += `👤 *Empleado:* ${emp.nombre}\n`;
            mensaje += `⏱️ Horas Totales: ${emp.horasTrabajadas.toFixed(1)} hrs\n`;
            mensaje += `🚀 Horas Extras: ${extras} hrs\n`;
            mensaje += `🚨 Retardos: ${emp.retardos}\n`;
            if (emp.fueraRango > 0) mensaje += `📍 Salidas Fuera de Rango: ${emp.fueraRango}\n`;
            mensaje += `----------------------------\n`;
        }

        return mensaje;
    } catch (error) {
        console.error('Error calculando resumen semanal:', error);
        return '⚠️ Error al generar el resumen semanal desde la base de datos.';
    }
}

/**
 * Consulta la pestaña 'Administradores' y devuelve una lista con los números autorizados.
 */
async function obtenerAdministradores() {
    try {
        const vistaAdmin = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Administradores!A2:A', // Lee únicamente la columna de teléfonos
        });

        const filas = vistaAdmin.data.values || [];
        // Convertir la matriz bidimensional en una lista plana de strings: ['525512345678', '525587654321']
        return filas.map(fila => fila[0].trim());
} catch (error) {
        console.error('Error al obtener administradores desde Sheets:', error);
        return []; // Si hay error, devuelve lista vacía por seguridad
    }
}

// 🚪 ESTO VA AL FINAL DE TODO EL ARCHIVO:
module.exports = {
    obtenerHorarioEmpleado,
    guardarAsistencia,
    generarSiguienteSemana,
    obtenerResumenSemanal,
    obtenerAdministradores
};

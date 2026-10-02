"use strict";
import ExcelJS from 'exceljs'
import fs from 'fs'
import path from 'path'
import os from 'os'
import { SPREADSHEET_ID, SUCURSALES, MAP_SUCURSAL_A_CHECADOR, getChecadorDeSucursal } from './config.js'
import { getRows, sheetsClient, getHorarioBaseMap } from './sheets.js'
import { fechaLaboral, horaMX, minutos, parseFechaMX, normaliza, parseHorarioRango, getRangoSemana, sucursalCoincideConFiltro, scoreEmpleado, calcularExtra } from './utils.js'

export async function asistenciaHoy(filtroSucursal, jid, sock){
  "use strict";
  const sClient = await sheetsClient();
  const [baseRows, asisRows] = await Promise.all([ getRows('Horario_Base!A2:K'), sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'}).then(r=>r.data.values||[]) ]);
  const fLab = fechaLaboral();
  const ahoraMin = minutos(horaMX());
  const ahoraMXDate = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'}));
  const diaNum = ahoraMXDate.getDay();

  // FIX: mapeo exacto por sucursal, no includes mezclados
  const fNorm = (filtroSucursal||'').toLowerCase();
  let permitidas = [];
  if(fNorm.includes('bucareli') || fNorm.includes('juarez')) permitidas = ['bucareli','juarez'];
  else if(fNorm.includes('coyo') || fNorm.includes('hotel') || fNorm.includes('trinidad')) permitidas = ['coyoacan','hotel','trinidad'];
  else permitidas = [fNorm];

  let llego=[], retardo=[], falta=[], futuro=[];
  for(const r of baseRows){
    const sucBaseLower = (r[2]||'').toLowerCase();
    const sucBaseOriginal = r[2]||'';
    let inc = permitidas.some(p => sucBaseLower.includes(p));
    if(!inc) continue;

    const nombre = r[1]||'';
    const tel = (r[0]||'').replace(/\D/g,'').slice(-10);
    const mapa = {1:r[3],2:r[4],3:r[5],4:r[6],5:r[7],6:r[8],0:r[9]};
    const v = (mapa[diaNum]||'').toString().trim();
    const parsed = parseHorarioRango(v);
    if(!parsed) continue;
    const esLibre = parsed.entrada === 'LIBRE';
    const horaProg = esLibre? 'LIBRE' : parsed.entrada;
    const registro = asisRows.find(a => a[0]?.replace(/\D/g,'').slice(-10)===tel && a[2]===fLab);
    if(registro && registro[3]){
      const entrada = registro[3];
      const sucEnt = registro[6]||'';
      const dif = esLibre? 0 : minutos(entrada) - minutos(horaProg);
      if(esLibre){ llego.push(`• ${nombre} - Entró ${entrada} en ${sucEnt} ✅`) }
      else if(dif > 15){ retardo.push(`• ${nombre} - [${sucBaseOriginal}] Prog ${horaProg} - Entró ${entrada} - ⏰ ${dif}m tarde - ${sucEnt}`) }
      else llego.push(`• ${nombre} - [${sucBaseOriginal}] Prog ${horaProg} - Entró ${entrada} ✅ - ${sucEnt}`)
    } else {
      if(esLibre){ if(ahoraMin >= 20*60) falta.push(`• ${nombre} - [${sucBaseOriginal}] - ❌ sin llegar`); continue }
      const dif = ahoraMin - minutos(horaProg);
      if(dif < 0){ futuro.push(`• ${nombre} - [${sucBaseOriginal}] - Prog ${horaProg}`) }
      else falta.push(`• ${nombre} - [${sucBaseOriginal}] - Prog ${horaProg} - ❌ ${dif}m sin llegar`)
    }
  }
  let txt = `📍 *ASISTENCIA HOY ${fLab} - ${filtroSucursal.toUpperCase()}* ${horaMX()}\n\n✅ *A TIEMPO (${llego.length}):*\n${llego.join('\n')||'-'}\n\n⏰ *RETARDOS (${retardo.length}):*\n${retardo.join('\n')||'-'}\n\n❌ *NO HAN LLEGADO (${falta.length}):*\n${falta.join('\n')||'Todos llegaron'}\n\n⏳ *PRÓXIMOS (${futuro.length}):*\n${futuro.join('\n')||'-'}`;
  await sock.sendMessage(jid,{text:txt})
}

// NUEVA FUNCION - RETARDOS 20 MIN POR GRUPO
export async function checarRetardos20Min(sock){
  "use strict";
  const sClient = await sheetsClient();
  const baseRows = await getRows('Horario_Base!A2:K');
  const asisRows = (await sClient.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range:'Asistencia!A2:M'})).data.values||[];
  const fLab = fechaLaboral();
  const ahoraMin = minutos(horaMX());
  const diaNum = new Date(new Date().toLocaleString('en-US',{timeZone:'America/Mexico_City'})).getDay();
  const porGrupo = {};

  for(const r of baseRows){
    const sucBase = (r[2]||'').toLowerCase();
    const nombre = r[1]||'';
    const tel = (r[0]||'').replace(/\D/g,'').slice(-10);
    const mapa = {1:r[3],2:r[4],3:r[5],4:r[6],5:r[7],6:r[8],0:r[9]};
    const parsed = parseHorarioRango((mapa[diaNum]||'').toString().trim());
    if(!parsed || parsed.entrada === 'LIBRE') continue;
    const registro = asisRows.find(a => a[0]?.replace(/\D/g,'').slice(-10)===tel && a[2]===fLab);
    if(registro && registro[3]) continue;
    const dif = ahoraMin - minutos(parsed.entrada);
    if(dif >= 20){
      const keySuc = (sucBase.includes('bucareli')||sucBase.includes('juarez'))? 'bucareli' : 'coyoacan';
      const gid = MAP_SUCURSAL_A_CHECADOR[keySuc];
      if(!porGrupo[gid]) porGrupo[gid]=[];
      porGrupo[gid].push(`• ${nombre} [${r[2]}] Prog ${parsed.entrada} - ${dif}m sin llegar`);
    }
  }
  for(const [gid, lista] of Object.entries(porGrupo)){
    if(lista.length && gid){
      await sock.sendMessage(gid, {text: `⏰ *RETARDOS 20min - ${fLab} ${horaMX()}*\n\n${lista.join('\n')}`});
    }
  }
}

// resto de tus funciones generarExcelEmpleado, resumenEmpleado, reporteSucursal, generarExcelSemanaYEnviar, handleReportes quedan igual - solo agregales "use strict"; al inicio
export async function generarExcelEmpleado(nombreBuscarRaw, jid, sock, tipo='actual'){ "use strict"; /* tu codigo igual */ }
export async function resumenEmpleado(nombreBuscar, jidRespuesta, sock, tipo='actual'){ "use strict"; /* tu codigo igual */ }
export async function reporteSucursal(filtroSucursal, jid, sock, tipo='pasada'){ "use strict"; /* tu codigo igual */ }
export async function generarExcelSemanaYEnviar(filtroSucursal, jid, sock, tipo='pasada'){ "use strict"; /* tu codigo igual */ }
export async function handleReportes({ texto, jid, sock, filtroGrupo }){ "use strict"; /* tu codigo igual */ }

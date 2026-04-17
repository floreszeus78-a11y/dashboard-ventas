// ============================================================
// EXPORTACIONES.JS - Sistema completo de exportación
// Versión 2.0 - Con filtros avanzados por Estado GPV y Mes de Venta
// CORREGIDO - Sin errores de sintaxis
// ============================================================

// ============================================================
// FUNCIÓN CRÍTICA: LIMPIAR NÚMEROS DE FORMATO DE MONEDA
// ============================================================
function limpiarNumero(valor) {
    if (valor === null || valor === undefined || valor === '') return 0;
    if (typeof valor === 'number') return valor;
    
    let str = String(valor);
    str = str.replace(/[S\/\$]/g, '').trim();
    
    if (str.match(/\d+\.\d{3},\d{2}$/) || str.match(/\d+,\d{2}$/)) {
        str = str.replace(/\./g, '');
        str = str.replace(/,/g, '.');
    } else {
        str = str.replace(/,/g, '');
    }
    
    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
}

// ============================================================
// OBTENER DATOS GLOBALES
// ============================================================
function getDatosGlobales() {
    if (typeof DATOS_COMPLETOS !== 'undefined' && DATOS_COMPLETOS) return DATOS_COMPLETOS;
    if (typeof datosCompletos !== 'undefined' && datosCompletos) return datosCompletos;
    if (typeof datosGlobales !== 'undefined' && datosGlobales) return datosGlobales;
    return null;
}

// ============================================================
// FUNCIÓN PARA OBTENER LA FECHA DE ÚLTIMA ACTUALIZACIÓN
// ============================================================
function obtenerFechaUltimaActualizacion() {
    const datos = getDatosGlobales();
    if (datos?.equipos) {
        let fechaMax = null;
        
        datos.equipos.forEach(equipo => {
            const ultimaTrx = equipo.ultima_transaccion;
            if (!ultimaTrx || ultimaTrx === '' || ultimaTrx === null) return;
            
            let fechaObj = null;
            
            if (typeof ultimaTrx === 'string' && ultimaTrx.includes('-')) {
                const partes = ultimaTrx.split('-');
                if (partes.length === 3) {
                    const anio = parseInt(partes[0], 10);
                    const mes = parseInt(partes[1], 10) - 1;
                    const dia = parseInt(partes[2], 10);
                    if (!isNaN(anio) && !isNaN(mes) && !isNaN(dia)) {
                        fechaObj = new Date(anio, mes, dia);
                    }
                }
            } else if (typeof ultimaTrx === 'string' && ultimaTrx.includes('/')) {
                const partes = ultimaTrx.split('/');
                if (partes.length === 3) {
                    const v0 = parseInt(partes[0], 10);
                    const v1 = parseInt(partes[1], 10);
                    const anio = parseInt(partes[2].length === 2 ? '20' + partes[2] : partes[2], 10);
                    const intentoMM = new Date(anio, v0 - 1, v1);
                    if (!isNaN(intentoMM.getTime()) && v0 >= 1 && v0 <= 12 && v1 >= 1 && v1 <= 31) {
                        fechaObj = intentoMM;
                    } else {
                        const intentoDD = new Date(anio, v1 - 1, v0);
                        if (!isNaN(intentoDD.getTime())) {
                            fechaObj = intentoDD;
                        }
                    }
                }
            } else if (ultimaTrx instanceof Date) {
                fechaObj = ultimaTrx;
            }
            
            if (fechaObj && !isNaN(fechaObj.getTime())) {
                if (!fechaMax || fechaObj > fechaMax) {
                    fechaMax = fechaObj;
                }
            }
        });
        
        if (fechaMax) {
            const meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
            return `${fechaMax.getDate().toString().padStart(2, '0')} de ${meses[fechaMax.getMonth()]} del ${fechaMax.getFullYear()}`;
        }
    }
    
    const hoy = new Date();
    const meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
    return `${hoy.getDate().toString().padStart(2, '0')} de ${meses[hoy.getMonth()]} del ${hoy.getFullYear()}`;
}

// ============================================================
// FUNCIÓN DE ESTADO BASADA EN GPV
// ============================================================
function getEstadoPorGPV(gpv) {
    const valor = limpiarNumero(gpv);
    
    if (valor === 0 || valor === null || valor === undefined) return 'INACTIVO';
    if (valor >= 700) return 'ACTIVO';
    if (valor >= 400 && valor <= 699) return 'REGULAR';
    if (valor >= 1 && valor <= 399) return 'SIN TANTO USO';
    return 'INACTIVO';
}

// ============================================================
// FUNCIONES AUXILIARES
// ============================================================
function formatearNumeroExport(num, dec = 2) {
    if (num === null || num === undefined || isNaN(num)) return dec === 0 ? '0' : '0.00';
    const numero = Number(num);
    if (isNaN(numero)) return dec === 0 ? '0' : '0.00';
    return numero.toLocaleString('es-PE', { minimumFractionDigits: dec, maximumFractionDigits: dec });
}

function formatearFechaExport(fecha) {
    if (!fecha) return '-';
    if (typeof fecha === 'string' && fecha.includes('-')) {
        const partes = fecha.split('-');
        if (partes.length === 3) {
            return `${partes[2]}/${partes[1]}/${partes[0]}`;
        }
    }
    if (fecha instanceof Date) {
        return `${fecha.getDate().toString().padStart(2, '0')}/${(fecha.getMonth() + 1).toString().padStart(2, '0')}/${fecha.getFullYear()}`;
    }
    return fecha;
}

function obtenerNombreMesExport(fecha) {
    if (!fecha) return '-';
    const meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 
                   'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
    if (typeof fecha === 'string' && fecha.includes('-')) {
        const mes = parseInt(fecha.split('-')[1]);
        return meses[mes - 1];
    }
    return '-';
}

function obtenerMesDesdeFechaExport(fecha) {
    if (!fecha) return null;
    if (typeof fecha === 'string' && fecha.includes('-')) {
        const mesNum = parseInt(fecha.split('-')[1]);
        const meses = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SETIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
        if (!isNaN(mesNum) && mesNum >= 1 && mesNum <= 12) return meses[mesNum - 1];
    }
    return null;
}

// ============================================================
// DIÁLOGO DE EXPORTACIÓN
// ============================================================
function mostrarDialogoExportacion() {
    const datos = getDatosGlobales();
    const ejecutivos = [...new Set(datos?.equipos?.map(e => e.responsable_real).filter(e => e))] || [];
    
    const mesesSet = new Set();
    if (datos?.equipos) {
        datos.equipos.forEach(e => {
            const mes = obtenerMesDesdeFechaExport(e.fecha_venta);
            if (mes) mesesSet.add(mes);
        });
    }
    const ordenMeses = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SETIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
    const mesesDisponibles = ordenMeses.filter(m => mesesSet.has(m));
    
    const modal = document.createElement('div');
    modal.id = 'modalExportacion';
    modal.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(0,0,0,0.8);
        backdrop-filter: blur(4px);
        z-index: 10000;
        display: flex;
        align-items: center;
        justify-content: center;
        font-family: 'Inter', sans-serif;
    `;
    
    modal.innerHTML = `
        <div style="
            background: var(--superficie);
            border-radius: 24px;
            padding: 32px;
            max-width: 520px;
            width: 90%;
            max-height: 90vh;
            overflow-y: auto;
            border: 1px solid var(--borde);
            box-shadow: 0 20px 40px rgba(0,0,0,0.3);
        ">
            <h2 style="font-size: 24px; font-weight: 700; margin-bottom: 24px; color: var(--texto);">
                📥 Exportar Reporte
            </h2>
            
            <div style="margin-bottom: 20px;">
                <label style="display: block; margin-bottom: 8px; font-weight: 600; color: var(--texto);">
                    📄 Formato:
                </label>
                <select id="formatoExportacion" style="width:100%; padding:12px; border-radius:12px; border:1px solid var(--borde); background:var(--superficie2); color:var(--texto); margin-bottom:12px;">
                    <option value="excel">📊 Excel (CSV)</option>
                    <option value="pdf">📄 PDF (Imprimir / Guardar)</option>
                    <option value="txt">📝 Texto (TXT)</option>
                </select>
            </div>
            
            <div style="margin-bottom: 20px;">
                <label style="display: block; margin-bottom: 8px; font-weight: 600; color: var(--texto);">
                    🎯 Rango base:
                </label>
                <select id="rangoExportacion" style="width:100%; padding:12px; border-radius:12px; border:1px solid var(--borde); background:var(--superficie2); color:var(--texto); margin-bottom:12px;">
                    <option value="todos">📋 Todos los datos</option>
                    <option value="filtros_actuales">✅ Datos con filtros actuales del dashboard</option>
                    <option value="ejecutivo">👤 Por ejecutivo específico</option>
                    <option value="mes">📅 Por mes específico</option>
                </select>
                
                <div id="opcionEjecutivo" style="display: none; margin-top: 12px;">
                    <select id="ejecutivoSeleccionado" style="width:100%; padding:12px; border-radius:12px; border:1px solid var(--borde); background:var(--superficie2); color:var(--texto);">
                        <option value="">Seleccionar ejecutivo...</option>
                        ${ejecutivos.map(exec => `<option value="${exec}">${exec}</option>`).join('')}
                    </select>
                </div>
                
                <div id="opcionMes" style="display: none; margin-top: 12px;">
                    <select id="mesSeleccionado" style="width:100%; padding:12px; border-radius:12px; border:1px solid var(--borde); background:var(--superficie2); color:var(--texto);">
                        <option value="">Seleccionar mes...</option>
                        ${mesesDisponibles.map(mes => `<option value="${mes}">${mes}</option>`).join('')}
                    </select>
                </div>
            </div>
            
            <div style="margin-bottom: 20px;">
                <label style="display: block; margin-bottom: 12px; font-weight: 600; color: var(--texto);">
                    🎯 Filtrar por Estado GPV
                </label>
                <select id="filtroEstadoGPV" style="width:100%; padding:12px; border-radius:12px; border:1px solid var(--borde); background:var(--superficie2); color:var(--texto);">
                    <option value="todos">📋 Todos los estados</option>
                    <option value="meta">🎯 Meta (≥ S/ 700)</option>
                    <option value="promedio">📊 Promedio (S/ 400 - 699)</option>
                    <option value="bajo">⚠️ Bajo (S/ 1 - 399)</option>
                    <option value="riesgo">🚨 Riesgo (S/ 0)</option>
                </select>
            </div>
            
            <div style="margin-bottom: 20px;">
                <label style="display: block; margin-bottom: 12px; font-weight: 600; color: var(--texto);">
                    📅 Filtrar por Mes de Venta
                </label>
                <select id="filtroMesVenta" style="width:100%; padding:12px; border-radius:12px; border:1px solid var(--borde); background:var(--superficie2); color:var(--texto);">
                    <option value="todos">📅 Todos los meses</option>
                    ${mesesDisponibles.map(mes => `<option value="${mes}">${mes}</option>`).join('')}
                </select>
            </div>
            
            <div style="margin-bottom: 24px;">
                <label style="display: block; margin-bottom: 8px; font-weight: 600; color: var(--texto);">
                    📊 Contenido del reporte:
                </label>
                <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
                    <input type="checkbox" id="incluirTabla" checked>
                    <span>Incluir tabla detallada</span>
                </label>
            </div>
            
            <div style="display: flex; gap: 12px; justify-content: flex-end;">
                <button onclick="cerrarModalExportacion()" style="padding:12px 24px; border-radius:12px; border:1px solid var(--borde); background:var(--superficie2); color:var(--texto); cursor:pointer;">Cancelar</button>
                <button onclick="ejecutarExportacion()" style="padding:12px 24px; border-radius:12px; border:none; background:linear-gradient(135deg, var(--acento), #7c3aed); color:white; cursor:pointer;">Exportar</button>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    const rangoSelect = document.getElementById('rangoExportacion');
    const opcionEjecutivo = document.getElementById('opcionEjecutivo');
    const opcionMes = document.getElementById('opcionMes');
    
    rangoSelect.addEventListener('change', function(e) {
        opcionEjecutivo.style.display = e.target.value === 'ejecutivo' ? 'block' : 'none';
        opcionMes.style.display = e.target.value === 'mes' ? 'block' : 'none';
    });
}

function cerrarModalExportacion() {
    const modal = document.getElementById('modalExportacion');
    if (modal) modal.remove();
}

function obtenerDatosConFiltrosAplicados() {
    let datos = getDatosGlobales()?.equipos || [];
    if (datos.length === 0) return datos;
    
    if (typeof mesActual !== 'undefined' && mesActual !== 'TODOS') {
        datos = datos.filter(function(e) { return obtenerMesDesdeFechaExport(e.fecha_venta) === mesActual; });
    }
    if (typeof ejecutivoActual !== 'undefined' && ejecutivoActual !== 'TODOS') {
        datos = datos.filter(function(e) { return e.responsable_real === ejecutivoActual; });
    }
    if (typeof filtroGPVActual !== 'undefined' && filtroGPVActual !== 'todos') {
        datos = datos.filter(function(e) {
            var gpvActual = limpiarNumero(e.gpv_mes_actual_corriendo);
            switch(filtroGPVActual) {
                case 'meta': return gpvActual >= 700;
                case 'promedio': return gpvActual >= 400 && gpvActual < 700;
                case 'bajo': return gpvActual > 0 && gpvActual < 400;
                case 'riesgo': return gpvActual === 0;
                default: return true;
            }
        });
    }
    return datos;
}

// ============================================================
// EJECUTAR EXPORTACIÓN
// ============================================================
function ejecutarExportacion() {
    var formato = document.getElementById('formatoExportacion').value;
    var rango = document.getElementById('rangoExportacion').value;
    var filtroEstado = document.getElementById('filtroEstadoGPV')?.value || 'todos';
    var filtroMes = document.getElementById('filtroMesVenta')?.value || 'todos';
    var incluirTabla = document.getElementById('incluirTabla')?.checked !== false;
    
    var datos = [];
    var ejecutivoEspecifico = null;
    var mesEspecifico = null;
    var textoFiltroEstado = null;
    var textoFiltroMes = 'Todos los meses';
    
    var mapEstadoTexto = {
        'todos': null,
        'meta': '🎯 Meta (≥ S/ 700)',
        'promedio': '📊 Promedio (S/ 400 - 699)',
        'bajo': '⚠️ Bajo (S/ 1 - 399)',
        'riesgo': '🚨 Riesgo (S/ 0)'
    };
    textoFiltroEstado = mapEstadoTexto[filtroEstado] || null;
    
    switch(rango) {
        case 'filtros_actuales':
            datos = obtenerDatosConFiltrosAplicados();
            break;
        case 'todos':
            datos = (getDatosGlobales()?.equipos || []).slice();
            break;
        case 'ejecutivo':
            ejecutivoEspecifico = document.getElementById('ejecutivoSeleccionado').value;
            if (!ejecutivoEspecifico) {
                alert('Por favor selecciona un ejecutivo');
                return;
            }
            datos = (getDatosGlobales()?.equipos || []).filter(function(e) { return e.responsable_real === ejecutivoEspecifico; });
            break;
        case 'mes':
            mesEspecifico = document.getElementById('mesSeleccionado').value;
            if (!mesEspecifico) {
                alert('Por favor selecciona un mes');
                return;
            }
            datos = (getDatosGlobales()?.equipos || []).filter(function(e) { return obtenerMesDesdeFechaExport(e.fecha_venta) === mesEspecifico; });
            textoFiltroMes = mesEspecifico;
            break;
    }
    
    if (filtroEstado !== 'todos') {
        datos = datos.filter(function(e) {
            var gpvActual = limpiarNumero(e.gpv_mes_actual_corriendo);
            switch(filtroEstado) {
                case 'meta': return gpvActual >= 700;
                case 'promedio': return gpvActual >= 400 && gpvActual < 700;
                case 'bajo': return gpvActual > 0 && gpvActual < 400;
                case 'riesgo': return gpvActual === 0;
                default: return true;
            }
        });
    }
    
    if (filtroMes !== 'todos') {
        datos = datos.filter(function(e) { return obtenerMesDesdeFechaExport(e.fecha_venta) === filtroMes; });
        textoFiltroMes = filtroMes;
    }
    
    if (datos.length === 0) {
        alert('No hay datos para exportar con los filtros seleccionados');
        return;
    }
    
    cerrarModalExportacion();
    
    var nombreBase = 'reporte_ventas';
    if (ejecutivoEspecifico) {
        nombreBase = 'REPORTE-' + ejecutivoEspecifico.replace(/[^a-zA-Z0-9]/g, '_');
    }
    
    var filtrosReporte = {
        ejecutivo: ejecutivoEspecifico,
        mes: textoFiltroMes,
        estado: textoFiltroEstado
    };
    
    if (formato === 'excel') {
        exportarExcelCompleto(datos, incluirTabla, nombreBase, filtrosReporte);
    } else if (formato === 'pdf') {
        exportarPDFCompleto(datos, incluirTabla, nombreBase, filtrosReporte);
    } else if (formato === 'txt') {
        exportarTXTCompleto(datos, incluirTabla, nombreBase, filtrosReporte);
    }
}

// ============================================================
// EXPORTAR EXCEL (CSV)
// ============================================================
function exportarExcelCompleto(datos, incluirTabla, nombreBase, filtrosReporte) {
    filtrosReporte = filtrosReporte || {};
    var ejecutivo = filtrosReporte.ejecutivo;
    var mes = filtrosReporte.mes;
    var estado = filtrosReporte.estado;
    
    var nombreArchivo = nombreBase + '_' + new Date().toISOString().split('T')[0];
    var fechaActualizacion = obtenerFechaUltimaActualizacion();
    var contenido = [];
    
    var tituloReporte = 'REPORTE DE VENTAS';
    if (ejecutivo) tituloReporte += ' - EJECUTIVO: ' + ejecutivo;
    if (estado) tituloReporte += ' - ' + estado;
    if (mes && mes !== 'Todos los meses') tituloReporte += ' - MES: ' + mes;
    
    contenido.push('=== ' + tituloReporte + ' ===');
    contenido.push('Fecha de exportación: ' + new Date().toLocaleString());
    contenido.push('📅 ÚLTIMA ACTUALIZACIÓN DE DATOS: ' + fechaActualizacion);
    contenido.push('Total de registros: ' + datos.length);
    contenido.push('');
    
    if (incluirTabla) {
        contenido.push('=== TABLA DETALLADA ===');
        contenido.push('');
        
        var headers = ['#', 'Fecha Venta', 'Mes Venta', 'Comercio', 'Serie', 'RUC', 'Ejecutivo',
            'Fecha Activación', 'Última Transacción', 'GPV M0', 'TRX M0', 'GPV M1', 'TRX M1', 
            'GPV M2', 'TRX M2', 'Mes Actual', 'GPV Actual', 'TRX Actual', 'Estado'];
        
        contenido.push(headers.map(function(h) { return '"' + h + '"'; }).join(','));
        
        for (var i = 0; i < datos.length; i++) {
            var item = datos[i];
            var gpvActual = limpiarNumero(item.gpv_mes_actual_corriendo);
            var estadoItem = getEstadoPorGPV(gpvActual);
            var estadoTexto = '';
            if (estadoItem === 'ACTIVO') estadoTexto = 'Activo (≥ S/700)';
            else if (estadoItem === 'REGULAR') estadoTexto = 'Regular (S/400-699)';
            else if (estadoItem === 'SIN TANTO USO') estadoTexto = 'Sin tanto uso (S/1-399)';
            else estadoTexto = 'Inactivo (S/0)';
            
            var fechaVentaFormateada = formatearFechaExport(item.fecha_venta);
            var mesVenta = obtenerNombreMesExport(item.fecha_venta);
            var fechaActivacionFormateada = formatearFechaExport(item.dia_activo);
            var ultimaTransaccionFormateada = formatearFechaExport(item.ultima_transaccion);
            
            var row = [
                i + 1,
                '"' + fechaVentaFormateada + '"',
                '"' + mesVenta + '"',
                '"' + (item.comercio || '').replace(/"/g, '""') + '"',
                item.numero_serie || '',
                item.ruc || '',
                item.responsable_real || '',
                '"' + fechaActivacionFormateada + '"',
                '"' + ultimaTransaccionFormateada + '"',
                limpiarNumero(item.gpv_m0),
                item.trx_m0 || 0,
                limpiarNumero(item.gpv_m1),
                item.trx_m1 || 0,
                limpiarNumero(item.gpv_m2),
                item.trx_m2 || 0,
                item.etiqueta_mes_actual || '',
                gpvActual,
                item.trx_mes_actual_corriendo || 0,
                estadoTexto
            ];
            contenido.push(row.join(','));
        }
    }
    
    var blob = new Blob(['\ufeff' + contenido.join('\n')], { type: 'text/csv;charset=utf-8;' });
    descargarArchivo(blob, nombreArchivo + '.csv');
}

// ============================================================
// EXPORTAR TXT
// ============================================================
function exportarTXTCompleto(datos, incluirTabla, nombreBase, filtrosReporte) {
    filtrosReporte = filtrosReporte || {};
    var ejecutivo = filtrosReporte.ejecutivo;
    var mes = filtrosReporte.mes;
    var estado = filtrosReporte.estado;
    
    var nombreArchivo = nombreBase + '_' + new Date().toISOString().split('T')[0];
    var fechaActualizacion = obtenerFechaUltimaActualizacion();
    var contenido = [];
    var separador = '='.repeat(100);
    
    var tituloReporte = 'REPORTE DE VENTAS';
    if (ejecutivo) tituloReporte += ' - EJECUTIVO: ' + ejecutivo;
    if (estado) tituloReporte += ' - ' + estado;
    if (mes && mes !== 'Todos los meses') tituloReporte += ' - MES: ' + mes;
    
    contenido.push(separador);
    contenido.push(tituloReporte);
    contenido.push(separador);
    contenido.push('Fecha de exportación: ' + new Date().toLocaleString());
    contenido.push('📅 ÚLTIMA ACTUALIZACIÓN DE DATOS: ' + fechaActualizacion);
    contenido.push('Total de registros: ' + datos.length);
    contenido.push('');
    
    if (incluirTabla) {
        contenido.push(separador);
        contenido.push('TABLA DETALLADA DE EQUIPOS');
        contenido.push(separador);
        contenido.push('');
        
        function pad(texto, ancho) {
            var str = String(texto || '-');
            return str.length > ancho ? str.substring(0, ancho - 3) + '...' : str.padEnd(ancho);
        }
        
        contenido.push(pad('#', 5) + pad('Fecha Venta', 12) + pad('Mes Venta', 12) +
            pad('Comercio', 30) + pad('Serie', 15) + pad('RUC', 12) + pad('Ejecutivo', 20) +
            pad('Fecha Activación', 16) + pad('Última Transacción', 18) + pad('GPV M0', 10) +
            pad('TRX M0', 8) + pad('GPV M1', 10) + pad('TRX M1', 8) + pad('GPV M2', 10) +
            pad('TRX M2', 8) + pad('GPV Act', 10) + pad('TRX Act', 8) + pad('Estado', 18));
        contenido.push('-'.repeat(250));
        
        for (var i = 0; i < datos.length; i++) {
            var item = datos[i];
            var gpvActual = limpiarNumero(item.gpv_mes_actual_corriendo);
            var estadoItem = getEstadoPorGPV(gpvActual);
            var estadoTexto = '';
            if (estadoItem === 'ACTIVO') estadoTexto = 'Activo (≥700)';
            else if (estadoItem === 'REGULAR') estadoTexto = 'Regular (400-699)';
            else if (estadoItem === 'SIN TANTO USO') estadoTexto = 'Sin tanto uso (1-399)';
            else estadoTexto = 'Inactivo (0)';
            
            var fechaVentaFormateada = formatearFechaExport(item.fecha_venta);
            var mesVenta = obtenerNombreMesExport(item.fecha_venta);
            var fechaActivacionFormateada = formatearFechaExport(item.dia_activo);
            var ultimaTransaccionFormateada = formatearFechaExport(item.ultima_transaccion);
            
            contenido.push(pad(i + 1, 5) + pad(fechaVentaFormateada, 12) + pad(mesVenta, 12) +
                pad(item.comercio || '-', 30) + pad(item.numero_serie || '-', 15) + pad(item.ruc || '-', 12) +
                pad(item.responsable_real || '-', 20) + pad(fechaActivacionFormateada, 16) +
                pad(ultimaTransaccionFormateada, 18) + pad(limpiarNumero(item.gpv_m0), 10) +
                pad(item.trx_m0 || 0, 8) + pad(limpiarNumero(item.gpv_m1), 10) + pad(item.trx_m1 || 0, 8) +
                pad(limpiarNumero(item.gpv_m2), 10) + pad(item.trx_m2 || 0, 8) + pad(gpvActual, 10) +
                pad(item.trx_mes_actual_corriendo || 0, 8) + pad(estadoTexto, 18));
        }
        
        contenido.push('');
        contenido.push(separador);
        contenido.push('RESUMEN POR ESTADO');
        contenido.push(separador);
        
        var activos = 0, regulares = 0, sinUso = 0, inactivos = 0;
        for (var j = 0; j < datos.length; j++) {
            var gpv = limpiarNumero(datos[j].gpv_mes_actual_corriendo);
            var est = getEstadoPorGPV(gpv);
            if (est === 'ACTIVO') activos++;
            else if (est === 'REGULAR') regulares++;
            else if (est === 'SIN TANTO USO') sinUso++;
            else inactivos++;
        }
        
        contenido.push('✅ ACTIVOS (≥ S/700): ' + activos + ' equipos');
        contenido.push('🟠 REGULARES (S/400-699): ' + regulares + ' equipos');
        contenido.push('🟡 SIN TANTO USO (S/1-399): ' + sinUso + ' equipos');
        contenido.push('🔴 INACTIVOS (S/0): ' + inactivos + ' equipos');
    }
    
    var blob = new Blob([contenido.join('\n')], { type: 'text/plain;charset=utf-8;' });
    descargarArchivo(blob, nombreArchivo + '.txt');
}

// ============================================================
// EXPORTAR PDF
// ============================================================
function exportarPDFCompleto(datos, incluirTabla, nombreBase, filtrosReporte) {
    filtrosReporte = filtrosReporte || {};
    var ejecutivo = filtrosReporte.ejecutivo;
    var mes = filtrosReporte.mes;
    var estado = filtrosReporte.estado;
    
    var ventanaReporte = window.open('', '_blank');
    if (!ventanaReporte) {
        alert('Por favor, permite las ventanas emergentes para generar el reporte en PDF.');
        return;
    }

    var fechaActualizacion = obtenerFechaUltimaActualizacion();
    
    var tituloPrincipal = 'REPORTE DE VENTAS';
    if (ejecutivo) tituloPrincipal += ' - ' + ejecutivo;
    
    var activos = 0, regulares = 0, sinUso = 0, inactivos = 0;
    for (var j = 0; j < datos.length; j++) {
        var gpv = limpiarNumero(datos[j].gpv_mes_actual_corriendo);
        var est = getEstadoPorGPV(gpv);
        if (est === 'ACTIVO') activos++;
        else if (est === 'REGULAR') regulares++;
        else if (est === 'SIN TANTO USO') sinUso++;
        else inactivos++;
    }

    var contenidoHTML = '<!DOCTYPE html>\n';
    contenidoHTML += '<html lang="es">\n<head>\n';
    contenidoHTML += '<meta charset="UTF-8">\n';
    contenidoHTML += '<title>Reporte de Ventas - ' + nombreBase + '</title>\n';
    contenidoHTML += '<style>\n';
    contenidoHTML += '* { margin: 0; padding: 0; box-sizing: border-box; }\n';
    contenidoHTML += 'body { font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background: white; color: #1f2937; font-size: 9px; line-height: 1.3; }\n';
    contenidoHTML += '.report-container { max-width: 100%; margin: 0 auto; }\n';
    contenidoHTML += '.header { text-align: center; margin-bottom: 15px; padding-bottom: 15px; border-bottom: 3px solid #0ea5e9; }\n';
    contenidoHTML += '.header-icon { font-size: 32px; margin-bottom: 8px; }\n';
    contenidoHTML += 'h1 { color: #0ea5e9; font-size: 22px; font-weight: 800; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.5px; }\n';
    contenidoHTML += '.subheader { color: #6b7280; font-size: 10px; margin: 4px 0; }\n';
    contenidoHTML += '.fecha-actualizacion { display: inline-block; background: #dcfce7; border: 1px solid #22c55e; padding: 6px 16px; border-radius: 20px; font-weight: 700; font-size: 10px; color: #16a34a; margin-top: 8px; }\n';
    contenidoHTML += '.total-registros { display: inline-block; background: #f0f9ff; border: 1px solid #0ea5e9; padding: 6px 16px; border-radius: 20px; font-weight: 700; font-size: 11px; color: #0ea5e9; margin-top: 8px; margin-left: 8px; }\n';
    contenidoHTML += '.filtros-badge { margin-top: 10px; display: flex; gap: 8px; flex-wrap: wrap; justify-content: center; }\n';
    contenidoHTML += '.filtro-tag { display: inline-flex; align-items: center; gap: 6px; padding: 4px 12px; border-radius: 20px; font-size: 10px; font-weight: 700; }\n';
    contenidoHTML += '.resumen-section { margin: 20px 0; }\n';
    contenidoHTML += '.resumen-title { font-size: 14px; font-weight: 700; color: #374151; margin-bottom: 12px; display: flex; align-items: center; gap: 8px; }\n';
    contenidoHTML += '.resumen-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 20px; }\n';
    contenidoHTML += '.resumen-card { text-align: center; padding: 15px 10px; border-radius: 8px; border: 1px solid #e5e7eb; }\n';
    contenidoHTML += '.resumen-card.activos { background: #f0fdf4; border-color: #22c55e; }\n';
    contenidoHTML += '.resumen-card.regulares { background: #fff7ed; border-color: #f97316; }\n';
    contenidoHTML += '.resumen-card.sin-uso { background: #fefce8; border-color: #eab308; }\n';
    contenidoHTML += '.resumen-card.inactivos { background: #fef2f2; border-color: #ef4444; }\n';
    contenidoHTML += '.resumen-numero { font-size: 28px; font-weight: 800; margin-bottom: 4px; }\n';
    contenidoHTML += '.resumen-card.activos .resumen-numero { color: #22c55e; }\n';
    contenidoHTML += '.resumen-card.regulares .resumen-numero { color: #f97316; }\n';
    contenidoHTML += '.resumen-card.sin-uso .resumen-numero { color: #eab308; }\n';
    contenidoHTML += '.resumen-card.inactivos .resumen-numero { color: #ef4444; }\n';
    contenidoHTML += '.resumen-label { font-size: 9px; color: #6b7280; font-weight: 600; }\n';
    contenidoHTML += '.tabla-section { margin-top: 20px; }\n';
    contenidoHTML += '.tabla-title { font-size: 14px; font-weight: 700; color: #374151; margin-bottom: 12px; display: flex; align-items: center; gap: 8px; padding-left: 8px; border-left: 4px solid #0ea5e9; }\n';
    contenidoHTML += 'table { width: 100%; border-collapse: collapse; font-size: 8px; margin-top: 10px; }\n';
    contenidoHTML += 'th { background: #f8fafc; padding: 8px 6px; text-align: center; font-weight: 700; color: #374151; border: 1px solid #d1d5db; font-size: 8px; white-space: nowrap; }\n';
    contenidoHTML += 'td { padding: 6px; border: 1px solid #e5e7eb; vertical-align: middle; text-align: center; }\n';
    contenidoHTML += 'tr:nth-child(even) { background-color: #fafafa; }\n';
    contenidoHTML += '.col-num { width: 3%; }\n';
    contenidoHTML += '.col-fecha { width: 7%; }\n';
    contenidoHTML += '.col-mes { width: 6%; }\n';
    contenidoHTML += '.col-comercio { width: 14%; text-align: left !important; }\n';
    contenidoHTML += '.col-serie { width: 10%; font-family: monospace; font-size: 7px; }\n';
    contenidoHTML += '.col-ruc { width: 8%; font-family: monospace; }\n';
    contenidoHTML += '.col-ejecutivo { width: 8%; }\n';
    contenidoHTML += '.col-fecha-act { width: 7%; }\n';
    contenidoHTML += '.col-ultima { width: 7%; }\n';
    contenidoHTML += '.col-gpv { width: 6%; text-align: right !important; }\n';
    contenidoHTML += '.col-trx { width: 4%; }\n';
    contenidoHTML += '.col-estado { width: 10%; }\n';
    contenidoHTML += '.estado-badge { display: inline-flex; align-items: center; gap: 4px; padding: 3px 8px; border-radius: 12px; font-size: 8px; font-weight: 700; white-space: nowrap; }\n';
    contenidoHTML += '.estado-activo { background: #dcfce7; color: #16a34a; }\n';
    contenidoHTML += '.estado-regular { background: #ffedd5; color: #ea580c; }\n';
    contenidoHTML += '.estado-sin-uso { background: #fef9c3; color: #ca8a04; }\n';
    contenidoHTML += '.estado-inactivo { background: #fee2e2; color: #dc2626; }\n';
    contenidoHTML += '.text-right { text-align: right !important; }\n';
    contenidoHTML += '.text-left { text-align: left !important; }\n';
    contenidoHTML += '.font-bold { font-weight: 700; }\n';
    contenidoHTML += '.footer { margin-top: 30px; padding-top: 15px; text-align: center; font-size: 8px; color: #9ca3af; border-top: 1px solid #e5e7eb; }\n';
    contenidoHTML += '@media print { body { margin: 0; padding: 10px; } th { background: #f8fafc !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; } tr:nth-child(even) { background-color: #fafafa !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; } .resumen-card { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; } .estado-badge { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; } .header { border-bottom-color: #0ea5e9 !important; -webkit-print-color-adjust: exact !important; } .filtro-tag { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; } }\n';
    contenidoHTML += '@page { size: landscape; margin: 10mm; }\n';
    contenidoHTML += '</style>\n</head>\n<body>\n';
    contenidoHTML += '<div class="report-container">\n';
    contenidoHTML += '<div class="header">\n';
    contenidoHTML += '<div class="header-icon">📊</div>\n';
    contenidoHTML += '<h1>' + tituloPrincipal + '</h1>\n';
    contenidoHTML += '<div class="subheader">Fecha de exportación: ' + new Date().toLocaleString() + '</div>\n';
    contenidoHTML += '<div><span class="fecha-actualizacion">📅 ACTUALIZADO HASTA: ' + fechaActualizacion + '</span>';
    contenidoHTML += '<span class="total-registros">📋 Total de registros: ' + datos.length + '</span></div>\n';
    contenidoHTML += '<div class="filtros-badge">\n';
    if (ejecutivo) contenidoHTML += '<span class="filtro-tag" style="background:#f0f9ff; border:1px solid #0ea5e9; color:#0ea5e9;">👤 ' + ejecutivo + '</span>\n';
    if (estado) contenidoHTML += '<span class="filtro-tag" style="background:#fff7ed; border:1px solid #f97316; color:#f97316;">' + estado + '</span>\n';
    if (mes && mes !== 'Todos los meses') contenidoHTML += '<span class="filtro-tag" style="background:#f0fdf4; border:1px solid #22c55e; color:#16a34a;">📅 ' + mes + '</span>\n';
    contenidoHTML += '</div>\n</div>\n';

    if (incluirTabla) {
        contenidoHTML += '<div class="resumen-section">\n';
        contenidoHTML += '<div class="resumen-title">📊 Resumen por Estado</div>\n';
        contenidoHTML += '<div class="resumen-grid">\n';
        contenidoHTML += '<div class="resumen-card activos"><div class="resumen-numero">' + activos + '</div><div class="resumen-label">Activos (≥ S/700)</div></div>\n';
        contenidoHTML += '<div class="resumen-card regulares"><div class="resumen-numero">' + regulares + '</div><div class="resumen-label">Regulares (S/400-699)</div></div>\n';
        contenidoHTML += '<div class="resumen-card sin-uso"><div class="resumen-numero">' + sinUso + '</div><div class="resumen-label">Sin tanto uso (S/1-399)</div></div>\n';
        contenidoHTML += '<div class="resumen-card inactivos"><div class="resumen-numero">' + inactivos + '</div><div class="resumen-label">Inactivos (S/0)</div></div>\n';
        contenidoHTML += '</div>\n</div>\n';
        
        contenidoHTML += '<div class="tabla-section">\n';
        contenidoHTML += '<div class="tabla-title">📋 Tabla Detallada de Equipos</div>\n';
        contenidoHTML += '<div style="overflow-x: auto;">\n';
        contenidoHTML += '<table>\n<thead>\n<tr>\n';
        contenidoHTML += '<th class="col-num">#</th><th class="col-fecha">Fecha<br>Venta</th><th class="col-mes">Mes<br>Venta</th>';
        contenidoHTML += '<th class="col-comercio text-left">Comercio</th><th class="col-serie">Serie</th><th class="col-ruc">RUC</th>';
        contenidoHTML += '<th class="col-ejecutivo">Ejecutivo</th><th class="col-fecha-act">Fecha<br>Activación</th>';
        contenidoHTML += '<th class="col-ultima">Última<br>Transacción</th><th class="col-gpv">GPV<br>M0</th><th class="col-trx">TRX<br>M0</th>';
        contenidoHTML += '<th class="col-gpv">GPV<br>M1</th><th class="col-trx">TRX<br>M1</th><th class="col-gpv">GPV<br>M2</th>';
        contenidoHTML += '<th class="col-trx">TRX<br>M2</th><th class="col-gpv">GPV<br>Actual</th><th class="col-trx">TRX<br>Actual</th>';
        contenidoHTML += '<th class="col-estado">Estado</th>\n</tr>\n</thead>\n<tbody>\n';
        
        for (var i = 0; i < datos.length; i++) {
            var item = datos[i];
            var gpvActual = limpiarNumero(item.gpv_mes_actual_corriendo);
            var gpvM0 = limpiarNumero(item.gpv_m0);
            var gpvM1 = limpiarNumero(item.gpv_m1);
            var gpvM2 = limpiarNumero(item.gpv_m2);
            
            var estadoActual = getEstadoPorGPV(gpvActual);
            var estadoClass = '';
            var estadoTexto = '';
            var estadoIcono = '';
            
            if (estadoActual === 'ACTIVO') {
                estadoClass = 'estado-activo';
                estadoTexto = 'Activo';
                estadoIcono = '✅';
            } else if (estadoActual === 'REGULAR') {
                estadoClass = 'estado-regular';
                estadoTexto = 'Regular';
                estadoIcono = '🟠';
            } else if (estadoActual === 'SIN TANTO USO') {
                estadoClass = 'estado-sin-uso';
                estadoTexto = 'Sin uso';
                estadoIcono = '🟡';
            } else {
                estadoClass = 'estado-inactivo';
                estadoTexto = 'Inactivo';
                estadoIcono = '🔴';
            }
            
            var fechaVenta = formatearFechaExport(item.fecha_venta);
            var mesVenta = obtenerNombreMesExport(item.fecha_venta);
            var fechaActivacion = formatearFechaExport(item.dia_activo);
            var ultimaTransaccion = formatearFechaExport(item.ultima_transaccion);
            
            contenidoHTML += '<tr>\n';
            contenidoHTML += '<td class="col-num">' + (i + 1) + '</td>\n';
            contenidoHTML += '<td class="col-fecha">' + fechaVenta + '</td>\n';
            contenidoHTML += '<td class="col-mes">' + mesVenta + '</td>\n';
            contenidoHTML += '<td class="col-comercio text-left">' + (item.comercio || '-').substring(0, 35) + '</td>\n';
            contenidoHTML += '<td class="col-serie"><code>' + (item.numero_serie || '-') + '</code></td>\n';
            contenidoHTML += '<td class="col-ruc"><code>' + (item.ruc || '-') + '</code></td>\n';
            contenidoHTML += '<td class="col-ejecutivo">' + (item.responsable_real || '-') + '</td>\n';
            contenidoHTML += '<td class="col-fecha-act">' + fechaActivacion + '</td>\n';
            contenidoHTML += '<td class="col-ultima">' + ultimaTransaccion + '</td>\n';
            contenidoHTML += '<td class="col-gpv font-bold">S/ ' + formatearNumeroExport(gpvM0, 0) + '</td>\n';
            contenidoHTML += '<td class="col-trx">' + (item.trx_m0 || 0) + '</td>\n';
            contenidoHTML += '<td class="col-gpv font-bold">S/ ' + formatearNumeroExport(gpvM1, 0) + '</td>\n';
            contenidoHTML += '<td class="col-trx">' + (item.trx_m1 || 0) + '</td>\n';
            contenidoHTML += '<td class="col-gpv font-bold">S/ ' + formatearNumeroExport(gpvM2, 0) + '</td>\n';
            contenidoHTML += '<td class="col-trx">' + (item.trx_m2 || 0) + '</td>\n';
            contenidoHTML += '<td class="col-gpv font-bold" style="color: #0ea5e9;">S/ ' + formatearNumeroExport(gpvActual, 0) + '</td>\n';
            contenidoHTML += '<td class="col-trx font-bold">' + (item.trx_mes_actual_corriendo || 0) + '</td>\n';
            contenidoHTML += '<td class="col-estado"><span class="estado-badge ' + estadoClass + '">' + estadoIcono + ' ' + estadoTexto + '</span></td>\n';
            contenidoHTML += '</tr>\n';
        }
        
        contenidoHTML += '</tbody>\n</table>\n</div>\n</div>\n';
    }

    contenidoHTML += '<div class="footer">Reporte generado automáticamente el ' + new Date().toLocaleString() + ' | Sistema de Gestión de Ventas Culqi</div>\n';
    contenidoHTML += '</div>\n';
    contenidoHTML += '<script>window.onload = function() { setTimeout(function() { window.print(); }, 500); };<\/script>\n';
    contenidoHTML += '</body>\n</html>';

    ventanaReporte.document.write(contenidoHTML);
    ventanaReporte.document.close();
}

function descargarArchivo(blob, nombreArchivo) {
    var link = document.createElement('a');
    var url = URL.createObjectURL(blob);
    link.href = url;
    link.setAttribute('download', nombreArchivo);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

// Hacer la función disponible globalmente
window.mostrarDialogoExportacion = mostrarDialogoExportacion;
window.cerrarModalExportacion = cerrarModalExportacion;
window.ejecutarExportacion = ejecutarExportacion;

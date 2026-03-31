// ============================================================
// EXPORTACIONES.JS - Sistema completo de exportación
// Versión con lógica de estado basada en GPV (según especificación)
// ACTIVO: ≥ 700 | REGULAR: 400-699 | SIN TANTO USO: 1-399 | INACTIVO: 0
// ============================================================

// ============================================================
// FUNCIÓN DE ESTADO BASADA EN GPV (ACTUALIZADA)
// ============================================================
function getEstadoPorGPV(gpv) {
    if (gpv === 0 || gpv === null || gpv === undefined) return 'INACTIVO';
    if (gpv >= 700) return 'ACTIVO';
    if (gpv >= 400 && gpv <= 699) return 'REGULAR';
    if (gpv >= 1 && gpv <= 399) return 'SIN TANTO USO';
    return 'INACTIVO';
}

// ============================================================
// FUNCIÓN PARA OBTENER TEXTO Y COLOR DEL ESTADO
// ============================================================
function obtenerEstadoTexto(estado) {
    switch(estado) {
        case 'ACTIVO': return '✅ Activo (≥ S/700)';
        case 'REGULAR': return '⚠️ Regular (S/400-699)';
        case 'SIN TANTO USO': return '📉 Sin tanto uso (S/1-399)';
        case 'INACTIVO': return '❌ Inactivo (S/0)';
        default: return '⚪ Sin datos';
    }
}

function obtenerColorEstado(estado) {
    switch(estado) {
        case 'ACTIVO': return '#22c55e';
        case 'REGULAR': return '#f97316';
        case 'SIN TANTO USO': return '#eab308';
        case 'INACTIVO': return '#ef4444';
        default: return '#64748b';
    }
}

// Función auxiliar para formatear números
function formatearNumeroExport(num, dec = 2) {
    if (num === null || num === undefined || isNaN(num)) return dec === 0 ? '0' : '0.00';
    const numero = Number(num);
    if (isNaN(numero)) return dec === 0 ? '0' : '0.00';
    return numero.toLocaleString('es-PE', { minimumFractionDigits: dec, maximumFractionDigits: dec });
}

// Función para formatear fecha a DD/MM/YYYY
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

// Función para obtener nombre del mes
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

// Función principal que muestra el diálogo de exportación
function mostrarDialogoExportacion() {
    const ejecutivos = [...new Set(DATOS_COMPLETOS?.equipos.map(e => e.responsable_real).filter(e => e))];
    const meses = DATOS_COMPLETOS?.meses_disponibles || [];
    
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
            max-width: 500px;
            width: 90%;
            max-height: 90vh;
            overflow-y: auto;
            border: 1px solid var(--borde);
            box-shadow: 0 20px 40px rgba(0,0,0,0.3);
        ">
            <h2 style="font-size: 24px; font-weight: 700; margin-bottom: 24px; color: var(--texto);">
                📥 Exportar Reporte
            </h2>
            
            <div style="margin-bottom: 24px;">
                <label style="display: block; margin-bottom: 8px; font-weight: 600; color: var(--texto);">
                    📄 Formato de exportación:
                </label>
                <select id="formatoExportacion" style="
                    width: 100%;
                    padding: 12px;
                    border-radius: 12px;
                    border: 1px solid var(--borde);
                    background: var(--superficie2);
                    color: var(--texto);
                    font-size: 14px;
                    cursor: pointer;
                ">
                    <option value="excel">📊 Excel (CSV)</option>
                    <option value="pdf">📄 PDF</option>
                    <option value="txt">📝 Texto (TXT)</option>
                </select>
            </div>
            
            <div style="margin-bottom: 24px;">
                <label style="display: block; margin-bottom: 8px; font-weight: 600; color: var(--texto);">
                    🎯 Rango de datos:
                </label>
                <select id="rangoExportacion" style="
                    width: 100%;
                    padding: 12px;
                    border-radius: 12px;
                    border: 1px solid var(--borde);
                    background: var(--superficie2);
                    color: var(--texto);
                    font-size: 14px;
                    cursor: pointer;
                    margin-bottom: 12px;
                ">
                    <option value="filtros_actuales">✅ Datos con filtros actuales</option>
                    <option value="todos">📋 Todos los datos</option>
                    <option value="ejecutivo">👤 Por ejecutivo específico</option>
                    <option value="mes">📅 Por mes específico</option>
                </select>
                
                <div id="opcionEjecutivo" style="display: none; margin-top: 12px;">
                    <select id="ejecutivoSeleccionado" style="
                        width: 100%;
                        padding: 12px;
                        border-radius: 12px;
                        border: 1px solid var(--borde);
                        background: var(--superficie2);
                        color: var(--texto);
                        font-size: 14px;
                    ">
                        <option value="">Seleccionar ejecutivo...</option>
                        ${ejecutivos.map(exec => `<option value="${exec}">${exec}</option>`).join('')}
                    </select>
                </div>
                
                <div id="opcionMes" style="display: none; margin-top: 12px;">
                    <select id="mesSeleccionado" style="
                        width: 100%;
                        padding: 12px;
                        border-radius: 12px;
                        border: 1px solid var(--borde);
                        background: var(--superficie2);
                        color: var(--texto);
                        font-size: 14px;
                    ">
                        <option value="">Seleccionar mes...</option>
                        ${meses.map(mes => `<option value="${mes}">${mes}</option>`).join('')}
                    </select>
                </div>
            </div>
            
            <div style="margin-bottom: 24px;">
                <label style="display: block; margin-bottom: 8px; font-weight: 600; color: var(--texto);">
                    📊 Información adicional:
                </label>
                <div style="display: flex; gap: 16px; flex-wrap: wrap;">
                    <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
                        <input type="checkbox" id="incluirTabla" checked>
                        <span>Incluir tabla detallada</span>
                    </label>
                </div>
            </div>
            
            <div style="display: flex; gap: 12px; justify-content: flex-end;">
                <button onclick="cerrarModalExportacion()" style="
                    padding: 12px 24px;
                    border-radius: 12px;
                    border: 1px solid var(--borde);
                    background: var(--superficie2);
                    color: var(--texto);
                    cursor: pointer;
                    font-weight: 600;
                ">Cancelar</button>
                <button onclick="ejecutarExportacion()" style="
                    padding: 12px 24px;
                    border-radius: 12px;
                    border: none;
                    background: linear-gradient(135deg, var(--acento), #7c3aed);
                    color: white;
                    cursor: pointer;
                    font-weight: 600;
                ">Exportar</button>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    const rangoSelect = document.getElementById('rangoExportacion');
    const opcionEjecutivo = document.getElementById('opcionEjecutivo');
    const opcionMes = document.getElementById('opcionMes');
    
    rangoSelect.addEventListener('change', (e) => {
        opcionEjecutivo.style.display = e.target.value === 'ejecutivo' ? 'block' : 'none';
        opcionMes.style.display = e.target.value === 'mes' ? 'block' : 'none';
    });
}

function cerrarModalExportacion() {
    const modal = document.getElementById('modalExportacion');
    if (modal) modal.remove();
}

function obtenerDatosConFiltrosAplicados() {
    let equipos = [...(DATOS_COMPLETOS?.equipos || [])];
    
    if (typeof mesActual !== 'undefined' && mesActual !== 'TODOS') {
        equipos = equipos.filter(e => e.mes_venta === mesActual);
    }
    if (typeof ejecutivoActual !== 'undefined' && ejecutivoActual !== 'TODOS') {
        equipos = equipos.filter(e => e.responsable_real === ejecutivoActual);
    }
    if (typeof filtroGPVActual !== 'undefined' && filtroGPVActual !== 'todos') {
        equipos = equipos.filter(e => {
            const gpvActual = parseFloat(e.gpv_mes_actual_corriendo) || 0;
            switch(filtroGPVActual) {
                case 'meta': return gpvActual >= 700;
                case 'promedio': return gpvActual >= 400 && gpvActual < 700;
                case 'bajo': return gpvActual > 0 && gpvActual < 400;
                case 'riesgo': return gpvActual === 0;
                default: return true;
            }
        });
    }
    
    return equipos;
}

function ejecutarExportacion() {
    const formato = document.getElementById('formatoExportacion').value;
    const rango = document.getElementById('rangoExportacion').value;
    const incluirTabla = document.getElementById('incluirTabla').checked;
    
    let datos = [];
    let ejecutivoEspecifico = null;
    let mesEspecifico = null;
    
    switch(rango) {
        case 'filtros_actuales':
            datos = obtenerDatosConFiltrosAplicados();
            if (typeof ejecutivoActual !== 'undefined' && ejecutivoActual !== 'TODOS') {
                ejecutivoEspecifico = ejecutivoActual;
            }
            if (typeof mesActual !== 'undefined' && mesActual !== 'TODOS') {
                mesEspecifico = mesActual;
            }
            break;
        case 'todos':
            datos = [...(DATOS_COMPLETOS?.equipos || [])];
            break;
        case 'ejecutivo':
            ejecutivoEspecifico = document.getElementById('ejecutivoSeleccionado').value;
            if (!ejecutivoEspecifico) {
                alert('Por favor selecciona un ejecutivo');
                return;
            }
            datos = DATOS_COMPLETOS.equipos.filter(e => e.responsable_real === ejecutivoEspecifico);
            break;
        case 'mes':
            mesEspecifico = document.getElementById('mesSeleccionado').value;
            if (!mesEspecifico) {
                alert('Por favor selecciona un mes');
                return;
            }
            datos = DATOS_COMPLETOS.equipos.filter(e => e.mes_venta === mesEspecifico);
            break;
    }
    
    if (datos.length === 0) {
        alert('No hay datos para exportar');
        return;
    }
    
    cerrarModalExportacion();
    
    let nombreBase = 'reporte_ventas';
    if (ejecutivoEspecifico) {
        const nombreLimpio = ejecutivoEspecifico.replace(/[^a-zA-Z0-9]/g, '_');
        nombreBase = `REPORTE-${nombreLimpio}`;
    }
    
    switch(formato) {
        case 'excel':
            exportarExcelCompleto(datos, incluirTabla, nombreBase, ejecutivoEspecifico, mesEspecifico);
            break;
        case 'pdf':
            exportarPDFCompleto(datos, incluirTabla, nombreBase, ejecutivoEspecifico, mesEspecifico);
            break;
        case 'txt':
            exportarTXTCompleto(datos, incluirTabla, nombreBase, ejecutivoEspecifico, mesEspecifico);
            break;
    }
}

// ============================================================
// EXPORTAR EXCEL (CSV)
// ============================================================
function exportarExcelCompleto(datos, incluirTabla, nombreBase, ejecutivoEspecifico, mesEspecifico) {
    const nombreArchivo = `${nombreBase}_${new Date().toISOString().split('T')[0]}`;
    let contenido = [];
    
    let tituloReporte = 'REPORTE DE VENTAS';
    if (ejecutivoEspecifico) tituloReporte += ` - EJECUTIVO: ${ejecutivoEspecifico}`;
    if (mesEspecifico) tituloReporte += ` - MES: ${mesEspecifico}`;
    
    contenido.push(`=== ${tituloReporte} ===`);
    contenido.push(`Fecha de exportación: ${new Date().toLocaleString()}`);
    contenido.push(`Total de registros: ${datos.length}`);
    contenido.push('');
    contenido.push('');
    
    if (incluirTabla) {
        contenido.push('=== TABLA DETALLADA ===');
        contenido.push('');
        
        const headers = [
            '#', 
            'Fecha Venta', 
            'Mes Venta',
            'Comercio', 
            'Serie', 
            'RUC', 
            'Ejecutivo',
            'Fecha Activación',
            'Última Transacción',
            'GPV M0', 
            'TRX M0', 
            'GPV M1', 
            'TRX M1', 
            'GPV M2', 
            'TRX M2',
            'Mes Actual',
            'GPV Actual', 
            'TRX Actual', 
            'Estado'
        ];
        
        contenido.push(headers.map(h => `"${h}"`).join(','));
        
        datos.forEach((item, i) => {
            const gpvActual = parseFloat(item.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            let estadoTexto = '';
            if (estado === 'ACTIVO') estadoTexto = 'Activo (≥ S/700)';
            else if (estado === 'REGULAR') estadoTexto = 'Regular (S/400-699)';
            else if (estado === 'SIN TANTO USO') estadoTexto = 'Sin tanto uso (S/1-399)';
            else estadoTexto = 'Inactivo (S/0)';
            
            const fechaVentaFormateada = formatearFechaExport(item.fecha_venta);
            const mesVenta = obtenerNombreMesExport(item.fecha_venta);
            const fechaActivacionFormateada = formatearFechaExport(item.dia_activo);
            const ultimaTransaccionFormateada = formatearFechaExport(item.ultima_transaccion);
            
            const row = [
                i + 1,
                `"${fechaVentaFormateada}"`,
                `"${mesVenta}"`,
                `"${(item.comercio || '').replace(/"/g, '""')}"`,
                item.numero_serie || '',
                item.ruc || '',
                item.responsable_real || '',
                `"${fechaActivacionFormateada}"`,
                `"${ultimaTransaccionFormateada}"`,
                item.gpv_m0 || 0,
                item.trx_m0 || 0,
                item.gpv_m1 || 0,
                item.trx_m1 || 0,
                item.gpv_m2 || 0,
                item.trx_m2 || 0,
                item.etiqueta_mes_actual || '',
                gpvActual,
                item.trx_mes_actual_corriendo || 0,
                estadoTexto
            ];
            contenido.push(row.join(','));
        });
    }
    
    const blob = new Blob(['\ufeff' + contenido.join('\n')], { type: 'text/csv;charset=utf-8;' });
    descargarArchivo(blob, `${nombreArchivo}.csv`);
}

// ============================================================
// EXPORTAR TXT
// ============================================================
function exportarTXTCompleto(datos, incluirTabla, nombreBase, ejecutivoEspecifico, mesEspecifico) {
    const nombreArchivo = `${nombreBase}_${new Date().toISOString().split('T')[0]}`;
    let contenido = [];
    const separador = '='.repeat(100);
    
    let tituloReporte = 'REPORTE DE VENTAS';
    if (ejecutivoEspecifico) tituloReporte += ` - EJECUTIVO: ${ejecutivoEspecifico}`;
    if (mesEspecifico) tituloReporte += ` - MES: ${mesEspecifico}`;
    
    contenido.push(separador);
    contenido.push(tituloReporte);
    contenido.push(separador);
    contenido.push(`Fecha de exportación: ${new Date().toLocaleString()}`);
    contenido.push(`Total de registros: ${datos.length}`);
    contenido.push('');
    contenido.push('');
    
    if (incluirTabla) {
        contenido.push(separador);
        contenido.push('TABLA DETALLADA DE EQUIPOS');
        contenido.push(separador);
        contenido.push('');
        
        const pad = (texto, ancho) => {
            const str = String(texto || '-');
            return str.length > ancho ? str.substring(0, ancho - 3) + '...' : str.padEnd(ancho);
        };
        
        contenido.push(
            pad('#', 5) +
            pad('Fecha Venta', 12) +
            pad('Mes Venta', 12) +
            pad('Comercio', 30) +
            pad('Serie', 15) +
            pad('RUC', 12) +
            pad('Ejecutivo', 20) +
            pad('Fecha Activación', 16) +
            pad('Última Transacción', 18) +
            pad('GPV M0', 10) +
            pad('TRX M0', 8) +
            pad('GPV M1', 10) +
            pad('TRX M1', 8) +
            pad('GPV M2', 10) +
            pad('TRX M2', 8) +
            pad('GPV Act', 10) +
            pad('TRX Act', 8) +
            pad('Estado', 12)
        );
        contenido.push('-'.repeat(250));
        
        datos.forEach((item, i) => {
            const gpvActual = parseFloat(item.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            let estadoTexto = '';
            if (estado === 'ACTIVO') estadoTexto = 'Activo';
            else if (estado === 'REGULAR') estadoTexto = 'Regular';
            else if (estado === 'SIN TANTO USO') estadoTexto = 'Sin tanto uso';
            else estadoTexto = 'Inactivo';
            
            const fechaVentaFormateada = formatearFechaExport(item.fecha_venta);
            const mesVenta = obtenerNombreMesExport(item.fecha_venta);
            const fechaActivacionFormateada = formatearFechaExport(item.dia_activo);
            const ultimaTransaccionFormateada = formatearFechaExport(item.ultima_transaccion);
            
            contenido.push(
                pad(i + 1, 5) +
                pad(fechaVentaFormateada, 12) +
                pad(mesVenta, 12) +
                pad(item.comercio || '-', 30) +
                pad(item.numero_serie || '-', 15) +
                pad(item.ruc || '-', 12) +
                pad(item.responsable_real || '-', 20) +
                pad(fechaActivacionFormateada, 16) +
                pad(ultimaTransaccionFormateada, 18) +
                pad(item.gpv_m0 || 0, 10) +
                pad(item.trx_m0 || 0, 8) +
                pad(item.gpv_m1 || 0, 10) +
                pad(item.trx_m1 || 0, 8) +
                pad(item.gpv_m2 || 0, 10) +
                pad(item.trx_m2 || 0, 8) +
                pad(gpvActual, 10) +
                pad(item.trx_mes_actual_corriendo || 0, 8) +
                pad(estadoTexto, 12)
            );
        });
        
        // Agregar resumen de estados
        contenido.push('');
        contenido.push(separador);
        contenido.push('RESUMEN POR ESTADO');
        contenido.push(separador);
        
        const activos = datos.filter(e => {
            const gpv = parseFloat(e.gpv_mes_actual_corriendo) || 0;
            return getEstadoPorGPV(gpv) === 'ACTIVO';
        }).length;
        const regulares = datos.filter(e => {
            const gpv = parseFloat(e.gpv_mes_actual_corriendo) || 0;
            return getEstadoPorGPV(gpv) === 'REGULAR';
        }).length;
        const sinUso = datos.filter(e => {
            const gpv = parseFloat(e.gpv_mes_actual_corriendo) || 0;
            return getEstadoPorGPV(gpv) === 'SIN TANTO USO';
        }).length;
        const inactivos = datos.filter(e => {
            const gpv = parseFloat(e.gpv_mes_actual_corriendo) || 0;
            return getEstadoPorGPV(gpv) === 'INACTIVO';
        }).length;
        
        contenido.push(`✅ ACTIVOS (≥ S/700): ${activos} equipos`);
        contenido.push(`⚠️ REGULARES (S/400-699): ${regulares} equipos`);
        contenido.push(`📉 SIN TANTO USO (S/1-399): ${sinUso} equipos`);
        contenido.push(`❌ INACTIVOS (S/0): ${inactivos} equipos`);
        
        // Agregar detalle de transacciones por mes
        contenido.push('');
        contenido.push(separador);
        contenido.push('DETALLE DE TRANSACCIONES POR MES (PRIMEROS 20 REGISTROS)');
        contenido.push(separador);
        contenido.push('');
        
        const padDetalle = (texto, ancho) => {
            const str = String(texto || '-');
            return str.length > ancho ? str.substring(0, ancho - 3) + '...' : str.padEnd(ancho);
        };
        
        contenido.push(
            padDetalle('#', 5) +
            padDetalle('Serie', 15) +
            padDetalle('Comercio', 30) +
            padDetalle('TRX M0', 8) +
            padDetalle('TRX M1', 8) +
            padDetalle('TRX M2', 8) +
            padDetalle('TRX Actual', 12)
        );
        contenido.push('-'.repeat(86));
        
        datos.slice(0, 20).forEach((item, i) => {
            contenido.push(
                padDetalle(i + 1, 5) +
                padDetalle(item.numero_serie || '-', 15) +
                padDetalle(item.comercio || '-', 30) +
                padDetalle(item.trx_m0 || 0, 8) +
                padDetalle(item.trx_m1 || 0, 8) +
                padDetalle(item.trx_m2 || 0, 8) +
                padDetalle(item.trx_mes_actual_corriendo || 0, 12)
            );
        });
    }
    
    const blob = new Blob([contenido.join('\n')], { type: 'text/plain;charset=utf-8;' });
    descargarArchivo(blob, `${nombreArchivo}.txt`);
}

// ============================================================
// EXPORTAR PDF - VERSIÓN OPTIMIZADA CON WINDOW.PRINT()
// ============================================================
function exportarPDFCompleto(datos, incluirTabla, nombreBase, ejecutivoEspecifico, mesEspecifico) {
    const ventanaReporte = window.open('', '_blank');
    if (!ventanaReporte) {
        alert('Por favor, permite las ventanas emergentes para generar el reporte en PDF.');
        return;
    }

    let tituloPrincipal = 'REPORTE DE VENTAS';
    if (ejecutivoEspecifico) tituloPrincipal += ` - EJECUTIVO: ${ejecutivoEspecifico}`;
    if (mesEspecifico) tituloPrincipal += ` - MES: ${mesEspecifico}`;

    let contenidoHTML = `
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Reporte de Ventas - ${nombreBase}</title>
        <style>
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body {
                font-family: 'Inter', 'Segoe UI', Arial, sans-serif;
                margin: 0;
                padding: 20px;
                background: white;
                color: #1f2937;
                font-size: 10px;
                line-height: 1.3;
            }
            .report-container { max-width: 100%; margin: 0 auto; }
            .header {
                text-align: center;
                margin-bottom: 30px;
                padding-bottom: 20px;
                border-bottom: 3px solid #0ea5e9;
            }
            h1 {
                color: #0ea5e9;
                font-size: 24px;
                font-weight: 800;
                margin-bottom: 10px;
            }
            .subheader {
                color: #6b7280;
                font-size: 11px;
                margin: 5px 0;
            }
            .total-registros {
                font-weight: 700;
                font-size: 13px;
                color: #0ea5e9;
                margin-top: 10px;
            }
            h2 {
                font-size: 16px;
                font-weight: 700;
                margin: 20px 0 15px 0;
                color: #374151;
                border-left: 4px solid #0ea5e9;
                padding-left: 12px;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                margin: 15px 0;
                font-size: 9px;
            }
            th {
                background: #f3f4f6;
                padding: 8px 6px;
                text-align: left;
                font-weight: 700;
                color: #1f2937;
                border: 1px solid #e5e7eb;
            }
            td {
                padding: 6px;
                border: 1px solid #e5e7eb;
                vertical-align: top;
            }
            tr:nth-child(even) { background-color: #f9fafb; }
            .footer {
                margin-top: 30px;
                padding-top: 15px;
                text-align: center;
                font-size: 8px;
                color: #9ca3af;
                border-top: 1px solid #e5e7eb;
            }
            .text-right { text-align: right; }
            .text-center { text-align: center; }
            .font-bold { font-weight: 700; }
            .estado-activo { color: #16a34a; font-weight: 700; }
            .estado-regular { color: #ea580c; font-weight: 700; }
            .estado-sin-uso { color: #eab308; font-weight: 700; }
            .estado-inactivo { color: #dc2626; font-weight: 700; }
            .page-break { page-break-before: always; }
            @media print {
                body { margin: 0; padding: 10px; }
                th { background: #f3f4f6; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                tr:nth-child(even) { background-color: #f9fafb; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
            }
        </style>
    </head>
    <body>
        <div class="report-container">
            <div class="header">
                <h1>📊 ${tituloPrincipal}</h1>
                <div class="subheader">Fecha de exportación: ${new Date().toLocaleString()}</div>
                <div class="total-registros">📋 Total de registros: ${datos.length}</div>
            </div>
    `;

    if (incluirTabla) {
        // Calcular resumen de estados
        const activos = datos.filter(e => {
            const gpv = parseFloat(e.gpv_mes_actual_corriendo) || 0;
            return getEstadoPorGPV(gpv) === 'ACTIVO';
        }).length;
        const regulares = datos.filter(e => {
            const gpv = parseFloat(e.gpv_mes_actual_corriendo) || 0;
            return getEstadoPorGPV(gpv) === 'REGULAR';
        }).length;
        const sinUso = datos.filter(e => {
            const gpv = parseFloat(e.gpv_mes_actual_corriendo) || 0;
            return getEstadoPorGPV(gpv) === 'SIN TANTO USO';
        }).length;
        const inactivos = datos.filter(e => {
            const gpv = parseFloat(e.gpv_mes_actual_corriendo) || 0;
            return getEstadoPorGPV(gpv) === 'INACTIVO';
        }).length;
        
        contenidoHTML += `
            <div style="background: #f9fafb; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
                <h3 style="margin-bottom: 10px;">📊 Resumen por Estado</h3>
                <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px;">
                    <div style="background: #dcfce7; padding: 10px; border-radius: 8px; text-align: center;">
                        <div style="font-size: 20px; font-weight: 700; color: #16a34a;">${activos}</div>
                        <div style="font-size: 10px;">Activos (≥ S/700)</div>
                    </div>
                    <div style="background: #ffedd5; padding: 10px; border-radius: 8px; text-align: center;">
                        <div style="font-size: 20px; font-weight: 700; color: #ea580c;">${regulares}</div>
                        <div style="font-size: 10px;">Regulares (S/400-699)</div>
                    </div>
                    <div style="background: #fef9c3; padding: 10px; border-radius: 8px; text-align: center;">
                        <div style="font-size: 20px; font-weight: 700; color: #eab308;">${sinUso}</div>
                        <div style="font-size: 10px;">Sin tanto uso (S/1-399)</div>
                    </div>
                    <div style="background: #fee2e2; padding: 10px; border-radius: 8px; text-align: center;">
                        <div style="font-size: 20px; font-weight: 700; color: #dc2626;">${inactivos}</div>
                        <div style="font-size: 10px;">Inactivos (S/0)</div>
                    </div>
                </div>
            </div>
            
            <h2>📋 Tabla Detallada de Equipos</h2>
            <div style="overflow-x: auto;">
                <table>
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Fecha Venta</th>
                            <th>Mes Venta</th>
                            <th>Comercio</th>
                            <th>Serie</th>
                            <th>RUC</th>
                            <th>Ejecutivo</th>
                            <th>Fecha Activación</th>
                            <th>Última Transacción</th>
                            <th class="text-right">GPV M0</th>
                            <th class="text-right">TRX M0</th>
                            <th class="text-right">GPV M1</th>
                            <th class="text-right">TRX M1</th>
                            <th class="text-right">GPV M2</th>
                            <th class="text-right">TRX M2</th>
                            <th class="text-right">GPV Actual</th>
                            <th class="text-right">TRX Actual</th>
                            <th>Estado</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        datos.forEach((item, i) => {
            const gpvActual = parseFloat(item.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            let estadoClass = '';
            let estadoTexto = '';
            
            if (estado === 'ACTIVO') {
                estadoClass = 'estado-activo';
                estadoTexto = '✅ Activo';
            } else if (estado === 'REGULAR') {
                estadoClass = 'estado-regular';
                estadoTexto = '⚠️ Regular';
            } else if (estado === 'SIN TANTO USO') {
                estadoClass = 'estado-sin-uso';
                estadoTexto = '📉 Sin tanto uso';
            } else {
                estadoClass = 'estado-inactivo';
                estadoTexto = '❌ Inactivo';
            }
            
            contenidoHTML += `
                <tr>
                    <td class="text-center">${i + 1}</td>
                    <td>${formatearFechaExport(item.fecha_venta)}</td>
                    <td>${obtenerNombreMesExport(item.fecha_venta)}</td>
                    <td>${(item.comercio || '-').substring(0, 35)}</td>
                    <td><code>${item.numero_serie || '-'}</code></td>
                    <td>${item.ruc || '-'}</td>
                    <td>${item.responsable_real || '-'}</td>
                    <td>${formatearFechaExport(item.dia_activo)}</td>
                    <td>${formatearFechaExport(item.ultima_transaccion)}</td>
                    <td class="text-right">${formatearNumeroExport(item.gpv_m0 || 0, 0)}</td>
                    <td class="text-right">${item.trx_m0 || 0}</td>
                    <td class="text-right">${formatearNumeroExport(item.gpv_m1 || 0, 0)}</td>
                    <td class="text-right">${item.trx_m1 || 0}</td>
                    <td class="text-right">${formatearNumeroExport(item.gpv_m2 || 0, 0)}</td>
                    <td class="text-right">${item.trx_m2 || 0}</td>
                    <td class="text-right font-bold">S/ ${formatearNumeroExport(gpvActual, 0)}</td>
                    <td class="text-right">${item.trx_mes_actual_corriendo || 0}</td>
                    <td class="text-center ${estadoClass}">${estadoTexto}</td>
                 </tr>
            `;
        });

        contenidoHTML += `
                    </tbody>
                </table>
            </div>
        `;
        
        // Agregar detalle de transacciones
        contenidoHTML += `
            <div class="page-break"></div>
            <h2>📊 Detalle de Transacciones por Mes</h2>
            <div style="overflow-x: auto;">
                <table>
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Serie</th>
                            <th>Comercio</th>
                            <th class="text-right">TRX M0</th>
                            <th class="text-right">TRX M1</th>
                            <th class="text-right">TRX M2</th>
                            <th class="text-right">TRX Actual</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        datos.slice(0, 50).forEach((item, i) => {
            contenidoHTML += `
                <tr>
                    <td class="text-center">${i + 1}</td>
                    <td><code>${item.numero_serie || '-'}</code></td>
                    <td>${(item.comercio || '-').substring(0, 35)}</td>
                    <td class="text-right">${item.trx_m0 || 0}</td>
                    <td class="text-right">${item.trx_m1 || 0}</td>
                    <td class="text-right">${item.trx_m2 || 0}</td>
                    <td class="text-right font-bold">${item.trx_mes_actual_corriendo || 0}</td>
                </tr>
            `;
        });
        
        contenidoHTML += `
                    </tbody>
                </table>
                ${datos.length > 50 ? '<p style="font-size: 9px; color: #6b7280; margin-top: 10px;">* Mostrando primeros 50 registros de transacciones</p>' : ''}
            </div>
        `;
    }

    contenidoHTML += `
            <div class="footer">
                Reporte generado automáticamente el ${new Date().toLocaleString()} | Sistema de Gestión de Ventas Culqi
            </div>
        </div>
    </body>
    </html>
    `;

    ventanaReporte.document.write(contenidoHTML);
    ventanaReporte.document.close();
    
    ventanaReporte.onload = function() {
        ventanaReporte.print();
    };
}

function descargarArchivo(blob, nombreArchivo) {
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.href = url;
    link.setAttribute('download', nombreArchivo);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

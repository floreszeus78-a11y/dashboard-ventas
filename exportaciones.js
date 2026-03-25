// ============================================================
// EXPORTACIONES.JS - Sistema completo de exportación
// Versión con lógica de estado basada en GPV (no días)
// ============================================================

// ============================================================
// FUNCIÓN DE ESTADO BASADA EN GPV (MISMA QUE EL DASHBOARD)
// ============================================================
function getEstadoPorGPV(gpv) {
    if (gpv === 0 || gpv === null || gpv === undefined) return 'INACTIVO';
    if (gpv < 400) return 'REGULAR';
    return 'ACTIVO';
}

// ============================================================
// FUNCIÓN PARA OBTENER TEXTO Y COLOR DEL ESTADO
// ============================================================
function obtenerEstadoTexto(estado) {
    switch(estado) {
        case 'ACTIVO': return '✅ Activo';
        case 'REGULAR': return '⚠️ Regular';
        case 'INACTIVO': return '❌ Inactivo';
        default: return '⚪ Sin datos';
    }
}

function obtenerColorEstado(estado) {
    switch(estado) {
        case 'ACTIVO': return '#22c55e';
        case 'REGULAR': return '#f97316';
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
            '#', 'Fecha Venta', 'Mes Venta', 'Fecha Activación', 'Última Transacción',
            'Comercio', 'Serie', 'RUC', 'Ejecutivo', 
            'GPV M0', 'TRX M0', 'GPV M1', 'TRX M1', 'GPV M2', 'TRX M2', 
            'Mes Actual', 'GPV Actual', 'TRX Actual', 'Estado'
        ];
        
        contenido.push(headers.map(h => `"${h}"`).join(','));
        
        datos.forEach((item, i) => {
            const gpvActual = parseFloat(item.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            const estadoTexto = estado === 'ACTIVO' ? 'Activo' : (estado === 'REGULAR' ? 'Regular' : 'Inactivo');
            const fechaVentaFormateada = formatearFechaExport(item.fecha_venta);
            const mesVenta = obtenerNombreMesExport(item.fecha_venta);
            const fechaActivacionFormateada = formatearFechaExport(item.dia_activo);
            const ultimaTransaccionFormateada = formatearFechaExport(item.ultima_transaccion);
            
            const row = [
                i + 1,
                `"${fechaVentaFormateada}"`,
                `"${mesVenta}"`,
                `"${fechaActivacionFormateada}"`,
                `"${ultimaTransaccionFormateada}"`,
                `"${(item.comercio || '').replace(/"/g, '""')}"`,
                item.numero_serie || '',
                item.ruc || '',
                item.responsable_real || '',
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
            pad('Fecha Activación', 14) +
            pad('Última Transacción', 16) +
            pad('Comercio', 28) +
            pad('Serie', 15) +
            pad('RUC', 12) +
            pad('Ejecutivo', 20) +
            pad('GPV M0', 10) +
            pad('TRX M0', 8) +
            pad('GPV M1', 10) +
            pad('TRX M1', 8) +
            pad('GPV M2', 10) +
            pad('TRX M2', 8) +
            pad('GPV Act', 10) +
            pad('Estado', 10)
        );
        contenido.push('-'.repeat(220));
        
        datos.forEach((item, i) => {
            const gpvActual = parseFloat(item.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            const estadoTexto = estado === 'ACTIVO' ? 'Activo' : (estado === 'REGULAR' ? 'Regular' : 'Inactivo');
            const fechaVentaFormateada = formatearFechaExport(item.fecha_venta);
            const mesVenta = obtenerNombreMesExport(item.fecha_venta);
            const fechaActivacionFormateada = formatearFechaExport(item.dia_activo);
            const ultimaTransaccionFormateada = formatearFechaExport(item.ultima_transaccion);
            
            contenido.push(
                pad(i + 1, 5) +
                pad(fechaVentaFormateada, 12) +
                pad(mesVenta, 12) +
                pad(fechaActivacionFormateada, 14) +
                pad(ultimaTransaccionFormateada, 16) +
                pad(item.comercio || '-', 28) +
                pad(item.numero_serie || '-', 15) +
                pad(item.ruc || '-', 12) +
                pad(item.responsable_real || '-', 20) +
                pad(item.gpv_m0 || 0, 10) +
                pad(item.trx_m0 || 0, 8) +
                pad(item.gpv_m1 || 0, 10) +
                pad(item.trx_m1 || 0, 8) +
                pad(item.gpv_m2 || 0, 10) +
                pad(item.trx_m2 || 0, 8) +
                pad(gpvActual, 10) +
                pad(estadoTexto, 10)
            );
        });
    }
    
    const blob = new Blob([contenido.join('\n')], { type: 'text/plain;charset=utf-8;' });
    descargarArchivo(blob, `${nombreArchivo}.txt`);
}

// ============================================================
// EXPORTAR PDF - TABLA GRANDE Y LEGIBLE
// ============================================================
function exportarPDFCompleto(datos, incluirTabla, nombreBase, ejecutivoEspecifico, mesEspecifico) {
    const nombreArchivo = `${nombreBase}_${new Date().toISOString().split('T')[0]}`;
    
    // Aumentamos el ancho para que quepa más texto
    const elementoPDF = document.createElement('div');
    elementoPDF.style.cssText = `
        position: absolute;
        top: -9999px;
        left: -9999px;
        width: 2600px;
        background: white;
        padding: 60px 50px;
        font-family: 'Inter', 'Segoe UI', Arial, sans-serif;
        color: #000000;
    `;
    
    let tituloPrincipal = '📊 REPORTE DE VENTAS';
    let subtitulo = '';
    if (ejecutivoEspecifico) {
        tituloPrincipal = `📊 REPORTE DE VENTAS - ${ejecutivoEspecifico}`;
        subtitulo = `👤 Ejecutivo: ${ejecutivoEspecifico}`;
    }
    if (mesEspecifico) {
        subtitulo += subtitulo ? ` | 📅 Mes: ${mesEspecifico}` : `📅 Mes: ${mesEspecifico}`;
    }
    
    let contenidoHTML = `
        <div style="margin-bottom: 50px; text-align: center; border-bottom: 4px solid #0ea5e9; padding-bottom: 30px;">
            <h1 style="color: #0ea5e9; margin-bottom: 20px; font-size: 48px; font-weight: 800;">${tituloPrincipal}</h1>
            ${subtitulo ? `<p style="color: #0ea5e9; margin: 10px 0; font-size: 24px; font-weight: 600;">${subtitulo}</p>` : ''}
            <p style="color: #4a5568; margin: 12px 0; font-size: 18px;">📅 Fecha de exportación: ${new Date().toLocaleString()}</p>
            <p style="color: #4a5568; margin: 8px 0; font-size: 18px; font-weight: 600;">📋 Total de registros: ${datos.length}</p>
        </div>
    `;
    
    if (incluirTabla) {
        contenidoHTML += `
            <div style="margin-top: 35px;">
                <h2 style="color: #1e293b; margin-bottom: 30px; font-size: 32px; font-weight: 700; border-left: 6px solid #0ea5e9; padding-left: 20px;">📋 TABLA DETALLADA DE EQUIPOS</h2>
                <table style="width: 100%; border-collapse: collapse; font-size: 14px; background: white; box-shadow: 0 4px 12px rgba(0,0,0,0.1);">
                    <thead>
                        <tr style="background: #1e293b; border-bottom: 3px solid #0ea5e9;">
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: center; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">#</th>
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: center; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">FECHA VENTA</th>
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: center; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">MES VENTA</th>
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: center; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">FECHA ACTIVACIÓN</th>
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: center; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">ÚLTIMA TRANSACCIÓN</th>
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: left; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">COMERCIO</th>
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: left; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">SERIE</th>
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: left; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">RUC</th>
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: left; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">EJECUTIVO</th>
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: right; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">GPV M0</th>
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: right; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">TRX M0</th>
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: right; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">GPV M1</th>
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: right; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">TRX M1</th>
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: right; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">GPV M2</th>
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: right; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">TRX M2</th>
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: right; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">GPV ACTUAL</th>
                            <th style="padding: 16px 12px; border: 1px solid #334155; text-align: center; font-weight: 700; color: white; background: #1e293b; font-size: 14px;">ESTADO</th>
                         </tr>
                    </thead>
                    <tbody>
        `;
        
        datos.forEach((item, i) => {
            const gpvActual = parseFloat(item.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            const estadoTexto = estado === 'ACTIVO' ? '✅ ACTIVO' : (estado === 'REGULAR' ? '⚠️ REGULAR' : '❌ INACTIVO');
            const estadoColor = estado === 'ACTIVO' ? '#16a34a' : (estado === 'REGULAR' ? '#ea580c' : '#dc2626');
            
            const fechaVentaFormateada = formatearFechaExport(item.fecha_venta);
            const mesVenta = obtenerNombreMesExport(item.fecha_venta);
            const fechaActivacionFormateada = formatearFechaExport(item.dia_activo);
            const ultimaTransaccionFormateada = formatearFechaExport(item.ultima_transaccion);
            
            const bgColor = i % 2 === 0 ? '#ffffff' : '#f8fafc';
            
            contenidoHTML += `
                <tr style="background: ${bgColor};">
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; text-align: center; color: #1f2937; font-weight: 600; font-size: 13px;">${i + 1}</td>
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; text-align: center; color: #1f2937; font-size: 13px;">${fechaVentaFormateada}</td>
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; text-align: center; color: #1f2937; font-weight: 600; font-size: 13px;">${mesVenta}</td>
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; text-align: center; color: #1f2937; font-size: 13px;">${fechaActivacionFormateada}</td>
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; text-align: center; color: #1f2937; font-size: 13px;">${ultimaTransaccionFormateada}</td>
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; color: #1f2937; font-weight: 600; font-size: 13px;">${(item.comercio || '-').substring(0, 45)}</td>
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; color: #4b5563; font-family: monospace; font-size: 11px;">${item.numero_serie || '-'}</td>
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; color: #4b5563; font-size: 12px;">${item.ruc || '-'}</td>
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; color: #4b5563; font-weight: 600; font-size: 13px;">${item.responsable_real || '-'}</td>
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; text-align: right; color: #1f2937; font-size: 13px;">${formatearNumeroExport(item.gpv_m0 || 0, 0)}</td>
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; text-align: right; color: #1f2937; font-weight: ${(item.trx_m0 || 0) > 0 ? '600' : 'normal'}; font-size: 13px;">${item.trx_m0 || 0}</td>
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; text-align: right; color: #1f2937; font-size: 13px;">${formatearNumeroExport(item.gpv_m1 || 0, 0)}</td>
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; text-align: right; color: #1f2937; font-weight: ${(item.trx_m1 || 0) > 0 ? '600' : 'normal'}; font-size: 13px;">${item.trx_m1 || 0}</td>
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; text-align: right; color: #1f2937; font-size: 13px;">${formatearNumeroExport(item.gpv_m2 || 0, 0)}</td>
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; text-align: right; color: #1f2937; font-weight: ${(item.trx_m2 || 0) > 0 ? '600' : 'normal'}; font-size: 13px;">${item.trx_m2 || 0}</td>
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; text-align: right; font-weight: 800; color: #0f172a; font-size: 14px;">S/ ${formatearNumeroExport(gpvActual, 0)}</td>
                    <td style="padding: 14px 10px; border: 1px solid #e2e8f0; text-align: center; color: ${estadoColor}; font-weight: 700; font-size: 13px;">${estadoTexto}</td>
                </tr>
            `;
        });
        
        contenidoHTML += `
                    </tbody>
                </table>
                <div style="margin-top: 35px; padding: 18px; background: #f1f5f9; border-radius: 16px; text-align: center; color: #475569; font-size: 14px; border: 1px solid #e2e8f0;">
                    📊 Reporte generado el ${new Date().toLocaleString()} | Total: ${datos.length} registros
                </div>
            </div>
        `;
    }
    
    elementoPDF.innerHTML = contenidoHTML;
    document.body.appendChild(elementoPDF);
    
    const scriptHtml2canvas = document.createElement('script');
    scriptHtml2canvas.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js';
    
    const scriptJspdf = document.createElement('script');
    scriptJspdf.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
    
    let scriptsLoaded = 0;
    
    function checkAndGeneratePDF() {
        if (scriptsLoaded === 2 && window.html2canvas && window.jspdf) {
            html2canvas(elementoPDF, {
                scale: 3.5,
                useCORS: true,
                logging: false,
                backgroundColor: '#ffffff'
            }).then(canvas => {
                const imgData = canvas.toDataURL('image/png', 1.0);
                const { jsPDF } = window.jspdf;
                const pdf = new jsPDF('p', 'mm', 'a4');
                const imgWidth = 190;
                const pageHeight = 277;
                const imgHeight = (canvas.height * imgWidth) / canvas.width;
                let heightLeft = imgHeight;
                let position = 0;
                
                pdf.addImage(imgData, 'PNG', 10, position, imgWidth, imgHeight);
                heightLeft -= pageHeight;
                
                while (heightLeft > 0) {
                    position = heightLeft - imgHeight;
                    pdf.addPage();
                    pdf.addImage(imgData, 'PNG', 10, position, imgWidth, imgHeight);
                    heightLeft -= pageHeight;
                }
                
                pdf.save(`${nombreArchivo}.pdf`);
                document.body.removeChild(elementoPDF);
            }).catch(err => {
                console.error('Error generando PDF:', err);
                alert('Error al generar el PDF. Por favor intenta con formato CSV o TXT.');
                document.body.removeChild(elementoPDF);
            });
        }
    }
    
    scriptHtml2canvas.onload = () => {
        scriptsLoaded++;
        checkAndGeneratePDF();
    };
    
    scriptJspdf.onload = () => {
        scriptsLoaded++;
        checkAndGeneratePDF();
    };
    
    document.head.appendChild(scriptHtml2canvas);
    document.head.appendChild(scriptJspdf);
    
    setTimeout(() => {
        if (scriptsLoaded < 2) {
            alert('Error: No se pudieron cargar las librerías para generar PDF. Por favor verifica tu conexión a internet.');
            document.body.removeChild(elementoPDF);
        }
    }, 10000);
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

// ============================================================
// EXPORTACIONES.JS - Sistema completo de exportación
// Versión corregida: PDF con fondo blanco, fechas DD/MM/YYYY, columna Mes Venta
// ============================================================

// Función auxiliar para obtener estado
function getEstadoPorGPV(gpv) {
    if (gpv === 0 || gpv === null || gpv === undefined) return 'INACTIVO';
    if (gpv < 400) return 'REGULAR';
    return 'ACTIVO';
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
    // Obtener lista de ejecutivos únicos
    const ejecutivos = [...new Set(DATOS_COMPLETOS?.equipos.map(e => e.responsable_real).filter(e => e))];
    const meses = DATOS_COMPLETOS?.meses_disponibles || [];
    
    // Crear modal personalizado
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
                        <input type="checkbox" id="incluirResumen" checked>
                        <span>Incluir resumen (KPIs)</span>
                    </label>
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
    
    // Mostrar/ocultar opciones según selección
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
    
    // Aplicar filtros actuales del dashboard
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
    const incluirResumen = document.getElementById('incluirResumen').checked;
    const incluirTabla = document.getElementById('incluirTabla').checked;
    
    let datos = [];
    
    // Obtener datos según el rango seleccionado
    switch(rango) {
        case 'filtros_actuales':
            datos = obtenerDatosConFiltrosAplicados();
            break;
        case 'todos':
            datos = [...(DATOS_COMPLETOS?.equipos || [])];
            break;
        case 'ejecutivo':
            const ejecutivo = document.getElementById('ejecutivoSeleccionado').value;
            if (!ejecutivo) {
                alert('Por favor selecciona un ejecutivo');
                return;
            }
            datos = DATOS_COMPLETOS.equipos.filter(e => e.responsable_real === ejecutivo);
            break;
        case 'mes':
            const mes = document.getElementById('mesSeleccionado').value;
            if (!mes) {
                alert('Por favor selecciona un mes');
                return;
            }
            datos = DATOS_COMPLETOS.equipos.filter(e => e.mes_venta === mes);
            break;
    }
    
    if (datos.length === 0) {
        alert('No hay datos para exportar');
        return;
    }
    
    // Cerrar modal
    cerrarModalExportacion();
    
    // Exportar según formato
    switch(formato) {
        case 'excel':
            exportarExcelCompleto(datos, incluirResumen, incluirTabla);
            break;
        case 'pdf':
            exportarPDFCompleto(datos, incluirResumen, incluirTabla);
            break;
        case 'txt':
            exportarTXTCompleto(datos, incluirResumen, incluirTabla);
            break;
    }
}

// ============================================================
// EXPORTAR EXCEL (CSV)
// ============================================================
function exportarExcelCompleto(datos, incluirResumen, incluirTabla) {
    const nombreArchivo = `reporte_ventas_${new Date().toISOString().split('T')[0]}`;
    let contenido = [];
    
    if (incluirResumen) {
        // Calcular totales para resumen
        let totalGPV = 0, totalGPV_M0 = 0, totalGPV_M1 = 0, totalGPV_M2 = 0;
        let totalTRX = 0, activos = 0, regulares = 0, inactivos = 0;
        
        datos.forEach(e => {
            const m0 = parseFloat(e.gpv_m0) || 0;
            const m1 = parseFloat(e.gpv_m1) || 0;
            const m2 = parseFloat(e.gpv_m2) || 0;
            totalGPV += (m0 + m1 + m2);
            totalGPV_M0 += m0;
            totalGPV_M1 += m1;
            totalGPV_M2 += m2;
            totalTRX += (parseInt(e.trx_m0) || 0) + (parseInt(e.trx_m1) || 0) + (parseInt(e.trx_m2) || 0);
            
            const gpvActual = parseFloat(e.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            if (estado === 'ACTIVO') activos++;
            else if (estado === 'REGULAR') regulares++;
            else inactivos++;
        });
        
        contenido.push('=== RESUMEN DEL REPORTE ===');
        contenido.push(`Fecha de exportación: ${new Date().toLocaleString()}`);
        contenido.push(`Total de registros: ${datos.length}`);
        contenido.push(`GPV Total: S/ ${totalGPV.toLocaleString('es-PE')}`);
        contenido.push(`GPV M0: S/ ${totalGPV_M0.toLocaleString('es-PE')}`);
        contenido.push(`GPV M1: S/ ${totalGPV_M1.toLocaleString('es-PE')}`);
        contenido.push(`GPV M2: S/ ${totalGPV_M2.toLocaleString('es-PE')}`);
        contenido.push(`Transacciones Totales: ${totalTRX.toLocaleString('es-PE')}`);
        contenido.push(`Equipos Activos: ${activos}`);
        contenido.push(`Equipos Regulares: ${regulares}`);
        contenido.push(`Equipos Inactivos: ${inactivos}`);
        contenido.push('');
        contenido.push('');
    }
    
    if (incluirTabla) {
        contenido.push('=== TABLA DETALLADA ===');
        contenido.push('');
        
        const headers = [
            '#', 'Fecha Venta', 'Mes Venta', 'Comercio', 'Serie', 'RUC', 
            'Fecha Activación', 'Ejecutivo', 'GPV M0', 'TRX M0', 
            'GPV M1', 'TRX M1', 'GPV M2', 'TRX M2', 'Mes Actual', 
            'GPV Actual', 'TRX Actual', 'Última Transacción', 'Estado'
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
                `"${(item.comercio || '').replace(/"/g, '""')}"`,
                item.numero_serie || '',
                item.ruc || '',
                `"${fechaActivacionFormateada}"`,
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
                `"${ultimaTransaccionFormateada}"`,
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
function exportarTXTCompleto(datos, incluirResumen, incluirTabla) {
    const nombreArchivo = `reporte_ventas_${new Date().toISOString().split('T')[0]}`;
    let contenido = [];
    const separador = '='.repeat(100);
    
    if (incluirResumen) {
        // Calcular totales
        let totalGPV = 0, totalGPV_M0 = 0, totalGPV_M1 = 0, totalGPV_M2 = 0;
        let totalTRX = 0, activos = 0, regulares = 0, inactivos = 0;
        
        datos.forEach(e => {
            const m0 = parseFloat(e.gpv_m0) || 0;
            const m1 = parseFloat(e.gpv_m1) || 0;
            const m2 = parseFloat(e.gpv_m2) || 0;
            totalGPV += (m0 + m1 + m2);
            totalGPV_M0 += m0;
            totalGPV_M1 += m1;
            totalGPV_M2 += m2;
            totalTRX += (parseInt(e.trx_m0) || 0) + (parseInt(e.trx_m1) || 0) + (parseInt(e.trx_m2) || 0);
            
            const gpvActual = parseFloat(e.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            if (estado === 'ACTIVO') activos++;
            else if (estado === 'REGULAR') regulares++;
            else inactivos++;
        });
        
        contenido.push(separador);
        contenido.push('RESUMEN DEL REPORTE');
        contenido.push(separador);
        contenido.push('');
        contenido.push(`Fecha de exportación: ${new Date().toLocaleString()}`);
        contenido.push(`Total de registros: ${datos.length}`);
        contenido.push('');
        contenido.push('📊 INDICADORES PRINCIPALES:');
        contenido.push(`   • GPV Total (Cohorte): S/ ${totalGPV.toLocaleString('es-PE')}`);
        contenido.push(`   • GPV M0 (Prueba): S/ ${totalGPV_M0.toLocaleString('es-PE')}`);
        contenido.push(`   • GPV M1: S/ ${totalGPV_M1.toLocaleString('es-PE')}`);
        contenido.push(`   • GPV M2: S/ ${totalGPV_M2.toLocaleString('es-PE')}`);
        contenido.push(`   • Transacciones Totales: ${totalTRX.toLocaleString('es-PE')}`);
        contenido.push('');
        contenido.push('📈 ESTADO DE EQUIPOS:');
        contenido.push(`   • Activos: ${activos}`);
        contenido.push(`   • Regulares: ${regulares}`);
        contenido.push(`   • Inactivos: ${inactivos}`);
        contenido.push('');
    }
    
    if (incluirTabla) {
        contenido.push(separador);
        contenido.push('TABLA DETALLADA DE EQUIPOS');
        contenido.push(separador);
        contenido.push('');
        
        // Función para formatear texto con ancho fijo
        const pad = (texto, ancho) => {
            const str = String(texto || '-');
            return str.length > ancho ? str.substring(0, ancho - 3) + '...' : str.padEnd(ancho);
        };
        
        contenido.push(
            pad('#', 5) +
            pad('Fecha Venta', 12) +
            pad('Mes Venta', 12) +
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
        contenido.push('-'.repeat(190));
        
        datos.forEach((item, i) => {
            const gpvActual = parseFloat(item.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            const estadoTexto = estado === 'ACTIVO' ? 'Activo' : (estado === 'REGULAR' ? 'Regular' : 'Inactivo');
            const fechaVentaFormateada = formatearFechaExport(item.fecha_venta);
            const mesVenta = obtenerNombreMesExport(item.fecha_venta);
            
            contenido.push(
                pad(i + 1, 5) +
                pad(fechaVentaFormateada, 12) +
                pad(mesVenta, 12) +
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
// EXPORTAR PDF - CORREGIDO (Fondo blanco, fechas DD/MM/YYYY, columna Mes Venta)
// ============================================================
// ============================================================
// EXPORTAR PDF - CORREGIDO (Con TRX M0, M1, M2 y fondo blanco)
// ============================================================
function exportarPDFCompleto(datos, incluirResumen, incluirTabla) {
    // Crear un elemento temporal para renderizar el PDF con colores fijos
    const elementoPDF = document.createElement('div');
    elementoPDF.style.cssText = `
        position: absolute;
        top: -9999px;
        left: -9999px;
        width: 1600px;
        background: white;
        padding: 40px;
        font-family: 'Inter', 'Segoe UI', Arial, sans-serif;
        color: #000000;
    `;
    
    let contenidoHTML = `
        <div style="margin-bottom: 30px; text-align: center; border-bottom: 3px solid #0ea5e9; padding-bottom: 20px;">
            <h1 style="color: #0ea5e9; margin-bottom: 10px; font-size: 32px; font-weight: 700;">📊 Reporte de Ventas</h1>
            <p style="color: #4a5568; margin: 6px 0; font-size: 14px;">Fecha de exportación: ${new Date().toLocaleString()}</p>
            <p style="color: #4a5568; margin: 6px 0; font-size: 14px; font-weight: 500;">Total de registros: ${datos.length}</p>
        </div>
    `;
    
    if (incluirResumen) {
        // Calcular totales
        let totalGPV = 0, totalGPV_M0 = 0, totalGPV_M1 = 0, totalGPV_M2 = 0;
        let totalTRX = 0, totalTRX_M0 = 0, totalTRX_M1 = 0, totalTRX_M2 = 0;
        let activos = 0, regulares = 0, inactivos = 0;
        
        datos.forEach(e => {
            const m0 = parseFloat(e.gpv_m0) || 0;
            const m1 = parseFloat(e.gpv_m1) || 0;
            const m2 = parseFloat(e.gpv_m2) || 0;
            const t0 = parseInt(e.trx_m0) || 0;
            const t1 = parseInt(e.trx_m1) || 0;
            const t2 = parseInt(e.trx_m2) || 0;
            
            totalGPV += (m0 + m1 + m2);
            totalGPV_M0 += m0;
            totalGPV_M1 += m1;
            totalGPV_M2 += m2;
            totalTRX += (t0 + t1 + t2);
            totalTRX_M0 += t0;
            totalTRX_M1 += t1;
            totalTRX_M2 += t2;
            
            const gpvActual = parseFloat(e.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            if (estado === 'ACTIVO') activos++;
            else if (estado === 'REGULAR') regulares++;
            else inactivos++;
        });
        
        contenidoHTML += `
            <div style="margin-bottom: 35px;">
                <h2 style="color: #1e293b; margin-bottom: 20px; font-size: 22px; font-weight: 600; border-left: 4px solid #0ea5e9; padding-left: 15px;">📈 Resumen General</h2>
                <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px;">
                    <div style="background: #f8fafc; padding: 15px 20px; border-radius: 16px; border: 1px solid #e2e8f0;">
                        <div style="color: #64748b; font-size: 13px; margin-bottom: 6px;">💰 GPV Total Cohortes</div>
                        <div style="color: #0f172a; font-size: 28px; font-weight: 700;">S/ ${formatearNumeroExport(totalGPV)}</div>
                    </div>
                    <div style="background: #f8fafc; padding: 15px 20px; border-radius: 16px; border: 1px solid #e2e8f0;">
                        <div style="color: #64748b; font-size: 13px; margin-bottom: 6px;">🧪 GPV M0 (Prueba)</div>
                        <div style="color: #0f172a; font-size: 28px; font-weight: 700;">S/ ${formatearNumeroExport(totalGPV_M0)}</div>
                    </div>
                    <div style="background: #f8fafc; padding: 15px 20px; border-radius: 16px; border: 1px solid #e2e8f0;">
                        <div style="color: #64748b; font-size: 13px; margin-bottom: 6px;">🚀 GPV M1</div>
                        <div style="color: #0f172a; font-size: 28px; font-weight: 700;">S/ ${formatearNumeroExport(totalGPV_M1)}</div>
                    </div>
                    <div style="background: #f8fafc; padding: 15px 20px; border-radius: 16px; border: 1px solid #e2e8f0;">
                        <div style="color: #64748b; font-size: 13px; margin-bottom: 6px;">📈 GPV M2</div>
                        <div style="color: #0f172a; font-size: 28px; font-weight: 700;">S/ ${formatearNumeroExport(totalGPV_M2)}</div>
                    </div>
                    <div style="background: #f8fafc; padding: 15px 20px; border-radius: 16px; border: 1px solid #e2e8f0;">
                        <div style="color: #64748b; font-size: 13px; margin-bottom: 6px;">🛒 Transacciones Totales</div>
                        <div style="color: #0f172a; font-size: 28px; font-weight: 700;">${formatearNumeroExport(totalTRX, 0)}</div>
                        <div style="color: #94a3b8; font-size: 10px; margin-top: 5px;">M0: ${totalTRX_M0} | M1: ${totalTRX_M1} | M2: ${totalTRX_M2}</div>
                    </div>
                    <div style="background: #f8fafc; padding: 15px 20px; border-radius: 16px; border: 1px solid #e2e8f0;">
                        <div style="color: #64748b; font-size: 13px; margin-bottom: 6px;">📊 Distribución de Equipos</div>
                        <div style="display: flex; gap: 15px; margin-top: 8px;">
                            <div><span style="color: #22c55e;">✅</span> ${activos}</div>
                            <div><span style="color: #f97316;">⚠️</span> ${regulares}</div>
                            <div><span style="color: #ef4444;">❌</span> ${inactivos}</div>
                        </div>
                    </div>
                </div>
            </div>
        `;
    }
    
    if (incluirTabla) {
        contenidoHTML += `
            <div style="margin-top: 25px;">
                <h2 style="color: #1e293b; margin-bottom: 20px; font-size: 22px; font-weight: 600; border-left: 4px solid #0ea5e9; padding-left: 15px;">📋 Tabla Detallada de Equipos</h2>
                <table style="width: 100%; border-collapse: collapse; font-size: 8px; background: white;">
                    <thead>
                        <tr style="background: #ffffff; border-bottom: 2px solid #0ea5e9;">
                            <th style="padding: 8px 4px; border: 1px solid #ddd; text-align: center; font-weight: 700; color: #000000; background: #ffffff;">#</th>
                            <th style="padding: 8px 4px; border: 1px solid #ddd; text-align: center; font-weight: 700; color: #000000; background: #ffffff;">Fecha Venta</th>
                            <th style="padding: 8px 4px; border: 1px solid #ddd; text-align: center; font-weight: 700; color: #000000; background: #ffffff;">Mes Venta</th>
                            <th style="padding: 8px 4px; border: 1px solid #ddd; text-align: left; font-weight: 700; color: #000000; background: #ffffff;">Comercio</th>
                            <th style="padding: 8px 4px; border: 1px solid #ddd; text-align: left; font-weight: 700; color: #000000; background: #ffffff;">Serie</th>
                            <th style="padding: 8px 4px; border: 1px solid #ddd; text-align: left; font-weight: 700; color: #000000; background: #ffffff;">RUC</th>
                            <th style="padding: 8px 4px; border: 1px solid #ddd; text-align: left; font-weight: 700; color: #000000; background: #ffffff;">Ejecutivo</th>
                            <th style="padding: 8px 4px; border: 1px solid #ddd; text-align: right; font-weight: 700; color: #000000; background: #ffffff;">GPV M0</th>
                            <th style="padding: 8px 4px; border: 1px solid #ddd; text-align: right; font-weight: 700; color: #000000; background: #ffffff;">TRX M0</th>
                            <th style="padding: 8px 4px; border: 1px solid #ddd; text-align: right; font-weight: 700; color: #000000; background: #ffffff;">GPV M1</th>
                            <th style="padding: 8px 4px; border: 1px solid #ddd; text-align: right; font-weight: 700; color: #000000; background: #ffffff;">TRX M1</th>
                            <th style="padding: 8px 4px; border: 1px solid #ddd; text-align: right; font-weight: 700; color: #000000; background: #ffffff;">GPV M2</th>
                            <th style="padding: 8px 4px; border: 1px solid #ddd; text-align: right; font-weight: 700; color: #000000; background: #ffffff;">TRX M2</th>
                            <th style="padding: 8px 4px; border: 1px solid #ddd; text-align: right; font-weight: 700; color: #000000; background: #ffffff;">GPV Act</th>
                            <th style="padding: 8px 4px; border: 1px solid #ddd; text-align: center; font-weight: 700; color: #000000; background: #ffffff;">Estado</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        datos.forEach((item, i) => {
            const gpvActual = parseFloat(item.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            let estadoTexto = '';
            let estadoColor = '';
            
            if (estado === 'ACTIVO') {
                estadoTexto = '✅ Activo';
                estadoColor = '#22c55e';
            } else if (estado === 'REGULAR') {
                estadoTexto = '⚠️ Regular';
                estadoColor = '#f97316';
            } else {
                estadoTexto = '❌ Inactivo';
                estadoColor = '#ef4444';
            }
            
            const fechaVentaFormateada = formatearFechaExport(item.fecha_venta);
            const mesVenta = obtenerNombreMesExport(item.fecha_venta);
            
            // Alternar colores de fila para mejor legibilidad
            const bgColor = i % 2 === 0 ? '#ffffff' : '#f9fafb';
            
            contenidoHTML += `
                <tr style="background: ${bgColor};">
                    <td style="padding: 6px 4px; border: 1px solid #ddd; text-align: center; color: #1f2937;">${i + 1}</td>
                    <td style="padding: 6px 4px; border: 1px solid #ddd; text-align: center; color: #1f2937;">${fechaVentaFormateada}</td>
                    <td style="padding: 6px 4px; border: 1px solid #ddd; text-align: center; color: #1f2937; font-weight: 500;">${mesVenta}</td>
                    <td style="padding: 6px 4px; border: 1px solid #ddd; color: #1f2937; font-weight: 500;">${(item.comercio || '-').substring(0, 30)}</td>
                    <td style="padding: 6px 4px; border: 1px solid #ddd; color: #4b5563; font-family: monospace; font-size: 7px;">${item.numero_serie || '-'}</td>
                    <td style="padding: 6px 4px; border: 1px solid #ddd; color: #4b5563;">${item.ruc || '-'}</td>
                    <td style="padding: 6px 4px; border: 1px solid #ddd; color: #4b5563;">${item.responsable_real || '-'}</td>
                    <td style="padding: 6px 4px; border: 1px solid #ddd; text-align: right; color: #1f2937;">${formatearNumeroExport(item.gpv_m0 || 0, 0)}</td>
                    <td style="padding: 6px 4px; border: 1px solid #ddd; text-align: right; color: #1f2937; font-weight: ${(item.trx_m0 || 0) > 0 ? '600' : 'normal'};">${item.trx_m0 || 0}</td>
                    <td style="padding: 6px 4px; border: 1px solid #ddd; text-align: right; color: #1f2937;">${formatearNumeroExport(item.gpv_m1 || 0, 0)}</td>
                    <td style="padding: 6px 4px; border: 1px solid #ddd; text-align: right; color: #1f2937; font-weight: ${(item.trx_m1 || 0) > 0 ? '600' : 'normal'};">${item.trx_m1 || 0}</td>
                    <td style="padding: 6px 4px; border: 1px solid #ddd; text-align: right; color: #1f2937;">${formatearNumeroExport(item.gpv_m2 || 0, 0)}</td>
                    <td style="padding: 6px 4px; border: 1px solid #ddd; text-align: right; color: #1f2937; font-weight: ${(item.trx_m2 || 0) > 0 ? '600' : 'normal'};">${item.trx_m2 || 0}</td>
                    <td style="padding: 6px 4px; border: 1px solid #ddd; text-align: right; font-weight: 700; color: #0f172a;">${formatearNumeroExport(gpvActual, 0)}</td>
                    <td style="padding: 6px 4px; border: 1px solid #ddd; text-align: center; color: ${estadoColor}; font-weight: 600;">${estadoTexto}</td>
                </tr>
            `;
        });
        
        contenidoHTML += `
                    </tbody>
                </table>
                <div style="margin-top: 20px; padding: 12px; background: #f8fafc; border-radius: 12px; text-align: center; color: #64748b; font-size: 10px; border: 1px solid #e2e8f0;">
                    📊 Reporte generado el ${new Date().toLocaleString()} | Total: ${datos.length} registros
                </div>
            </div>
        `;
    }
    
    elementoPDF.innerHTML = contenidoHTML;
    document.body.appendChild(elementoPDF);
    
    // Cargar librerías para PDF
    const scriptHtml2canvas = document.createElement('script');
    scriptHtml2canvas.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js';
    
    const scriptJspdf = document.createElement('script');
    scriptJspdf.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
    
    let scriptsLoaded = 0;
    
    function checkAndGeneratePDF() {
        if (scriptsLoaded === 2 && window.html2canvas && window.jspdf) {
            html2canvas(elementoPDF, {
                scale: 2.5,
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
                
                pdf.save(`reporte_ventas_${new Date().toISOString().split('T')[0]}.pdf`);
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
    
    // Timeout por si falla la carga
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

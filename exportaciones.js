// ============================================================
// EXPORTACIONES.JS - Sistema completo de exportación
// Versión corregida - Colores fijos para PDF
// ============================================================

// Función auxiliar para obtener estado (debe coincidir con dashboard)
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

// Función principal que muestra el diálogo de exportación
function mostrarDialogoExportacion() {
    // Verificar que DATOS_COMPLETOS esté disponible
    if (typeof DATOS_COMPLETOS === 'undefined' || !DATOS_COMPLETOS) {
        alert('Error: No hay datos cargados para exportar');
        return;
    }
    
    // Obtener lista de ejecutivos únicos
    const ejecutivos = [...new Set(DATOS_COMPLETOS.equipos.map(e => e.responsable_real).filter(e => e))];
    const meses = DATOS_COMPLETOS.meses_disponibles || [];
    
    // Crear modal personalizado
    const modal = document.createElement('div');
    modal.id = 'modalExportacion';
    modal.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(0,0,0,0.85);
        backdrop-filter: blur(8px);
        z-index: 10000;
        display: flex;
        align-items: center;
        justify-content: center;
        font-family: 'Inter', system-ui, -apple-system, sans-serif;
    `;
    
    modal.innerHTML = `
        <div style="
            background: var(--superficie, #ffffff);
            border-radius: 28px;
            padding: 32px;
            max-width: 520px;
            width: 90%;
            max-height: 85vh;
            overflow-y: auto;
            border: 1px solid var(--borde, #e2e8f0);
            box-shadow: 0 25px 50px -12px rgba(0,0,0,0.25);
        ">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 24px;">
                <h2 style="font-size: 26px; font-weight: 700; margin: 0; background: linear-gradient(135deg, var(--acento, #0ea5e9), #7c3aed); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">
                    📥 Exportar Reporte
                </h2>
                <button onclick="cerrarModalExportacion()" style="
                    background: none;
                    border: none;
                    font-size: 28px;
                    cursor: pointer;
                    color: var(--texto-suave, #64748b);
                    padding: 4px 12px;
                    border-radius: 40px;
                    transition: all 0.2s;
                " onmouseover="this.style.background='var(--superficie2, #f1f5f9)'" onmouseout="this.style.background='none'">✕</button>
            </div>
            
            <div style="margin-bottom: 24px;">
                <label style="display: block; margin-bottom: 10px; font-weight: 600; color: var(--texto, #1f2937); font-size: 14px;">
                    📄 Formato de exportación:
                </label>
                <select id="formatoExportacion" style="
                    width: 100%;
                    padding: 14px 16px;
                    border-radius: 16px;
                    border: 1px solid var(--borde, #e2e8f0);
                    background: var(--superficie2, #f8fafc);
                    color: var(--texto, #1f2937);
                    font-size: 14px;
                    font-weight: 500;
                    cursor: pointer;
                    transition: all 0.2s;
                ">
                    <option value="excel">📊 Excel / CSV (Recomendado)</option>
                    <option value="pdf">📄 PDF (Para impresión)</option>
                    <option value="txt">📝 Texto Plano (TXT)</option>
                </select>
            </div>
            
            <div style="margin-bottom: 24px;">
                <label style="display: block; margin-bottom: 10px; font-weight: 600; color: var(--texto, #1f2937); font-size: 14px;">
                    🎯 Rango de datos:
                </label>
                <select id="rangoExportacion" style="
                    width: 100%;
                    padding: 14px 16px;
                    border-radius: 16px;
                    border: 1px solid var(--borde, #e2e8f0);
                    background: var(--superficie2, #f8fafc);
                    color: var(--texto, #1f2937);
                    font-size: 14px;
                    font-weight: 500;
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
                        padding: 14px 16px;
                        border-radius: 16px;
                        border: 1px solid var(--borde, #e2e8f0);
                        background: var(--superficie2, #f8fafc);
                        color: var(--texto, #1f2937);
                        font-size: 14px;
                        font-weight: 500;
                        cursor: pointer;
                    ">
                        <option value="">📌 Seleccionar ejecutivo...</option>
                        ${ejecutivos.map(exec => `<option value="${exec}">👤 ${exec}</option>`).join('')}
                    </select>
                </div>
                
                <div id="opcionMes" style="display: none; margin-top: 12px;">
                    <select id="mesSeleccionado" style="
                        width: 100%;
                        padding: 14px 16px;
                        border-radius: 16px;
                        border: 1px solid var(--borde, #e2e8f0);
                        background: var(--superficie2, #f8fafc);
                        color: var(--texto, #1f2937);
                        font-size: 14px;
                        font-weight: 500;
                        cursor: pointer;
                    ">
                        <option value="">📅 Seleccionar mes...</option>
                        ${meses.map(mes => `<option value="${mes}">📆 ${mes}</option>`).join('')}
                    </select>
                </div>
            </div>
            
            <div style="margin-bottom: 28px;">
                <label style="display: block; margin-bottom: 10px; font-weight: 600; color: var(--texto, #1f2937); font-size: 14px;">
                    📊 Información adicional:
                </label>
                <div style="display: flex; gap: 24px; flex-wrap: wrap;">
                    <label style="display: flex; align-items: center; gap: 10px; cursor: pointer;">
                        <input type="checkbox" id="incluirResumen" checked style="width: 18px; height: 18px; cursor: pointer;">
                        <span style="color: var(--texto, #1f2937);">📈 Incluir resumen (KPIs)</span>
                    </label>
                    <label style="display: flex; align-items: center; gap: 10px; cursor: pointer;">
                        <input type="checkbox" id="incluirTabla" checked style="width: 18px; height: 18px; cursor: pointer;">
                        <span style="color: var(--texto, #1f2937);">📋 Incluir tabla detallada</span>
                    </label>
                </div>
            </div>
            
            <div style="display: flex; gap: 12px; justify-content: flex-end; border-top: 1px solid var(--borde, #e2e8f0); padding-top: 24px;">
                <button onclick="cerrarModalExportacion()" style="
                    padding: 12px 28px;
                    border-radius: 40px;
                    border: 1px solid var(--borde, #e2e8f0);
                    background: var(--superficie2, #f1f5f9);
                    color: var(--texto, #1f2937);
                    cursor: pointer;
                    font-weight: 600;
                    font-size: 14px;
                    transition: all 0.2s;
                " onmouseover="this.style.transform='translateY(-1px)'" onmouseout="this.style.transform='translateY(0)'">
                    Cancelar
                </button>
                <button onclick="ejecutarExportacion()" style="
                    padding: 12px 32px;
                    border-radius: 40px;
                    border: none;
                    background: linear-gradient(135deg, var(--acento, #0ea5e9), #7c3aed);
                    color: white;
                    cursor: pointer;
                    font-weight: 600;
                    font-size: 14px;
                    transition: all 0.2s;
                    box-shadow: 0 4px 12px rgba(14,165,233,0.3);
                " onmouseover="this.style.transform='translateY(-1px)'; this.style.boxShadow='0 6px 16px rgba(14,165,233,0.4)'" onmouseout="this.style.transform='translateY(0)'; this.style.boxShadow='0 4px 12px rgba(14,165,233,0.3)'">
                    Exportar
                </button>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    // Mostrar/ocultar opciones según selección
    const rangoSelect = document.getElementById('rangoExportacion');
    const opcionEjecutivo = document.getElementById('opcionEjecutivo');
    const opcionMes = document.getElementById('opcionMes');
    
    if (rangoSelect) {
        rangoSelect.addEventListener('change', (e) => {
            opcionEjecutivo.style.display = e.target.value === 'ejecutivo' ? 'block' : 'none';
            opcionMes.style.display = e.target.value === 'mes' ? 'block' : 'none';
        });
    }
}

function cerrarModalExportacion() {
    const modal = document.getElementById('modalExportacion');
    if (modal) modal.remove();
}

function obtenerDatosConFiltrosAplicados() {
    // Usar variables globales del dashboard
    let equipos = [...(window.DATOS_COMPLETOS?.equipos || [])];
    
    // Aplicar filtros actuales del dashboard
    if (typeof window.mesActual !== 'undefined' && window.mesActual !== 'TODOS') {
        equipos = equipos.filter(e => e.mes_venta === window.mesActual);
    }
    if (typeof window.ejecutivoActual !== 'undefined' && window.ejecutivoActual !== 'TODOS') {
        equipos = equipos.filter(e => e.responsable_real === window.ejecutivoActual);
    }
    if (typeof window.filtroGPVActual !== 'undefined' && window.filtroGPVActual !== 'todos') {
        equipos = equipos.filter(e => {
            const gpvActual = parseFloat(e.gpv_mes_actual_corriendo) || 0;
            switch(window.filtroGPVActual) {
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
            datos = [...(window.DATOS_COMPLETOS?.equipos || [])];
            break;
        case 'ejecutivo':
            const ejecutivo = document.getElementById('ejecutivoSeleccionado').value;
            if (!ejecutivo) {
                alert('⚠️ Por favor selecciona un ejecutivo');
                return;
            }
            datos = (window.DATOS_COMPLETOS?.equipos || []).filter(e => e.responsable_real === ejecutivo);
            break;
        case 'mes':
            const mes = document.getElementById('mesSeleccionado').value;
            if (!mes) {
                alert('⚠️ Por favor selecciona un mes');
                return;
            }
            datos = (window.DATOS_COMPLETOS?.equipos || []).filter(e => e.mes_venta === mes);
            break;
    }
    
    if (datos.length === 0) {
        alert('⚠️ No hay datos para exportar con los filtros seleccionados');
        return;
    }
    
    // Cerrar modal
    cerrarModalExportacion();
    
    // Mostrar indicador de carga
    mostrarMensajeCarga(`Exportando ${datos.length} registros a ${formato.toUpperCase()}...`);
    
    // Exportar según formato
    setTimeout(() => {
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
        ocultarMensajeCarga();
    }, 100);
}

function mostrarMensajeCarga(mensaje) {
    const div = document.createElement('div');
    div.id = 'mensajeCargaExport';
    div.style.cssText = `
        position: fixed;
        bottom: 30px;
        right: 30px;
        background: linear-gradient(135deg, #0ea5e9, #7c3aed);
        color: white;
        padding: 12px 24px;
        border-radius: 50px;
        font-weight: 600;
        font-size: 14px;
        z-index: 10001;
        box-shadow: 0 4px 20px rgba(0,0,0,0.2);
        animation: slideIn 0.3s ease;
    `;
    div.innerHTML = `⏳ ${mensaje}`;
    document.body.appendChild(div);
}

function ocultarMensajeCarga() {
    const div = document.getElementById('mensajeCargaExport');
    if (div) div.remove();
}

// ============================================================
// EXPORTAR EXCEL (CSV) - CORREGIDO
// ============================================================
function exportarExcelCompleto(datos, incluirResumen, incluirTabla) {
    const nombreArchivo = `reporte_ventas_${new Date().toISOString().split('T')[0]}`;
    let lineas = [];
    
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
        
        lineas.push('=== RESUMEN DEL REPORTE ===');
        lineas.push(`"Fecha de exportación","${new Date().toLocaleString()}"`);
        lineas.push(`"Total de registros",${datos.length}`);
        lineas.push(`"GPV Total",S/ ${formatearNumeroExport(totalGPV)}`);
        lineas.push(`"GPV M0",S/ ${formatearNumeroExport(totalGPV_M0)}`);
        lineas.push(`"GPV M1",S/ ${formatearNumeroExport(totalGPV_M1)}`);
        lineas.push(`"GPV M2",S/ ${formatearNumeroExport(totalGPV_M2)}`);
        lineas.push(`"Transacciones Totales",${formatearNumeroExport(totalTRX, 0)}`);
        lineas.push(`"Equipos Activos",${activos}`);
        lineas.push(`"Equipos Regulares",${regulares}`);
        lineas.push(`"Equipos Inactivos",${inactivos}`);
        lineas.push('');
        lineas.push('');
    }
    
    if (incluirTabla) {
        lineas.push('=== TABLA DETALLADA ===');
        lineas.push('');
        
        const headers = [
            '#', 'Fecha Venta', 'Comercio', 'Serie', 'RUC', 
            'Fecha Activación', 'Ejecutivo', 'GPV M0', 'TRX M0', 
            'GPV M1', 'TRX M1', 'GPV M2', 'TRX M2', 'Mes Actual', 
            'GPV Actual', 'TRX Actual', 'Última Transacción', 'Estado'
        ];
        
        lineas.push(headers.map(h => `"${h}"`).join(','));
        
        datos.forEach((item, i) => {
            const gpvActual = parseFloat(item.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            const estadoTexto = estado === 'ACTIVO' ? 'Activo' : (estado === 'REGULAR' ? 'Regular' : 'Inactivo');
            
            const row = [
                i + 1,
                item.fecha_venta || '',
                `"${(item.comercio || '').replace(/"/g, '""')}"`,
                item.numero_serie || '',
                item.ruc || '',
                item.dia_activo || '',
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
                item.ultima_transaccion || '',
                estadoTexto
            ];
            lineas.push(row.join(','));
        });
    }
    
    const blob = new Blob(['\ufeff' + lineas.join('\n')], { type: 'text/csv;charset=utf-8;' });
    descargarArchivo(blob, `${nombreArchivo}.csv`);
    mostrarMensajeExito(`${datos.length} registros exportados a CSV`);
}

// ============================================================
// EXPORTAR TXT - CORREGIDO
// ============================================================
function exportarTXTCompleto(datos, incluirResumen, incluirTabla) {
    const nombreArchivo = `reporte_ventas_${new Date().toISOString().split('T')[0]}`;
    let contenido = [];
    const separador = '═'.repeat(100);
    const separadorLinea = '─'.repeat(100);
    
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
        contenido.push('                    📊 REPORTE DE VENTAS - RESUMEN GENERAL');
        contenido.push(separador);
        contenido.push('');
        contenido.push(`📅 Fecha de exportación: ${new Date().toLocaleString()}`);
        contenido.push(`📋 Total de registros: ${datos.length}`);
        contenido.push('');
        contenido.push('┌─────────────────────────────────────────────────────────────────────────────┐');
        contenido.push('│                         📈 INDICADORES PRINCIPALES                         │');
        contenido.push('├─────────────────────────────────────────────────────────────────────────────┤');
        contenido.push(`│   💰 GPV Total (Cohorte):     S/ ${formatearNumeroExport(totalGPV).padStart(20)}               │`);
        contenido.push(`│   🧪 GPV M0 (Prueba):         S/ ${formatearNumeroExport(totalGPV_M0).padStart(20)}               │`);
        contenido.push(`│   🚀 GPV M1:                  S/ ${formatearNumeroExport(totalGPV_M1).padStart(20)}               │`);
        contenido.push(`│   📈 GPV M2:                  S/ ${formatearNumeroExport(totalGPV_M2).padStart(20)}               │`);
        contenido.push(`│   🛒 Transacciones Totales:   ${formatearNumeroExport(totalTRX, 0).padStart(20)}               │`);
        contenido.push('├─────────────────────────────────────────────────────────────────────────────┤');
        contenido.push('│                         📊 ESTADO DE EQUIPOS                                 │');
        contenido.push('├─────────────────────────────────────────────────────────────────────────────┤');
        contenido.push(`│   ✅ Equipos Activos:         ${activos.toString().padStart(20)}               │`);
        contenido.push(`│   ⚠️  Equipos Regulares:       ${regulares.toString().padStart(20)}               │`);
        contenido.push(`│   ❌ Equipos Inactivos:       ${inactivos.toString().padStart(20)}               │`);
        contenido.push('└─────────────────────────────────────────────────────────────────────────────┘');
        contenido.push('');
    }
    
    if (incluirTabla) {
        contenido.push(separador);
        contenido.push('                    📋 TABLA DETALLADA DE EQUIPOS');
        contenido.push(separador);
        contenido.push('');
        
        // Función para formatear texto con ancho fijo
        const pad = (texto, ancho) => {
            const str = String(texto === null || texto === undefined ? '-' : texto);
            return str.length > ancho ? str.substring(0, ancho - 3) + '...' : str.padEnd(ancho);
        };
        
        // Encabezados
        const headerLine = 
            pad('#', 5) +
            pad('Fecha Venta', 12) +
            pad('Comercio', 30) +
            pad('Serie', 16) +
            pad('RUC', 12) +
            pad('Ejecutivo', 20) +
            pad('GPV M0', 10) +
            pad('TRX M0', 8) +
            pad('GPV M1', 10) +
            pad('TRX M1', 8) +
            pad('GPV M2', 10) +
            pad('TRX M2', 8) +
            pad('GPV Act', 10) +
            pad('Estado', 10);
        
        contenido.push(headerLine);
        contenido.push('─'.repeat(headerLine.length));
        
        datos.forEach((item, i) => {
            const gpvActual = parseFloat(item.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            const estadoTexto = estado === 'ACTIVO' ? 'Activo' : (estado === 'REGULAR' ? 'Regular' : 'Inactivo');
            const estadoIcono = estado === 'ACTIVO' ? '✅' : (estado === 'REGULAR' ? '⚠️' : '❌');
            
            contenido.push(
                pad(i + 1, 5) +
                pad(item.fecha_venta || '-', 12) +
                pad(item.comercio || '-', 30) +
                pad(item.numero_serie || '-', 16) +
                pad(item.ruc || '-', 12) +
                pad(item.responsable_real || '-', 20) +
                pad(formatearNumeroExport(item.gpv_m0 || 0, 0), 10) +
                pad(item.trx_m0 || 0, 8) +
                pad(formatearNumeroExport(item.gpv_m1 || 0, 0), 10) +
                pad(item.trx_m1 || 0, 8) +
                pad(formatearNumeroExport(item.gpv_m2 || 0, 0), 10) +
                pad(item.trx_m2 || 0, 8) +
                pad(formatearNumeroExport(gpvActual, 0), 10) +
                pad(`${estadoIcono} ${estadoTexto}`, 10)
            );
        });
        
        contenido.push('');
        contenido.push(separadorLinea);
        contenido.push(`📊 Total de registros mostrados: ${datos.length}`);
    }
    
    const blob = new Blob([contenido.join('\n')], { type: 'text/plain;charset=utf-8;' });
    descargarArchivo(blob, `${nombreArchivo}.txt`);
    mostrarMensajeExito(`${datos.length} registros exportados a TXT`);
}

// ============================================================
// EXPORTAR PDF - VERSIÓN CORREGIDA (COLORES FIJOS)
// ============================================================
function exportarPDFCompleto(datos, incluirResumen, incluirTabla) {
    // Crear un elemento temporal para renderizar el PDF con estilos fijos
    const elementoPDF = document.createElement('div');
    elementoPDF.style.cssText = `
        position: absolute;
        top: -9999px;
        left: -9999px;
        width: 1100px;
        background: white;
        padding: 40px 50px;
        font-family: 'Segoe UI', 'Inter', Arial, sans-serif;
        color: #1e293b;
        line-height: 1.5;
    `;
    
    let contenidoHTML = `
        <div style="margin-bottom: 35px; text-align: center; border-bottom: 3px solid #0ea5e9; padding-bottom: 20px;">
            <h1 style="color: #0ea5e9; margin-bottom: 12px; font-size: 32px; font-weight: 700;">📊 Reporte de Ventas</h1>
            <p style="color: #475569; margin: 6px 0; font-size: 14px;">Fecha de exportación: ${new Date().toLocaleString()}</p>
            <p style="color: #475569; margin: 6px 0; font-size: 14px; font-weight: 500;">Total de registros: ${datos.length}</p>
        </div>
    `;
    
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
        
        contenidoHTML += `
            <div style="margin-bottom: 35px;">
                <h2 style="color: #1e293b; margin-bottom: 20px; font-size: 22px; font-weight: 600; border-left: 4px solid #0ea5e9; padding-left: 15px;">📈 Resumen General</h2>
                <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px;">
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
                <table style="width: 100%; border-collapse: collapse; font-size: 10px; background: white;">
                    <thead>
                        <tr style="background: #f1f5f9;">
                            <th style="padding: 10px 8px; border: 1px solid #e2e8f0; text-align: center; font-weight: 600; color: #1e293b;">#</th>
                            <th style="padding: 10px 8px; border: 1px solid #e2e8f0; text-align: left; font-weight: 600; color: #1e293b;">Fecha Venta</th>
                            <th style="padding: 10px 8px; border: 1px solid #e2e8f0; text-align: left; font-weight: 600; color: #1e293b;">Comercio</th>
                            <th style="padding: 10px 8px; border: 1px solid #e2e8f0; text-align: left; font-weight: 600; color: #1e293b;">Serie</th>
                            <th style="padding: 10px 8px; border: 1px solid #e2e8f0; text-align: left; font-weight: 600; color: #1e293b;">RUC</th>
                            <th style="padding: 10px 8px; border: 1px solid #e2e8f0; text-align: left; font-weight: 600; color: #1e293b;">Ejecutivo</th>
                            <th style="padding: 10px 8px; border: 1px solid #e2e8f0; text-align: right; font-weight: 600; color: #1e293b;">GPV M0</th>
                            <th style="padding: 10px 8px; border: 1px solid #e2e8f0; text-align: right; font-weight: 600; color: #1e293b;">GPV M1</th>
                            <th style="padding: 10px 8px; border: 1px solid #e2e8f0; text-align: right; font-weight: 600; color: #1e293b;">GPV M2</th>
                            <th style="padding: 10px 8px; border: 1px solid #e2e8f0; text-align: right; font-weight: 600; color: #1e293b;">GPV Act</th>
                            <th style="padding: 10px 8px; border: 1px solid #e2e8f0; text-align: center; font-weight: 600; color: #1e293b;">Estado</th>
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
            
            // Alternar colores de fila
            const bgColor = i % 2 === 0 ? '#ffffff' : '#fafafa';
            
            contenidoHTML += `
                <tr style="background: ${bgColor};">
                    <td style="padding: 8px 6px; border: 1px solid #e2e8f0; text-align: center; color: #475569;">${i + 1}</td>
                    <td style="padding: 8px 6px; border: 1px solid #e2e8f0; color: #475569;">${item.fecha_venta || '-'}</td>
                    <td style="padding: 8px 6px; border: 1px solid #e2e8f0; color: #1e293b; font-weight: 500;">${(item.comercio || '-').substring(0, 30)}</td>
                    <td style="padding: 8px 6px; border: 1px solid #e2e8f0; color: #475569; font-family: monospace; font-size: 9px;">${item.numero_serie || '-'}</td>
                    <td style="padding: 8px 6px; border: 1px solid #e2e8f0; color: #475569;">${item.ruc || '-'}</td>
                    <td style="padding: 8px 6px; border: 1px solid #e2e8f0; color: #475569;">${item.responsable_real || '-'}</td>
                    <td style="padding: 8px 6px; border: 1px solid #e2e8f0; text-align: right; color: #475569;">${formatearNumeroExport(item.gpv_m0 || 0, 0)}</td>
                    <td style="padding: 8px 6px; border: 1px solid #e2e8f0; text-align: right; color: #475569;">${formatearNumeroExport(item.gpv_m1 || 0, 0)}</td>
                    <td style="padding: 8px 6px; border: 1px solid #e2e8f0; text-align: right; color: #475569;">${formatearNumeroExport(item.gpv_m2 || 0, 0)}</td>
                    <td style="padding: 8px 6px; border: 1px solid #e2e8f0; text-align: right; font-weight: 600; color: #0f172a;">${formatearNumeroExport(gpvActual, 0)}</td>
                    <td style="padding: 8px 6px; border: 1px solid #e2e8f0; text-align: center; color: ${estadoColor}; font-weight: 600;">${estadoTexto}</td>
                </tr>
            `;
        });
        
        contenidoHTML += `
                    </tbody>
                </table>
                <div style="margin-top: 20px; padding: 12px; background: #f8fafc; border-radius: 12px; text-align: center; color: #64748b; font-size: 11px;">
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
                mostrarMensajeExito(`${datos.length} registros exportados a PDF`);
            }).catch(err => {
                console.error('Error generando PDF:', err);
                alert('Error al generar el PDF. Por favor intenta con formato CSV o TXT.');
                document.body.removeChild(elementoPDF);
                ocultarMensajeCarga();
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
            ocultarMensajeCarga();
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

function mostrarMensajeExito(mensaje) {
    const div = document.createElement('div');
    div.style.cssText = `
        position: fixed;
        bottom: 30px;
        right: 30px;
        background: #22c55e;
        color: white;
        padding: 12px 24px;
        border-radius: 50px;
        font-weight: 600;
        font-size: 14px;
        z-index: 10001;
        box-shadow: 0 4px 20px rgba(0,0,0,0.2);
        animation: slideIn 0.3s ease;
    `;
    div.innerHTML = `✅ ${mensaje}`;
    document.body.appendChild(div);
    setTimeout(() => {
        if (div) div.remove();
    }, 3000);
}

// Agregar animación CSS si no existe
if (!document.querySelector('#exportAnimationStyle')) {
    const style = document.createElement('style');
    style.id = 'exportAnimationStyle';
    style.textContent = `
        @keyframes slideIn {
            from {
                transform: translateX(100%);
                opacity: 0;
            }
            to {
                transform: translateX(0);
                opacity: 1;
            }
        }
    `;
    document.head.appendChild(style);
}

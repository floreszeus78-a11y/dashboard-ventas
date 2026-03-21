// ============================================================
// EXPORTACIONES.JS - Sistema completo de exportación
// ============================================================

// Función principal que muestra el diálogo de exportación
function mostrarDialogoExportacion() {
    // Obtener datos actuales con filtros aplicados
    const datosActuales = obtenerDatosConFiltrosAplicados();
    
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
                    <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
                        <input type="checkbox" id="incluirGraficas" checked>
                        <span>Incluir gráficas (solo PDF)</span>
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
    if (mesActual !== 'TODOS') {
        equipos = equipos.filter(e => e.mes_venta === mesActual);
    }
    if (ejecutivoActual !== 'TODOS') {
        equipos = equipos.filter(e => e.responsable_real === ejecutivoActual);
    }
    if (filtroGPVActual !== 'todos') {
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
    const incluirGraficas = document.getElementById('incluirGraficas').checked;
    
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
            exportarPDFCompleto(datos, incluirResumen, incluirTabla, incluirGraficas);
            break;
        case 'txt':
            exportarTXTCompleto(datos, incluirResumen, incluirTabla);
            break;
    }
}

// ============================================================
// EXPORTAR EXCEL (CSV con múltiples hojas)
// ============================================================
function exportarExcelCompleto(datos, incluirResumen, incluirTabla) {
    const nombreArchivo = `reporte_ventas_${new Date().toISOString().split('T')[0]}`;
    
    if (incluirResumen && incluirTabla) {
        // Crear CSV con resumen y tabla
        let contenido = [];
        
        // Resumen
        contenido.push('=== RESUMEN DEL REPORTE ===');
        contenido.push('');
        contenido.push(`Fecha de exportación: ${new Date().toLocaleString()}`);
        contenido.push(`Total de registros: ${datos.length}`);
        
        // Calcular totales
        let totalGPV = 0;
        let totalTRX = 0;
        let activos = 0;
        
        datos.forEach(e => {
            totalGPV += (parseFloat(e.gpv_m0) || 0) + (parseFloat(e.gpv_m1) || 0) + (parseFloat(e.gpv_m2) || 0);
            totalTRX += (parseInt(e.trx_m0) || 0) + (parseInt(e.trx_m1) || 0) + (parseInt(e.trx_m2) || 0);
            const gpvActual = parseFloat(e.gpv_mes_actual_corriendo) || 0;
            if (getEstadoPorGPV(gpvActual) === 'ACTIVO') activos++;
        });
        
        contenido.push(`GPV Total: S/ ${totalGPV.toLocaleString('es-PE')}`);
        contenido.push(`TRX Total: ${totalTRX.toLocaleString('es-PE')}`);
        contenido.push(`Equipos Activos: ${activos}/${datos.length}`);
        contenido.push('');
        contenido.push('');
        
        // Tabla detallada
        contenido.push('=== TABLA DETALLADA ===');
        contenido.push('');
        
        const headers = [
            '#', 'Fecha Venta', 'Comercio', 'Serie', 'RUC', 
            'Fecha Activación', 'Ejecutivo', 'GPV M0', 'TRX M0', 
            'GPV M1', 'TRX M1', 'GPV M2', 'TRX M2', 'Mes Actual', 
            'GPV Actual', 'TRX Actual', 'Última Transacción', 'Estado'
        ];
        
        contenido.push(headers.join(','));
        
        datos.forEach((item, i) => {
            const gpvActual = parseFloat(item.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            
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
                estado === 'ACTIVO' ? 'Activo' : (estado === 'REGULAR' ? 'Regular' : 'Inactivo')
            ];
            contenido.push(row.join(','));
        });
        
        const blob = new Blob(['\ufeff' + contenido.join('\n')], { type: 'text/csv;charset=utf-8;' });
        descargarArchivo(blob, `${nombreArchivo}.csv`);
        
    } else if (incluirTabla) {
        // Solo tabla
        const headers = [
            '#', 'Fecha Venta', 'Comercio', 'Serie', 'RUC', 
            'Fecha Activación', 'Ejecutivo', 'GPV M0', 'TRX M0', 
            'GPV M1', 'TRX M1', 'GPV M2', 'TRX M2', 'Mes Actual', 
            'GPV Actual', 'TRX Actual', 'Última Transacción', 'Estado'
        ];
        
        const rows = [headers.join(',')];
        
        datos.forEach((item, i) => {
            const gpvActual = parseFloat(item.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            
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
                estado === 'ACTIVO' ? 'Activo' : (estado === 'REGULAR' ? 'Regular' : 'Inactivo')
            ];
            rows.push(row.join(','));
        });
        
        const blob = new Blob(['\ufeff' + rows.join('\n')], { type: 'text/csv;charset=utf-8;' });
        descargarArchivo(blob, `${nombreArchivo}_tabla.csv`);
    }
}

// ============================================================
// EXPORTAR TXT (formato legible)
// ============================================================
function exportarTXTCompleto(datos, incluirResumen, incluirTabla) {
    const nombreArchivo = `reporte_ventas_${new Date().toISOString().split('T')[0]}`;
    let contenido = [];
    
    // Línea separadora
    const separador = '='.repeat(100);
    
    if (incluirResumen) {
        contenido.push(separador);
        contenido.push('RESUMEN DEL REPORTE');
        contenido.push(separador);
        contenido.push('');
        contenido.push(`Fecha de exportación: ${new Date().toLocaleString()}`);
        contenido.push(`Total de registros: ${datos.length}`);
        contenido.push('');
        
        // Calcular totales
        let totalGPV = 0;
        let totalGPV_M0 = 0;
        let totalGPV_M1 = 0;
        let totalGPV_M2 = 0;
        let totalTRX = 0;
        let activos = 0;
        let regulares = 0;
        let inactivos = 0;
        
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
        
        // Encabezado formateado
        const formatoEncabezado = (texto, ancho) => texto.padEnd(ancho);
        contenido.push(
            formatoEncabezado('#', 5) +
            formatoEncabezado('Fecha Venta', 12) +
            formatoEncabezado('Comercio', 30) +
            formatoEncabezado('Serie', 15) +
            formatoEncabezado('RUC', 12) +
            formatoEncabezado('Ejecutivo', 20) +
            formatoEncabezado('GPV M0', 10) +
            formatoEncabezado('TRX M0', 8) +
            formatoEncabezado('GPV M1', 10) +
            formatoEncabezado('TRX M1', 8) +
            formatoEncabezado('GPV M2', 10) +
            formatoEncabezado('TRX M2', 8) +
            formatoEncabezado('GPV Act', 10) +
            formatoEncabezado('Estado', 10)
        );
        contenido.push('-'.repeat(180));
        
        datos.forEach((item, i) => {
            const gpvActual = parseFloat(item.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            const estadoTexto = estado === 'ACTIVO' ? 'Activo' : (estado === 'REGULAR' ? 'Regular' : 'Inactivo');
            
            contenido.push(
                formatoEncabezado(String(i + 1), 5) +
                formatoEncabezado(item.fecha_venta || '-', 12) +
                formatoEncabezado((item.comercio || '-').substring(0, 28), 30) +
                formatoEncabezado(item.numero_serie || '-', 15) +
                formatoEncabezado(item.ruc || '-', 12) +
                formatoEncabezado((item.responsable_real || '-').substring(0, 18), 20) +
                formatoEncabezado(String(item.gpv_m0 || 0), 10) +
                formatoEncabezado(String(item.trx_m0 || 0), 8) +
                formatoEncabezado(String(item.gpv_m1 || 0), 10) +
                formatoEncabezado(String(item.trx_m1 || 0), 8) +
                formatoEncabezado(String(item.gpv_m2 || 0), 10) +
                formatoEncabezado(String(item.trx_m2 || 0), 8) +
                formatoEncabezado(String(gpvActual), 10) +
                formatoEncabezado(estadoTexto, 10)
            );
        });
    }
    
    const blob = new Blob([contenido.join('\n')], { type: 'text/plain;charset=utf-8;' });
    descargarArchivo(blob, `${nombreArchivo}.txt`);
}

// ============================================================
// EXPORTAR PDF (usando html2canvas + jsPDF)
// ============================================================
function exportarPDFCompleto(datos, incluirResumen, incluirTabla, incluirGraficas) {
    // Crear un elemento temporal para renderizar el PDF
    const elementoPDF = document.createElement('div');
    elementoPDF.style.cssText = `
        position: absolute;
        top: -9999px;
        left: -9999px;
        width: 1200px;
        background: white;
        padding: 40px;
        font-family: 'Inter', sans-serif;
        color: #000;
    `;
    
    let contenidoHTML = `
        <div style="margin-bottom: 30px; text-align: center;">
            <h1 style="color: #0ea5e9; margin-bottom: 10px;">📊 Reporte de Ventas</h1>
            <p style="color: #666;">Fecha de exportación: ${new Date().toLocaleString()}</p>
            <p style="color: #666;">Total de registros: ${datos.length}</p>
        </div>
    `;
    
    if (incluirResumen) {
        // Calcular totales
        let totalGPV = 0;
        let totalGPV_M0 = 0;
        let totalGPV_M1 = 0;
        let totalGPV_M2 = 0;
        let totalTRX = 0;
        let activos = 0;
        let regulares = 0;
        let inactivos = 0;
        
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
            <div style="margin-bottom: 30px;">
                <h2 style="color: #333; margin-bottom: 15px;">📈 Resumen General</h2>
                <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
                    <tr style="background: #f0f0f0;">
                        <th style="padding: 10px; border: 1px solid #ddd; text-align: left;">Métrica</th>
                        <th style="padding: 10px; border: 1px solid #ddd; text-align: right;">Valor</th>
                    </tr>
                    <tr><td style="padding: 8px; border: 1px solid #ddd;">💰 GPV Total Cohortes</td><td style="padding: 8px; border: 1px solid #ddd; text-align: right;">S/ ${totalGPV.toLocaleString('es-PE')}</td></tr>
                    <tr><td style="padding: 8px; border: 1px solid #ddd;">🧪 GPV M0 (Prueba)</td><td style="padding: 8px; border: 1px solid #ddd; text-align: right;">S/ ${totalGPV_M0.toLocaleString('es-PE')}</td></tr>
                    <tr><td style="padding: 8px; border: 1px solid #ddd;">🚀 GPV M1</td><td style="padding: 8px; border: 1px solid #ddd; text-align: right;">S/ ${totalGPV_M1.toLocaleString('es-PE')}</td></tr>
                    <tr><td style="padding: 8px; border: 1px solid #ddd;">📈 GPV M2</td><td style="padding: 8px; border: 1px solid #ddd; text-align: right;">S/ ${totalGPV_M2.toLocaleString('es-PE')}</td></tr>
                    <tr><td style="padding: 8px; border: 1px solid #ddd;">🛒 Transacciones Totales</td><td style="padding: 8px; border: 1px solid #ddd; text-align: right;">${totalTRX.toLocaleString('es-PE')}</td></tr>
                    <tr><td style="padding: 8px; border: 1px solid #ddd;">✅ Equipos Activos</td><td style="padding: 8px; border: 1px solid #ddd; text-align: right;">${activos}</td></tr>
                    <tr><td style="padding: 8px; border: 1px solid #ddd;">⚠️ Equipos Regulares</td><td style="padding: 8px; border: 1px solid #ddd; text-align: right;">${regulares}</td></tr>
                    <tr><td style="padding: 8px; border: 1px solid #ddd;">❌ Equipos Inactivos</td><td style="padding: 8px; border: 1px solid #ddd; text-align: right;">${inactivos}</td></tr>
                </table>
            </div>
        `;
    }
    
    if (incluirTabla) {
        contenidoHTML += `
            <div>
                <h2 style="color: #333; margin-bottom: 15px;">📋 Tabla Detallada de Equipos</h2>
                <table style="width: 100%; border-collapse: collapse; font-size: 10px;">
                    <thead>
                        <tr style="background: #f0f0f0;">
                            <th style="padding: 8px; border: 1px solid #ddd;">#</th>
                            <th style="padding: 8px; border: 1px solid #ddd;">Fecha Venta</th>
                            <th style="padding: 8px; border: 1px solid #ddd;">Comercio</th>
                            <th style="padding: 8px; border: 1px solid #ddd;">Serie</th>
                            <th style="padding: 8px; border: 1px solid #ddd;">RUC</th>
                            <th style="padding: 8px; border: 1px solid #ddd;">Ejecutivo</th>
                            <th style="padding: 8px; border: 1px solid #ddd;">GPV M0</th>
                            <th style="padding: 8px; border: 1px solid #ddd;">GPV M1</th>
                            <th style="padding: 8px; border: 1px solid #ddd;">GPV M2</th>
                            <th style="padding: 8px; border: 1px solid #ddd;">GPV Act</th>
                            <th style="padding: 8px; border: 1px solid #ddd;">Estado</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        datos.forEach((item, i) => {
            const gpvActual = parseFloat(item.gpv_mes_actual_corriendo) || 0;
            const estado = getEstadoPorGPV(gpvActual);
            const estadoTexto = estado === 'ACTIVO' ? 'Activo' : (estado === 'REGULAR' ? 'Regular' : 'Inactivo');
            const colorEstado = estado === 'ACTIVO' ? '#22c55e' : (estado === 'REGULAR' ? '#f97316' : '#ef4444');
            
            contenidoHTML += `
                <tr>
                    <td style="padding: 6px; border: 1px solid #ddd;">${i + 1}</td>
                    <td style="padding: 6px; border: 1px solid #ddd;">${item.fecha_venta || '-'}</td>
                    <td style="padding: 6px; border: 1px solid #ddd;">${(item.comercio || '-').substring(0, 25)}</td>
                    <td style="padding: 6px; border: 1px solid #ddd;">${item.numero_serie || '-'}</td>
                    <td style="padding: 6px; border: 1px solid #ddd;">${item.ruc || '-'}</td>
                    <td style="padding: 6px; border: 1px solid #ddd;">${item.responsable_real || '-'}</td>
                    <td style="padding: 6px; border: 1px solid #ddd; text-align: right;">${item.gpv_m0 || 0}</td>
                    <td style="padding: 6px; border: 1px solid #ddd; text-align: right;">${item.gpv_m1 || 0}</td>
                    <td style="padding: 6px; border: 1px solid #ddd; text-align: right;">${item.gpv_m2 || 0}</td>
                    <td style="padding: 6px; border: 1px solid #ddd; text-align: right;">${gpvActual}</td>
                    <td style="padding: 6px; border: 1px solid #ddd; color: ${colorEstado}; font-weight: bold;">${estadoTexto}</td>
                </tr>
            `;
        });
        
        contenidoHTML += `
                    </tbody>
                </table>
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
    
    scriptHtml2canvas.onload = () => {
        scriptJspdf.onload = () => {
            html2canvas(elementoPDF, {
                scale: 2,
                useCORS: true,
                logging: false
            }).then(canvas => {
                const imgData = canvas.toDataURL('image/png');
                const { jsPDF } = window.jspdf;
                const pdf = new jsPDF('p', 'mm', 'a4');
                const imgWidth = 210;
                const pageHeight = 297;
                const imgHeight = (canvas.height * imgWidth) / canvas.width;
                let heightLeft = imgHeight;
                let position = 0;
                
                pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
                heightLeft -= pageHeight;
                
                while (heightLeft > 0) {
                    position = heightLeft - imgHeight;
                    pdf.addPage();
                    pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
                    heightLeft -= pageHeight;
                }
                
                pdf.save(`reporte_ventas_${new Date().toISOString().split('T')[0]}.pdf`);
                document.body.removeChild(elementoPDF);
            });
        };
        document.head.appendChild(scriptJspdf);
    };
    document.head.appendChild(scriptHtml2canvas);
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

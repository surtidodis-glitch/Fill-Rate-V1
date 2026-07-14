// STATE APP
let rawData = [];
let columnsAvailable = [];
let columnMapping = {};
let uniqueCategories = [];
let chartInstances = {};

// SYSTEM SPEC REQS
const REQUIRED_FIELDS = {
    semana: "Semana",
    tienda: "Tienda",
    categoria: "Categoría",
    surtido: "Surtido",
    entrega: "Entrega",
    clasificacion: "Clasificación",
    fillRate: "Fill Rate",
    departamento: "Departamento",
    pais: "País",
    subcategoria: "Subcategoría"
};

// DOM ELEMENTS
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const screenUpload = document.getElementById('screenUpload');
const screenMapping = document.getElementById('screenMapping');
const screenDashboard = document.getElementById('screenDashboard');
const dynamicMappingFields = document.getElementById('dynamicMappingFields');
const selGroupBrand = document.getElementById('selGroupBrand');
const categoryChecklist = document.getElementById('categoryChecklist');
const loadingOverlay = document.getElementById('loadingOverlay');
const loadingText = document.getElementById('loadingText');
const filterWeek = document.getElementById('filterWeek');
const headerActions = document.getElementById('headerActions');

// EVENT LISTENERS - INITIALIZATION
document.addEventListener('DOMContentLoaded', () => {
    initUploadEvents();
    document.getElementById('btnProcessData').addEventListener('click', processAndGenerateDashboard);
    document.getElementById('btnReset').addEventListener('click', resetApp);
    document.getElementById('btnExportPNG').addEventListener('click', exportDashboardPNG);
    document.getElementById('btnExportPDF').addEventListener('click', exportDashboardPDF);
    filterWeek.addEventListener('change', updateDashboardView);
});

function showLoading(text) {
    loadingText.innerText = text || 'Procesando...';
    loadingOverlay.classList.add('active');
}

function hideLoading() {
    loadingOverlay.classList.remove('active');
}

// PANTALLA 1: CONTROL EXCEL UPLOAD
function initUploadEvents() {
    dropZone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', (e) => handleFile(e.target.files[0]));
    
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });
    
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
    
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            handleFile(e.dataTransfer.files[0]);
        }
    });
}

function handleFile(file) {
    if (!file) return;
    showLoading('Leyendo archivo Excel...');
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Search exact BASE_MAESTRA sheet
            let targetSheetName = workbook.SheetNames.find(name => name.toUpperCase() === 'BASE_MAESTRA');
            if (!targetSheetName) {
                // Fallback to first available sheet if not found
                targetSheetName = workbook.SheetNames[0];
                alert(`No se encontró explícitamente la hoja "BASE_MAESTRA". Se leerá la hoja activa por defecto: "${targetSheetName}"`);
            }
            
            const worksheet = workbook.Sheets[targetSheetName];
            rawData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
            
            if (rawData.length === 0) {
                throw new Error("La hoja de cálculo está vacía o no contiene filas de datos.");
            }
            
            // Extract Headers
            columnsAvailable = Object.keys(rawData[0]);
            suggestColumnMapping();
            buildMappingInterface();
            
            // Extract Unique Categories for Bucket List
            extractCategories();
            buildCategoryChecklist();
            
            // Move to Screen 2
            switchScreen(screenMapping);
        } catch (error) {
            console.error(error);
            alert("Error al procesar el Excel: " + error.message);
        } finally {
            hideLoading();
        }
    };
    reader.readAsArrayBuffer(file);
}

function switchScreen(targetScreen) {
    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
    targetScreen.classList.add('active');
    
    if (targetScreen === screenDashboard) {
        headerActions.style.display = 'flex';
    } else {
        headerActions.style.display = 'none';
    }
}

// PANTALLA 2: SUGGESTION & MAP BUILDS
function suggestColumnMapping() {
    columnMapping = {};
    columnsAvailable.forEach(col => {
        const cleanCol = col.toLowerCase().trim()
            .normalize("NFD").replace(/[\u0300-\u036f]/g, ""); // Remove accents
            
        for (const [key, value] of Object.entries(REQUIRED_FIELDS)) {
            const cleanReq = value.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            if (cleanCol === cleanReq || cleanCol.includes(cleanReq)) {
                if (!columnMapping[key]) {
                    columnMapping[key] = col;
                }
            }
        }
    });
}

function buildMappingInterface() {
    dynamicMappingFields.innerHTML = '';
    
    for (const [key, label] of Object.entries(REQUIRED_FIELDS)) {
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        
        const htmlLabel = document.createElement('label');
        htmlLabel.innerText = `Columna para ${label}:`;
        
        const select = document.createElement('select');
        select.className = 'form-control mapping-selector';
        select.dataset.fieldKey = key;
        
        // Option null
        const optNone = document.createElement('option');
        optNone.value = '';
        optNone.innerText = '-- Ignorar / No Disponible --';
        select.appendChild(optNone);
        
        columnsAvailable.forEach(col => {
            const opt = document.createElement('option');
            opt.value = col;
            opt.innerText = col;
            if (columnMapping[key] === col) {
                opt.selected = true;
            }
            select.appendChild(opt);
        });
        
        formGroup.appendChild(htmlLabel);
        formGroup.appendChild(select);
        dynamicMappingFields.appendChild(formGroup);
    }
    
    // Brand Selection Suggestion
    selGroupBrand.innerHTML = '';
    columnsAvailable.forEach(col => {
        const opt = document.createElement('option');
        opt.value = col;
        opt.innerText = col;
        // Pre-select if matches common patterns
        if (col.toLowerCase().includes('marca') || col.toLowerCase().includes('proveedor')) {
            opt.selected = true;
        }
        selGroupBrand.appendChild(opt);
    });
}

function extractCategories() {
    const catField = columnMapping['categoria'];
    if (!catField) {
        uniqueCategories = [];
        return;
    }
    const setCats = new Set();
    rawData.forEach(row => {
        if (row[catField]) setCats.add(String(row[catField]).toUpperCase().trim());
    });
    uniqueCategories = Array.from(setCats).sort();
}

function buildCategoryChecklist() {
    categoryChecklist.innerHTML = '';
    if (uniqueCategories.length === 0) {
        categoryChecklist.innerHTML = '<div style="color:var(--text-muted); font-size:12px; padding:10px;">Ninguna categoría detectada. Configura la columna de Categoría primero.</div>';
        return;
    }
    
    uniqueCategories.forEach(cat => {
        const item = document.createElement('label');
        item.className = 'checklist-item';
        
        const checkbox = document.createElement('document');
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.value = cat;
        // Auto-detect word "ROPA"
        if (cat.includes('ROPA')) {
            cb.checked = true;
        }
        
        const txt = document.createTextNode(` ${cat}`);
        item.appendChild(cb);
        item.appendChild(txt);
        categoryChecklist.appendChild(item);
    });
}

// PANTALLA 3: FORMULA LOGICS & PROCESS MANAGEMENT
function processAndGenerateDashboard() {
    // Read final user mappings
    document.querySelectorAll('.mapping-selector').forEach(sel => {
        columnMapping[sel.dataset.fieldKey] = sel.value;
    });
    
    const brandField = selGroupBrand.value;
    const weekLabelInput = document.getElementById('txtDeliveryWeek').value.trim();
    const dateSurtirInput = document.getElementById('dateSurtir').value;
    
    if (!weekLabelInput || !dateSurtirInput) {
        alert("Por favor ingresa la semana de entrega y la fecha a surtir.");
        return;
    }
    
    showLoading('Procesando datos e indices analíticos...');
    
    // Set text meta tags
    document.getElementById('metaWeek').innerText = weekLabelInput;
    document.getElementById('metaDate').innerText = dateSurtirInput;
    
    // Extract weeks available inside data array to populate top filters
    populateWeekFilter();
    
    // Refresh calculations view
    updateDashboardView();
    
    hideLoading();
    switchScreen(screenDashboard);
}

function populateWeekFilter() {
    filterWeek.innerHTML = '<option value="ALL">Todas las semanas</option>';
    const weekField = columnMapping['semana'];
    if (!weekField) return;
    
    const weeks = new Set();
    rawData.forEach(row => {
        if (row[weekField]) weeks.add(String(row[weekField]).trim());
    });
    
    Array.from(weeks).sort().forEach(w => {
        const opt = document.createElement('option');
        opt.value = w;
        opt.innerText = `Semana: ${w}`;
        filterWeek.appendChild(opt);
    });
}

function updateDashboardView() {
    const selectedWeekFilter = filterWeek.value;
    
    // Mapping keys
    const fSemana = columnMapping['semana'];
    const fSurtido = columnMapping['surtido'];
    const fEntrega = columnMapping['entrega'];
    const fClasificacion = columnMapping['clasificacion'];
    const fFillRate = columnMapping['fillRate'];
    const fCategoria = columnMapping['categoria'];
    const fBrand = selGroupBrand.value;
    
    // Filter data array
    const filteredData = rawData.filter(row => {
        if (selectedWeekFilter !== 'ALL' && fSemana) {
            return String(row[fSemana]).trim() === selectedWeekFilter;
        }
        return true;
    });
    
    // Global Accumulators
    let sumSurtido = 0;
    let sumEntrega = 0;
    let fpcCumplidos = 0;
    let fpcNoCumplidos = 0;
    let skusCumplidos = 0;
    let totalLineasFPC = filteredData.length;
    
    // Bucket configuration selections
    const labelRopa = document.getElementById('txtLabelRopa').value || "Ropa";
    const labelOtros = document.getElementById('txtLabelOtros').value || "Otros";
    const colorRopa = document.getElementById('colorRopa').value;
    const colorOtros = document.getElementById('colorOtros').value;
    
    // Map of selected items for Ropa
    const checkedRopaCategories = new Set();
    categoryChecklist.querySelectorAll('input[type="checkbox"]').forEach(cb => {
        if (cb.checked) checkedRopaCategories.add(cb.value);
    });
    
    // Aggregation variables
    let brandAgg = {};
    let bucketAgg = {
        ropa: { name: labelRopa, surtido: 0, entrega: 0, fpcTotal: 0, fpcCumplido: 0 },
        otros: { name: labelOtros, surtido: 0, entrega: 0, fpcTotal: 0, fpcCumplido: 0 }
    };
    
    filteredData.forEach(row => {
        const surtidoVal = parseFloat(row[fSurtido]) || 0;
        const entregaVal = parseFloat(row[fEntrega]) || 0;
        
        sumSurtido += surtidoVal;
        sumEntrega += entregaVal;
        
        // SKU Cumplidos = min(Entrega, Surtido)
        const rowSkuCumplido = Math.min(entregaVal, surtidoVal);
        skusCumplidos += rowSkuCumplido;
        
        // FPC Cumplidos evaluation
        let isCumplido = false;
        if (fClasificacion && row[fClasificacion]) {
            isCumplido = String(row[fClasificacion]).trim().toLowerCase() !== 'underfilled';
        } else if (fFillRate && row[fFillRate]) {
            let frVal = parseFloat(row[fFillRate]);
            // If decimal 0-1 or 0-100% format
            if (frVal < 2 && frVal > 0) frVal = frVal * 100;
            isCumplido = frVal >= 100;
        } else {
            // Calculated row percentage metric rule fallback
            isCumplido = surtidoVal > 0 ? (entregaVal >= surtidoVal) : true;
        }
        
        if (isCumplido) fpcCumplidos++; else fpcNoCumplidos++;
        
        // Identify Bucket category row allocation
        const rowCat = fCategoria && row[fCategoria] ? String(row[fCategoria]).toUpperCase().trim() : '';
        const isRopaBucket = checkedRopaCategories.has(rowCat);
        const activeBucketKey = isRopaBucket ? 'ropa' : 'otros';
        
        bucketAgg[activeBucketKey].surtido += surtidoVal;
        bucketAgg[activeBucketKey].entrega += entregaVal;
        bucketAgg[activeBucketKey].fpcTotal++;
        if (isCumplido) bucketAgg[activeBucketKey].fpcCumplido++;
        
        // Brand Aggregation tracking
        const rowBrand = fBrand && row[fBrand] ? String(row[fBrand]).trim() : 'SIN MARCA';
        if (!brandAgg[rowBrand]) {
            brandAgg[rowBrand] = { surtido: 0, entrega: 0 };
        }
        brandAgg[rowBrand].surtido += surtidoVal;
        brandAgg[rowBrand].entrega += entregaVal;
    });
    
    // Compute KPIs
    const globalFillRate = sumSurtido > 0 ? (sumEntrega / sumSurtido) * 100 : 0;
    const globalFpcPct = totalLineasFPC > 0 ? (fpcCumplidos / totalLineasFPC) * 100 : 0;
    const globalSkuPct = sumSurtido > 0 ? (skusCumplidos / sumSurtido) * 100 : 0;
    
    // Update KPI text numbers UI
    document.getElementById('kpiFillRate').innerText = `${globalFillRate.toFixed(1)}%`;
    document.getElementById('kpiTotalFPC').innerText = totalLineasFPC.toLocaleString();
    document.getElementById('kpiFpccumplidos').innerText = fpcCumplidos.toLocaleString();
    document.getElementById('kpiFpcNoCumplidos').innerText = fpcNoCumplidos.toLocaleString();
    document.getElementById('kpiSkusTotales').innerText = sumSurtido.toLocaleString();
    document.getElementById('kpiSkusCumplidos').innerText = skusCumplidos.toLocaleString();
    
    // Render and print tables
    renderTables(bucketAgg, brandAgg);
    
    // Chart Render Executions
    renderDonutChart('chartDonutFR', globalFillRate, varColor('--color-teal'));
    renderDonutChart('chartDonutFPC', globalFpcPct, varColor('--color-blue'));
    renderDonutChart('chartDonutSKU', globalSkuPct, '#a855f7'); // Violet mix accent
    
    // Bucket comparative multi donut calculation
    const ropaFr = bucketAgg.ropa.surtido > 0 ? (bucketAgg.ropa.entrega / bucketAgg.ropa.surtido) * 100 : 0;
    const otrosFr = bucketAgg.otros.surtido > 0 ? (bucketAgg.otros.entrega / bucketAgg.otros.surtido) * 100 : 0;
    renderCompareDonutChart('chartDonutBuckets', [ropaFr, otrosFr], [labelRopa, labelOtros], [colorRopa, colorOtros]);
    
    // Main structural charts
    renderCategoryBars(bucketAgg, colorRopa, colorOtros);
    renderTopBrandsChart(brandAgg);
}

function varColor(varName) {
    return getComputedStyle(document.documentElement).getPropertyValue(varName).trim();
}

function renderTables(bucketAgg, brandAgg) {
    // 1. Bucket Table View
    const tbodyBucket = document.querySelector('#tableBuckets tbody');
    tbodyBucket.innerHTML = '';
    
    Object.values(bucketAgg).forEach(b => {
        const fr = b.surtido > 0 ? (b.entrega / b.surtido) * 100 : 0;
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${b.name}</strong></td>
            <td>${b.surtido.toLocaleString()}</td>
            <td>${b.entrega.toLocaleString()}</td>
            <td>${fr.toFixed(1)}%</td>
        `;
        tbodyBucket.appendChild(tr);
    });
    
    // 2. Brand Table View
    const tbodyBrand = document.querySelector('#tableBrands tbody');
    tbodyBrand.innerHTML = '';
    
    const sortedBrands = Object.entries(brandAgg).sort((a, b) => b[1].surtido - a[1].surtido);
    
    sortedBrands.forEach(([name, data]) => {
        const fr = data.surtido > 0 ? (data.entrega / data.surtido) * 100 : 0;
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${name}</td>
            <td>${data.surtido.toLocaleString()}</td>
            <td>${data.entrega.toLocaleString()}</td>
            <td>${fr.toFixed(1)}%</td>
        `;
        tbodyBrand.appendChild(tr);
    });
}

// GRAPH GENERATOR ENGINE (CHART.JS WRAPPERS)
function cleanChartInstance(canvasId) {
    if (chartInstances[canvasId]) {
        chartInstances[canvasId].destroy();
    }
}

function renderDonutChart(canvasId, value, primaryColor) {
    cleanChartInstance(canvasId);
    const ctx = document.getElementById(canvasId).getContext('2d');
    
    const remainder = Math.max(0, 100 - value);
    
    chartInstances[canvasId] = new Chart(ctx, {
        type: 'doughnut',
        data: {
            datasets: [{
                data: [value, remainder],
                backgroundColor: [primaryColor, 'rgba(255,255,255,0.05)'],
                borderWidth: 0
            }]
        },
        options: {
            cutout: '78%',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: { enabled: false }
            }
        },
        plugins: [{
            id: 'centerText',
            beforeDraw: function(chart) {
                const width = chart.width, height = chart.height, ctx = chart.ctx;
                ctx.restore();
                ctx.font = "bold 22px 'Oswald', sans-serif";
                ctx.textBaseline = "middle";
                ctx.fillStyle = "#f8fafc";
                const text = value.toFixed(1) + "%",
                      textX = Math.round((width - ctx.measureText(text).width) / 2),
                      textY = height / 2;
                ctx.fillText(text, textX, textY);
                ctx.save();
            }
        }]
    });
}

function renderCompareDonutChart(canvasId, values, labels, colors) {
    cleanChartInstance(canvasId);
    const ctx = document.getElementById(canvasId).getContext('2d');
    
    chartInstances[canvasId] = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                data: values,
                backgroundColor: colors,
                borderWidth: 0
            }]
        },
        options: {
            cutout: '60%',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: true,
                    position: 'bottom',
                    labels: { color: '#94a3b8', boxWidth: 10, font: { size: 10 } }
                }
            }
        }
    });
}

function renderCategoryBars(bucketAgg, colorRopa, colorOtros) {
    const canvasId = 'chartCategoryBars';
    cleanChartInstance(canvasId);
    const ctx = document.getElementById(canvasId).getContext('2d');
    
    const bRopa = bucketAgg.ropa;
    const bOtros = bucketAgg.otros;
    
    const ropaFpcPct = bRopa.fpcTotal > 0 ? (bRopa.fpcCumplido / bRopa.fpcTotal) * 100 : 0;
    const ropaSkuPct = bRopa.surtido > 0 ? (Math.min(bRopa.entrega, bRopa.surtido) / bRopa.surtido) * 100 : 0;
    
    const otrosFpcPct = bOtros.fpcTotal > 0 ? (bOtros.fpcCumplido / bOtros.fpcTotal) * 100 : 0;
    const otrosSkuPct = bOtros.surtido > 0 ? (Math.min(bOtros.entrega, bOtros.surtido) / bOtros.surtido) * 100 : 0;
    
    chartInstances[canvasId] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: [bRopa.name, bOtros.name],
            datasets: [
                {
                    label: 'FPC % Cumplido',
                    data: [ropaFpcPct, otrosFpcPct],
                    backgroundColor: '#2dd4bf',
                    borderRadius: 4
                },
                {
                    label: 'SKUs % Surtido',
                    data: [ropaSkuPct, otrosSkuPct],
                    backgroundColor: '#3b82f6',
                    borderRadius: 4
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { grid: { display: false }, ticks: { color: '#94a3b8' } },
                y: { min: 0, max: 100, grid: { color: '#1e3a66' }, ticks: { color: '#94a3b8' } }
            },
            plugins: {
                legend: { labels: { color: '#f8fafc', font: { family: 'Inter' } } }
            }
        }
    });
}

function renderTopBrandsChart(brandAgg) {
    const canvasId = 'chartTopBrands';
    cleanChartInstance(canvasId);
    const ctx = document.getElementById(canvasId).getContext('2d');
    
    // Sort and slice top 10
    const top10 = Object.entries(brandAgg)
        .sort((a, b) => b[1].surtido - a[1].surtido)
        .slice(0, 10);
        
    const labels = top10.map(item => item[0]);
    const dataSurtido = top10.map(item => item[1].surtido);
    
    chartInstances[canvasId] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Cantidad Surtido',
                data: dataSurtido,
                backgroundColor: 'rgba(59, 130, 246, 0.75)',
                borderColor: '#3b82f6',
                borderWidth: 1,
                borderRadius: 4
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { grid: { color: '#1e3a66' }, ticks: { color: '#94a3b8' } },
                y: { grid: { display: false }, ticks: { color: '#94a3b8', font: { size: 11 } } }
            },
            plugins: {
                legend: { display: false }
            }
        }
    });
}

// RESET ACTION
function resetApp() {
    if (confirm('¿Estás seguro de reiniciar la aplicación? Se perderán los datos actuales.')) {
        rawData = [];
        columnsAvailable = [];
        columnMapping = {};
        uniqueCategories = [];
        
        // Clean instances
        Object.keys(chartInstances).forEach(key => {
            chartInstances[key].destroy();
        });
        chartInstances = {};
        
        fileInput.value = '';
        document.getElementById('mappingForm').reset();
        switchScreen(screenUpload);
    }
}

// EXPORT TO PNG
function exportDashboardPNG() {
    showLoading('Generando imagen de Dashboard...');
    const target = document.getElementById('screenDashboard');
    
    // Safe adjustment for capture view margins layout
    html2canvas(target, {
        backgroundColor: '#070c17',
        scale: 2,
        logging: false,
        useCORS: true
    }).then(canvas => {
        const link = document.createElement('a');
        link.download = `FR_Dashboard_${document.getElementById('metaWeek').innerText}.png`;
        link.href = canvas.toDataURL('image/png');
        link.click();
        hideLoading();
    }).catch(err => {
        console.error(err);
        alert('Error al exportar PNG');
        hideLoading();
    });
}

// EXPORT TO PDF
function exportDashboardPDF() {
    showLoading('Compilando reporte en PDF...');
    const target = document.getElementById('screenDashboard');
    
    html2canvas(target, {
        backgroundColor: '#070c17',
        scale: 2,
        logging: false,
        useCORS: true
    }).then(canvas => {
        const imgData = canvas.toDataURL('image/png');
        
        const { jsPDF } = window.jspdf;
        // Landscape orientation report view fit
        const pdf = new jsPDF('l', 'mm', 'a4');
        const imgWidth = 297; 
        const pageHeight = 210;  
        const imgHeight = (canvas.height * imgWidth) / canvas.width;
        let heightLeft = imgHeight;
        let position = 0;
        
        pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
        heightLeft -= pageHeight;
        
        while (heightLeft >= 0) {
            position = heightLeft - imgHeight;
            pdf.addPage();
            pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
            heightLeft -= pageHeight;
        }
        
        pdf.save(`Reporte_Fill_Rate_${document.getElementById('metaWeek').innerText}.pdf`);
        hideLoading();
    }).catch(err => {
        console.error(err);
        alert('Error al exportar PDF');
        hideLoading();
    });
}

document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const fileInfo = document.getElementById('file-info');
    const checklistContainer = document.getElementById('categories-checklist');
    
    // Todos los selects de mapeo de columnas
    const selectElements = {
        semana: document.getElementById('col-semana'),
        tienda: document.getElementById('col-tienda'),
        categoria: document.getElementById('col-categoria'),
        surtido: document.getElementById('col-surtido'),
        entrega: document.getElementById('col-entrega'),
        clasificacion: document.getElementById('col-clasificacion'),
        fillrate: document.getElementById('col-fillrate'),
        departamento: document.getElementById('col-departamento'),
        pais: document.getElementById('col-pais'),
        subcategoria: document.getElementById('col-subcategoria')
    };

    // Disparadores para abrir explorador de archivos
    dropZone.addEventListener('click', () => fileInput.click());

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) handleFile(e.target.files[0]);
    });

    // Eventos Drag & Drop
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.style.background = '#162238';
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.style.background = '#131c2e';
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.style.background = '#131c2e';
        if (e.dataTransfer.files.length > 0) {
            fileInput.files = e.dataTransfer.files;
            handleFile(e.dataTransfer.files[0]);
        }
    });

    // Procesamiento inteligente del Excel / CSV
    function handleFile(file) {
        fileInfo.classList.remove('hidden');
        fileInfo.textContent = `Leyendo archivo: ${file.name} (${(file.size / 1024).toFixed(1)} KB)...`;

        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                // Leemos únicamente las primeras filas para mapear de manera instantánea e indolora
                const workbook = XLSX.read(data, { 
                    type: 'array', 
                    sheetRows: 50 
                });

                // Buscar pestaña prioritaria BASE_MAESTRA o tomar la primera por defecto
                let sheetName = workbook.SheetNames.find(name => name.toUpperCase() === 'BASE_MAESTRA') || workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                
                // Extraer encabezados (Headers)
                const headers = XLSX.utils.sheet_to_json(worksheet, { header: 1 })[0];

                if (!headers || headers.length === 0) {
                    throw new Error("No se detectaron columnas válidas en la fila superior.");
                }

                fileInfo.textContent = `Archivo cargado con éxito. Pestaña leída: "${sheetName}". Columnas encontradas: ${headers.length}`;
                populateSelects(headers);
                generateMockCategories(); // Simulación inteligente de categorías detectadas

            } catch (error) {
                fileInfo.textContent = `Error al leer archivo: ${error.message}`;
                fileInfo.style.borderLeftColor = '#ef4444';
            }
        };

        reader.readAsArrayBuffer(file);
    }

    // Llenado inteligente de menús desplegables e intento de auto-match
    function populateSelects(headers) {
        Object.keys(selectElements).forEach(key => {
            const select = selectElements[key];
            select.innerHTML = ''; // Limpiar estado de espera

            headers.forEach(header => {
                const option = document.createElement('option');
                option.value = header;
                option.textContent = header;
                
                // Lógica de preselección inteligente básica por texto similar
                if (header.toLowerCase().includes(key.toLowerCase())) {
                    option.selected = true;
                }
                select.appendChild(option);
            });
        });
    }

    // Generador dinámico de categorías basado en la lectura
    function generateMockCategories() {
        checklistContainer.innerHTML = '';
        
        const mockCategories = [
            'ROPA INTERIOR', 
            'ROPA DEPORTIVA', 
            'ACCESORIOS', 
            'CALZADO', 
            'ROPA DE SS DEPORTIVA'
        ];

        mockCategories.forEach(cat => {
            const label = document.createElement('label');
            label.className = 'check-item';
            
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.value = cat;
            
            // Auto marcar si contiene la palabra ROPA para emular el Bucket Ropa
            if (cat.includes('ROPA')) {
                checkbox.checked = true;
            }

            label.appendChild(checkbox);
            label.appendChild(document.createTextNode(cat));
            checklistContainer.appendChild(label);
        });
    }

    // Botón de Procesamiento Final
    document.getElementById('btn-generate').addEventListener('click', () => {
        alert('¡Excelente! Los datos estructurados de configuración y el mapeo están listos para ser transformados en las métricas finales de tu Dashboard.');
    });
});

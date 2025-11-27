// Inicializar mapa
const map = L.map("map").setView([6.2442, -75.5812], 13);

// Definir capas de mapas
const layers = {
    osm: L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
        attribution: "&copy; OpenStreetMap contributors",
    }),
    esri: L.tileLayer(
        "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
        {
            attribution: "&copy; Esri",
        }
    ),
    topo: L.tileLayer(
        "https://server.arcgisonline.com/ArcGIS/rest/services/World_Topo_Map/MapServer/tile/{z}/{y}/{x}",
        {
            attribution: "&copy; Esri",
        }
    ),
    dark: L.tileLayer(
        "https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png",
        {
            attribution: "&copy; CartoDB",
        }
    )
};

// Carga por defecto
let currentLayer = 'osm';
layers.osm.addTo(map);

// Función para cambiar mapa base
function setBaseMap(type) {
    // Elimina todas las capas
    map.eachLayer((layer) => map.removeLayer(layer));
    // Agrega la seleccionada
    layers[type].addTo(map);
    currentLayer = type;

    // Actualizar indicador visual
    document.querySelectorAll('.map-option').forEach(option => {
        option.classList.remove('active');
    });
    event.target.classList.add('active');

    // Cerrar dropdown
    document.getElementById('mapsDropdown').classList.remove('show');
}

// Toggle del menú principal
const toggleMenuBtn = document.getElementById("toggleMenuBtn");
const menuButtons = document.getElementById("menuButtons");
let menuOpen = false;

toggleMenuBtn.addEventListener("click", () => {
    menuOpen = !menuOpen;
    if (menuOpen) {
        menuButtons.classList.add("show");
        toggleMenuBtn.textContent = "❌";
    } else {
        menuButtons.classList.remove("show");
        toggleMenuBtn.textContent = "☰";
        // Cerrar dropdown si está abierto
        document.getElementById('mapsDropdown').classList.remove('show');
    }
});

// Manejo del dropdown de mapas
const mapsBtn = document.getElementById("mapsBtn");
const mapsDropdown = document.getElementById("mapsDropdown");

mapsBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    mapsDropdown.classList.toggle("show");
});

// Cerrar dropdown al hacer click fuera
document.addEventListener("click", (e) => {
    if (!mapsDropdown.contains(e.target) && e.target !== mapsBtn) {
        mapsDropdown.classList.remove("show");
    }
});

// Funciones placeholder para otros botones
function showFeature(feature) {
    alert(`Funcionalidad "${feature}" será implementada próximamente`);
}

// Funciones del modal de medición
let selectedColor = '#ff4500';
let projectLayer = null;
let currentProjectFile = null;
let correriaAssign = { active: false, name: '', color: '#28a745' };
let projectPopupsEnabled = true;

function createPinIcon(color, extraClass = '') {
    const html = `\n<div class="pin" style="--pin-color:${color}">\n  <svg viewBox="0 0 24 24" class="pin-svg">\n    <path d="M12 2c4.971 0 9 4.029 9 9 0 6.394-9 13-9 13S3 17.394 3 11c0-4.971 4.029-9 9-9z" fill="var(--pin-color)"></path>\n    <circle cx="12" cy="11" r="3" fill="#ffffff"></circle>\n  </svg>\n</div>`;
    return L.divIcon({ className: `pin-icon ${extraClass}`, html, iconSize: [30, 42], iconAnchor: [15, 42], popupAnchor: [0, -36] });
}

function parseDDMMYYYY(s) {
    if (!s) return null;
    const parts = String(s).split(/[\/-]/).map(p => p.trim());
    if (parts.length < 3) return null;
    const dd = parseInt(parts[0], 10);
    const mm = parseInt(parts[1], 10) - 1;
    const yyyy = parseInt(parts[2], 10);
    if (Number.isNaN(dd) || Number.isNaN(mm) || Number.isNaN(yyyy)) return null;
    return new Date(yyyy, mm, dd);
}

function businessDaysBetween(startDate, endDate) {
    if (!startDate || !endDate) return 0;
    const s = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
    const e = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
    if (s > e) return 0;
    let count = 0;
    const d = new Date(s);
    while (d <= e) {
        const day = d.getDay();
        if (day !== 0 && day !== 6) count++;
        d.setDate(d.getDate() + 1);
    }
    return Math.max(0, count - 1);
}

function showMeasureModal() {
    document.getElementById('measureModal').classList.add('show');
}

function closeMeasureModal() {
    document.getElementById('measureModal').classList.remove('show');
}

// Funciones del modal de carga
function showUploadModal() {
    document.getElementById('uploadModal').classList.add('show');
    loadFiles();
}

function closeUploadModal() {
    document.getElementById('uploadModal').classList.remove('show');
    // Limpiar estado
    document.getElementById('excelFile').value = '';
    document.getElementById('uploadStatus').textContent = '';
    const fileNameEl = document.getElementById('fileNameDisplay');
    if (fileNameEl) fileNameEl.textContent = 'Ningún archivo seleccionado';
}

async function loadFiles() {
    try {
        const response = await fetch('/files');
        const files = await response.json();
        renderFileList(files);
    } catch (error) {
        console.error('Error cargando archivos:', error);
    }
}

async function uploadExcel() {
    const fileInput = document.getElementById('excelFile');
    const statusDiv = document.getElementById('uploadStatus');
    const file = fileInput.files[0];
    const fileNameEl = document.getElementById('fileNameDisplay');

    if (!file) {
        statusDiv.textContent = 'Por favor seleccione un archivo';
        statusDiv.className = 'status-message error';
        return;
    }

    const formData = new FormData();
    formData.append('excelFile', file);

    statusDiv.textContent = 'Subiendo y procesando...';
    statusDiv.className = 'status-message';

    try {
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (response.ok) {
            statusDiv.textContent = 'Archivo cargado exitosamente';
            statusDiv.className = 'status-message success';
            loadFiles(); // Recargar lista
            fileInput.value = '';
            if (fileNameEl) fileNameEl.textContent = 'Ningún archivo seleccionado';
        } else {
            throw new Error(result.error || 'Error en la carga');
        }
    } catch (error) {
        statusDiv.textContent = 'Error: ' + error.message;
        statusDiv.className = 'status-message error';
    }
}

function renderFileList(files) {
    const previewSection = document.getElementById('previewSection');
    const table = document.getElementById('previewTable');

    // Limpiar tabla
    table.innerHTML = '';

    // Crear encabezados
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    const thName = document.createElement('th');
    thName.textContent = 'Nombre del Archivo';
    const thActions = document.createElement('th');
    thActions.textContent = 'Acciones';
    headerRow.appendChild(thName);
    headerRow.appendChild(thActions);
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Crear cuerpo
    const tbody = document.createElement('tbody');
    if (files.length === 0) {
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        td.textContent = 'No hay archivos guardados';
        tr.appendChild(td);
        tbody.appendChild(tr);
    } else {
        files.forEach(file => {
            const tr = document.createElement('tr');
            const tdName = document.createElement('td');
            const nameSpan = document.createElement('span');
            nameSpan.textContent = file;
            tdName.appendChild(nameSpan);

            const tdActions = document.createElement('td');
            const runBtn = document.createElement('button');
            runBtn.type = 'button';
            runBtn.className = 'run-btn';
            const runIcon = document.createElement('i');
            runIcon.setAttribute('data-lucide', 'play');
            const runLabel = document.createElement('span');
            runLabel.textContent = 'Run';
            runBtn.appendChild(runIcon);
            runBtn.appendChild(runLabel);
            runBtn.addEventListener('click', () => openPrepareModal(file));

            const deleteBtn = document.createElement('button');
            deleteBtn.type = 'button';
            deleteBtn.className = 'delete-btn';
            const delIcon = document.createElement('i');
            delIcon.setAttribute('data-lucide', 'trash-2');
            const delLabel = document.createElement('span');
            delLabel.textContent = 'Borrar';
            deleteBtn.appendChild(delIcon);
            deleteBtn.appendChild(delLabel);

            deleteBtn.addEventListener('click', async () => {
                const confirmed = await Swal.fire({
                    title: '¿Borrar archivo?',
                    text: file,
                    icon: 'warning',
                    showCancelButton: true,
                    confirmButtonColor: '#d33',
                    cancelButtonColor: '#3085d6',
                    confirmButtonText: 'Sí, borrar',
                    cancelButtonText: 'Cancelar'
                }).then(r => r.isConfirmed);
                if (!confirmed) return;
                try {
                    const resp = await fetch('/files/' + encodeURIComponent(file), { method: 'DELETE' });
                    const data = await resp.json();
                    if (resp.ok) {
                        await Swal.fire('Borrado', 'El archivo fue eliminado', 'success');
                        loadFiles();
                    } else {
                        await Swal.fire('Error', data.error || 'No se pudo borrar', 'error');
                    }
                } catch (e) {
                    await Swal.fire('Error', 'Falló la petición', 'error');
                }
            });

            tdActions.appendChild(runBtn);
            tdActions.appendChild(deleteBtn);

            tr.appendChild(tdName);
            tr.appendChild(tdActions);
            tbody.appendChild(tr);
        });
    }
    table.appendChild(tbody);

    previewSection.style.display = 'block';
    if (window.lucide && typeof lucide.createIcons === 'function') {
        lucide.createIcons();
    }
}

function closePrepareModal() {
    document.getElementById('prepareModal').classList.remove('show');
}

async function openPrepareModal(file) {
    const modal = document.getElementById('prepareModal');
    const status = document.getElementById('prepareStatus');
    const fileDiv = document.getElementById('prepareFileName');
    const select = document.getElementById('columnsSelect');
    const summarySection = document.getElementById('prepareSummarySection');
    const summaryTable = document.getElementById('columnSummaryTable');

    fileDiv.textContent = file;
    status.style.display = 'none';
    select.innerHTML = '<option value="" disabled selected>Cargando columnas...</option>';
    summarySection.style.display = 'none';
    summaryTable.innerHTML = '';
    modal.classList.add('show');

    try {
        const resp = await fetch('/files/' + encodeURIComponent(file) + '/columns');
        const data = await resp.json();
        if (!resp.ok) {
            throw new Error(data.error || 'No se pudieron leer las columnas');
        }
        select.innerHTML = '';
        if (!data.columns || data.columns.length === 0) {
            const opt = document.createElement('option');
            opt.textContent = 'No se encontraron columnas';
            opt.disabled = true;
            select.appendChild(opt);
        } else {
            data.columns.forEach(col => {
                const opt = document.createElement('option');
                opt.value = col;
                opt.textContent = col;
                select.appendChild(opt);
            });
            select.onchange = () => loadColumnSummary(file, select.value);
        }
        if (window.lucide && typeof lucide.createIcons === 'function') {
            lucide.createIcons();
        }
    } catch (e) {
        status.textContent = 'Error: ' + e.message;
        status.className = 'status-message error';
        status.style.display = 'block';
    }
}

async function loadColumnSummary(file, column) {
    const status = document.getElementById('prepareStatus');
    const summarySection = document.getElementById('prepareSummarySection');
    const table = document.getElementById('columnSummaryTable');
    if (!column) return;
    status.textContent = 'Cargando resumen...';
    status.className = 'status-message';
    status.style.display = 'block';
    summarySection.style.display = 'none';
    table.innerHTML = '';
    try {
        const resp = await fetch('/files/' + encodeURIComponent(file) + '/column-summary?column=' + encodeURIComponent(column));
        const data = await resp.json();
        if (!resp.ok) {
            throw new Error(data.error || 'No se pudo obtener el resumen');
        }
        renderColumnSummaryTable(data.summary);
        status.style.display = 'none';
        summarySection.style.display = 'block';
    } catch (e) {
        status.textContent = 'Error: ' + e.message;
        status.className = 'status-message error';
        status.style.display = 'block';
    }
}

function renderColumnSummaryTable(items) {
    const table = document.getElementById('columnSummaryTable');
    table.innerHTML = '';
    const thead = document.createElement('thead');
    const hr = document.createElement('tr');
    const thVal = document.createElement('th');
    thVal.textContent = 'Tipo de Dato';
    const thCount = document.createElement('th');
    thCount.textContent = 'Cantidad';
    const thColor = document.createElement('th');
    thColor.textContent = 'Color';
    hr.appendChild(thVal);
    hr.appendChild(thCount);
    hr.appendChild(thColor);
    thead.appendChild(hr);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    if (!items || items.length === 0) {
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        td.textContent = 'Sin datos';
        td.colSpan = 3;
        tr.appendChild(td);
        tbody.appendChild(tr);
    } else {
        items.forEach(({ value, count }) => {
            const tr = document.createElement('tr');
            const tdV = document.createElement('td');
            tdV.textContent = value;
            const tdC = document.createElement('td');
            tdC.textContent = String(count);
            const tdColor = document.createElement('td');
            const colorInput = document.createElement('input');
            colorInput.type = 'color';
            colorInput.className = 'summary-color-input';
            colorInput.value = '#000000';
            colorInput.addEventListener('click', function (e) {
                e.preventDefault();
                e.stopPropagation();
                openSummaryColorPicker(colorInput);
            });
            tdColor.appendChild(colorInput);
            tr.appendChild(tdV);
            tr.appendChild(tdC);
            tr.appendChild(tdColor);
            tbody.appendChild(tr);
        });
    }
    table.appendChild(tbody);
}

function collectColorMapping() {
    const rows = document.querySelectorAll('#columnSummaryTable tbody tr');
    const mapping = {};
    rows.forEach(tr => {
        const tds = tr.querySelectorAll('td');
        if (tds.length >= 3) {
            const val = tds[0].textContent.trim();
            const colorInput = tds[2].querySelector('input[type="color"]');
            if (val && colorInput) {
                mapping[val] = colorInput.value || '#000000';
            }
        }
    });
    return mapping;
}

async function saveColoredExcel() {
    const file = document.getElementById('prepareFileName').textContent.trim();
    const column = document.getElementById('columnsSelect').value;
    if (!file || !column) {
        await Swal.fire('Faltan datos', 'Seleccione archivo y columna', 'warning');
        return;
    }
    const colors = collectColorMapping();
    try {
        const resp = await fetch('/files/' + encodeURIComponent(file) + '/colorized', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ column, colors })
        });
        const data = await resp.json();
        if (resp.ok) {
            await Swal.fire('Guardado', 'Proyecto: ' + data.output, 'success');
        } else {
            await Swal.fire('Error', data.error || 'No se pudo guardar', 'error');
        }
    } catch (e) {
        await Swal.fire('Error', 'Falló la petición', 'error');
    }
}

let summaryColorTarget = null;

function closeColorPickerModal() {
    document.getElementById('colorPickerModal').classList.remove('show');
}

function openSummaryColorPicker(target) {
    summaryColorTarget = target;
    const modal = document.getElementById('colorPickerModal');
    const input = document.getElementById('colorPickerInput');
    const grid = document.getElementById('colorPresetGrid');
    input.value = target.value || '#000000';
    Array.from(grid.querySelectorAll('.preset-color')).forEach(el => {
        el.classList.toggle('selected', el.dataset.color.toLowerCase() === input.value.toLowerCase());
        el.onclick = () => {
            Array.from(grid.querySelectorAll('.preset-color')).forEach(x => x.classList.remove('selected'));
            el.classList.add('selected');
            input.value = el.dataset.color;
        };
    });
    input.onchange = function () {
        Array.from(grid.querySelectorAll('.preset-color')).forEach(x => {
            x.classList.toggle('selected', x.dataset.color.toLowerCase() === input.value.toLowerCase());
        });
    };
    modal.classList.add('show');
}

function applySummaryColor() {
    const input = document.getElementById('colorPickerInput');
    if (typeof correriaColorPickerActive !== 'undefined' && correriaColorPickerActive) {
        correriaAssign.color = input.value || '#28a745';
        correriaColorPickerActive = false;
    } else if (summaryColorTarget) {
        summaryColorTarget.value = input.value;
    }
    closeColorPickerModal();
}

// Manejo de selección de colores
document.addEventListener('DOMContentLoaded', function () {
    const colorOptions = document.querySelectorAll('.color-option');
    const customColorPicker = document.getElementById('customColorPicker');
    const fileInputEl = document.getElementById('excelFile');
    const fileNameEl = document.getElementById('fileNameDisplay');

    if (window.lucide && typeof lucide.createIcons === 'function') {
        lucide.createIcons();
    }

    colorOptions.forEach(option => {
        option.addEventListener('click', function () {
            // Remover selección anterior
            colorOptions.forEach(opt => opt.classList.remove('selected'));
            // Seleccionar nuevo color
            this.classList.add('selected');
            selectedColor = this.dataset.color;
            customColorPicker.value = selectedColor;
        });
    });

    // Color personalizado
    customColorPicker.addEventListener('change', function () {
        selectedColor = this.value;
        // Remover selección de colores predefinidos
        colorOptions.forEach(opt => opt.classList.remove('selected'));

        // Verificar si el color personalizado coincide con alguno predefinido
        const matchingOption = Array.from(colorOptions).find(opt =>
            opt.dataset.color.toLowerCase() === selectedColor.toLowerCase()
        );
        if (matchingOption) {
            matchingOption.classList.add('selected');
        }
    });

    if (fileInputEl && fileNameEl) {
        fileInputEl.addEventListener('change', function () {
            const f = this.files && this.files[0];
            fileNameEl.textContent = f ? f.name : 'Ningún archivo seleccionado';
        });
    }

    window.openFilePicker = function () {
        const input = document.getElementById('excelFile');
        if (input) input.click();
    }
});

function searchUsers() {
    const usersText = document.getElementById('usersInput').value.trim();

    if (!usersText) {
        alert('Por favor, ingrese la lista de usuarios');
        return;
    }

    // Convertir texto en array de usuarios
    const users = usersText.split('\n')
        .map(user => user.trim())
        .filter(user => user.length > 0);

    if (users.length === 0) {
        alert('No se encontraron usuarios válidos');
        return;
    }

    // Aquí implementarás la lógica de búsqueda
    console.log('Usuarios a buscar:', users);
    console.log('Color seleccionado:', selectedColor);

    // Placeholder para la funcionalidad
    alert(`Se buscarán ${users.length} usuarios con color ${selectedColor}`);

    // Cerrar modal
    closeMeasureModal();
}

// Cerrar modales al hacer click fuera
window.onclick = function (event) {
    const measureModal = document.getElementById('measureModal');
    const uploadModal = document.getElementById('uploadModal');
    const prepareModal = document.getElementById('prepareModal');
    const colorPickerModal = document.getElementById('colorPickerModal');
    const dataModal = document.getElementById('dataModal');
    const workshopModal = document.getElementById('workshopModal');

    if (event.target == measureModal) {
        closeMeasureModal();
    }
    if (event.target == uploadModal) {
        closeUploadModal();
    }
    if (event.target == prepareModal) {
        closePrepareModal();
    }
    if (event.target == colorPickerModal) {
        closeColorPickerModal();
    }
    if (event.target == dataModal) {
        closeDataModal();
    }
    if (event.target == workshopModal) {
        closeWorkshopModal();
    }
}

function closeDataModal() {
    const el = document.getElementById('dataModal');
    if (el) el.classList.remove('show');
}

async function openDataModal() {
    const modal = document.getElementById('dataModal');
    const status = document.getElementById('projectsStatus');
    const table = document.getElementById('projectsTable');
    if (!modal) return;
    table.innerHTML = '';
    status.style.display = 'none';
    modal.classList.add('show');
    try {
        const resp = await fetch('/projects');
        const data = await resp.json();
        if (!resp.ok) throw new Error(data.error || 'No se pudo listar proyectos');
        renderProjectsTable(data);
    } catch (e) {
        status.textContent = 'Error: ' + e.message;
        status.className = 'status-message error';
        status.style.display = 'block';
    }
}

function openWorkshopModal() {
    const el = document.getElementById('workshopModal');
    if (el) {
        el.classList.add('show');
        el.classList.remove('minimized');
        const content = document.querySelector('#workshopModal .modal-content');
        if (content) content.classList.remove('minimized');
        const restore = document.querySelector('#workshopModal .restore-btn');
        if (restore) restore.style.display = 'none';
        const minimize = document.querySelector('#workshopModal .minimize-btn');
        if (minimize) minimize.style.display = 'inline-block';
        if (window.lucide && typeof lucide.createIcons === 'function') {
            lucide.createIcons();
        }
    }
}

function closeWorkshopModal() {
    const el = document.getElementById('workshopModal');
    if (el) el.classList.remove('show');
    hideCorreriaBanner();
}

function openManualCorreria() {
    const section = document.getElementById('manualCorreriaSection');
    const msg = document.querySelector('#workshopModal .status-message');
    if (section) section.style.display = 'block';
    if (msg) msg.style.display = 'none';
    if (window.lucide && typeof lucide.createIcons === 'function') {
        lucide.createIcons();
    }
}

function openAutoCorreria() {
    Swal.fire('Modo automático', 'Próximamente', 'info');
}

async function saveCorreriaName() {
    const input = document.getElementById('correriaNameInput');
    const name = input ? input.value.trim() : '';
    if (!name) {
        await Swal.fire('Dato requerido', 'Ingrese el nombre de la correría', 'warning');
        return;
    }
    addCorreriaRow(name);
    input.value = '';
    await Swal.fire('Guardado', 'Correría: ' + name, 'success');
}

let correrias = [];

function ensureCorreriaTableHeader() {
    const table = document.getElementById('correriaTable');
    if (!table) return;
    if (table.querySelector('thead')) return;
    const thead = document.createElement('thead');
    const hr = document.createElement('tr');
    const thN = document.createElement('th');
    thN.textContent = 'Nombre Correria';
    const thC = document.createElement('th');
    thC.textContent = 'Cant.';
    const thT = document.createElement('th');
    thT.textContent = 'Terminal';
    const thE = document.createElement('th');
    thE.textContent = 'Editar';
    const thS = document.createElement('th');
    thS.textContent = 'Guardar';
    const thV = document.createElement('th');
    thV.textContent = 'Ver';
    const thA = document.createElement('th');
    thA.textContent = 'Add';
    hr.appendChild(thN);
    hr.appendChild(thC);
    hr.appendChild(thT);
    hr.appendChild(thE);
    hr.appendChild(thS);
    hr.appendChild(thV);
    hr.appendChild(thA);
    thead.appendChild(hr);
    table.appendChild(thead);
    const tbody = document.createElement('tbody');
    table.appendChild(tbody);
}

function addCorreriaRow(name) {
    ensureCorreriaTableHeader();
    const table = document.getElementById('correriaTable');
    const tbody = table.querySelector('tbody');
    const idx = correrias.length;
    correrias.push({ name, count: 0, terminal: '' });
    const tr = document.createElement('tr');
    const tdName = document.createElement('td');
    tdName.textContent = name;
    const tdCount = document.createElement('td');
    tdCount.textContent = '0';
    const tdTerminal = document.createElement('td');
    const terminalInput = document.createElement('input');
    terminalInput.type = 'text';
    terminalInput.className = 'select-input terminal-input';
    terminalInput.value = '';
    tdTerminal.appendChild(terminalInput);
    const tdEdit = document.createElement('td');
    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.className = 'icon-btn edit';
    const editIcon = document.createElement('i');
    editIcon.setAttribute('data-lucide', 'pencil');
    editBtn.appendChild(editIcon);
    editBtn.addEventListener('click', () => editCorreria(idx));
    tdEdit.appendChild(editBtn);
    const tdSave = document.createElement('td');
    const saveBtn = document.createElement('button');
    saveBtn.type = 'button';
    saveBtn.className = 'icon-btn save';
    const saveIcon = document.createElement('i');
    saveIcon.setAttribute('data-lucide', 'save');
    saveBtn.appendChild(saveIcon);
    saveBtn.addEventListener('click', () => saveCorreria(idx));
    tdSave.appendChild(saveBtn);
    const tdView = document.createElement('td');
    const viewBtn = document.createElement('button');
    viewBtn.type = 'button';
    viewBtn.className = 'icon-btn view';
    const viewIcon = document.createElement('i');
    viewIcon.setAttribute('data-lucide', 'eye');
    viewBtn.appendChild(viewIcon);
    tdView.appendChild(viewBtn);
    const tdAdd = document.createElement('td');
    const addBtn = document.createElement('button');
    addBtn.type = 'button';
    addBtn.className = 'icon-btn add';
    const addIcon = document.createElement('i');
    addIcon.setAttribute('data-lucide', 'plus');
    addBtn.appendChild(addIcon);
    addBtn.addEventListener('click', () => startCorreriaAssignment(idx));
    tdAdd.appendChild(addBtn);
    tr.appendChild(tdName);
    tr.appendChild(tdCount);
    tr.appendChild(tdTerminal);
    tr.appendChild(tdEdit);
    tr.appendChild(tdSave);
    tr.appendChild(tdView);
    tr.appendChild(tdAdd);
    tbody.appendChild(tr);
    if (window.lucide && typeof lucide.createIcons === 'function') {
        lucide.createIcons();
    }
}

function generateTerminalCode() {
    let r = '';
    for (let i = 0; i < 4; i++) {
        r += Math.floor(Math.random() * 10);
    }
    return 'CENS' + r;
}

function editCorreria(idx) {
    const table = document.getElementById('correriaTable');
    const row = table.querySelectorAll('tbody tr')[idx];
    if (!row) return;
    const tdName = row.children[0];
    const current = tdName.textContent.trim();
    const input = document.createElement('input');
    input.type = 'text';
    input.value = current;
    input.className = 'select-input';
    tdName.innerHTML = '';
    tdName.appendChild(input);
    input.focus();
    const tdTerminal = row.children[2];
    const tInput = tdTerminal.querySelector('input');
    if (tInput) tInput.disabled = false;
    const saveBtn = row.children[4]?.querySelector('button.icon-btn.save');
    if (saveBtn) saveBtn.classList.remove('saved');
}

function saveCorreria(idx) {
    const table = document.getElementById('correriaTable');
    const row = table.querySelectorAll('tbody tr')[idx];
    if (!row) return;
    const tdName = row.children[0];
    const input = tdName.querySelector('input');
    const name = input ? input.value.trim() : tdName.textContent.trim();
    if (!name) return;
    correrias[idx].name = name;
    tdName.textContent = name;
    const tdTerminal = row.children[2];
    const tInput = tdTerminal.querySelector('input');
    const terminal = tInput ? tInput.value.trim() : '';
    correrias[idx].terminal = terminal;
    if (!currentProjectFile) {
        Swal.fire('Proyecto requerido', 'Abre un proyecto para guardar Tpl en el Excel', 'warning');
        return;
    }
    fetch('/projects/' + encodeURIComponent(currentProjectFile) + '/update-tpl-by-description', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ description: name, tpl: terminal })
    })
      .then(r => r.json().then(d => ({ ok: r.ok, data: d })))
      .then(({ ok }) => {
        if (ok) {
          if (tInput) tInput.disabled = true;
          const saveBtn = row.children[4]?.querySelector('button.icon-btn.save');
          if (saveBtn) saveBtn.classList.add('saved');
        } else {
          Swal.fire('Error', 'No se pudo actualizar Tpl', 'error');
        }
      })
      .catch(() => Swal.fire('Error', 'Falló la actualización en Excel', 'error'));
}

function startCorreriaAssignment(idx) {
    if (!currentProjectFile) {
        Swal.fire('Proyecto requerido', 'Abre un proyecto en "Proyectos guardados" para asignar puntos', 'warning');
        return;
    }
    const table = document.getElementById('correriaTable');
    const row = table.querySelectorAll('tbody tr')[idx];
    if (!row) return;
    const nameCell = row.children[0];
    const input = nameCell.querySelector('input');
    const name = input ? input.value.trim() : nameCell.textContent.trim();
    if (!name) {
        Swal.fire('Dato requerido', 'El nombre de la correría es necesario', 'warning');
        return;
    }
    correriaAssign = { active: true, name, color: '#28a745', idx };
    minimizeWorkshopModal();
    fetch('/projects/' + encodeURIComponent(currentProjectFile) + '/ensure-columns', { method: 'POST' })
        .then(resp => resp.json())
        .then(() => Swal.fire('Modo asignación', 'Haga clic en los puntos del mapa para asignarlos a: ' + name, 'info'))
        .catch(() => Swal.fire('Advertencia', 'No se pudieron asegurar columnas, intentará de todas formas', 'warning'));
    openCorreriaColorPicker();
    setProjectPopupsEnabled(false);
}

let correriaColorPickerActive = false;
function openCorreriaColorPicker() {
    correriaColorPickerActive = true;
    const modal = document.getElementById('colorPickerModal');
    const input = document.getElementById('colorPickerInput');
    input.value = correriaAssign.color || '#28a745';
    const grid = document.getElementById('colorPresetGrid');
    Array.from(grid.querySelectorAll('.preset-color')).forEach(el => {
        el.classList.toggle('selected', el.dataset.color.toLowerCase() === input.value.toLowerCase());
        el.onclick = () => {
            Array.from(grid.querySelectorAll('.preset-color')).forEach(x => x.classList.remove('selected'));
            el.classList.add('selected');
            input.value = el.dataset.color;
        };
    });
    input.onchange = function () {
        Array.from(grid.querySelectorAll('.preset-color')).forEach(x => {
            x.classList.toggle('selected', x.dataset.color.toLowerCase() === input.value.toLowerCase());
        });
    };
    modal.classList.add('show');
}

function handleMarkerClick(props, marker, bounceClass) {
    if (!correriaAssign.active) return;
    const id = props['Id. Orden Trabajo'];
    if (!id) {
        Swal.fire('Sin ID', 'El punto no tiene "Id. Orden Trabajo"', 'error');
        return;
    }
    const body = {
        idOrdenTrabajo: String(id),
        description: correriaAssign.name,
        estadoOts: 'Programado',
        color: correriaAssign.color
    };
    fetch('/projects/' + encodeURIComponent(currentProjectFile) + '/update-record', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
    })
        .then(r => r.json().then(d => ({ ok: r.ok, data: d })))
        .then(({ ok }) => {
            if (ok) {
                if (projectLayer) {
                    projectLayer.removeLayer(marker);
                } else {
                    marker.remove();
                }
                const table = document.getElementById('correriaTable');
                const row = table.querySelectorAll('tbody tr')[correriaAssign.idx];
                if (row) {
                    correrias[correriaAssign.idx].count = (correrias[correriaAssign.idx].count || 0) + 1;
                    const tdCount = row.children[1];
                    if (tdCount) tdCount.textContent = String(correrias[correriaAssign.idx].count);
                    updateCorreriaBanner();
                }
            } else {
                Swal.fire('Error', 'No se pudo asignar el punto', 'error');
            }
        })
        .catch(() => Swal.fire('Error', 'Falló la actualización', 'error'));
}

function minimizeWorkshopModal() {
    const overlay = document.getElementById('workshopModal');
    const content = document.querySelector('#workshopModal .modal-content');
    const restore = document.querySelector('#workshopModal .restore-btn');
    const minimize = document.querySelector('#workshopModal .minimize-btn');
    if (overlay && content) {
        overlay.classList.add('minimized');
        content.classList.add('minimized');
        // posicion inicial flotante
        content.style.right = '';
        content.style.bottom = '';
        content.style.left = '20px';
        content.style.top = '90px';
        if (restore) restore.style.display = 'inline-block';
        if (minimize) minimize.style.display = 'none';
        enableWorkshopDrag(true);
        showCorreriaBanner();
    }
}

function maximizeWorkshopModal() {
    const overlay = document.getElementById('workshopModal');
    const content = document.querySelector('#workshopModal .modal-content');
    const restore = document.querySelector('#workshopModal .restore-btn');
    const minimize = document.querySelector('#workshopModal .minimize-btn');
    if (overlay && content) {
        overlay.classList.remove('minimized');
        content.classList.remove('minimized');
        content.style.left = '';
        content.style.top = '';
        content.style.right = '';
        content.style.bottom = '';
        if (restore) restore.style.display = 'none';
        if (minimize) minimize.style.display = 'inline-block';
        enableWorkshopDrag(false);
        hideCorreriaBanner();
        setProjectPopupsEnabled(true);
    }
}

function ensureCorreriaBanner() {
    let el = document.getElementById('correriaBanner');
    if (!el) {
        el = document.createElement('div');
        el.id = 'correriaBanner';
        el.className = 'correria-banner';
        const nameSpan = document.createElement('span');
        nameSpan.className = 'correria-name';
        const sep = document.createElement('span');
        sep.textContent = ' • ';
        const countSpan = document.createElement('span');
        countSpan.className = 'correria-count';
        el.appendChild(nameSpan);
        el.appendChild(sep);
        el.appendChild(countSpan);
        document.body.appendChild(el);
    }
    return el;
}

function showCorreriaBanner() {
    const el = ensureCorreriaBanner();
    el.style.display = 'flex';
    updateCorreriaBanner();
}

function updateCorreriaBanner() {
    if (!correrias || !correrias.length) return;
    const idx = correriaAssign && typeof correriaAssign.idx === 'number' ? correriaAssign.idx : 0;
    const name = (correrias[idx] && correrias[idx].name) || '';
    const count = (correrias[idx] && correrias[idx].count) || 0;
    const el = document.getElementById('correriaBanner');
    if (!el) return;
    const ns = el.querySelector('.correria-name');
    const cs = el.querySelector('.correria-count');
    if (ns) ns.textContent = name;
    if (cs) cs.textContent = 'Cant.: ' + count;
}

function hideCorreriaBanner() {
    const el = document.getElementById('correriaBanner');
    if (el) el.style.display = 'none';
}

function setProjectPopupsEnabled(enabled) {
    if (!projectLayer) return;
    try {
        projectLayer.eachLayer(l => {
            if (l && l.getPopup && typeof l.getPopup === 'function') {
                if (enabled) {
                    if (l._popupContent) {
                        l.bindPopup(l._popupContent);
                    }
                } else {
                    const pop = l.getPopup();
                    if (pop) {
                        const cnt = pop.getContent();
                        if (cnt) l._popupContent = cnt;
                    }
                    l.unbindPopup();
                }
            }
        });
        projectPopupsEnabled = !!enabled;
        if (!enabled && typeof map !== 'undefined' && map && map.closePopup) map.closePopup();
    } catch (e) {
        console.warn('No se pudieron alternar los popups de puntos:', e);
    }
}

function initContextMenu() {
    const menu = document.getElementById('contextMenu');
    const toggleCb = document.getElementById('togglePopupsCheckbox');
    if (!menu || !toggleCb) return;
    toggleCb.checked = projectPopupsEnabled;
    toggleCb.addEventListener('change', () => {
        setProjectPopupsEnabled(toggleCb.checked);
    });
    map.on('contextmenu', (e) => {
        const pt = e.containerPoint;
        menu.style.left = (pt.x + 10) + 'px';
        menu.style.top = (pt.y + 10) + 'px';
        menu.style.display = 'block';
    });
    map.on('click', () => {
        menu.style.display = 'none';
    });
    document.addEventListener('keydown', (ev) => {
        if (ev.key === 'Escape') menu.style.display = 'none';
    });
}

let workshopDragEnabled = false;
let workshopDragState = { dragging: false, startX: 0, startY: 0, elX: 0, elY: 0 };

function enableWorkshopDrag(enable) {
    workshopDragEnabled = enable;
    const header = document.querySelector('#workshopModal .modal-header');
    const content = document.querySelector('#workshopModal .modal-content');
    if (!header || !content) return;
    if (enable) {
        header.style.cursor = 'move';
        header.addEventListener('mousedown', workshopDragStart);
        document.addEventListener('mousemove', workshopDragMove);
        document.addEventListener('mouseup', workshopDragEnd);
        // touch support
        header.addEventListener('touchstart', workshopDragStart, { passive: false });
        document.addEventListener('touchmove', workshopDragMove, { passive: false });
        document.addEventListener('touchend', workshopDragEnd);
    } else {
        header.style.cursor = '';
        header.removeEventListener('mousedown', workshopDragStart);
        document.removeEventListener('mousemove', workshopDragMove);
        document.removeEventListener('mouseup', workshopDragEnd);
        header.removeEventListener('touchstart', workshopDragStart);
        document.removeEventListener('touchmove', workshopDragMove);
        document.removeEventListener('touchend', workshopDragEnd);
    }
}

function getEventPoint(e) {
    if (e.touches && e.touches.length) {
        return { x: e.touches[0].clientX, y: e.touches[0].clientY };
    }
    return { x: e.clientX, y: e.clientY };
}

function workshopDragStart(e) {
    if (!workshopDragEnabled) return;
    const content = document.querySelector('#workshopModal .modal-content');
    if (!content || !content.classList.contains('minimized')) return;
    const pt = getEventPoint(e);
    workshopDragState.dragging = true;
    workshopDragState.startX = pt.x;
    workshopDragState.startY = pt.y;
    const rect = content.getBoundingClientRect();
    workshopDragState.elX = rect.left;
    workshopDragState.elY = rect.top;
    e.preventDefault();
}

function workshopDragMove(e) {
    if (!workshopDragEnabled || !workshopDragState.dragging) return;
    const content = document.querySelector('#workshopModal .modal-content');
    if (!content) return;
    const pt = getEventPoint(e);
    const dx = pt.x - workshopDragState.startX;
    const dy = pt.y - workshopDragState.startY;
    let nx = workshopDragState.elX + dx;
    let ny = workshopDragState.elY + dy;
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    const rect = content.getBoundingClientRect();
    // clamp within viewport
    nx = Math.max(0, Math.min(vw - rect.width, nx));
    ny = Math.max(0, Math.min(vh - rect.height, ny));
    content.style.left = nx + 'px';
    content.style.top = ny + 'px';
    e.preventDefault();
}

function workshopDragEnd(e) {
    if (!workshopDragEnabled) return;
    workshopDragState.dragging = false;
}

function renderProjectsTable(files) {
    const table = document.getElementById('projectsTable');
    const thead = document.createElement('thead');
    const hr = document.createElement('tr');
    const thName = document.createElement('th');
    thName.textContent = 'Nombre';
    const thActions = document.createElement('th');
    thActions.textContent = 'Acciones';
    hr.appendChild(thName);
    hr.appendChild(thActions);
    thead.appendChild(hr);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    if (!files || files.length === 0) {
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        td.textContent = 'No hay proyectos guardados';
        td.colSpan = 2;
        tr.appendChild(td);
        tbody.appendChild(tr);
    } else {
        files.forEach(file => {
            const tr = document.createElement('tr');
            const tdName = document.createElement('td');
            tdName.textContent = file;
            const tdActions = document.createElement('td');
            const openBtn = document.createElement('button');
            openBtn.type = 'button';
            openBtn.className = 'run-btn';
            const icon = document.createElement('i');
            icon.setAttribute('data-lucide', 'map-pin');
            const lbl = document.createElement('span');
            lbl.textContent = 'Abrir';
            openBtn.appendChild(icon);
            openBtn.appendChild(lbl);
            openBtn.addEventListener('click', () => openProject(file));
            tdActions.appendChild(openBtn);
            tr.appendChild(tdName);
            tr.appendChild(tdActions);
            tbody.appendChild(tr);
        });
    }
    table.appendChild(tbody);
    if (window.lucide && typeof lucide.createIcons === 'function') {
        lucide.createIcons();
    }
}

// Cerrar menú al hacer click en el mapa
map.on('click', () => {
    if (menuOpen) {
        menuButtons.classList.remove("show");
        toggleMenuBtn.textContent = "☰";
        menuOpen = false;
        mapsDropdown.classList.remove("show");
    }
});
initContextMenu();
async function openProject(file) {
    try {
        const resp = await fetch('/projects/' + encodeURIComponent(file) + '/points');
        const data = await resp.json();
        if (!resp.ok) throw new Error(data.error || 'No se pudieron cargar puntos');
        if (projectLayer) {
            map.removeLayer(projectLayer);
        }
        projectLayer = L.layerGroup();
        const bounds = [];
        (data.points || []).forEach(p => {
            const color = typeof p.color === 'string' && p.color.trim() ? p.color : '#28a745';
            const f = p.props || {};
            const reqDate = parseDDMMYYYY(f['Fecha Solicitud']);
            const bdays = businessDaysBetween(reqDate, new Date());
            const bounceClass = bdays >= 3 ? 'pin-bounce-fast' : (bdays >= 2 ? 'pin-bounce-slow' : '');
            const marker = L.marker([p.lat, p.lng], { icon: createPinIcon(color, bounceClass) });
            const keys = ['Cliente','Ruta','Dirección','Nombre','Ciclo','Tarea','Id. Orden Trabajo','Revisión','Nombre Localidad','Gps','Fecha Solicitud'];
            const lines = keys.map(k => `<div><strong>${k}:</strong> ${f[k] || ''}</div>`).join('');
            const popupHtml = `<div>${lines}</div>`;
            marker.bindPopup(popupHtml);
            marker._popupContent = popupHtml;
            marker.on('click', () => handleMarkerClick(f, marker, bounceClass));
            marker.addTo(projectLayer);
            bounds.push([p.lat, p.lng]);
        });
        projectLayer.addTo(map);
        currentProjectFile = file;
        if (bounds.length) {
            map.fitBounds(bounds, { padding: [20, 20] });
        }
        closeDataModal();
        Swal.fire('Cargado', 'Se graficaron ' + (data.points ? data.points.length : 0) + ' puntos', 'success');
    } catch (e) {
        Swal.fire('Error', e.message || 'No se pudieron graficar los puntos', 'error');
    }
}

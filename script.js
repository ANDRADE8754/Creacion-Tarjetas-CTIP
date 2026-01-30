let workbookData = null;
let tarjetasGeneradas = [];
let tarjetasFiltradas = [];
let vistaActual = 'todas';
let paginaActual = 0;

const DIAS_SEMANA_ARRAY = ['L', 'M', 'MI', 'J', 'V', 'S', 'D'];
const MESES_CONST = {
    'ENERO': 0, 'FEBRERO': 1, 'MARZO': 2, 'ABRIL': 3, 'MAYO': 4, 'JUNIO': 5,
    'JULIO': 6, 'AGOSTO': 7, 'SEPTIEMBRE': 8, 'OCTUBRE': 9, 'NOVIEMBRE': 10, 'DICIEMBRE': 11
};

document.getElementById('fileInput').addEventListener('change', handleFileUpload);

window.addEventListener('beforeunload', function(e) {
    if (workbookData || tarjetasGeneradas.length > 0) {
        e.preventDefault();
        e.returnValue = '¿Salir?';
        return e.returnValue;
    }
});

const scrollToTopBtn = document.getElementById('scrollToTop');
window.addEventListener('scroll', function() {
    if (window.scrollY > 500) scrollToTopBtn.classList.add('show');
    else scrollToTopBtn.classList.remove('show');
});
scrollToTopBtn.addEventListener('click', function() {
    window.scrollTo({ top: 0, behavior: 'smooth' });
});

async function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    if (workbookData || tarjetasGeneradas.length > 0) {
        if (!confirm('⚠️ Ya hay datos. ¿Cargar nuevo archivo?')) {
            e.target.value = '';
            return;
        }
    }
    
    const uploadSection = document.querySelector('.upload-section');
    uploadSection.innerHTML = '<div class="loading">Procesando archivo...</div>';
    
    try {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                workbookData = workbook;
                validarYProcesarExcel(workbook);
            } catch (error) {
                mostrarError('Error de lectura', error.message, '');
            }
        };
        reader.readAsArrayBuffer(file);
    } catch (error) {
        mostrarError('Error', error.message, '');
    }
}

function mostrarError(titulo, mensaje, detalle) {
    const uploadSection = document.querySelector('.upload-section');
    uploadSection.innerHTML = `
        <div style="background: #fee; border: 2px solid #fcc; padding: 2rem; border-radius: 12px; text-align: center;">
            <h3 style="color: #c41e3a;">${titulo}</h3>
            <p>${mensaje}</p>
            <p style="font-size:0.8rem;color:#666">${detalle}</p>
            <button class="upload-btn" onclick="location.reload()">Reintentar</button>
        </div>
    `;
}



function detectarInicio(hoja) {
    // Busca el número 1 en las primeras 10 columnas y 40 filas
    for(let r=3; r<=40; r++) {
        for(let c=0; c<=10; c++) { 
            const cell = hoja[XLSX.utils.encode_cell({r:r, c:c})];
            if(cell && String(cell.v).trim() == '1') {
                return {r: r, c: c};
            }
        }
    }
    return null;
}

// Detectar si una columna es una columna de días (números consecutivos 1-31)
function esColumnaDias(hoja, col, startRow) {
    let contadorConsecutivos = 0;
    let ultimoNumero = 0;
    
    for(let r = startRow; r <= startRow + 31; r++) {
        const cell = hoja[XLSX.utils.encode_cell({r:r, c:col})];
        if(cell && cell.v) {
            const num = parseInt(String(cell.v).trim());
            if(!isNaN(num) && num >= 1 && num <= 31) {
                if(num === ultimoNumero + 1 || ultimoNumero === 0) {
                    contadorConsecutivos++;
                    ultimoNumero = num;
                }
            }
        }
    }
    
    // Si tiene al menos 20 números consecutivos, es una columna de días
    return contadorConsecutivos >= 20;
}

function validarYProcesarExcel(workbook) {
    try {
        let hoja = buscarHojaCuadro(workbook);
        if (!hoja) throw new Error("No se encontró hoja válida.");

        // 1. ENCONTRAR DÍA 1
        const inicio = detectarInicio(hoja);
        if (!inicio) {
            mostrarError('Faltan Días', 'No se encontró el día 1 en las primeras columnas.', '');
            return;
        }

        // 2. ESCANEO DE DISCOS (EXCLUYENDO COLUMNAS DE DÍAS)
        const discosUnicos = new Set();
        const startRow = inicio.r;
        const colDia = inicio.c;
        
        // Detectar dónde empiezan los datos (después de la columna de días y letras)
        let colDatos = colDia + 1;
        const cellNext = hoja[XLSX.utils.encode_cell({r:startRow, c:colDia + 1})];
        if (cellNext && DIAS_SEMANA_ARRAY.includes(String(cellNext.v).toUpperCase())) {
            colDatos = colDia + 2;
        }
        
        // Identificar columnas de días para excluirlas
        const columnasExcluidas = new Set();
        columnasExcluidas.add(colDia); // La columna principal de días
        for (let c = colDatos; c <= 150; c++) {
            if (esColumnaDias(hoja, c, startRow)) {
                columnasExcluidas.add(c);
            }
        }
        
        for (let r = startRow; r <= startRow + 32; r++) {
            for (let c = colDatos; c <= 150; c++) {
                // Saltar columnas de días
                if (columnasExcluidas.has(c)) continue;
                
                const cell = hoja[XLSX.utils.encode_cell({r:r, c:c})];
                if (!cell || !cell.v) continue;

                const valStr = String(cell.v).trim().toUpperCase();
                
                // Ignorar letras de días
                if (DIAS_SEMANA_ARRAY.includes(valStr)) continue;
                
                const valNum = parseInt(valStr);
                
                // Validar si es un disco real
                if (!isNaN(valNum) && valNum > 0 && valNum < 1000) {
                    discosUnicos.add(valNum);
                }
            }
        }

        if (discosUnicos.size === 0) {
            mostrarError('Sin Discos', 'No se encontraron números de disco válidos.', '');
            return;
        }

        procesarExcel(workbook, inicio);

    } catch (e) {
        mostrarError('Error Validación', e.message, '');
    }
}

function buscarHojaCuadro(workbook) {
    for(let name of workbook.SheetNames) {
        if(name.toUpperCase().includes('CUADRO')) return workbook.Sheets[name];
    }
    for(let name of workbook.SheetNames) {
        if(name.match(/20\d{2}/)) return workbook.Sheets[name];
    }
    // Si no encuentra, busca meses
    for(let name of workbook.SheetNames) {
        if(MESES_CONST[name.toUpperCase()] !== undefined) return workbook.Sheets[name];
    }
    return workbook.Sheets[workbook.SheetNames[0]];
}

function procesarExcel(workbook, inicio) {
    try {
        workbookData.modo = 'desde_cuadro';
        let hoja = buscarHojaCuadro(workbook);
        
        const startPoint = inicio || detectarInicio(hoja);
        const colDia = startPoint.c;
        
        // Detectar columna de inicio de datos
        let colDatos = colDia + 1;
        const cellNext = hoja[XLSX.utils.encode_cell({r:startPoint.r, c:colDia + 1})];
        if (cellNext && DIAS_SEMANA_ARRAY.includes(String(cellNext.v).toUpperCase())) {
            colDatos = colDia + 2;
        }
        
        // Identificar columnas de días para excluirlas
        const columnasExcluidas = new Set();
        columnasExcluidas.add(colDia);
        for (let c = colDatos; c <= 150; c++) {
            if (esColumnaDias(hoja, c, startPoint.r)) {
                columnasExcluidas.add(c);
            }
        }
        
        // Recolectar discos nuevamente para la configuración
        const discosUnicos = new Set();
        for (let r = startPoint.r; r <= startPoint.r + 32; r++) {
            for (let c = colDatos; c <= 150; c++) {
                // Saltar columnas de días
                if (columnasExcluidas.has(c)) continue;
                
                const cell = hoja[XLSX.utils.encode_cell({r:r, c:c})];
                if (cell && cell.v) {
                    const valStr = String(cell.v).trim().toUpperCase();
                    // Ignorar letras de días
                    if (DIAS_SEMANA_ARRAY.includes(valStr)) continue;
                    
                    const val = parseInt(valStr);
                    if (!isNaN(val) && val > 0 && val < 1000) discosUnicos.add(val);
                }
            }
        }
        
        let sociosConfig = {};
        Array.from(discosUnicos).sort((a,b)=>a-b).forEach(d => sociosConfig[d] = `Socio ${d}`);
        
        intentarDetectarDiaInicio(hoja);
        mostrarConfiguracion(sociosConfig);
        
        document.querySelector('.upload-section').innerHTML = `
            <div class="status-message">✅ Archivo cargado. ${discosUnicos.size} discos detectados.</div>
        `;

    } catch (e) {
        mostrarError('Error Procesamiento', e.message, '');
    }
}

function intentarDetectarDiaInicio(hoja) {
    let diaDetectado = 3; 
    let exito = false;
    const rangos = [];
    for(let r=0; r<20; r++) {
        for(let c=0; c<10; c++) rangos.push(XLSX.utils.encode_cell({r:r, c:c}));
    }
    
    for(let ref of rangos) {
        if(hoja[ref] && hoja[ref].v) {
            const txt = String(hoja[ref].v).toUpperCase();
            let anio = 2026;
            let mes = -1;
            
            const mAnio = txt.match(/20\d{2}/);
            if(mAnio) anio = parseInt(mAnio[0]);
            
            for(let k in MESES_CONST) {
                if(txt.includes(k)) { mes = MESES_CONST[k]; break; }
            }
            
            if(mes !== -1) {
                const fecha = new Date(anio, mes, 1);
                const map = [6,0,1,2,3,4,5];
                diaDetectado = map[fecha.getDay()];
                exito = true;
                
                const label = document.querySelector('#configSection label');
                if(label) {
                    label.innerHTML = `Día 1 detectado: <strong>${obtenerNombreDia(diaDetectado)}</strong>`;
                    label.style.color = "#2d5016";
                }
                break;
            }
        }
    }
    
    const select = document.getElementById('diaInicio');
    if(select) {
        select.value = diaDetectado;
        if(exito) select.style.backgroundColor = "#e8f5e9";
    }
}

function obtenerNombreDia(val) {
    return ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"][val] || "";
}

function mostrarConfiguracion(sociosConfig) {
    const grid = document.getElementById('sociosGrid');
    grid.innerHTML = '';
    const discos = Object.keys(sociosConfig).sort((a, b) => parseInt(a) - parseInt(b));
    discos.forEach(disco => {
        const div = document.createElement('div');
        div.className = 'socio-item';
        div.innerHTML = `<label>Disco ${disco}</label><input type="text" id="socio_${disco}" value="${sociosConfig[disco]}">`;
        grid.appendChild(div);
    });
    document.getElementById('configSection').style.display = 'block';
}

function generarTarjetas() {
    const diaInicio = document.getElementById('diaInicio').value;
    if (!diaInicio) return alert('Seleccione el día de inicio');

    const cardsContainer = document.getElementById('cardsContainer');
    cardsContainer.innerHTML = '<div class="loading">Generando...</div>';
    
    setTimeout(() => {
        try {
            tarjetasGeneradas = [];
            generarDesdeCuadroInteligente();
            
            if(tarjetasGeneradas.length > 0) {
                filtrarTarjetas();
                cambiarVista('todas');
                document.getElementById('previewSection').classList.add('active');
                document.getElementById('previewSection').scrollIntoView({ behavior: 'smooth' });
            } else {
                cardsContainer.innerHTML = '<div class="loading" style="color:red">No se generaron tarjetas.</div>';
            }
        } catch (error) {
            console.error(error);
            cardsContainer.innerHTML = `<div class="loading" style="color:red">Error: ${error.message}</div>`;
        }
    }, 100);
}

// --- LÓGICA DE GENERACIÓN ---
function generarDesdeCuadroInteligente() {
    const hoja = buscarHojaCuadro(workbookData);
    const offsetDia = parseInt(document.getElementById('diaInicio').value);
    
    // 1. Periodo
    let periodoTexto = "MES DE TRABAJO";
    for(let r=0; r<10; r++) {
        for(let c=0; c<5; c++) {
            const cell = hoja[XLSX.utils.encode_cell({r:r, c:c})];
            if(cell && cell.v && String(cell.v).length > 12 && String(cell.v).toUpperCase().includes('20')) {
                periodoTexto = cell.v;
            }
        }
    }

    // 2. Encontrar inicio
    const inicio = detectarInicio(hoja);
    if(!inicio) throw new Error("No se encontró el día 1.");
    
    const startRow = inicio.r;
    const colDia = inicio.c; // Columna principal de días

    // 3. Detectar dónde empiezan datos (Saltar letras)
    let colDatos = colDia + 1;
    const cellNext = hoja[XLSX.utils.encode_cell({r:startRow, c:colDia + 1})];
    if (cellNext && DIAS_SEMANA_ARRAY.includes(String(cellNext.v).toUpperCase())) {
        colDatos = colDia + 2;
    }

    // 4. DETECTAR FILA DE RUTAS
    let filaRutas = -1;
    let maxScore = -1;

    for(let r = 0; r < startRow; r++) {
        let score = 0;
        let esNumerica = true;
        for(let c = colDatos; c < colDatos + 15; c++) {
            const cell = hoja[XLSX.utils.encode_cell({r:r, c:c})];
            if (cell && cell.v) {
                if (isNaN(cell.v)) {
                    esNumerica = false;
                    if(String(cell.v).length > 3) score++;
                }
            }
        }
        if (!esNumerica && score > maxScore) {
            maxScore = score;
            filaRutas = r;
        }
    }
    if (filaRutas === -1) filaRutas = startRow - 1;

    // 5. OBTENER RUTAS (con limpieza)
    const rutas = [];
    const columnasExcluidas = new Set(); // Columnas que son de días
    let contadorVacios = 0;
    
    // Primero, identificar qué columnas son columnas de días
    for(let c = colDatos; c < colDatos + 150; c++) {
        if(esColumnaDias(hoja, c, startRow)) {
            columnasExcluidas.add(c);
        }
    }
    
    for(let c = colDatos; c < colDatos + 150; c++) { 
        // Si esta columna es una columna de días, no tiene ruta
        if(columnasExcluidas.has(c)) {
            rutas.push('');
            continue;
        }
        
        const cell = hoja[XLSX.utils.encode_cell({r:filaRutas, c:c})];
        const cellUp = hoja[XLSX.utils.encode_cell({r:filaRutas-1, c:c})];
        
        let valor = '';
        if (cell && cell.v) valor = cell.v;
        else if (cellUp && cellUp.v) valor = cellUp.v; 

        if (valor) {
            rutas.push(valor);
            contadorVacios = 0;
        } else {
            rutas.push('');
            contadorVacios++;
        }
        if (contadorVacios > 20) break;
    }
    const numCols = rutas.length;

    // 6. BARRIDO DE DATOS
    const cuadroCompleto = [];
    for(let r = startRow; r <= startRow + 32; r++) { 
        const cellDia = hoja[XLSX.utils.encode_cell({r:r, c:colDia})];
        if(!cellDia || cellDia.v == null) continue;
        
        const valStr = String(cellDia.v).toUpperCase();
        if(valStr.includes('NOTA') || valStr.includes('ELABORADO') || valStr.includes('OBSERV')) break;
        
        const diaNum = parseInt(valStr);
        if(isNaN(diaNum) || diaNum > 31) continue;

        let letraDia = "";
        let idxDia = (diaNum - 1 + offsetDia) % 7;
        if(idxDia < 0) idxDia += 7;
        letraDia = DIAS_SEMANA_ARRAY[idxDia];

        const asignaciones = [];
        
        for(let i = 0; i < numCols; i++) {
            const currentC = colDatos + i;
            const cell = hoja[XLSX.utils.encode_cell({r:r, c:currentC})];
            
            if(cell && cell.v) {
                const valStr = String(cell.v).trim();
                const dVal = parseInt(valStr);
                
                // VALIDACIONES CORREGIDAS:
                // Quitamos "dVal !== diaNum". 
                // Ahora confiamos en que si hay una ruta válida (rutas[i]), el número es un disco.
                if(!isNaN(dVal) && dVal > 0 && !DIAS_SEMANA_ARRAY.includes(valStr)) {
                    if (rutas[i] && rutas[i].trim() !== '') { // Solo si hay ruta válida
                        asignaciones.push({
                            disco: dVal,
                            ruta: rutas[i]
                        });
                    }
                }
            }
        }
        cuadroCompleto.push({ dia: diaNum, abrev: letraDia, asignaciones });
    }

    // 7. GENERAR OBJETOS
    const discosSet = new Set();
    cuadroCompleto.forEach(d => d.asignaciones.forEach(a => discosSet.add(a.disco)));
    
    const ordenSelect = document.getElementById('ordenDiscos');
    const orden = ordenSelect ? ordenSelect.value : 'numerico';
    let discos = Array.from(discosSet);
    if(orden === 'numerico') discos.sort((a,b)=>a-b);

    discos.forEach(disco => {
        const nombreSocio = (document.getElementById(`socio_${disco}`) || {}).value || `Socio ${disco}`;
        const datos = [];

        cuadroCompleto.forEach(diaData => {
            const asig = diaData.asignaciones.find(a => a.disco === disco);
            if(asig) {
                datos.push({
                    dia: diaData.dia,
                    abreviatura: diaData.abrev,
                    ruta: asig.ruta,
                    disco: disco
                });
            }
        });

        if(datos.length > 0) {
            tarjetasGeneradas.push({disco, nombre: nombreSocio, periodo: periodoTexto, datos: datos});
        }
    });
}

function cambiarVista(vista) {
    vistaActual = vista;
    paginaActual = 0;
    document.getElementById('btnVistaTodas').classList.toggle('active', vista === 'todas');
    document.getElementById('btnVistaIndividual').classList.toggle('active', vista !== 'todas');
    document.querySelectorAll('.pagination').forEach(p => p.style.display = vista === 'todas' ? 'none' : 'flex');
    filtrarTarjetas();
}

function filtrarTarjetas() {
    const term = document.getElementById('searchInput').value.toLowerCase().trim();
    if(term === '') tarjetasFiltradas = [...tarjetasGeneradas];
    else tarjetasFiltradas = tarjetasGeneradas.filter(t => t.nombre.toLowerCase().includes(term) || String(t.disco).includes(term));
    paginaActual = 0;
    renderizarTarjetas();
}

function renderizarTarjetas() {
    const container = document.getElementById('cardsContainer');
    container.innerHTML = '';
    if(tarjetasFiltradas.length === 0) {
        container.innerHTML = '<div class="loading">No hay resultados</div>';
        return;
    }
    if(vistaActual === 'todas') {
        container.style.display = 'grid';
        tarjetasFiltradas.forEach(t => container.appendChild(crearTarjetaHTML(t.disco, t.nombre, t.periodo, t.datos)));
    } else {
        container.style.display = 'flex';
        container.style.justifyContent = 'center';
        const t = tarjetasFiltradas[paginaActual];
        container.appendChild(crearTarjetaHTML(t.disco, t.nombre, t.periodo, t.datos));
        actualizarPaginacion();
    }
}

function crearTarjetaHTML(disco, nombre, periodo, datos) {
    const div = document.createElement('div');
    div.className = 'card-wrapper';
    let html = `
        <div class="card-logo"><img src="img/image.png" alt="Logo"></div>
        <div class="card-header"><div class="card-name">${nombre.toUpperCase()}</div><div class="card-disco">${disco}</div></div>
        <div style="text-align:center;margin-bottom:1rem;font-weight:600;color:var(--primary);">${periodo}</div>
        <table class="card-table">
            <thead><tr><th style="width:8%">DÍA</th><th style="width:8%">DÍA</th><th style="width:70%">RUTA</th><th style="width:14%">DISCO</th></tr></thead>
            <tbody>`;
    datos.forEach(f => {
        const esEsp = ['DISPONIBLE', 'LIBRE', 'PARADA'].some(x => f.ruta && f.ruta.toUpperCase().includes(x));
        html += `<tr><td class="col-dia">${f.dia}</td><td class="col-abrev">${f.abreviatura}</td><td class="col-ruta ${esEsp ? 'ruta-especial' : ''}">${f.ruta}</td><td>${f.disco}</td></tr>`;
    });
    html += '</tbody></table>';
    div.innerHTML = html;
    return div;
}

function cambiarPagina(dir) {
    if(dir === 'prev' && paginaActual > 0) paginaActual--;
    else if(dir === 'next' && paginaActual < tarjetasFiltradas.length - 1) paginaActual++;
    renderizarTarjetas();
}

function actualizarPaginacion() {
    const txt = `${paginaActual + 1} / ${tarjetasFiltradas.length}`;
    document.getElementById('pageInfo').textContent = txt;
    document.getElementById('pageInfoBottom').textContent = txt;
}

async function exportarExcel() {
    if(tarjetasGeneradas.length === 0) return alert('No hay datos');
    const workbook = new ExcelJS.Workbook();
    let list = tarjetasFiltradas;
    if(tarjetasFiltradas.length < tarjetasGeneradas.length) {
        if(confirm('¿Exportar TODAS? (Cancelar = Solo filtradas)')) list = tarjetasGeneradas;
    }
    let logoBuffer = null;
    try {
        const resp = await fetch('img/image.png');
        logoBuffer = await resp.arrayBuffer();
    } catch(e) {}

    for(const t of list) {
        const sheet = workbook.addWorksheet(`D${t.disco}`);
        if(logoBuffer) {
            const imgId = workbook.addImage({buffer: logoBuffer, extension: 'png'});
            sheet.addImage(imgId, {tl:{col:0,row:0}, ext:{width:500,height:60}});
            sheet.getRow(1).height = 45;
        }
        sheet.columns = [{width:8}, {width:8}, {width:60}, {width:10}];
        
        // Fusionar celdas A3:D3 para nombre del socio
        sheet.mergeCells('A3:D3');
        const celNom = sheet.getCell('A3');
        celNom.value = t.nombre.toUpperCase();
        celNom.font = {bold:true, size:14, name:'Bookman Old Style'};
        celNom.alignment = {horizontal:'center', vertical:'middle'};
        celNom.fill = {type:'pattern', pattern:'solid', fgColor:{argb:'FFFFF8DC'}};
        celNom.border = {
            top: {style:'thin', color:{argb:'FF000000'}},
            left: {style:'thin', color:{argb:'FF000000'}},
            bottom: {style:'thin', color:{argb:'FF000000'}},
            right: {style:'thin', color:{argb:'FF000000'}}
        };
        
        // Encabezados de la tabla
        const rowHead = sheet.getRow(4);
        ['DÍA','DÍA','RUTA','DISCO'].forEach((h, i) => {
            const c = rowHead.getCell(i+1);
            c.value = h;
            c.font = {bold:true, color:{argb:'FFFFFFFF'}, name:'Bookman Old Style'};
            c.fill = {type:'pattern', pattern:'solid', fgColor:{argb:'FFDC143C'}};
            c.alignment = {horizontal:'center', vertical:'middle'};
            c.border = {
                top: {style:'thin', color:{argb:'FF000000'}},
                left: {style:'thin', color:{argb:'FF000000'}},
                bottom: {style:'thin', color:{argb:'FF000000'}},
                right: {style:'thin', color:{argb:'FF000000'}}
            };
        });
        
        // Filas de datos
        t.datos.forEach((d, i) => {
            const r = sheet.getRow(i+5);
            r.values = [d.dia, d.abreviatura, d.ruta, d.disco];
            
            // Columna DÍA (número)
            r.getCell(1).fill = {type:'pattern', pattern:'solid', fgColor:{argb:'FF5b9ad5'}};
            r.getCell(1).font = {bold:true, color:{argb:'FFFFFFFF'}, name:'Bookman Old Style'};
            r.getCell(1).alignment = {horizontal:'center', vertical:'middle'};
            r.getCell(1).border = {
                top: {style:'thin', color:{argb:'FF000000'}},
                left: {style:'thin', color:{argb:'FF000000'}},
                bottom: {style:'thin', color:{argb:'FF000000'}},
                right: {style:'thin', color:{argb:'FF000000'}}
            };
            
            // Columna DÍA (letra)
            r.getCell(2).fill = {type:'pattern', pattern:'solid', fgColor:{argb:'FFDC143C'}};
            r.getCell(2).font = {bold:true, color:{argb:'FFFFFFFF'}, name:'Bookman Old Style'};
            r.getCell(2).alignment = {horizontal:'center', vertical:'middle'};
            r.getCell(2).border = {
                top: {style:'thin', color:{argb:'FF000000'}},
                left: {style:'thin', color:{argb:'FF000000'}},
                bottom: {style:'thin', color:{argb:'FF000000'}},
                right: {style:'thin', color:{argb:'FF000000'}}
            };
            
            // Columna RUTA
            const esEspecial = ['DISPONIBLE', 'LIBRE', 'PARADA'].some(x => d.ruta && d.ruta.toUpperCase().includes(x));
            r.getCell(3).font = {
                color:{argb: esEspecial ? 'FFFF0000' : 'FF000000'},
                name:'Bookman Old Style'
            };
            r.getCell(3).alignment = {wrapText:true, vertical:'middle'};
            r.getCell(3).border = {
                top: {style:'thin', color:{argb:'FF000000'}},
                left: {style:'thin', color:{argb:'FF000000'}},
                bottom: {style:'thin', color:{argb:'FF000000'}},
                right: {style:'thin', color:{argb:'FF000000'}}
            };
            
            // Columna DISCO
            r.getCell(4).font = {name:'Bookman Old Style'};
            r.getCell(4).alignment = {horizontal:'center', vertical:'middle'};
            r.getCell(4).border = {
                top: {style:'thin', color:{argb:'FF000000'}},
                left: {style:'thin', color:{argb:'FF000000'}},
                bottom: {style:'thin', color:{argb:'FF000000'}},
                right: {style:'thin', color:{argb:'FF000000'}}
            };
        });
    }
    const buf = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Tarjetas.xlsx';
    a.click();
}

function mostrarModalImpresion() {
    if(confirm('¿Imprimir ahora? Recuerde quitar márgenes y activar gráficos de fondo.')) {
        setTimeout(() => window.print(), 100);
    }
}
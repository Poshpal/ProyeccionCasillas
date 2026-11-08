/**
 * Catálogo de municipios de Chiapas (clave INEGI → nombre).
 * Claves según catálogo geoestadístico INEGI / IEPC.
 */
const MUNICIPIOS_CHIAPAS = {
    1: 'ACACOYAGUA', 2: 'ACALA', 3: 'ACAPETAHUA', 4: 'ALTAMIRANO',
    5: 'AMATÁN', 6: 'AMATENANGO DE LA FRONTERA', 7: 'AMATENANGO DEL VALLE',
    8: 'ÁNGEL ALBINO CORZO', 9: 'ARRIAGA', 10: 'BEJUCAL DE OCAMPO',
    11: 'BELLA VISTA', 12: 'BERRIOZÁBAL', 13: 'BOCHIL', 14: 'EL BOSQUE',
    15: 'CACAHOATÁN', 16: 'CATAZAJÁ', 17: 'CINTALAPA', 18: 'COAPILLA',
    19: 'COMITÁN DE DOMÍNGUEZ', 20: 'LA CONCORDIA', 21: 'COPAINALÁ',
    22: 'CHALCHIHUITÁN', 23: 'CHAMULA', 24: 'CHANAL', 25: 'CHAPULTENANGO',
    26: 'CHENALHÓ', 27: 'CHIAPA DE CORZO', 28: 'CHIAPILLA', 29: 'CHICOASÉN',
    30: 'CHICOMUSELO', 31: 'CHILÓN', 32: 'ESCUINTLA', 33: 'FRANCISCO LEÓN',
    34: 'FRONTERA COMALAPA', 35: 'FRONTERA HIDALGO', 36: 'LA GRANDEZA',
    37: 'HUEHUETÁN', 38: 'HUIXTÁN', 39: 'HUITIUPÁN', 40: 'HUIXTLA',
    41: 'LA INDEPENDENCIA', 42: 'IXHUATÁN', 43: 'IXTACOMITÁN', 44: 'IXTAPA',
    45: 'IXTAPANGAJOYA', 46: 'JIQUIPILAS', 47: 'JITOTOL', 48: 'JUÁREZ',
    49: 'LARRÁINZAR', 50: 'LA LIBERTAD', 51: 'MAPASTEPEC', 52: 'LAS MARGARITAS',
    53: 'MAZAPA DE MADERO', 54: 'MAZATÁN', 55: 'METAPA', 56: 'MITONTIC',
    57: 'MOTOZINTLA', 58: 'NICOLÁS RUIZ', 59: 'OCOSINGO', 60: 'OCOTEPEC',
    61: 'OCOZOCOAUTLA DE ESPINOSA', 62: 'OSTUACÁN', 63: 'OSUMACINTA',
    64: 'OXCHUC', 65: 'PALENQUE', 66: 'PANTELHÓ', 67: 'PANTEPEC',
    68: 'PICHUCALCO', 69: 'PIJIJIAPAN', 70: 'EL PORVENIR', 71: 'VILLA COMALTITLÁN',
    72: 'PUEBLO NUEVO SOLISTAHUACÁN', 73: 'RAYÓN', 74: 'REFORMA',
    75: 'LAS ROSAS', 76: 'SABANILLA', 77: 'SALTO DE AGUA',
    78: 'SAN CRISTÓBAL DE LAS CASAS', 79: 'SAN FERNANDO', 80: 'SILTEPEC',
    81: 'SIMOJOVEL', 82: 'SITALÁ', 83: 'SOCOLTENANGO', 84: 'SOLOSUCHIAPA',
    85: 'SOYALÓ', 86: 'SUCHIAPA', 87: 'SUCHIATE', 88: 'SUNUAPA',
    89: 'TAPACHULA', 90: 'TAPALAPA', 91: 'TAPILULA', 92: 'TECPATÁN',
    93: 'TENEJAPA', 94: 'TEOPISCA', 96: 'TILA', 97: 'TONALÁ',
    98: 'TOTOLAPA', 99: 'LA TRINITARIA', 100: 'TUMBALÁ',
    101: 'TUXTLA GUTIÉRREZ', 102: 'TUXTLA CHICO', 103: 'TUZANTÁN',
    104: 'TZIMOL', 105: 'UNIÓN JUÁREZ', 106: 'VENUSTIANO CARRANZA',
    107: 'VILLA CORZO', 108: 'VILLAFLORES', 109: 'YAJALÓN',
    110: 'SAN LUCAS', 111: 'ZINACANTÁN', 112: 'SAN JUAN CANCUC',
    113: 'ALDAMA', 114: 'BENEMÉRITO DE LAS AMÉRICAS', 115: 'MARAVILLA TENEJAPA',
    116: 'MARQUÉS DE COMILLAS', 117: 'MONTECRISTO DE GUERRERO',
    118: 'SAN ANDRÉS DURAZNAL', 119: 'SANTIAGO EL PINAR',
    120: 'CAPITÁN LUIS ÁNGEL VIDAL', 121: 'RINCÓN CHAMULA SAN PEDRO',
    122: 'EL PARRAL', 123: 'EMILIANO ZAPATA', 124: 'MEZCALAPA',
    125: 'HONDURAS DE LA SIERRA'
};

/** Devuelve el nombre del municipio o la clave si no está en el catálogo */
function nombreMunicipio(clave) {
    const n = Number(String(clave).trim());
    if (!n) return String(clave).trim().toUpperCase();
    return MUNICIPIOS_CHIAPAS[n] || String(clave).trim().toUpperCase();
}

/**
 * Convierte el código de tipo de casilla a texto legible para la etiqueta.
 * Ejemplos:
 *   B      → BÁSICA
 *   C1     → CONTIGUA 1
 *   C2     → CONTIGUA 2
 *   E1     → EXTRAORDINARIA 1
 *   E1C1   → EXTRAORDINARIA 1 CONTIGUA 1
 *   E2C3   → EXTRAORDINARIA 2 CONTIGUA 3
 *   S      → ESPECIAL
 */
function tipoATexto(tipo) {
    const t = String(tipo || '').trim().toUpperCase();
    if (!t) return '';

    if (t === 'B') return 'BÁSICA';
    if (t === 'S') return 'ESPECIAL';

    // Contigua pura: C1, C2, C12...
    const mCont = t.match(/^C(\d+)$/);
    if (mCont) return `CONTIGUA ${mCont[1]}`;

    // Extraordinaria pura: E1, E2...
    const mExt = t.match(/^E(\d+)$/);
    if (mExt) return `EXTRAORDINARIA ${mExt[1]}`;

    // Extraordinaria + Contigua: E1C1, E2C3...
    const mExtCont = t.match(/^E(\d+)C(\d+)$/);
    if (mExtCont) return `EXTRAORDINARIA ${mExtCont[1]} CONTIGUA ${mExtCont[2]}`;

    // Si no coincide, devolver tal cual
    return t;
}

/** Rellena la sección a 4 dígitos: 35 → 0035 */
function seccionPadded(sec) {
    const s = String(sec || '').trim();
    if (!s) return '';
    return s.padStart(4, '0');
}

/** Plantillas de hojas de etiquetas autoadheribles */
const PLANTILLAS = {
    '5163':     { nombre: 'Avery 5163 / envío', porHoja: 10, cols: 2, filas: 5, desc: '≈ 10 × 5 cm — ideal para casillas' },
    '5395':     { nombre: 'Avery 5395 / gafete', porHoja: 8,  cols: 2, filas: 4, desc: '≈ 9 × 6 cm' },
    '5164':     { nombre: 'Avery 5164 / grande', porHoja: 6,  cols: 2, filas: 3, desc: '≈ 10 × 8.5 cm — etiquetas grandes' },
    '5162':     { nombre: 'Avery 5162 / mediana', porHoja: 14, cols: 2, filas: 7, desc: '≈ 10 × 3.4 cm — más compacta' },
    'custom10': { nombre: 'Genérica 10/hoja', porHoja: 10, cols: 2, filas: 5, desc: '2 columnas × 5 filas' },
    'custom8':  { nombre: 'Genérica 8/hoja',  porHoja: 8,  cols: 2, filas: 4, desc: '2 columnas × 4 filas' },
    'custom6':  { nombre: 'Genérica 6/hoja',  porHoja: 6,  cols: 2, filas: 3, desc: '2 columnas × 3 filas' }
};

function actualizarInfoPlantilla() {
    const key = document.getElementById('plantilla-etiqueta').value;
    const p = PLANTILLAS[key];
    if (!p) return;
    document.getElementById('info-plantilla').textContent =
        `${p.nombre}: ${p.porHoja} etiquetas por hoja (${p.cols} columnas × ${p.filas} filas). ${p.desc}`;
}

// Inicializar info al cargar
document.addEventListener('DOMContentLoaded', actualizarInfoPlantilla);

function generarEtiquetas() {
    const archivo = document.getElementById('archivo-proyeccion').files[0];
    if (!archivo) {
        alert('Por favor, sube el archivo de proyección de casillas.');
        return;
    }

    const entidad = (document.getElementById('entidad').value || 'CHIAPAS').trim().toUpperCase();
    const numEtiquetas = Math.min(20, Math.max(1, Number(document.getElementById('num-etiquetas').value) || 1));
    const plantillaKey = document.getElementById('plantilla-etiqueta').value;
    const plantilla = PLANTILLAS[plantillaKey] || PLANTILLAS['5163'];

    const loading = document.getElementById('loading');
    loading.style.display = 'flex';

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = e.target.result;
            const wb = XLSX.read(data, { type: 'binary' });
            const sheet = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });

            if (sheet.length < 2) {
                throw new Error('El archivo está vacío o no tiene datos.');
            }

            const header = sheet[0].map(h => String(h || '').trim().toLowerCase());

            let idxDist = header.findIndex(h => h.includes('distrito'));
            let idxMun  = header.findIndex(h => h.includes('municipio'));
            let idxSec  = header.findIndex(h => h.includes('sección') || h.includes('seccion'));
            let idxTipo = header.findIndex(h => h.includes('tipo'));

            if (idxDist === -1) idxDist = 0;
            if (idxMun  === -1) idxMun  = 1;
            if (idxSec  === -1) idxSec  = 2;
            if (idxTipo === -1) idxTipo = 3;

            // Recolectar casillas
            const casillas = [];
            for (let i = 1; i < sheet.length; i++) {
                const row = sheet[i];
                if (!row || row.length < 4) continue;

                const distrito = String(row[idxDist] ?? '').trim();
                const mun      = String(row[idxMun]  ?? '').trim();
                const sec      = String(row[idxSec]  ?? '').trim();
                const tipo     = String(row[idxTipo] ?? '').trim().toUpperCase();

                if (!mun || !sec || !tipo) continue;

                casillas.push({ distrito, mun, sec, tipo });
            }

            if (casillas.length === 0) {
                throw new Error('No se encontraron casillas válidas en el archivo.');
            }

            // Agrupar por municipio
            const porMunicipio = new Map();
            for (const c of casillas) {
                if (!porMunicipio.has(c.mun)) porMunicipio.set(c.mun, []);
                porMunicipio.get(c.mun).push(c);
            }

            // Expandir copias
            function expandir(lista) {
                const out = [];
                for (const c of lista) {
                    for (let k = 0; k < numEtiquetas; k++) out.push(c);
                }
                return out;
            }

            function htmlEtiqueta(c) {
                const tipoTxt = tipoATexto(c.tipo);
                const secPad  = seccionPadded(c.sec);
                const munNom  = nombreMunicipio(c.mun);
                return `
                    <div class="etiqueta">
                        <div class="linea"><strong>ENTIDAD: ${entidad}</strong></div>
                        <div class="linea"><strong>DISTRITO: ${c.distrito}</strong></div>
                        <div class="linea"><strong>MUNICIPIO: ${munNom}</strong></div>
                        <div class="linea"><strong>SECCIÓN: ${secPad}</strong></div>
                        <div class="linea"><strong>CASILLA: ${tipoTxt}</strong></div>
                    </div>
                `;
            }

            // Construye HTML con TABLAS (más fiable para PDF que CSS grid)
            function construirHojas(listaCasillas) {
                const cols = plantilla.cols;
                const filas = plantilla.filas;
                const porHoja = plantilla.porHoja;
                let html = '';
                let numHojas = 0;
                let i = 0;

                while (i < listaCasillas.length) {
                    numHojas++;
                    // Cada hoja es una tabla completa que no se parte
                    html += `<table class="hoja-etiquetas" style="width:100%;border-collapse:separate;border-spacing:6px;page-break-after:always;break-after:page;">`;
                    for (let f = 0; f < filas; f++) {
                        html += '<tr>';
                        for (let c = 0; c < cols; c++) {
                            if (i < listaCasillas.length) {
                                html += `<td style="width:${Math.floor(100 / cols)}%;vertical-align:top;padding:0;">${htmlEtiqueta(listaCasillas[i])}</td>`;
                                i++;
                            } else {
                                html += `<td style="width:${Math.floor(100 / cols)}%;"></td>`;
                            }
                        }
                        html += '</tr>';
                    }
                    html += '</table>';
                }
                return { html, numHojas, total: listaCasillas.length };
            }

            // Vista previa: solo primeras ~40 etiquetas para no saturar el navegador
            const todasPreview = expandir(casillas).slice(0, 40);
            const preview = construirHojas(todasPreview);

            // Datos por municipio para descarga
            const munsOrdenados = Array.from(porMunicipio.keys()).sort((a, b) => Number(a) - Number(b));
            const datosPorMun = [];
            let totalEtiquetas = 0;
            let totalHojasEst = 0;

            for (const mun of munsOrdenados) {
                const lista = expandir(porMunicipio.get(mun));
                const res = construirHojas(lista);
                totalEtiquetas += res.total;
                totalHojasEst += res.numHojas;
                datosPorMun.push({
                    clave: mun,
                    nombre: nombreMunicipio(mun),
                    casillas: porMunicipio.get(mun).length,
                    etiquetas: res.total,
                    hojas: res.numHojas,
                    html: res.html,
                    lista: lista   // casillas expandidas para regenerar PDF
                });
            }

            // Resumen
            document.getElementById('resumen-etiquetas').innerHTML = `
                <h3>Resumen</h3>
                <ul style="list-style:none;padding:0;margin:10px 0 0 0;">
                    <li><strong>Municipios:</strong> ${munsOrdenados.length}</li>
                    <li><strong>Casillas en el archivo:</strong> ${casillas.length.toLocaleString('es-MX')}</li>
                    <li><strong>Etiquetas por casilla:</strong> ${numEtiquetas}</li>
                    <li><strong>Total de etiquetas:</strong> ${totalEtiquetas.toLocaleString('es-MX')}</li>
                    <li><strong>Plantilla:</strong> ${plantilla.nombre} (${plantilla.porHoja} por hoja)</li>
                    <li><strong>Hojas estimadas (total):</strong> ${totalHojasEst.toLocaleString('es-MX')}</li>
                    <li><strong>Entidad:</strong> ${entidad}</li>
                </ul>
                <p style="margin:12px 0 0;font-size:0.9rem;color:#64748b;">
                    Usa el botón <strong>Imprimir / PDF</strong> de cada municipio: se abre una ventana lista para imprimir
                    o para <strong>Guardar como PDF</strong> (elige esa opción en el diálogo de impresión del navegador).
                    Permite ventanas emergentes si el navegador las bloquea.
                </p>
            `;

            // Lista de municipios con botones de descarga
            let listaHTML = `
                <div class="stats-card" style="opacity:1;transform:none;">
                    <h3>Descargar PDF por municipio</h3>
                    <table style="width:100%;border-collapse:collapse;margin-top:12px;font-size:0.92rem;">
                        <thead>
                            <tr style="background:#edf2f7;">
                                <th style="padding:8px;border:1px solid #e2e8f0;">Municipio</th>
                                <th style="padding:8px;border:1px solid #e2e8f0;">Casillas</th>
                                <th style="padding:8px;border:1px solid #e2e8f0;">Etiquetas</th>
                                <th style="padding:8px;border:1px solid #e2e8f0;">Hojas</th>
                                <th style="padding:8px;border:1px solid #e2e8f0;">Acción</th>
                            </tr>
                        </thead>
                        <tbody>
            `;
            datosPorMun.forEach((d, idx) => {
                listaHTML += `
                    <tr>
                        <td style="padding:8px;border:1px solid #e2e8f0;"><strong>${d.nombre}</strong> (${d.clave})</td>
                        <td style="padding:8px;border:1px solid #e2e8f0;text-align:center;">${d.casillas}</td>
                        <td style="padding:8px;border:1px solid #e2e8f0;text-align:center;">${d.etiquetas}</td>
                        <td style="padding:8px;border:1px solid #e2e8f0;text-align:center;">${d.hojas}</td>
                        <td style="padding:8px;border:1px solid #e2e8f0;text-align:center;">
                            <button class="btn" style="margin:0;padding:6px 12px;font-size:0.85rem;background:#c53030;"
                                onclick="descargarPDFMunicipio(${idx})">🖨️ Imprimir / PDF</button>
                        </td>
                    </tr>
                `;
            });
            listaHTML += `</tbody></table></div>`;
            document.getElementById('lista-municipios').innerHTML = listaHTML;

            // Vista previa limitada
            document.getElementById('hojas-etiquetas').innerHTML =
                `<p class="no-print" style="text-align:center;color:#64748b;margin-bottom:12px;">Vista previa (primeras etiquetas). Usa los botones de arriba para descargar PDF por municipio.</p>` +
                preview.html;

            document.getElementById('check-archivo').classList.add('visible');
            document.getElementById('etiquetas-container').classList.add('visible');

            // Guardar para descargas
            window._etiquetasPorMun = datosPorMun;
            window._etiquetasMeta = {
                entidad,
                plantilla: plantilla.nombre,
                plantillaKey,
                total: totalEtiquetas,
                hojas: totalHojasEst
            };
            // HTML completo solo si no es excesivo (< 300 etiquetas)
            window._etiquetasHTML = totalEtiquetas <= 300
                ? datosPorMun.map(d => d.html).join('')
                : null;

            document.getElementById('etiquetas-container').scrollIntoView({ behavior: 'smooth', block: 'start' });
            loading.style.display = 'none';
        } catch (err) {
            console.error(err);
            alert('Error al generar etiquetas:\n' + err.message);
            loading.style.display = 'none';
        }
    };
    reader.readAsBinaryString(archivo);
}

/** Descarga un archivo HTML autónomo listo para abrir e imprimir */
function descargarHTML() {
    let hojas = window._etiquetasHTML;
    if (!hojas && window._etiquetasPorMun) {
        // Reconstruir desde los municipios si no se guardó el completo
        hojas = window._etiquetasPorMun.map(d => d.html).join('');
    }
    if (!hojas) {
        alert('Primero genera las etiquetas. Si hay muchas, usa "Descargar PDF por municipio".');
        return;
    }
    const meta = window._etiquetasMeta || {};
    const estilos = `
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 8px; color: #000; }
        .hoja-etiquetas { display: grid; grid-template-columns: 1fr 1fr; gap: 10px 14px; page-break-after: always; margin-bottom: 16px; }
        .etiqueta { border: 2px solid #000; border-radius: 4px; padding: 14px 16px; background: #fff;
            display: flex; flex-direction: column; justify-content: center; box-sizing: border-box; min-height: 115px; }
        .etiqueta .linea { font-size: 11pt; line-height: 1.45; color: #000; }
        .etiqueta .linea strong { font-weight: 700; color: #000; }
        @media print {
            body { padding: 0; }
            .hoja-etiquetas { gap: 8px 12px; margin: 0; }
            .no-print { display: none !important; }
            @page { margin: 10mm 12mm; size: letter; }
        }
    `;
    const doc = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Etiquetas de Casillas — ${meta.entidad || ''}</title>
<style>${estilos}</style>
</head>
<body>
<div class="no-print" style="text-align:center;padding:12px;background:#edf2f7;margin-bottom:16px;">
    <strong>${meta.total || 0}</strong> etiquetas · <strong>${meta.hojas || 0}</strong> hojas · ${meta.plantilla || ''}
    &nbsp;|&nbsp; <button onclick="window.print()" style="padding:8px 16px;font-size:1rem;cursor:pointer;">🖨️ Imprimir</button>
</div>
${hojas}
</body>
</html>`;

    const blob = new Blob([doc], { type: 'text/html;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'etiquetas_casillas.html';
    a.click();
    URL.revokeObjectURL(a.href);
}

/** Descarga el PDF de un solo municipio (por índice en window._etiquetasPorMun) */
/**
 * Construye el documento HTML completo listo para imprimir / guardar como PDF.
 * Usa tablas + CSS de impresión nativo del navegador (sin html2pdf).
 */
function construirDocumentoImpresion(listaCasillas, entidad, plantilla, titulo) {
    const cols = plantilla.cols || 2;
    const filas = plantilla.filas || 5;
    let tablas = '';
    let i = 0;

    while (i < listaCasillas.length) {
        tablas += '<table class="hoja">';
        for (let f = 0; f < filas; f++) {
            tablas += '<tr>';
            for (let c = 0; c < cols; c++) {
                if (i < listaCasillas.length) {
                    const cas = listaCasillas[i];
                    const tipoTxt = tipoATexto(cas.tipo);
                    const secPad  = seccionPadded(cas.sec);
                    const munNom  = nombreMunicipio(cas.mun);
                    tablas +=
                        '<td><div class="etiqueta">' +
                        '<div>ENTIDAD: ' + entidad + '</div>' +
                        '<div>DISTRITO: ' + cas.distrito + '</div>' +
                        '<div>MUNICIPIO: ' + munNom + '</div>' +
                        '<div>SECCI\u00D3N: ' + secPad + '</div>' +
                        '<div>CASILLA: ' + tipoTxt + '</div>' +
                        '</div></td>';
                    i++;
                } else {
                    tablas += '<td></td>';
                }
            }
            tablas += '</tr>';
        }
        tablas += '</table>';
    }

    // 10 por hoja (2×5), medido en el PDF "Etiquetas buenas" (carta):
    // texto Aptos/Calibri 11 pt, interlineado 22.5 pt, salto entre filas 156 pt,
    // margen izq. 26.2 pt, columna derecha en x = 324 pt.
    const esDiez = filas === 5 && cols === 2;
    const pageMargin = esDiez ? '12mm 8pt 8pt 9mm' : '10mm 8mm 8mm 8mm';
    const sepH = esDiez ? '18pt' : '4mm';
    const sepV = esDiez ? '43.5pt' : (filas <= 3 ? '6mm' : (filas <= 4 ? '4mm' : '2.5mm'));
    const fontSize = esDiez ? '12pt' : (filas <= 5 ? '11pt' : '10pt');
    const lineH = esDiez ? '22.5pt' : (filas <= 5 ? '1.25' : '1.18');
    const tableW = esDiez ? '578pt' : '100%';
    const colW = esDiez ? '280pt' : (Math.floor(100 / cols) + '%');

    return '<!DOCTYPE html>\n<html lang="es">\n<head>\n<meta charset="UTF-8">\n' +
        '<title>' + (titulo || 'Etiquetas') + '</title>\n' +
        '<style>\n' +
        '  @page { size: letter; margin: ' + pageMargin + '; }\n' +
        '  * { box-sizing: border-box; margin: 0; padding: 0; }\n' +
        '  body { font-family: Aptos, Calibri, Arial, sans-serif; color: #000; background: #fff; }\n' +
        '  .toolbar { text-align: center; padding: 12px; background: #edf2f7; margin-bottom: 12px; font-family: Arial, sans-serif; }\n' +
        '  .toolbar button { padding: 10px 20px; font-size: 1rem; cursor: pointer; margin: 0 6px; }\n' +
        '  table.hoja {\n' +
        '    width: ' + tableW + ';\n' +
        '    border-collapse: separate;\n' +
        '    border-spacing: ' + sepH + ' ' + sepV + ';\n' +
        '    page-break-after: always;\n' +
        '    table-layout: fixed;\n' +
        '  }\n' +
        '  table.hoja:last-child { page-break-after: auto; }\n' +
        '  table.hoja td {\n' +
        '    width: ' + colW + ';\n' +
        '    vertical-align: top;\n' +
        '  }\n' +
        '  .etiqueta { border: none; padding: 0; margin: 0; background: transparent; }\n' +
        '  .etiqueta div {\n' +
        '    font-family: Aptos, Calibri, Arial, sans-serif;\n' +
        '    font-size: ' + fontSize + ';\n' +
        '    line-height: ' + lineH + ';\n' +
        '    font-weight: 700;\n' +
        '    margin: 0;\n' +
        '    padding: 0;\n' +
        '  }\n' +
        '  @media print {\n' +
        '    .toolbar { display: none !important; }\n' +
        '    body { margin: 0; }\n' +
        '  }\n' +
        '</style>\n</head>\n<body>\n' +
        '<div class="toolbar">' +
        '<strong>' + (titulo || '') + '</strong> &mdash; ' + listaCasillas.length + ' etiquetas &nbsp; ' +
        '<button onclick="window.print()">Imprimir / Guardar como PDF</button>' +
        '</div>\n' +
        tablas +
        '\n</body>\n</html>';
}

/**
 * Abre una ventana con las etiquetas del municipio lista para
 * Imprimir o Guardar como PDF (método nativo del navegador).
 */
function descargarPDFMunicipio(idx) {
    const datos = window._etiquetasPorMun;
    if (!datos || !datos[idx]) {
        alert('Datos no disponibles. Genera las etiquetas primero.');
        return Promise.resolve();
    }
    const d = datos[idx];
    const meta = window._etiquetasMeta || {};
    const entidad = meta.entidad || 'CHIAPAS';
    const plantillaKey = meta.plantillaKey || '5163';
    const plantilla = (typeof PLANTILLAS !== 'undefined' && PLANTILLAS[plantillaKey])
        ? PLANTILLAS[plantillaKey]
        : { cols: 2, filas: 5, porHoja: 10 };

    const listaCasillas = d.lista || [];
    if (listaCasillas.length === 0) {
        alert('No hay etiquetas para este municipio.');
        return Promise.resolve();
    }

    const titulo = 'Etiquetas ' + d.nombre + ' (' + d.clave + ')';
    const doc = construirDocumentoImpresion(listaCasillas, entidad, plantilla, titulo);

    const w = window.open('', '_blank');
    if (!w) {
        alert('El navegador bloqueó la ventana emergente. Permite pop-ups para este sitio e inténtalo de nuevo.');
        return Promise.resolve();
    }
    w.document.open();
    w.document.write(doc);
    w.document.close();
    // Dar tiempo a que cargue y luego abrir el diálogo de impresión
    setTimeout(function () {
        try { w.focus(); w.print(); } catch (e) { /* el usuario puede imprimir manualmente */ }
    }, 400);

    return Promise.resolve();
}

/**
 * Descarga un archivo HTML por cada municipio (listo para abrir e imprimir/guardar PDF).
 * Más fiable que abrir muchas ventanas de impresión a la vez.
 */
async function descargarTodosPDFPorMunicipio() {
    const datos = window._etiquetasPorMun;
    if (!datos || datos.length === 0) {
        alert('Primero genera las etiquetas.');
        return;
    }

    const meta = window._etiquetasMeta || {};
    const entidad = meta.entidad || 'CHIAPAS';
    const plantillaKey = meta.plantillaKey || '5163';
    const plantilla = (typeof PLANTILLAS !== 'undefined' && PLANTILLAS[plantillaKey])
        ? PLANTILLAS[plantillaKey]
        : { cols: 2, filas: 5, porHoja: 10 };

    const btn = document.querySelector('button[onclick="descargarTodosPDFPorMunicipio()"]');
    const textoOriginal = btn ? btn.innerHTML : '';
    if (btn) btn.disabled = true;

    try {
        for (let i = 0; i < datos.length; i++) {
            const d = datos[i];
            if (btn) btn.innerHTML = '⏳ HTML ' + (i + 1) + ' de ' + datos.length + ': ' + d.nombre + '...';

            const lista = d.lista || [];
            if (lista.length === 0) continue;

            const titulo = 'Etiquetas ' + d.nombre + ' (' + d.clave + ')';
            const doc = construirDocumentoImpresion(lista, entidad, plantilla, titulo);
            const blob = new Blob([doc], { type: 'text/html;charset=utf-8' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = 'etiquetas_' + d.nombre.replace(/\s+/g, '_') + '_' + d.clave + '.html';
            a.click();
            URL.revokeObjectURL(a.href);

            await new Promise(function (r) { setTimeout(r, 500); });
        }
        if (btn) {
            btn.innerHTML = '✅ HTML descargados';
            setTimeout(function () { btn.innerHTML = textoOriginal; btn.disabled = false; }, 2500);
        }
        alert('Se descargaron ' + datos.length + ' archivos HTML (uno por municipio).\n\nÁbrelos y usa Imprimir → Guardar como PDF.');
    } catch (err) {
        console.error(err);
        alert('Error al generar los archivos. Usa el botón individual de cada municipio.');
        if (btn) {
            btn.innerHTML = textoOriginal;
            btn.disabled = false;
        }
    }
}

// Mostrar / ocultar campos de independientes
function toggleIndependientes() {
    const si = document.querySelector('input[name="tiene-independientes"][value="si"]').checked;
    document.getElementById('independientes-campos').style.display = si ? 'grid' : 'none';
    if (!si) {
        document.getElementById('num-independientes').value = 0;
        document.getElementById('boletas-independientes').value = 0;
    }
    actualizarExtrasPreview();
}

// Mostrar / ocultar campos de TEPJF
function toggleTepjf() {
    const si = document.querySelector('input[name="tiene-tepjf"][value="si"]').checked;
    document.getElementById('tepjf-campos').style.display = si ? 'block' : 'none';
    if (si) {
        generarCamposCasosTepjf();
    } else {
        document.getElementById('casos-tepjf-lista').innerHTML = '';
    }
    actualizarExtrasPreview();
}

// Genera los campos dinámicos: Sección + Tipo de Casilla + Boletas
function generarCamposCasosTepjf() {
    const n = Math.min(100, Math.max(1, Number(document.getElementById('num-casos-tepjf').value) || 1));
    document.getElementById('num-casos-tepjf').value = n;

    const contenedor = document.getElementById('casos-tepjf-lista');
    contenedor.innerHTML = '';

    for (let i = 1; i <= n; i++) {
        const div = document.createElement('div');
        div.className = 'params-grid';
        div.style.marginBottom = '12px';
        div.style.padding = '12px';
        div.style.background = '#fff';
        div.style.border = '1px solid #e2e8f0';
        div.style.borderRadius = '8px';
        div.style.gridTemplateColumns = '1fr 1fr 1fr';
        div.innerHTML = `
            <div class="param-group">
                <label for="tepjf-sec-${i}">Sección</label>
                <input type="text" id="tepjf-sec-${i}" class="param-input" placeholder="Ej. 35" inputmode="numeric">
            </div>
            <div class="param-group">
                <label for="tepjf-tipo-${i}">Tipo de Casilla</label>
                <input type="text" id="tepjf-tipo-${i}" class="param-input" placeholder="Ej. B, C1, E1, E1C1">
            </div>
            <div class="param-group">
                <label for="tepjf-bol-${i}">Boletas adicionales</label>
                <input type="number" id="tepjf-bol-${i}" min="0" value="1" class="param-input">
            </div>
        `;
        contenedor.appendChild(div);

        div.querySelector(`#tepjf-sec-${i}`).addEventListener('input', actualizarExtrasPreview);
        div.querySelector(`#tepjf-tipo-${i}`).addEventListener('input', actualizarExtrasPreview);
        div.querySelector(`#tepjf-bol-${i}`).addEventListener('input', actualizarExtrasPreview);
    }
}

// Obtiene el mapa "sección|tipo" → boletas TEPJF
// Clave: "35|C2"  (sección + tipo en mayúsculas)
function obtenerMapaTepjf() {
    const mapa = {};
    if (!document.querySelector('input[name="tiene-tepjf"][value="si"]').checked) {
        return mapa;
    }
    const n = Number(document.getElementById('num-casos-tepjf').value) || 0;
    for (let i = 1; i <= n; i++) {
        const secEl  = document.getElementById(`tepjf-sec-${i}`);
        const tipoEl = document.getElementById(`tepjf-tipo-${i}`);
        const bolEl  = document.getElementById(`tepjf-bol-${i}`);
        if (!secEl || !tipoEl || !bolEl) continue;

        const sec  = String(secEl.value || '').trim();
        const tipo = String(tipoEl.value || '').trim().toUpperCase();
        const bol  = Number(bolEl.value) || 0;

        // No permitir tipo S (casilla especial)
        if (tipo === 'S') continue;

        if (sec && tipo && bol > 0) {
            const key = `${sec}|${tipo}`;
            mapa[key] = (mapa[key] || 0) + bol;
        }
    }
    return mapa;
}

// Calcula boletas extras base (partidos + independientes) — sin TEPJF
function calcularExtras() {
    const fed = Number(document.getElementById('partidos-federales').value) || 0;
    const bolFed = Number(document.getElementById('boletas-federales').value) || 0;
    const loc = Number(document.getElementById('partidos-locales').value) || 0;
    const bolLoc = Number(document.getElementById('boletas-locales').value) || 0;

    let ind = 0;
    let bolInd = 0;
    if (document.querySelector('input[name="tiene-independientes"][value="si"]').checked) {
        ind = Number(document.getElementById('num-independientes').value) || 0;
        bolInd = Number(document.getElementById('boletas-independientes').value) || 0;
    }

    return {
        extras: (fed * bolFed) + (loc * bolLoc) + (ind * bolInd),
        fed, bolFed, loc, bolLoc, ind, bolInd
    };
}

// Actualiza la vista previa de boletas extras en tiempo real
function actualizarExtrasPreview() {
    const { extras } = calcularExtras();
    const mapaTepjf = obtenerMapaTepjf();
    const numCasos = Object.keys(mapaTepjf).length;

    let texto = `Boletas extras base por casilla: <strong>${extras.toLocaleString('es-MX')}</strong>`;
    if (numCasos > 0) {
        texto += ` &nbsp;|&nbsp; TEPJF aplica en <strong>${numCasos}</strong> casilla(s) específica(s)`;
    }
    document.getElementById('extras-preview').innerHTML = texto;
}

// Escuchar cambios en los inputs
['partidos-federales', 'boletas-federales', 'partidos-locales', 'boletas-locales',
 'num-independientes', 'boletas-independientes'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', actualizarExtrasPreview);
});

function procesarFolios() {
    const archivo = document.getElementById('archivo-proyeccion').files[0];
    if (!archivo) {
        alert('Por favor, sube el archivo de proyección de casillas.');
        return;
    }

    const { extras: extrasBase, fed: partidosFed, bolFed: boletasFed,
            loc: partidosLoc, bolLoc: boletasLoc, ind: numInd, bolInd: boletasInd } = calcularExtras();

    const mapaTepjf = obtenerMapaTepjf(); // sección → boletas adicionales TEPJF

    const loading = document.getElementById('loading');
    loading.style.display = 'flex';

    const resultsDiv = document.getElementById('results');
    resultsDiv.style.display = 'none';
    resultsDiv.style.opacity = '0';

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = e.target.result;
            const wb = XLSX.read(data, { type: 'binary' });
            const sheet = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });

            if (sheet.length < 2) {
                throw new Error('El archivo está vacío o no tiene datos.');
            }

            const header = sheet[0].map(h => String(h || '').trim().toLowerCase());

            let idxMun  = header.findIndex(h => h.includes('municipio'));
            let idxSec  = header.findIndex(h => h.includes('sección') || h.includes('seccion'));
            let idxTipo = header.findIndex(h => h.includes('tipo'));
            let idxPad  = header.findIndex(h => h.includes('padrón') || h.includes('padron'));
            let idxLis  = header.findIndex(h => h.includes('listado') || h.includes('lista') || h.includes('nominal'));
            let idxDist = header.findIndex(h => h.includes('distrito'));

            if (idxMun === -1)  idxMun = 1;
            if (idxSec === -1)  idxSec = 2;
            if (idxTipo === -1) idxTipo = 3;
            if (idxPad === -1)  idxPad = 4;
            if (idxLis === -1)  idxLis = 5;
            if (idxDist === -1) idxDist = 0;

            const output = [[
                'Distrito', 'Municipio', 'Sección', 'Tipo de Casilla',
                'Padrón Electoral', 'Listado nominal',
                'Boletas Extras', 'Total Boletas',
                'Folio Inicial', 'Folio Final', 'Observaciones'
            ]];

            let currentMun = null;
            let currentFolio = 1;

            let totalCasillas = 0;
            let totalConFolios = 0;
            let totalEspeciales = 0;
            let totalBoletas = 0;
            let totalLN = 0;
            let totalTepjfAplicado = 0;
            let municipiosProcesados = new Set();
            const previewRows = [];

            for (let i = 1; i < sheet.length; i++) {
                const row = sheet[i];
                if (!row || row.length < 4) continue;

                const distrito = row[idxDist] ?? '';
                const mun  = String(row[idxMun] ?? '').trim();
                const sec  = String(row[idxSec] ?? '').trim();
                const tipo = String(row[idxTipo] ?? '').trim().toUpperCase();
                const pad  = Number(row[idxPad]) || 0;
                const lis  = Number(row[idxLis]) || 0;

                if (!mun) continue;

                // Reiniciar contador al cambiar de municipio
                if (mun !== currentMun) {
                    currentMun = mun;
                   // currentFolio = 1;
                    municipiosProcesados.add(mun);
                }

                // Boletas TEPJF solo para la casilla exacta (sección + tipo). Nunca para tipo S.
                const tepjfKey = `${sec}|${tipo}`;
                const tepjfExtra = (tipo !== 'S') ? (mapaTepjf[tepjfKey] || 0) : 0;

                let totalBoletasCasilla = 0;
                let extrasUsadas = 0;
                let observacion = '';
                let folioIni = '';
                let folioFin = '';

                // ── Casilla especial (tipo S) → 1000 + extras base (SIN TEPJF) ──
                if (tipo === 'S') {
                    extrasUsadas = extrasBase;
                    totalBoletasCasilla = 1000 + extrasUsadas;
                    observacion = `Casilla especial (1000 + ${extrasBase} base)`;
                    totalEspeciales++;
                } else {
                    // Casilla ordinaria (B, C*, E*, E*C*)
                    extrasUsadas = extrasBase + tepjfExtra;
                    totalBoletasCasilla = lis + extrasUsadas;
                    if (tepjfExtra > 0) {
                        observacion = `+${tepjfExtra} boleta(s) TEPJF`;
                    }
                }

                if (tepjfExtra > 0) totalTepjfAplicado += tepjfExtra;

                if (totalBoletasCasilla > 0) {
                    folioIni = currentFolio;
                    folioFin = currentFolio + totalBoletasCasilla - 1;
                    currentFolio = folioFin + 1;
                    totalConFolios++;
                    totalBoletas += totalBoletasCasilla;
                }

                totalCasillas++;
                totalLN += lis;

                const outRow = [
                    distrito, mun, sec, tipo,
                    pad, lis,
                    extrasUsadas, totalBoletasCasilla,
                    folioIni, folioFin, observacion
                ];
                output.push(outRow);

                if (previewRows.length < 30) {
                    previewRows.push(outRow);
                }
            }

            // Generar Excel
            const ws = XLSX.utils.aoa_to_sheet(output);
            ws['!cols'] = [
                { wch: 10 }, { wch: 12 }, { wch: 10 }, { wch: 16 },
                { wch: 16 }, { wch: 16 },
                { wch: 14 }, { wch: 14 },
                { wch: 14 }, { wch: 14 }, { wch: 28 }
            ];

            const wbout = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wbout, ws, 'Asignacion Folios');

            const blob = new Blob([XLSX.write(wbout, { bookType: 'xlsx', type: 'array' })],
                                 { type: 'application/octet-stream' });
            const download = document.getElementById('download');
            download.href = URL.createObjectURL(blob);
            download.style.display = 'block';

            // Estadísticas
            let indText = 'Ninguna';
            if (numInd > 0) {
                indText = `${numInd} × ${boletasInd} boletas = ${(numInd * boletasInd).toLocaleString('es-MX')}`;
            }

            const numCasosTepjf = Object.keys(mapaTepjf).length;
            let tepjfText = 'No aplica';
            if (numCasosTepjf > 0) {
                const detalle = Object.entries(mapaTepjf)
                    .map(([key, b]) => {
                        const [s, t] = key.split('|');
                        return `Sec ${s} Tipo ${t}: ${b}`;
                    })
                    .join(' | ');
                tepjfText = `${numCasosTepjf} casilla(s) → ${detalle}`;
            }

            document.getElementById('global-stats-content').innerHTML = `
                <ul>
                    <li><strong>Municipios procesados:</strong> ${municipiosProcesados.size}</li>
                    <li><strong>Total de casillas:</strong> ${totalCasillas.toLocaleString('es-MX')}</li>
                    <li><strong>Casillas con folios asignados:</strong> ${totalConFolios.toLocaleString('es-MX')}</li>
                    <li><strong>Total Listado Nominal:</strong> ${totalLN.toLocaleString('es-MX')}</li>
                    <li><strong>Casillas especiales (tipo S):</strong> ${totalEspeciales.toLocaleString('es-MX')} → 1000 + extras c/u</li>
                    <li><strong>Partidos federales:</strong> ${partidosFed} × ${boletasFed} boletas</li>
                    <li><strong>Partidos locales:</strong> ${partidosLoc} × ${boletasLoc} boletas</li>
                    <li><strong>Candidaturas independientes:</strong> ${indText}</li>
                    <li><strong>Boletas extras base (partidos + ind.):</strong> ${extrasBase.toLocaleString('es-MX')}</li>
                    <li><strong>Resolución TEPJF:</strong> ${tepjfText}</li>
                    <li><strong>Boletas TEPJF aplicadas (suma en casillas):</strong> ${totalTepjfAplicado.toLocaleString('es-MX')}</li>
                    <li><strong>Total de boletas asignadas:</strong> ${totalBoletas.toLocaleString('es-MX')}</li>

                </ul>
            `;

            // Tabla de vista previa
            let tableHTML = `
                <table>
                    <thead>
                        <tr>
                            <th>Municipio</th>
                            <th>Sección</th>
                            <th>Tipo</th>
                            <th>Lista Nominal</th>
                            <th>Extras</th>
                            <th>Total</th>
                            <th>Folio Inicial</th>
                            <th>Folio Final</th>
                            <th>Observaciones</th>
                        </tr>
                    </thead>
                    <tbody>
            `;
            for (const r of previewRows) {
                const esEspecial = String(r[3]).toUpperCase() === 'S';
                const rowStyle = esEspecial ? ' style="background-color:#fef3c7;"' : '';
                tableHTML += `
                    <tr${rowStyle}>
                        <td>${r[1]}</td>
                        <td>${r[2]}</td>
                        <td><strong>${r[3]}</strong></td>
                        <td>${Number(r[5]).toLocaleString('es-MX')}</td>
                        <td>${Number(r[6]).toLocaleString('es-MX')}</td>
                        <td>${Number(r[7]).toLocaleString('es-MX')}</td>
                        <td>${r[8] !== '' ? Number(r[8]).toLocaleString('es-MX') : '—'}</td>
                        <td>${r[9] !== '' ? Number(r[9]).toLocaleString('es-MX') : '—'}</td>
                        <td style="font-size:0.85em;color:#92400e;">${r[10] || ''}</td>
                    </tr>
                `;
            }
            tableHTML += `</tbody></table>`;
            if (output.length > 31) {
                tableHTML += `<p style="text-align:center;color:#64748b;margin-top:10px;">Mostrando las primeras 30 filas de ${output.length - 1} casillas. Descarga el Excel para ver el archivo completo.</p>`;
            }
            document.getElementById('preview-table-container').innerHTML = tableHTML;

            // Mostrar resultados
            document.getElementById('check-archivo').classList.add('visible');
            resultsDiv.style.display = 'block';
            setTimeout(() => {
                resultsDiv.style.opacity = '1';
                resultsDiv.style.transform = 'translateY(0)';
            }, 50);

            document.querySelectorAll('.stats-card').forEach((card, i) => {
                setTimeout(() => card.classList.add('visible'), 200 + i * 150);
            });

            loading.style.display = 'none';

        } catch (err) {
            console.error(err);
            alert('Ocurrió un error al procesar el archivo.\nRevisa la consola (F12) para más detalles.\n\n' + err.message);
            loading.style.display = 'none';
        }
    };
    reader.readAsBinaryString(archivo);
}

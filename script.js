// Forzar al navegador a que olvide la posición de scroll anterior al recargar
if ('scrollRestoration' in history) {
    history.scrollRestoration = 'manual';
}

// Al cargar la página, subir hasta arriba del todo
window.addEventListener('load', () => {
    window.scrollTo(0, 0);
});

document.addEventListener('DOMContentLoaded', () => {
    // ─── LÓGICA DE PALOMITAS Y NOMBRES DE ARCHIVO ───
    const fileInputs = [
        { id: 'conformacion', check: 'check-conformacion', name: 'name-conformacion' },
        { id: 'estadistico', check: 'check-estadistico', name: 'name-estadistico' },
        { id: 'catalogo', check: 'check-catalogo', name: 'name-catalogo' }
    ];

    fileInputs.forEach(item => {
        const input = document.getElementById(item.id);
        const check = document.getElementById(item.check);
        const nameDisplay = document.getElementById(item.name);

        if (input && check) {
            input.addEventListener('change', function() {
                if (this.files.length > 0) {
                    check.classList.add('visible');
                    if (nameDisplay) {
                        nameDisplay.innerHTML = '✓ ' + this.files[0].name;
                        nameDisplay.classList.add('visible'); // Muestra la cajita verde
                    }
                } else {
                    check.classList.remove('visible');
                    if (nameDisplay) {
                        nameDisplay.innerHTML = '';
                        nameDisplay.classList.remove('visible'); // Oculta la cajita
                    }
                }
            });
        }
    });

    // ─── CASILLAS ESPECIALES ───
    const inputEspeciales = document.getElementById('num-especiales');
    const containerEspeciales = document.getElementById('secciones-especiales-container');
    const inputsDiv = document.getElementById('secciones-inputs');

    if (inputEspeciales) {
        inputEspeciales.addEventListener('input', function(e) {
            let valStr = e.target.value.replace(/[^0-9]/g, '');
            let val = valStr === '' ? 0 : parseInt(valStr, 10);
            if (val > 10) val = 10;
            e.target.value = valStr === '' ? '' : val;

            inputsDiv.innerHTML = '';
            if (val > 0) {
                containerEspeciales.style.display = 'block';
                for (let i = 1; i <= val; i++) {
                    inputsDiv.innerHTML += `
                        <div style="display: flex; align-items: center; gap: 8px;">
                            <label style="font-size: 0.85em; color: #4a5568;">Sección ${i}:</label>
                            <input type="text" class="input-esp-sec" style="width: 80px; padding: 5px; border: 1px solid #cbd5e0; border-radius: 4px;">
                        </div>
                    `;
                }
            } else {
                containerEspeciales.style.display = 'none';
            }
        });
    }
});

let excelDetalleGlobal = null;
let excelResumenGlobal = null;

async function descargarZIP() {
    if (!excelDetalleGlobal || !excelResumenGlobal) {
        alert("Primero genera la proyección.");
        return;
    }

    try {
        const zip = new JSZip();
        zip.file("proyeccion_casillas.xlsx", excelDetalleGlobal);
        zip.file("resumen_casillas_por_localidad.xlsx", excelResumenGlobal);

        const content = await zip.generateAsync({ type: "blob" });
        const url = URL.createObjectURL(content);
        const a = document.createElement("a");
        a.style.display = "none";
        a.href = url;
        a.download = "Proyeccion_Casillas.zip";
        document.body.appendChild(a);
        a.click();
        
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);

    } catch (error) {
        console.error("Error al generar ZIP:", error);
        alert("Ocurrió un error al empaquetar los archivos.");
    }
}

function distribute(total, num) {
    if (num === 0) return [];
    const avg = Math.floor(total / num);
    const rem = total % num;
    const res = [];
    for (let i = 0; i < num; i++) {
        res.push(avg + (i < rem ? 1 : 0));
    }
    return res;
}

function processFiles() {
    const conformFile = document.getElementById('conformacion').files[0];
    const estadFile = document.getElementById('estadistico').files[0];
    
    const baseOptions = document.querySelectorAll('input[name="base"]');
    let base = 'lista';
    for (const option of baseOptions) {
        if (option.checked) { base = option.value; break; }
    }

    const inputsEspeciales = document.querySelectorAll('.input-esp-sec');
    const especialRequests = new Map();
    inputsEspeciales.forEach(input => {
        const secStr = input.value.trim();
        if (secStr !== '') {
            const sec = String(parseInt(secStr, 10));
            especialRequests.set(sec, (especialRequests.get(sec) || 0) + 1);
        }
    });

    if (!conformFile || !estadFile) {
        alert('Por favor, suba ambos archivos Excel.');
        return;
    }

    const loading = document.getElementById('loading');
    loading.style.display = 'flex';

    const resultsDiv = document.getElementById('results');
    resultsDiv.style.display = 'none';
    resultsDiv.style.opacity = '0';
    resultsDiv.style.transform = 'translateY(20px)';

    const btnZip = document.getElementById('btn-download-zip');
    if (btnZip) btnZip.style.display = 'none';

    const reader1 = new FileReader();
    reader1.onload = function(e) {
        try {
            const data1 = e.target.result;
            const wb1 = XLSX.read(data1, { type: 'binary' });
            const sheet1 = XLSX.utils.sheet_to_json(wb1.Sheets[wb1.SheetNames[0]], { header: 1 });

            const reader2 = new FileReader();
            reader2.onload = function(e) {
                try {
                    const data2 = e.target.result;
                    const wb2 = XLSX.read(data2, { type: 'binary' });
                    const sheet2 = XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[0]], { header: 1 });

                    let catalogMap = new Map();
                    const stats = new Map();
                    let totalPad = 0, totalLis = 0, totalSecs = new Set();
                    
                    for (let i = 1; i < sheet2.length; i++) {
                        const row = sheet2[i];
                        if (row.length < 8) continue;
                        const mun = String(row[2]).trim();
                        const sec = String(row[3]).trim();
                        const loc = String(row[4]).trim();
                        const manzana = String(row[5] || '**').trim();
                        const pad = Number(row[6]) || 0;
                        const lis = Number(row[7]) || 0;

                        const unit = (manzana === '**') ? loc : `${loc}-${manzana}`;
                        const key = `${mun}-${sec}-${unit}`;

                        if (!stats.has(key)) stats.set(key, { pad: 0, lis: 0 });
                        const val = stats.get(key);
                        val.pad += pad;
                        val.lis += lis;
                        totalPad += pad;
                        totalLis += lis;
                        totalSecs.add(`${mun}-${sec}`);
                    }

                    const localitySums = new Map();
                    for (const [key, v] of stats) {
                        const parts = key.split('-');
                        if (parts.length < 3) continue;
                        const locKey = `${parts[0]}-${parts[1]}-${parts[2]}`;
                        if (!localitySums.has(locKey)) localitySums.set(locKey, { pad: 0, lis: 0 });
                        const sum = localitySums.get(locKey);
                        sum.pad += v.pad;
                        sum.lis += v.lis;
                    }

                    const groups = new Map();
                    const sedeNames = new Map();
                    for (let i = 1; i < sheet1.length; i++) {
                        const row = sheet1[i];
                        if (row.length < 9) continue;
                        const mun = String(row[3]).trim();
                        const sec = String(row[4]).trim();
                        const loc = String(row[5]).trim();
                        const nombre = String(row[6] || '').trim();
                        const manzana = String(row[7] || '**').trim();
                        const asg = (row[8] || '').trim();
                        const sede = (row[9] || '').toString().trim();
                        if (!asg) continue;

                        const unit = (manzana === '**') ? loc : `${loc}-${manzana}`;
                        const secKey = `${mun}-${sec}`;
                        if (!groups.has(secKey)) groups.set(secKey, new Map());
                        const secGroups = groups.get(secKey);
                        if (!secGroups.has(asg)) secGroups.set(asg, new Set());
                        secGroups.get(asg).add(unit);

                        if (sede) {
                            if (!sedeNames.has(secKey)) sedeNames.set(secKey, new Map());
                            sedeNames.get(secKey).set(sede, nombre || ('Localidad ' + loc));
                        }
                    }

                    const allSecs = Array.from(new Set([
                        ...Array.from(stats.keys()).map(k => k.split('-').slice(0,2).join('-')),
                        ...Array.from(groups.keys())
                    ])).sort((a, b) => {
                        const [ma, sa] = a.split('-').map(Number);
                        const [mb, sb] = b.split('-').map(Number);
                        return ma - mb || sa - sb;
                    });

                    const munAgg = new Map();
                    const muns = new Set(allSecs.map(k => k.split('-')[0]));
                    for (const m of muns) {
                        munAgg.set(m, {
                            bas_pad: 0, bas_lis: 0, bas_count: 0, con_pad: 0, con_lis: 0, con_count: 0,
                            ext_pad: 0, ext_lis: 0, ext_count: 0, exc_pad: 0, exc_lis: 0, exc_count: 0, esp_count: 0
                        });
                    }

                    const output = [["Distrito", "Municipio", "Sección", "Tipo de Casilla", "Padrón Electoral", "Listado nominal"]];
                    const resumenMap = new Map();

                    for (const secKey of allSecs) {
                        const [mun, sec] = secKey.split('-').map(String);
                        const groupData = new Map();

                        let totalPadSec = 0, totalLisSec = 0;
                        for (const [key, v] of stats) {
                            if (key.startsWith(`${mun}-${sec}-`)) {
                                totalPadSec += v.pad;
                                totalLisSec += v.lis;
                            }
                        }

                        if (groups.has(secKey)) {
                            for (const [asg, units] of groups.get(secKey)) {
                                let pad = 0, lis = 0;
                                for (const assignedUnit of units) {
                                    if (!assignedUnit.includes('-')) {
                                        const locKey = `${mun}-${sec}-${assignedUnit}`;
                                        if (localitySums.has(locKey)) {
                                            const sum = localitySums.get(locKey);
                                            pad += sum.pad;
                                            lis += sum.lis;
                                        }
                                    } else {
                                        const lkey = `${mun}-${sec}-${assignedUnit}`;
                                        if (stats.has(lkey)) {
                                            const v = stats.get(lkey);
                                            pad += v.pad;
                                            lis += v.lis;
                                        }
                                    }
                                }
                                if (pad > 0 || lis > 0) {
                                    groupData.set(asg, { pad, lis });
                                }
                            }
                            let assignedPad = 0, assignedLis = 0;
                            for (const data of groupData.values()) {
                                assignedPad += data.pad;
                                assignedLis += data.lis;
                            }
                            const restPad = Math.max(0, totalPadSec - assignedPad);
                            const restLis = Math.max(0, totalLisSec - assignedLis);
                            if (restPad > 0 || restLis > 0) {
                                if (groupData.has('B')) {
                                    const b = groupData.get('B');
                                    b.pad += restPad;
                                    b.lis += restLis;
                                } else {
                                    groupData.set('B', { pad: restPad, lis: restLis });
                                }
                            }
                        } else {
                            if (totalPadSec > 0 || totalLisSec > 0) {
                                groupData.set('B', { pad: totalPadSec, lis: totalLisSec });
                            }
                        }

                        const agg = munAgg.get(mun);

                        const getSedeNombre = (asgTipo) => {
                            if (sedeNames.has(secKey) && sedeNames.get(secKey).has(asgTipo)) {
                                return sedeNames.get(secKey).get(asgTipo);
                            }
                            if (typeof catalogMap !== 'undefined' && catalogMap && catalogMap.size > 0) {
                                let bestLoc = null, bestPad = -1;
                                for (const [key, v] of stats) {
                                    if (key.startsWith(mun + '-' + sec + '-')) {
                                        const locNum = key.split('-')[2];
                                        const cKey = mun + '-' + sec + '-' + locNum;
                                        if (catalogMap.has(cKey) && v.pad > bestPad) {
                                            bestPad = v.pad;
                                            bestLoc = catalogMap.get(cKey);
                                        }
                                    }
                                }
                                if (bestLoc) return bestLoc;
                                const k1 = mun + '-' + sec + '-1';
                                if (catalogMap.has(k1)) return catalogMap.get(k1);
                            }
                            return 'Localidad N/D';
                        };

                        if (groupData.has('B')) {
                            const { pad, lis } = groupData.get('B');
                            const divValue = base === 'lista' ? lis : pad;
                            const num = divValue > 0 ? Math.ceil(divValue / 750) : 0;
                            const padPort = distribute(pad, num);
                            const lisPort = distribute(lis, num);
                            const tiposB = [];
                            for (let i = 0; i < num; i++) {
                                const tipo = i === 0 ? 'B' : `C${i}`;
                                tiposB.push(tipo);
                                output.push([11, mun, sec, tipo, padPort[i], lisPort[i]]);
                                if (i === 0) {
                                    agg.bas_pad += padPort[i]; agg.bas_lis += lisPort[i]; agg.bas_count += 1;
                                } else {
                                    agg.con_pad += padPort[i]; agg.con_lis += lisPort[i]; agg.con_count += 1;
                                }
                            }
                            if (num > 0) {
                                const nombreLoc = getSedeNombre('B');
                                const rKey = mun + '|' + sec + '|B|' + nombreLoc;
                                if (!resumenMap.has(rKey)) {
                                    resumenMap.set(rKey, { distrito: 11, mun, sec, nombre: nombreLoc, grupo: 'B', tipos: [] });
                                }
                                resumenMap.get(rKey).tipos.push(...tiposB);
                            }
                        }

                        const eAsgs = Array.from(groupData.keys())
                            .filter(k => k.startsWith('E'))
                            .sort((a, b) => parseInt(a.slice(1)) - parseInt(b.slice(1)));

                        for (const asg of eAsgs) {
                            const { pad, lis } = groupData.get(asg);
                            const divValue = base === 'lista' ? lis : pad;
                            const num = divValue > 0 ? Math.ceil(divValue / 750) : 0;
                            const padPort = distribute(pad, num);
                            const lisPort = distribute(lis, num);
                            const tiposE = [];
                            for (let i = 0; i < num; i++) {
                                const tipo = i === 0 ? asg : `${asg}C${i}`;
                                tiposE.push(tipo);
                                output.push([11, mun, sec, tipo, padPort[i], lisPort[i]]);
                                if (i === 0) {
                                    agg.ext_pad += padPort[i]; agg.ext_lis += lisPort[i]; agg.ext_count += 1;
                                } else {
                                    agg.exc_pad += padPort[i]; agg.exc_lis += lisPort[i]; agg.exc_count += 1;
                                }
                            }
                            if (num > 0) {
                                const nombreLoc = getSedeNombre(asg);
                                const rKey = mun + '|' + sec + '|' + asg + '|' + nombreLoc;
                                if (!resumenMap.has(rKey)) {
                                    resumenMap.set(rKey, { distrito: 11, mun, sec, nombre: nombreLoc, grupo: asg, tipos: [] });
                                }
                                resumenMap.get(rKey).tipos.push(...tiposE);
                            }
                        }

                        const normSec = String(parseInt(sec, 10));
                        if (especialRequests.has(normSec)) {
                            const count = especialRequests.get(normSec);
                            for (let i = 0; i < count; i++) {
                                const tipo = count === 1 ? 'S' : `S${i+1}`;
                                output.push([11, mun, sec, tipo, 0, 0]);
                                agg.esp_count += 1;
                            }
                            especialRequests.delete(normSec);
                        }
                    }

                    if (especialRequests.size > 0) {
                        if (!munAgg.has('N/A')) {
                            munAgg.set('N/A', {
                                bas_pad: 0, bas_lis: 0, bas_count: 0, con_pad: 0, con_lis: 0, con_count: 0,
                                ext_pad: 0, ext_lis: 0, ext_count: 0, exc_pad: 0, exc_lis: 0, exc_count: 0,
                                esp_count: 0
                            });
                        }
                        const aggNA = munAgg.get('N/A');
                        for (const [secNum, count] of especialRequests) {
                            for (let i = 0; i < count; i++) {
                                const tipo = count === 1 ? 'S' : `S${i+1}`;
                                output.push([11, 'N/A', secNum, tipo, 0, 0]);
                                aggNA.esp_count += 1;
                            }
                        }
                    }

                    const sortedMuns = Array.from(munAgg.keys()).sort((a, b) => {
                        if (a === 'N/A') return 1;
                        if (b === 'N/A') return -1;
                        return Number(a) - Number(b);
                    });

                    // ── EXCEL DETALLE ──
                    const ws = XLSX.utils.aoa_to_sheet(output);
                    const wbout = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wbout, ws, "Proyeccion");
                    excelDetalleGlobal = XLSX.write(wbout, { bookType: 'xlsx', type: 'array' });

                    // ── EXCEL RESUMEN ──
                    const resumenOutput = [["Distrito", "Municipio", "Sección", "Nombre Localidad", "Numero de casillas", "Tipo de Casillas"]];
                    const sortedResumen = Array.from(resumenMap.values()).sort(function(a, b) {
                        const grupoOrder = function(g) {
                            if (!g || g === 'B') return 0;
                            if (g.startsWith('E')) return 100 + (parseInt(g.slice(1), 10) || 0);
                            return 999;
                        };
                        return Number(a.mun) - Number(b.mun)
                            || Number(a.sec) - Number(b.sec)
                            || grupoOrder(a.grupo) - grupoOrder(b.grupo)
                            || a.nombre.localeCompare(b.nombre);
                    });
                    for (const r of sortedResumen) {
                        const uniqueTipos = [...new Set(r.tipos)];
                        uniqueTipos.sort(function(x, y) {
                            function getOrder(t) {
                                if (t === 'B') return 0;
                                if (t.startsWith('C')) return 100 + parseInt(t.slice(1) || '0', 10);
                                if (t.startsWith('E')) {
                                    const m = t.match(/^E(\d+)(C(\d+))?$/);
                                    if (m) return 1000 + parseInt(m[1], 10) * 100 + (m[3] ? parseInt(m[3], 10) : 0);
                                }
                                if (t.startsWith('S')) return 5000;
                                return 9000;
                            }
                            return getOrder(x) - getOrder(y);
                        });
                        resumenOutput.push([r.distrito, r.mun, r.sec, r.nombre, uniqueTipos.length, uniqueTipos.join(', ')]);
                    }
                    const wsRes = XLSX.utils.aoa_to_sheet(resumenOutput);
                    const wbRes = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wbRes, wsRes, "Resumen");
                    excelResumenGlobal = XLSX.write(wbRes, { bookType: 'xlsx', type: 'array' });

                    if (btnZip) btnZip.style.display = 'block';

                    // ── TABLA GLOBAL ──
                    document.getElementById('global-stats-content').innerHTML = `
                        <ul>
                            <li><strong>Total Padrón Electoral:</strong> ${totalPad.toLocaleString('es-MX')}</li>
                            <li><strong>Total Lista Nominal:</strong> ${totalLis.toLocaleString('es-MX')}</li>
                            <li><strong>Número de Secciones:</strong> ${totalSecs.size}</li>
                            <li><strong>Número de Municipios:</strong> ${muns.size}</li>
                            <li><strong>Base usada:</strong> ${base === 'padron' ? 'Padrón Electoral' : 'Lista Nominal'}</li>
                        </ul>
                    `;

                    // ── TABLA MUNICIPIOS ──
                    let tableHTML = `
                        <table>
                            <thead><tr><th>Municipio</th><th>Básicas</th><th>Contiguas</th><th>Extraordinarias</th><th>Ext. Contiguas</th><th>Especiales</th><th>Total Casillas</th></tr></thead>
                            <tbody>
                    `;
                    let grandBas = 0, grandCon = 0, grandExt = 0, grandExc = 0, grandEsp = 0, grandTotal = 0;
                    for (const mun of sortedMuns) {
                        const agg = munAgg.get(mun);
                        const totalCas = agg.bas_count + agg.con_count + agg.ext_count + agg.exc_count + agg.esp_count;
                        grandBas += agg.bas_count; grandCon += agg.con_count;
                        grandExt += agg.ext_count; grandExc += agg.exc_count; grandEsp += agg.esp_count; grandTotal += totalCas;
                        tableHTML += `<tr><td>${mun}</td><td>${agg.bas_count}</td><td>${agg.con_count}</td><td>${agg.ext_count}</td><td>${agg.exc_count}</td><td>${agg.esp_count}</td><td style="font-weight:bold;">${totalCas}</td></tr>`;
                    }
                    tableHTML += `</tbody><tfoot><tr class="total"><td>TOTAL</td><td>${grandBas}</td><td>${grandCon}</td><td>${grandExt}</td><td>${grandExc}</td><td>${grandEsp}</td><td>${grandTotal}</td></tr></tfoot></table>`;

                    // ── CÁLCULO DE UNIDADES SIN ASIGNAR ──
                    const unidadesSinAsignar = [];
                    for (const [key, val] of stats.entries()) {
                        const parts = key.split('-');
                        const mun = parts[0];
                        const sec = parts[1];
                        const unit = parts.slice(2).join('-');
                        const secKey = `${mun}-${sec}`;

                        if (!groups.has(secKey)) continue;

                        const asignaciones = groups.get(secKey);

                        let tieneExtra = false;
                        for (const asg of asignaciones.keys()) {
                            if (asg.startsWith('E')) { tieneExtra = true; break; }
                        }
                        if (!tieneExtra) continue;

                        let unitAsignada = false;
                        for (const unitSets of asignaciones.values()) {
                            for (const assigned of unitSets) {
                                if (assigned === unit) { unitAsignada = true; break; }
                                if (!assigned.includes('-')) {
                                    const assignedLoc = assigned;
                                    const unitLoc = unit.includes('-') ? unit.split('-')[0] : unit;
                                    if (assignedLoc === unitLoc) { unitAsignada = true; break; }
                                }
                            }
                            if (unitAsignada) break;
                        }
                        if (!unitAsignada) {
                            unidadesSinAsignar.push({ mun, sec, loc: unit, pad: val.pad, lis: val.lis });
                        }
                    }

                    unidadesSinAsignar.sort((a, b) => Number(a.mun) - Number(b.mun) || Number(a.sec) - Number(b.sec) || a.loc.localeCompare(b.loc));

                    let tableSinAsignarHTML = `
                        <div class="stats-card" style="margin-top: 20px;">
                            <h3 style="color: #d9534f; margin-bottom: 5px;">⚠️ Unidades Sin Asignar</h3>
                            <p style="font-size: 0.9em; color: #555; margin-bottom: 15px;">
                                Unidades que aparecen en el estadístico, su sección tiene al menos una casilla Extraordinaria, 
                                pero la unidad no tiene asignación explícita.
                            </p>
                    `;

                    if (unidadesSinAsignar.length > 0) {
                        tableSinAsignarHTML += `
                            <div class="table-responsive">
                                <table style="width: 100%; border-collapse: collapse; text-align: center;">
                                    <thead><tr style="background-color: #f9f2f2;"><th>Municipio</th><th>Sección</th><th>Unidad</th><th>Padrón</th><th>Lista</th></tr></thead><tbody>
                        `;
                        for (const item of unidadesSinAsignar) {
                            tableSinAsignarHTML += `<tr><td>${item.mun}</td><td>${item.sec}</td><td style="font-weight:bold;">${item.loc}</td><td>${item.pad}</td><td>${item.lis}</td></tr>`;
                        }
                        tableSinAsignarHTML += `</tbody></table></div>`;
                    } else {
                        tableSinAsignarHTML += `<div style="background-color:#d4edda;color:#155724;padding:12px;border-radius:4px;text-align:center;">✓ Todas las unidades están correctamente asignadas.</div>`;
                    }
                    tableSinAsignarHTML += `</div>`;

                    // Insertar ambas tablas en el contenedor
                    document.getElementById('municipal-table-container').innerHTML = tableHTML + tableSinAsignarHTML;

                    resultsDiv.style.display = 'block';
                    setTimeout(() => {
                        resultsDiv.style.opacity = '1';
                        resultsDiv.style.transform = 'translateY(0)';
                        resultsDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    }, 50);

                    document.querySelectorAll('.stats-card').forEach((card, i) => {
                        setTimeout(() => card.classList.add('visible'), 300 + i * 250);
                    });

                    loading.style.display = 'none';

                } catch (err) {
                    console.error("Error en procesamiento:", err);
                    alert("Ocurrió un error durante el procesamiento.\nRevisa la consola (F12) para detalles.");
                    loading.style.display = 'none';
                }
            };
            reader2.readAsBinaryString(estadFile);
        } catch (err) {
            console.error("Error al leer conformación:", err);
            alert("Error al leer el archivo de conformación.");
            loading.style.display = 'none';
        }
    };
    reader1.readAsBinaryString(conformFile);
}
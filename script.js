
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
        if (option.checked) {
            base = option.value;
            break;
        }
    }

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

    document.querySelectorAll('.file-check').forEach(check => check.classList.remove('visible'));
    document.querySelectorAll('.stats-card').forEach(card => card.classList.remove('visible'));

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

                    document.querySelectorAll('.file-check').forEach(check => check.classList.add('visible'));

                    // ────────────────────────────────────────────────
                    // ESTADÍSTICO → clave por unidad (loc o loc-manzana)
                    // ────────────────────────────────────────────────
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

                    // PRECOMPUTE LOCALITY SUMS (para que sea ultra-rápido y no se congele)
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

                    // ────────────────────────────────────────────────
                    // CONFORMACIÓN → clave por unidad (¡ÍNDICE CORREGIDO!)
                    // ────────────────────────────────────────────────
                    const groups = new Map();
                    for (let i = 1; i < sheet1.length; i++) {
                        const row = sheet1[i];
                        if (row.length < 9) continue;
                        const mun = String(row[3]).trim();
                        const sec = String(row[4]).trim();
                        const loc = String(row[5]).trim();
                        const manzana = String(row[7] || '**').trim();   // ← COLUMNA MANZANA = índice 7
                        const asg = (row[8] || '').trim();
                        if (!asg) continue;

                        const unit = (manzana === '**') ? loc : `${loc}-${manzana}`;
                        const secKey = `${mun}-${sec}`;
                        if (!groups.has(secKey)) groups.set(secKey, new Map());
                        const secGroups = groups.get(secKey);
                        if (!secGroups.has(asg)) secGroups.set(asg, new Set());
                        secGroups.get(asg).add(unit);
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
                            bas_pad: 0, bas_lis: 0, bas_count: 0,
                            con_pad: 0, con_lis: 0, con_count: 0,
                            ext_pad: 0, ext_lis: 0, ext_count: 0,
                            exc_pad: 0, exc_lis: 0, exc_count: 0
                        });
                    }

                    const output = [["Distrito", "Municipio", "Sección", "Tipo de Casilla", "Padrón Electoral", "Listado nominal"]];

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
                            // SECCIÓN CON CONFORMACIÓN → respeta E* y manzanas
                            for (const [asg, units] of groups.get(secKey)) {
                                let pad = 0, lis = 0;
                                for (const assignedUnit of units) {
                                    if (!assignedUnit.includes('-')) {
                                        // LOCALIDAD COMPLETA (**)
                                        const locKey = `${mun}-${sec}-${assignedUnit}`;
                                        if (localitySums.has(locKey)) {
                                            const sum = localitySums.get(locKey);
                                            pad += sum.pad;
                                            lis += sum.lis;
                                        }
                                    } else {
                                        // MANZANA ESPECÍFICA
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

                            // Resto de la sección → Básica
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
                            // SECCIÓN SIN CONFORMACIÓN
                            if (totalPadSec > 0 || totalLisSec > 0) {
                                groupData.set('B', { pad: totalPadSec, lis: totalLisSec });
                            }
                        }

                        const agg = munAgg.get(mun);

                        // Básicas + Contiguas
                        if (groupData.has('B')) {
                            const { pad, lis } = groupData.get('B');
                            const divValue = base === 'lista' ? lis : pad;
                            const num = divValue > 0 ? Math.ceil(divValue / 750) : 0;

                            const padPort = distribute(pad, num);
                            const lisPort = distribute(lis, num);

                            for (let i = 0; i < num; i++) {
                                const tipo = i === 0 ? 'B' : `C${i}`;
                                output.push([11, mun, sec, tipo, padPort[i], lisPort[i]]);

                                if (i === 0) {
                                    agg.bas_pad += padPort[i];
                                    agg.bas_lis += lisPort[i];
                                    agg.bas_count += 1;
                                } else {
                                    agg.con_pad += padPort[i];
                                    agg.con_lis += lisPort[i];
                                    agg.con_count += 1;
                                }
                            }
                        }

                        // Extraordinarias (E*)
                        const eAsgs = Array.from(groupData.keys())
                            .filter(k => k.startsWith('E'))
                            .sort((a, b) => parseInt(a.slice(1)) - parseInt(b.slice(1)));

                        for (const asg of eAsgs) {
                            const { pad, lis } = groupData.get(asg);
                            const divValue = base === 'lista' ? lis : pad;
                            const num = divValue > 0 ? Math.ceil(divValue / 750) : 0;

                            const padPort = distribute(pad, num);
                            const lisPort = distribute(lis, num);

                            for (let i = 0; i < num; i++) {
                                const tipo = i === 0 ? asg : `${asg}C${i}`;
                                output.push([11, mun, sec, tipo, padPort[i], lisPort[i]]);

                                if (i === 0) {
                                    agg.ext_pad += padPort[i];
                                    agg.ext_lis += lisPort[i];
                                    agg.ext_count += 1;
                                } else {
                                    agg.exc_pad += padPort[i];
                                    agg.exc_lis += lisPort[i];
                                    agg.exc_count += 1;
                                }
                            }
                        }
                    }

                    // ── Excel y tablas (sin cambios) ──
                    const sortedMuns = Array.from(munAgg.keys()).sort((a, b) => Number(a) - Number(b));

                    const ws = XLSX.utils.aoa_to_sheet(output);
                    const wbout = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wbout, ws, "Proyeccion");

                    const blob = new Blob([XLSX.write(wbout, { bookType: 'xlsx', type: 'array' })],
                                         { type: 'application/octet-stream' });
                    const download = document.getElementById('download');
                    download.href = URL.createObjectURL(blob);
                    download.style.display = 'block';

                    document.getElementById('global-stats-content').innerHTML = `
                        <ul>
                            <li><strong>Total Padrón Electoral:</strong> ${totalPad.toLocaleString('es-MX')}</li>
                            <li><strong>Total Lista Nominal:</strong> ${totalLis.toLocaleString('es-MX')}</li>
                            <li><strong>Número de Secciones:</strong> ${totalSecs.size}</li>
                            <li><strong>Número de Municipios:</strong> ${muns.size}</li>
                            <li><strong>Base usada:</strong> ${base === 'padron' ? 'Padrón Electoral' : 'Lista Nominal'}</li>
                        </ul>
                    `;

                    let tableHTML = `
                        <table>
                            <thead><tr><th>Municipio</th><th>Básicas</th><th>Contiguas</th><th>Extraordinarias</th><th>Ext. Contiguas</th><th>Total Casillas</th></tr></thead>
                            <tbody>
                    `;
                    let grandBas = 0, grandCon = 0, grandExt = 0, grandExc = 0, grandTotal = 0;
                    for (const mun of sortedMuns) {
                        const agg = munAgg.get(mun);
                        const totalCas = agg.bas_count + agg.con_count + agg.ext_count + agg.exc_count;
                        grandBas += agg.bas_count; grandCon += agg.con_count;
                        grandExt += agg.ext_count; grandExc += agg.exc_count; grandTotal += totalCas;
                        tableHTML += `<tr><td>${mun}</td><td>${agg.bas_count}</td><td>${agg.con_count}</td><td>${agg.ext_count}</td><td>${agg.exc_count}</td><td>${totalCas}</td></tr>`;
                    }
                    tableHTML += `</tbody><tfoot><tr class="total"><td>TOTAL</td><td>${grandBas}</td><td>${grandCon}</td><td>${grandExt}</td><td>${grandExc}</td><td>${grandTotal}</td></tr></tfoot></table>`;

                    // UNIDADES SIN ASIGNAR (ya no mostrará las que están en E*)
                    const unidadesSinAsignar = [];
                    for (const [key, val] of stats.entries()) {
                        const parts = key.split('-');
                        const mun = parts[0];
                        const sec = parts[1];
                        const unit = parts.slice(2).join('-');
                        const secKey = `${mun}-${sec}`;

                        if (groups.has(secKey)) {
                            let unitAsignada = false;
                            const asignaciones = groups.get(secKey);
                            for (const unitSets of asignaciones.values()) {
                                for (const assigned of unitSets) {
                                    if (assigned === unit) {
                                        unitAsignada = true;
                                        break;
                                    }
                                    if (!assigned.includes('-')) {
                                        const assignedLoc = assigned;
                                        const unitLoc = unit.includes('-') ? unit.split('-')[0] : unit;
                                        if (assignedLoc === unitLoc) {
                                            unitAsignada = true;
                                            break;
                                        }
                                    }
                                }
                                if (unitAsignada) break;
                            }
                            if (!unitAsignada) {
                                unidadesSinAsignar.push({ mun, sec, loc: unit, pad: val.pad, lis: val.lis });
                            }
                        }
                    }

                    unidadesSinAsignar.sort((a, b) => Number(a.mun) - Number(b.mun) || Number(a.sec) - Number(b.sec) || a.loc.localeCompare(b.loc));

                    let tableSinAsignarHTML = `
                        <div class="stats-card" style="margin-top: 20px;">
                            <h3 style="color: #d9534f; margin-bottom: 5px;">⚠️ Unidades Sin Asignar</h3>
                            <p style="font-size: 0.9em; color: #555; margin-bottom: 15px;">
                                Unidades (localidades o manzanas) que aparecen en el estadístico, su sección existe en conformación, 
                                pero no tienen asignación de casilla.
                            </p>
                    `;

                    if (unidadesSinAsignar.length > 0) {
                        tableSinAsignarHTML += `
                            <table style="width: 100%; border-collapse: collapse; text-align: center;">
                                <thead><tr style="background-color: #f9f2f2;"><th>Municipio</th><th>Sección</th><th>Unidad</th><th>Padrón</th><th>Lista</th></tr></thead><tbody>
                        `;
                        for (const item of unidadesSinAsignar) {
                            tableSinAsignarHTML += `<tr><td>${item.mun}</td><td>${item.sec}</td><td style="font-weight:bold;">${item.loc}</td><td>${item.pad}</td><td>${item.lis}</td></tr>`;
                        }
                        tableSinAsignarHTML += `</tbody></table>`;
                    } else {
                        tableSinAsignarHTML += `<div style="background-color:#d4edda;color:#155724;padding:12px;border-radius:4px;text-align:center;">✓ Todas las unidades están correctamente asignadas.</div>`;
                    }
                    tableSinAsignarHTML += `</div>`;

                    document.getElementById('municipal-table-container').innerHTML = tableHTML + tableSinAsignarHTML;

                    resultsDiv.style.display = 'block';
                    setTimeout(() => {
                        resultsDiv.style.opacity = '1';
                        resultsDiv.style.transform = 'translateY(0)';
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
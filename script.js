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
let distritoActualGlobal = null;
let entidadActualGlobal = null;

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
        const sufijoEntidad = entidadActualGlobal ? `_Ent_${entidadActualGlobal}` : "";
        const sufijoDistrito = distritoActualGlobal ? `_Dto_${distritoActualGlobal}` : "";
        a.download = `Proyeccion_Casillas${sufijoEntidad}${sufijoDistrito}.zip`;
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

// Función auxiliar para leer archivos Excel con Promesas
function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const wb = XLSX.read(e.target.result, { type: 'binary' });
                const sheet = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
                resolve(sheet);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = (err) => reject(err);
        reader.readAsBinaryString(file);
    });
}

// Normaliza códigos numéricos (ej. "03" -> 3, "07" -> 7) o texto limpio
function normalizarCodigo(val) {
    if (val === undefined || val === null) return '';
    const str = String(val).trim();
    if (str === '') return '';
    return /^\d+$/.test(str) ? Number(str) : str;
}

async function processFiles() {
    const conformFile = document.getElementById('conformacion').files[0];
    const estadFile = document.getElementById('estadistico').files[0];
    const catalogoFile = document.getElementById('catalogo').files[0];
    
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
        alert('Por favor, suba ambos archivos obligatorios (Conformación y Estadístico).');
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

    try {
        const [sheet1, sheet2, sheet3] = await Promise.all([
            readExcelFile(conformFile),
            readExcelFile(estadFile),
            catalogoFile ? readExcelFile(catalogoFile) : Promise.resolve(null)
        ]);

        // ─── 1. VALIDAR Y EXTRAER ENTIDAD Y DISTRITO OBLIGATORIOS EN TODOS LOS ARCHIVOS ───
        const entidadesConformacion = new Set();
        const distritosConformacion = new Set();
        for (let i = 1; i < sheet1.length; i++) {
            const row = sheet1[i];
            if (!row || row.length === 0 || row.every(cell => cell === undefined || String(cell).trim() === '')) continue;
            const ent = normalizarCodigo(row[0]); // Columna A: Entidad
            const dto = normalizarCodigo(row[1]); // Columna B: Distrito
            if (ent === '') {
                throw new Error(`El archivo de Conformación tiene filas sin valor en la columna obligatoria "Entidad" (Columna A, fila ${i + 1}).`);
            }
            if (dto === '') {
                throw new Error(`El archivo de Conformación tiene filas sin valor en la columna obligatoria "Distrito" (Columna B, fila ${i + 1}).`);
            }
            entidadesConformacion.add(ent);
            distritosConformacion.add(dto);
        }
        if (distritosConformacion.size === 0) {
            throw new Error('No se encontró información válida con "Entidad" y "Distrito" en el archivo de Conformación.');
        }

        const entidadesEstadistico = new Set();
        const distritosEstadistico = new Set();
        for (let i = 1; i < sheet2.length; i++) {
            const row = sheet2[i];
            if (!row || row.length === 0 || row.every(cell => cell === undefined || String(cell).trim() === '')) continue;
            const ent = normalizarCodigo(row[0]); // Columna A: Entidad
            const dto = normalizarCodigo(row[1]); // Columna B: Distrito
            if (ent === '') {
                throw new Error(`El archivo Estadístico tiene filas sin valor en la columna obligatoria "Entidad" (Columna A, fila ${i + 1}).`);
            }
            if (dto === '') {
                throw new Error(`El archivo Estadístico tiene filas sin valor en la columna obligatoria "Distrito" (Columna B, fila ${i + 1}).`);
            }
            entidadesEstadistico.add(ent);
            distritosEstadistico.add(dto);
        }
        if (distritosEstadistico.size === 0) {
            throw new Error('No se encontró información válida con "Entidad" y "Distrito" en el archivo Estadístico.');
        }

        // Catálogo de Localidades: A=Entidad, B=Distrito, C=Municipio, D=Sección, E=Localidad, F=Nombre
        const catalogMap = new Map();
        const entidadesCatalogo = new Set();
        const distritosCatalogo = new Set();
        if (sheet3) {
            for (let i = 1; i < sheet3.length; i++) {
                const row = sheet3[i];
                if (!row || row.length === 0 || row.every(cell => cell === undefined || String(cell).trim() === '')) continue;
                const ent = normalizarCodigo(row[0]); // Columna A: Entidad
                const dto = normalizarCodigo(row[1]); // Columna B: Distrito
                if (ent === '') {
                    throw new Error(`El Catálogo de Localidades tiene filas sin valor en la columna obligatoria "Entidad" (Columna A, fila ${i + 1}).`);
                }
                if (dto === '') {
                    throw new Error(`El Catálogo de Localidades tiene filas sin valor en la columna obligatoria "Distrito" (Columna B, fila ${i + 1}).`);
                }
                entidadesCatalogo.add(ent);
                distritosCatalogo.add(dto);

                const mun = String(row[2] ?? '').trim();
                const sec = String(row[3] ?? '').trim();
                const loc = String(row[4] ?? '').trim();
                const nomLoc = String(row[5] ?? '').trim();
                if (mun && sec && loc && nomLoc) {
                    catalogMap.set(`${ent}-${dto}-${mun}-${sec}-${loc}`, nomLoc);
                }
            }
        }

        // Verificar consistencia de Entidad entre los archivos subidos
        const entConfStr = Array.from(entidadesConformacion).join(', ');
        const entEstStr = Array.from(entidadesEstadistico).join(', ');

        if (entidadesConformacion.size > 1 || entidadesEstadistico.size > 1 || entConfStr !== entEstStr) {
            throw new Error(
                `Los archivos no corresponden a la misma Entidad:\n` +
                `• Conformación: Entidad ${entConfStr}\n` +
                `• Estadístico: Entidad ${entEstStr}\n` +
                `Verifique que ambos archivos pertenezcan a la misma Entidad Federativa.`
            );
        }

        if (sheet3 && entidadesCatalogo.size > 0) {
            const entCatStr = Array.from(entidadesCatalogo).join(', ');
            if (entCatStr !== entEstStr) {
                throw new Error(
                    `El Catálogo de Localidades (Entidad ${entCatStr}) no coincide con la Entidad de los demás archivos (Entidad ${entEstStr}).`
                );
            }
        }

        // Verificar consistencia de Distrito entre los archivos subidos
        const dtoConfStr = Array.from(distritosConformacion).join(', ');
        const dtoEstStr = Array.from(distritosEstadistico).join(', ');

        if (distritosConformacion.size > 1 || distritosEstadistico.size > 1 || dtoConfStr !== dtoEstStr) {
            throw new Error(
                `Los archivos no corresponden al mismo Distrito:\n` +
                `• Conformación: Distrito ${dtoConfStr}\n` +
                `• Estadístico: Distrito ${dtoEstStr}\n` +
                `Verifique que ambos archivos pertenezcan a la misma Junta Distrital.`
            );
        }

        if (sheet3 && distritosCatalogo.size > 0) {
            const dtoCatStr = Array.from(distritosCatalogo).join(', ');
            if (dtoCatStr !== dtoEstStr) {
                throw new Error(
                    `El Catálogo de Localidades (Distrito ${dtoCatStr}) no coincide con el Distrito de los demás archivos (Distrito ${dtoEstStr}).`
                );
            }
        }

        const entidadActiva = Array.from(entidadesEstadistico)[0];
        const distritoActivo = Array.from(distritosEstadistico)[0];
        entidadActualGlobal = entidadActiva;
        distritoActualGlobal = distritoActivo;

        // ─── 2. PROCESAMIENTO DE ESTADÍSTICO ───
        const stats = new Map();
        let totalPad = 0, totalLis = 0, totalSecs = new Set();
        
        for (let i = 1; i < sheet2.length; i++) {
            const row = sheet2[i];
            if (!row || row.length < 8) continue;
            const ent = normalizarCodigo(row[0]);
            const dto = normalizarCodigo(row[1]);
            const mun = String(row[2]).trim();
            const sec = String(row[3]).trim();
            const loc = String(row[4]).trim();
            const manzana = String(row[5] || '**').trim();
            const pad = Number(row[6]) || 0;
            const lis = Number(row[7]) || 0;

            const unit = (manzana === '**') ? loc : `${loc}-${manzana}`;
            const key = `${mun}-${sec}-${unit}`;

            if (!stats.has(key)) stats.set(key, { pad: 0, lis: 0, entidad: ent, distrito: dto });
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

        // ─── 3. PROCESAMIENTO DE CONFORMACIÓN ───
        const groups = new Map();
        const sedeNames = new Map();
        const locNamesFromConf = new Map();

        for (let i = 1; i < sheet1.length; i++) {
            const row = sheet1[i];
            if (!row || row.length < 9) continue;
            const mun = String(row[3]).trim();
            const sec = String(row[4]).trim();
            const loc = String(row[5]).trim();
            const nombre = String(row[6] || '').trim();
            const manzana = String(row[7] || '**').trim();
            const asg = String(row[8] || '').trim();
            const sede = String(row[9] || '').trim().toUpperCase();
            if (!asg) continue;

            const unit = (manzana === '**') ? loc : `${loc}-${manzana}`;
            const secKey = `${mun}-${sec}`;
            if (!groups.has(secKey)) groups.set(secKey, new Map());
            const secGroups = groups.get(secKey);
            if (!secGroups.has(asg)) secGroups.set(asg, new Set());
            secGroups.get(asg).add(unit);

            if (nombre && loc) {
                locNamesFromConf.set(`${mun}-${sec}-${loc}`, nombre);
            }

            if (sede) {
                if (!sedeNames.has(secKey)) sedeNames.set(secKey, new Map());
                const nombreFinal = nombre || (loc ? `Localidad ${loc}` : '');
                if (nombreFinal) {
                    sedeNames.get(secKey).set(sede, nombreFinal);
                }
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

            const buscarNombreCatalogo = (locNum) => {
                if (!catalogMap || catalogMap.size === 0 || !locNum) return null;
                return catalogMap.get(`${entidadActiva}-${distritoActivo}-${mun}-${sec}-${locNum}`) || null;
            };

            const getSedeNombre = (asgTipo) => {
                const tipoNorm = String(asgTipo || '').trim().toUpperCase();

                // 1. Si Conformación tiene nombre explícito en Sede para este tipo (B, E1, etc.) y no es genérico
                if (sedeNames.has(secKey) && sedeNames.get(secKey).has(tipoNorm)) {
                    const nomSede = sedeNames.get(secKey).get(tipoNorm);
                    if (!nomSede.startsWith('Localidad ')) {
                        return nomSede;
                    }
                }

                // 2. Buscar localidad asociada en el grupo de asignación (B o E1, etc.)
                let locCandidata = null;
                if (groups.has(secKey) && groups.get(secKey).has(tipoNorm)) {
                    const unidadesAsignadas = Array.from(groups.get(secKey).get(tipoNorm));
                    if (unidadesAsignadas.length > 0) {
                        locCandidata = unidadesAsignadas[0].split('-')[0];
                    }
                }

                // 3. Si no hay unidad asignada, tomar la de mayor padrón en la sección
                if (!locCandidata) {
                    let bestPad = -1;
                    for (const [key, v] of stats) {
                        if (key.startsWith(`${mun}-${sec}-`) && v.pad > bestPad) {
                            bestPad = v.pad;
                            locCandidata = key.split('-')[2];
                        }
                    }
                }

                // 4. Consultar Catálogo de Localidades si está disponible
                if (locCandidata) {
                    const nomCat = buscarNombreCatalogo(locCandidata);
                    if (nomCat) return nomCat;
                }

                // 5. Consultar nombre capturado en Conformación para esa localidad
                if (locCandidata && locNamesFromConf.has(`${mun}-${sec}-${locCandidata}`)) {
                    return locNamesFromConf.get(`${mun}-${sec}-${locCandidata}`);
                }

                // 6. Último recurso: si ya venía "Localidad X" en sedeNames o localidad 1 en catálogo
                if (sedeNames.has(secKey) && sedeNames.get(secKey).has(tipoNorm)) {
                    return sedeNames.get(secKey).get(tipoNorm);
                }
                const nomLoc1 = buscarNombreCatalogo('1');
                if (nomLoc1) return nomLoc1;

                return locCandidata ? `Localidad ${locCandidata}` : 'Localidad N/D';
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
                    output.push([distritoActivo, mun, sec, tipo, padPort[i], lisPort[i]]);
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
                        resumenMap.set(rKey, { distrito: distritoActivo, mun, sec, nombre: nombreLoc, grupo: 'B', tipos: [] });
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
                    output.push([distritoActivo, mun, sec, tipo, padPort[i], lisPort[i]]);
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
                        resumenMap.set(rKey, { distrito: distritoActivo, mun, sec, nombre: nombreLoc, grupo: asg, tipos: [] });
                    }
                    resumenMap.get(rKey).tipos.push(...tiposE);
                }
            }

            const normSec = String(parseInt(sec, 10));
            if (especialRequests.has(normSec)) {
                const count = especialRequests.get(normSec);
                const tiposS = [];
                for (let i = 0; i < count; i++) {
                    const tipo = count === 1 ? 'S' : `S${i+1}`;
                    tiposS.push(tipo);
                    output.push([distritoActivo, mun, sec, tipo, 0, 0]);
                    agg.esp_count += 1;
                }
                const nombreLocEsp = getSedeNombre('B');
                const rKeyEsp = `${mun}|${sec}|S|${nombreLocEsp}`;
                if (!resumenMap.has(rKeyEsp)) {
                    resumenMap.set(rKeyEsp, {
                        distrito: distritoActivo,
                        mun,
                        sec,
                        nombre: nombreLocEsp,
                        grupo: 'S',
                        tipos: []
                    });
                }
                resumenMap.get(rKeyEsp).tipos.push(...tiposS);
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
                const tiposS = [];
                for (let i = 0; i < count; i++) {
                    const tipo = count === 1 ? 'S' : `S${i+1}`;
                    tiposS.push(tipo);
                    output.push([distritoActivo, 'N/A', secNum, tipo, 0, 0]);
                    aggNA.esp_count += 1;
                }
                const rKeyNA = `N/A|${secNum}|S|Sección no localizada`;
                if (!resumenMap.has(rKeyNA)) {
                    resumenMap.set(rKeyNA, {
                        distrito: distritoActivo,
                        mun: 'N/A',
                        sec: secNum,
                        nombre: 'Sección no localizada',
                        grupo: 'S',
                        tipos: []
                    });
                }
                resumenMap.get(rKeyNA).tipos.push(...tiposS);
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
            const munOrder = function(m) {
                if (m === 'N/A') return Number.MAX_SAFE_INTEGER;
                return Number(m);
            };
            const grupoOrder = function(g) {
                if (!g || g === 'B') return 0;
                if (g.startsWith('E')) return 100 + (parseInt(g.slice(1), 10) || 0);
                if (g === 'S') return 900;
                return 999;
            };
            return munOrder(a.mun) - munOrder(b.mun)
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
                    if (t.startsWith('S')) return 5000 + (parseInt(t.slice(1) || '0', 10));
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
                <li><strong>Entidad Federativa:</strong> ${entidadActiva}</li>
                <li><strong>Distrito Electoral:</strong> ${distritoActivo}</li>
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
                <thead><tr><th>Distrito</th><th>Municipio</th><th>Básicas</th><th>Contiguas</th><th>Extraordinarias</th><th>Ext. Contiguas</th><th>Especiales</th><th>Total Casillas</th></tr></thead>
                <tbody>
        `;
        let grandBas = 0, grandCon = 0, grandExt = 0, grandExc = 0, grandEsp = 0, grandTotal = 0;
        for (const mun of sortedMuns) {
            const agg = munAgg.get(mun);
            const totalCas = agg.bas_count + agg.con_count + agg.ext_count + agg.exc_count + agg.esp_count;
            grandBas += agg.bas_count; grandCon += agg.con_count;
            grandExt += agg.ext_count; grandExc += agg.exc_count; grandEsp += agg.esp_count; grandTotal += totalCas;
            tableHTML += `<tr><td>${distritoActivo}</td><td>${mun}</td><td>${agg.bas_count}</td><td>${agg.con_count}</td><td>${agg.ext_count}</td><td>${agg.exc_count}</td><td>${agg.esp_count}</td><td style="font-weight:bold;">${totalCas}</td></tr>`;
        }
        tableHTML += `</tbody><tfoot><tr class="total"><td colspan="2">TOTAL</td><td>${grandBas}</td><td>${grandCon}</td><td>${grandExt}</td><td>${grandExc}</td><td>${grandEsp}</td><td>${grandTotal}</td></tr></tfoot></table>`;

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
                <h3 style="color: #d9534f; margin-bottom: 5px;">⚠️ Localidades o manzanas sin asignar</h3>
                <p style="font-size: 0.9em; color: #555; margin-bottom: 15px;">
                    Localidades o manzanas que aparecen en el estadístico, pero en el archivo de conformación
                    no tienen asignación explícita.
                </p>
        `;

        if (unidadesSinAsignar.length > 0) {
            tableSinAsignarHTML += `
                <div class="table-responsive">
                    <table style="width: 100%; border-collapse: collapse; text-align: center;">
                        <thead><tr style="background-color: #f9f2f2;"><th>Distrito</th><th>Municipio</th><th>Sección</th><th>Unidad</th><th>Padrón</th><th>Lista</th></tr></thead><tbody>
            `;
            for (const item of unidadesSinAsignar) {
                tableSinAsignarHTML += `<tr><td>${distritoActivo}</td><td>${item.mun}</td><td>${item.sec}</td><td style="font-weight:bold;">${item.loc}</td><td>${item.pad}</td><td>${item.lis}</td></tr>`;
            }
            tableSinAsignarHTML += `</tbody></table></div>`;
        } else {
            tableSinAsignarHTML += `<div style="background-color:#d4edda;color:#155724;padding:12px;border-radius:4px;text-align:center;">✓ Todas las localidades y/o manzanas están correctamente asignadas.</div>`;
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
        alert(err.message || "Ocurrió un error durante el procesamiento.\nRevisa la consola (F12) para detalles.");
        loading.style.display = 'none';
    }
}
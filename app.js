/**
 * NCs Dashboard - Premium Dark Logic (Multi-select, Vendedor Column & Status Logic)
 */

document.addEventListener('DOMContentLoaded', () => {
    // State
    const state = {
        dataGeral: [],
        dataPessoal: [],
        filteredGeral: [],
        currentTableData: [],
        charts: {
            pie: null,
            line: null,
            sparks: {}
        },
        filters: {
            date: '',
            selectedSellers: [], 
            tableStart: '',
            tableEnd: '',
            statusTab: 'TODOS'
        }
    };

    // DOM Elements
    const fileGeral = document.getElementById('fileGeral');
    const filePessoal = document.getElementById('filePessoal');
    const btnUpdate = document.getElementById('btnUpdate');
    const tbody = document.getElementById('tabelaNCs');
    const dateFilter = document.getElementById('dateFilter');
    const tableStart = document.getElementById('tableStartDate');
    const tableEnd = document.getElementById('tableEndDate');
    
    // Multi-select elements
    const btnSellerFilter = document.getElementById('btnSellerFilter');
    const sellerDropdown = document.getElementById('sellerDropdown');
    const selectedSellersText = document.getElementById('selectedSellersText');

    // Display today's date in input but don't filter state by it yet
    dateFilter.value = new Date().toISOString().split('T')[0];
    initCharts();

    // Event Listeners
    btnUpdate.addEventListener('click', processarArquivos);
    dateFilter.addEventListener('change', (e) => { state.filters.date = e.target.value; applyFilters(); });
    tableStart.addEventListener('change', (e) => { state.filters.tableStart = e.target.value; applyFilters(); });
    tableEnd.addEventListener('change', (e) => { state.filters.tableEnd = e.target.value; applyFilters(); });

    // Multi-select toggle
    btnSellerFilter.addEventListener('click', (e) => {
        e.stopPropagation();
        sellerDropdown.classList.toggle('active');
    });

    // Status filter tabs
    const statusFilterTabs = document.getElementById('statusFilterTabs');
    if (statusFilterTabs) {
        statusFilterTabs.addEventListener('click', (e) => {
            const tab = e.target.closest('.status-tab');
            if (!tab) return;
            statusFilterTabs.querySelectorAll('.status-tab').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            state.filters.statusTab = tab.dataset.status;
            applyFilters();
        });
    }

    document.addEventListener('click', (e) => {
        if (!sellerDropdown.contains(e.target) && e.target !== btnSellerFilter) {
            sellerDropdown.classList.remove('active');
        }
    });

    const btnExportExcel = document.getElementById('btnExportExcel');
    if (btnExportExcel) {
        btnExportExcel.addEventListener('click', () => {
            if (!state.currentTableData || state.currentTableData.length === 0) {
                alert("Não há dados para exportar.");
                return;
            }
            
            const exportData = state.currentTableData.map(nc => {
                const registroGeral = state.dataGeral.find(g => g.pedido.toString() == nc.pedido.toString());
                const status = computeStatus(nc, registroGeral);
                const dateBase = formatExcelDate(nc.data);
                const dateNC = registroGeral ? (formatExcelDate(registroGeral.data) || '') : '';
                let motivo = (nc.motivo && nc.motivo.toString().trim() !== "") ? nc.motivo.toString().trim() : "";

                return {
                    "PEDIDO": nc.pedido,
                    "VENDEDOR": nc.vendedor,
                    "CODIGO": nc.codigo,
                    "DATA R. LOGISTICA": dateBase ? dateBase.split('-').reverse().join('/') : '',
                    "DATA R. VENDAS": dateNC ? dateNC.split('-').reverse().join('/') : '',
                    "STATUS": status,
                    "DIVERGENCIA": motivo,
                    "CLIENTE": nc.cliente
                };
            });

            const ws = XLSX.utils.json_to_sheet(exportData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Detalhamento NCs");
            XLSX.writeFile(wb, "Detalhamento_NCs.xlsx");
        });
    }

    async function processarArquivos() {
        const inputGeral = fileGeral.files[0];
        const inputPessoal = filePessoal.files[0];

        if (!inputPessoal) {
            return alert("Por favor, selecione ao menos o ARQUIVO PRINCIPAL!");
        }

        try {
            const [geralDataRaw, pessoalDataRaw] = await Promise.all([
                inputGeral ? readDataFile(inputGeral) : Promise.resolve([]),
                readDataFile(inputPessoal)
            ]);


            state.dataGeral = normalizeData(geralDataRaw);
            state.dataPessoal = normalizeData(pessoalDataRaw);

            console.log("Geral normalizado:", state.dataGeral.length);
            console.log("Principal normalizado:", state.dataPessoal.length);

            populateSellersFilter();
            applyFilters();
            
            btnUpdate.style.background = 'var(--success)';
            setTimeout(() => btnUpdate.style.background = 'var(--primary)', 2000);
            
        } catch (error) {
            console.error("Erro ao processar arquivo:", error);
            alert("Erro ao ler os arquivos. Verifique os formatos.");
        }
    }

    function readDataFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                let workbook;
                try {
                    workbook = XLSX.read(data, { type: 'array', cellDates: true });
                } catch (err) {
                    console.error("Erro XLSX:", err);
                    return reject(err);
                }
                
                let bestSheetName = workbook.SheetNames[0];
                let maxScore = -1;

                for (const name of workbook.SheetNames) {
                    const sheet = workbook.Sheets[name];
                    const csv = XLSX.utils.sheet_to_csv(sheet).toUpperCase();
                    let score = 0;
                    
                    // Check for headers (including synonyms)
                    if (csv.includes('PEDIDO') || csv.includes('PED')) score += 15;
                    if (csv.includes('VENDEDOR') || csv.includes('REPR') || csv.includes('VEND')) score += 10;
                    if (csv.includes('CODIGO') || csv.includes('CÓDIGO') || csv.includes('COD')) score += 10;
                    if (csv.includes('CLIENTE') || csv.includes('RAZÃO') || csv.includes('RAZAO')) score += 10;
                    if (csv.includes('DATA') || csv.includes('DT.')) score += 10;

                    // Boost for common sheet names
                    const upperName = name.toUpperCase();
                    if (upperName.includes('PLAN1') || upperName === 'BASE' || upperName === 'DADOS') score += 30;

                    const rowsCount = XLSX.utils.sheet_to_json(sheet).length;
                    
                    // Prioritize sheets with more data if they have some headers, 
                    // but don't let it overwhelm the sheet name bonus.
                    if (score > 0) {
                        score += Math.min(rowsCount / 100, 20); 
                    }


                    if (score > maxScore) {
                        maxScore = score;
                        bestSheetName = name;
                    }
                }


                console.log(`Usando aba: ${bestSheetName} (score: ${maxScore})`);

                const worksheet = workbook.Sheets[bestSheetName];
                const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
                
                let headerIndex = 0;
                for (let i = 0; i < Math.min(rows.length, 30); i++) {
                    const rowValues = rows[i].map(c => String(c).toUpperCase().trim());
                    if (rowValues.includes('PEDIDO') || rowValues.includes('VENDEDOR') || rowValues.some(v => v.includes('PEDIDO'))) {
                        headerIndex = i;
                        break;
                    }
                }

                // Verifica se a linha detectada contém os nomes de colunas como VALORES
                // (caso do Plan1: linha 0 tem "STATUS GERAL", linha 1 tem "PEDIDO","VENDEDOR",...)
                // Nesse caso, o XLSX.utils.sheet_to_json com range não funciona corretamente
                // porque usa a linha 0 real como header e gera chaves __EMPTY.
                // Solução: construir o JSON manualmente com a linha detectada como chaves.
                const headerRow = rows[headerIndex].map(c => String(c).trim());
                const hasRealColumnNames = headerRow.some(h => 
                    ['PEDIDO','VENDEDOR','CODIGO','DATA','CLIENTE','DIVERGENCIA'].includes(h.toUpperCase())
                );

                let jsonData;
                if (hasRealColumnNames) {
                    // Constrói manualmente: headerRow = chaves, linhas seguintes = dados
                    jsonData = [];
                    for (let r = headerIndex + 1; r < rows.length; r++) {
                        const row = rows[r];
                        if (row.every(cell => cell === "" || cell === null || cell === undefined)) continue;
                        const obj = {};
                        headerRow.forEach((key, idx) => {
                            obj[key || `__COL${idx}`] = row[idx] !== undefined ? row[idx] : "";
                        });
                        jsonData.push(obj);
                    }
                    console.log(`Header manual detectado na linha ${headerIndex}:`, headerRow);
                } else {
                    jsonData = XLSX.utils.sheet_to_json(worksheet, { range: headerIndex, defval: "" });
                }
                resolve(jsonData);
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    function normalizeData(data) {
        if (!Array.isArray(data)) return [];
        return data.map(item => {
            const findValue = (keys) => {
                const normalize = (s) => String(s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim();
                const normalizedKeys = keys.map(normalize);
                const foundKey = Object.keys(item).find(k => normalizedKeys.includes(normalize(k)));
                return foundKey ? item[foundKey] : "";
            };

            let pedidoRaw = findValue(['PEDIDO', 'STATUS GERAL', 'PEDIDO NUMERO', 'NÚMERO DO PEDIDO', 'PEDIDO_NUMERO', 'PED']);
            let pedido = pedidoRaw;

            
            if (typeof pedido === 'string') {
                if (pedido.startsWith("Cód: ")) {
                    pedido = pedido.replace("Cód: ", "").trim();
                }
            }

            if (!pedido || pedido === "0" || pedido === 0) {
                const alt = item['PEDIDO NUMERO'] || item['PEDIDO_NUMERO'];
                if (alt && alt !== "0" && alt !== 0) pedido = alt;
            }

            return {
                pedido: pedido ? pedido.toString().trim() : "",
                vendedor: findValue(['VENDEDOR', 'DESC.REPR/PREP', 'REPRESENTANTE', 'VEND']),
                data: findValue(['DATA R. LOGISTICA', 'DATA R. LOGÍSTICA', 'DATA', 'DT.FATUR', 'DATA DE FATURAMENTO', 'DT FATUR', 'DT. FATUR']),
                codigo: findValue(['CODIGO', 'CÓDIGO', 'COD', 'ID', 'CÓDIGO DO PRODUTO', 'CODIGO DO PRODUTO', 'PRODUTO']),
                cliente: findValue(['RAZÃO SOCIAL', 'RAZAO SOCIAL', 'CLIENTE', 'NOME']),
                motivo: findValue(['DIVERGENCIA', 'DIVERGÊNCIA', 'STATUS', 'SITUAÇÃO', 'OCORRÊNCIA', 'MOTIVO', 'OBS.COMERC', 'SITUACAO', 'OCORRENCIA', 'DIV']),
                status_raw: findValue(['STATUS', 'SITUAÇÃO'])
            };

        }).filter(nc => nc.pedido && nc.pedido !== "0" && nc.pedido !== "Cód: 0");
    }

    function populateSellersFilter() {
        const allSellers = [
            ...state.dataPessoal.map(nc => nc.vendedor),
            ...state.dataGeral.map(nc => nc.vendedor)
        ];
        const uniqueSellers = [...new Set(allSellers.filter(v => v && v.toString().trim() !== ""))].sort();
        
        sellerDropdown.innerHTML = '';
        state.filters.selectedSellers = []; 

        uniqueSellers.forEach(seller => {
            const label = document.createElement('label');
            label.className = 'seller-item';
            
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.value = seller;
            checkbox.addEventListener('change', () => {
                updateSelectedSellers();
                applyFilters();
            });

            label.appendChild(checkbox);
            label.appendChild(document.createTextNode(seller));
            sellerDropdown.appendChild(label);
        });

        updateSelectedSellers();
    }

    function updateSelectedSellers() {
        const checked = Array.from(sellerDropdown.querySelectorAll('input:checked')).map(i => i.value);
        state.filters.selectedSellers = checked;
        
        if (checked.length === 0) {
            selectedSellersText.textContent = "Todos os Vendedores";
        } else if (checked.length === 1) {
            selectedSellersText.textContent = checked[0];
        } else {
            selectedSellersText.textContent = `${checked.length} Vendedores`;
        }
    }

    function applyFilters() {
        const sellerFilter = (nc) => {
            if (state.filters.selectedSellers.length === 0) return true;
            return state.filters.selectedSellers.includes(nc.vendedor);
        };

        let baseData = state.dataPessoal.filter(sellerFilter);
        
        let metricData = baseData;
        if (state.filters.date) {
            metricData = metricData.filter(nc => formatExcelDate(nc.data) === state.filters.date);
        }

        let tableData = baseData;
        if (state.filters.tableStart) {
            tableData = tableData.filter(nc => formatExcelDate(nc.data) >= state.filters.tableStart);
        }
        if (state.filters.tableEnd) {
            tableData = tableData.filter(nc => formatExcelDate(nc.data) <= state.filters.tableEnd);
        }

        // Filtro por status tab
        if (state.filters.statusTab && state.filters.statusTab !== 'TODOS') {
            tableData = tableData.filter(nc => {
                const registroGeral = state.dataGeral.find(g => g.pedido.toString() == nc.pedido.toString());
                const computedStatus = computeStatus(nc, registroGeral);
                return computedStatus === state.filters.statusTab;
            });
        }

        state.currentTableData = tableData;
        updateDashboard(metricData, tableData);
    }

    function formatExcelDate(excelDate) {
        if (!excelDate) return "";
        let date;
        if (excelDate instanceof Date) {
            // Already a date, but might be at midnight UTC. 
            // We want the local date representation.
            date = excelDate;
        } else if (typeof excelDate === 'number') {
            // Excel serial date to JS date
            date = new Date(Math.round((excelDate - 25569) * 86400 * 1000));
            // Adjust for timezone to keep the same calendar day
            date.setMinutes(date.getMinutes() + date.getTimezoneOffset());
        } else {
            const dateStr = String(excelDate).trim();
            if (dateStr.includes('/')) {
                const p = dateStr.split('/');
                if (p.length === 3) {
                    const day = p[0].padStart(2, '0');
                    const month = p[1].padStart(2, '0');
                    const year = p[2].length === 2 ? '20' + p[2] : p[2];
                    date = new Date(`${year}-${month}-${day}T12:00:00`); // Use T12:00 to avoid day shift
                }
            } else if (dateStr.includes('-')) {
                date = new Date(dateStr + "T12:00:00");
            }
        }
        if (!date || isNaN(date.getTime())) return "";
        
        // Return YYYY-MM-DD
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    // Calcula o status de um pedido dado o registro geral correspondente
    function computeStatus(nc, registroGeral) {
        const dateBase = formatExcelDate(nc.data);
        if (!registroGeral) return 'PENDENTE';
        const dateNC = formatExcelDate(registroGeral.data);
        if (!dateNC) return 'PENDENTE';
        if (dateNC < dateBase) return 'REINCIDENTE';   // Vendas registrou ANTES da logística
        if (dateNC === dateBase) return 'CONFORME';
        return 'DIVERGENCIA';                           // Vendas registrou DEPOIS da logística
    }

    function updateDashboard(metricData, tableData) {
        const totalBase = metricData.length;
        
        const presentInNC = metricData.filter(nc => 
            state.dataGeral.some(g => g.pedido.toString() == nc.pedido.toString())
        );
        
        let countConforme = 0, countReincidente = 0, countDivergencia = 0;
        presentInNC.forEach(nc => {
            const g = state.dataGeral.find(item => item.pedido.toString() == nc.pedido.toString());
            const s = computeStatus(nc, g);
            if (s === 'CONFORME') countConforme++;
            else if (s === 'REINCIDENTE') countReincidente++;
            else if (s === 'DIVERGENCIA') countDivergencia++;
        });

        const countPendente = totalBase - presentInNC.length;

        updateMetric('valTotal', totalBase);
        updateMetric('valMim', presentInNC.length); 
        updateMetric('valComercial', countConforme);
        updateMetric('valNaoComercial', countPendente);
        updateMetric('valReincidente', countReincidente);

        const getPerc = (val) => totalBase > 0 ? ((val/totalBase)*100).toFixed(1) + "%" : "0%";
        document.getElementById('percMim').innerText = getPerc(presentInNC.length);
        document.getElementById('percComercial').innerText = getPerc(countConforme);
        document.getElementById('percNaoComercial').innerText = getPerc(countPendente);
        document.getElementById('percReincidente').innerText = getPerc(countReincidente);

        renderTable(tableData);
        updateCharts(metricData);
    }

    function updateMetric(id, value) {
        const el = document.getElementById(id);
        if (el) {
            const start = parseInt(el.innerText) || 0;
            animateValue(el, start, value, 500);
        }
    }

    function animateValue(obj, start, end, duration) {
        let startTimestamp = null;
        const step = (timestamp) => {
            if (!startTimestamp) startTimestamp = timestamp;
            const progress = Math.min((timestamp - startTimestamp) / duration, 1);
            obj.innerHTML = Math.floor(progress * (end - start) + start);
            if (progress < 1) window.requestAnimationFrame(step);
        };
        window.requestAnimationFrame(step);
    }

    function renderTable(data) {
        tbody.innerHTML = "";
        
        const totalInTable = data.length;
        const presentInNC = data.filter(nc => state.dataGeral.some(g => g.pedido.toString() == nc.pedido.toString()));
        const pendentes = totalInTable - presentInNC.length;
        
        document.getElementById('badgeCount').innerText = `${pendentes} Pedidos Pendentes`;

        if (data.length === 0) {
            tbody.innerHTML = `<tr><td colspan="8" class="empty-state">Nenhuma informação encontrada na base para os filtros selecionados.</td></tr>`;
            return;
        }

        tbody.innerHTML = data.map(nc => {
            const registroGeral = state.dataGeral.find(g => g.pedido.toString() == nc.pedido.toString());
            
            const status = computeStatus(nc, registroGeral);
            const dateBase = formatExcelDate(nc.data);
            const dateNC = registroGeral ? (formatExcelDate(registroGeral.data) || '---') : '---';
            let motivo = (nc.motivo && nc.motivo.toString().trim() !== "") ? nc.motivo.toString().trim() : "---";

            let statusClass;
            switch(status) {
                case 'CONFORME':    statusClass = 'status-done'; break;
                case 'REINCIDENTE': statusClass = 'status-reincidente'; break;
                case 'DIVERGENCIA': statusClass = 'status-warning'; break;
                default:            statusClass = 'status-pending'; break;
            }
            
            return `
                <tr>
                    <td style="font-weight: 700;">${nc.pedido || '---'}</td>
                    <td>${nc.vendedor || '---'}</td>
                    <td style="font-family: monospace;">${nc.codigo || '---'}</td>
                    <td style="text-align: center;">${dateBase ? dateBase.split('-').reverse().join('/') : '---'}</td>
                    <td style="text-align: center;">${dateNC !== '---' ? dateNC.split('-').reverse().join('/') : '---'}</td>
                    <td style="text-align: center;"><span class="status-badge ${statusClass}">${status}</span></td>
                    <td style="font-size: 0.8rem; color: var(--text-dim);">${motivo}</td>
                    <td style="font-size: 0.75rem; color: var(--text-muted); max-width: 250px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${nc.cliente || ''}">${nc.cliente || '---'}</td>
                </tr>
            `;
        }).join('');
    }

    function initCharts() {
        const ctxPie = document.getElementById('pieChart').getContext('2d');
        state.charts.pie = new Chart(ctxPie, {
            type: 'doughnut',
            data: { labels: [], datasets: [{ data: [], backgroundColor: [], borderWidth: 0 }] },
            options: { 
                cutout: '80%', 
                plugins: { legend: { position: 'bottom', labels: { color: '#94a3b8', font: { size: 10 } } } }, 
                maintainAspectRatio: false,
                onClick: (event, elements) => {
                    if (elements.length > 0) {
                        const index = elements[0].index;
                        const label = state.charts.pie.data.labels[index];
                        
                        if (state.charts.pie.showingSellers) {
                            // Filter by this seller
                            state.filters.selectedSellers = [label];
                            Array.from(sellerDropdown.querySelectorAll('input[type="checkbox"]')).forEach(cb => {
                                cb.checked = (cb.value === label);
                            });
                            selectedSellersText.textContent = label;
                        } else {
                            // Clear seller filter
                            state.filters.selectedSellers = [];
                            Array.from(sellerDropdown.querySelectorAll('input[type="checkbox"]')).forEach(cb => {
                                cb.checked = false;
                            });
                            selectedSellersText.textContent = "Todos os Vendedores";
                        }
                        applyFilters();
                    }
                }
            }
        });

        const ctxLine = document.getElementById('lineChart').getContext('2d');
        state.charts.line = new Chart(ctxLine, {
            type: 'line',
            data: { labels: [], datasets: [{ label: 'Não Registrados', data: [], borderColor: '#f97316', backgroundColor: 'rgba(249, 115, 22, 0.1)', fill: true, tension: 0.4 }] },
            options: { plugins: { legend: { display: false } }, scales: { y: { grid: { color: '#1e293b' }, ticks: { color: '#64748b' } }, x: { grid: { display: false }, ticks: { color: '#64748b' } } }, maintainAspectRatio: false }
        });

        initSpark('sparkTotal', '#3b82f6');
        initSpark('sparkMim', '#10b981');
        initSpark('sparkComercial', '#8b5cf6');
        initSpark('sparkNaoComercial', '#f97316');
        initSpark('sparkReincidente', '#ec4899');
    }

    function initSpark(id, color) {
        const ctx = document.getElementById(id).getContext('2d');
        state.charts.sparks[id] = new Chart(ctx, {
            type: 'line',
            data: { labels: [1,2,3,4,5,6,7,8,9,10], datasets: [{ data: [0,0,0,0,0,0,0,0,0,0], borderColor: color, borderWidth: 2, pointRadius: 0, tension: 0.4, fill: false }] },
            options: { plugins: { legend: { display: false }, tooltip: { enabled: false } }, scales: { x: { display: false }, y: { display: false } }, maintainAspectRatio: false, responsive: true }
        });
    }

    function updateCharts(metricData) {
        // Pie Chart: Panorama de Vendas Atual
        const showingStatuses = state.filters.selectedSellers.length === 1;
        state.charts.pie.showingSellers = !showingStatuses;

        if (!showingStatuses) {
            const sellerCounts = {};
            metricData.forEach(nc => {
                const seller = nc.vendedor || 'Desconhecido';
                sellerCounts[seller] = (sellerCounts[seller] || 0) + 1;
            });
            state.charts.pie.data.labels = Object.keys(sellerCounts);
            state.charts.pie.data.datasets[0].data = Object.values(sellerCounts);
            
            // Generate some colors based on sellers
            const colors = ['#3b82f6', '#8b5cf6', '#10b981', '#f97316', '#ef4444', '#ec4899', '#06b6d4', '#eab308'];
            state.charts.pie.data.datasets[0].backgroundColor = Object.keys(sellerCounts).map((_, i) => colors[i % colors.length]);
        } else {
            let counts = { CONFORME: 0, PENDENTE: 0, DIVERGENCIA: 0, REINCIDENTE: 0 };
            metricData.forEach(nc => {
                const g = state.dataGeral.find(item => item.pedido.toString() == nc.pedido.toString());
                const s = computeStatus(nc, g);
                counts[s] = (counts[s] || 0) + 1;
            });
            
            const labels = [];
            const data = [];
            const bg = [];
            
            if (counts.CONFORME > 0) { labels.push('Conforme'); data.push(counts.CONFORME); bg.push('#10b981'); }
            if (counts.PENDENTE > 0) { labels.push('Pendente'); data.push(counts.PENDENTE); bg.push('#ef4444'); }
            if (counts.DIVERGENCIA > 0) { labels.push('Divergência'); data.push(counts.DIVERGENCIA); bg.push('#f97316'); }
            if (counts.REINCIDENTE > 0) { labels.push('Reincidente'); data.push(counts.REINCIDENTE); bg.push('#ec4899'); }

            state.charts.pie.data.labels = labels;
            state.charts.pie.data.datasets[0].data = data;
            state.charts.pie.data.datasets[0].backgroundColor = bg;
        }
        state.charts.pie.update();

        // Line chart (Evolução)
        const dayCounts = {};
        metricData.forEach(nc => {
            const day = formatExcelDate(nc.data);
            if (day) dayCounts[day] = (dayCounts[day] || 0) + 1;
        });
        const sortedDays = Object.keys(dayCounts).sort();
        state.charts.line.data.labels = sortedDays.map(d => d.split('-').slice(1).reverse().join('/'));
        state.charts.line.data.datasets[0].data = sortedDays.map(d => dayCounts[d]);
        state.charts.line.update();

        Object.keys(state.charts.sparks).forEach(id => {
            const chart = state.charts.sparks[id];
            chart.data.datasets[0].data = Array.from({length: 10}, () => Math.floor(Math.random() * 50));
            chart.update();
        });
    }
});

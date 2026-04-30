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
        lineSellersData: {},
        filters: {
            date: '',
            selectedSellers: [],
            tableStart: '',
            tableEnd: '',
            statusTab: 'TODOS'
        }
    };

    // DOM Elements
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
    btnUpdate.addEventListener('click', loadFromDatabase);
    dateFilter.addEventListener('change', (e) => { state.filters.date = e.target.value; applyFilters(); });
    tableStart.addEventListener('change', (e) => { state.filters.tableStart = e.target.value; applyFilters(); });
    tableEnd.addEventListener('change', (e) => { state.filters.tableEnd = e.target.value; applyFilters(); });
    
    // Auto-load on start
    loadFromDatabase();

    async function loadFromDatabase() {
        btnUpdate.classList.add('loading');
        btnUpdate.disabled = true;
        
        try {
            const response = await fetch('../api/naoconformidade.php');
            const result = await response.json();

            if (result.status === 'success') {
                // Mapeia a Base da Logística (Antigo 2026.xlsx)
                const dataBase = result.dataBase || [];
                state.dataPessoal = dataBase.map(item => ({
                    pedido: item.numero_pedido || "",
                    vendedor: item.vendedor || "",
                    data: item.data_logistica || "",
                    codigo: item.codigo_produto || "",
                    cliente: item.cliente || "",
                    motivo: item.divergencia || ""
                }));

                // Mapeia as Não Conformidades (Geral)
                const dataNC = result.dataNC || [];
                state.dataGeral = dataNC.map(item => ({
                    pedido: item.numero_pedido || "",
                    vendedor: item.vendedor_responsavel || "",
                    data: item.data_ocorrencia || "",
                    codigo: item.codigo || "",
                    cliente: item.cliente_afetado || "",
                    motivo: item.titulo || "",
                    status_raw: item.status || ""
                }));

                console.log("Base Logística:", state.dataPessoal.length);
                console.log("NCs Geral:", state.dataGeral.length);

                populateSellersFilter();
                applyFilters();
                
                // feedback visual
                btnUpdate.style.background = 'var(--success)';
                setTimeout(() => btnUpdate.style.background = 'var(--primary)', 2000);
            }
        } catch (error) {
            console.error("Erro ao carregar do DB:", error);
        } finally {
            btnUpdate.classList.remove('loading');
            btnUpdate.disabled = false;
        }
    }

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

    const btnClearFilters = document.getElementById('btnClearFilters');
    if (btnClearFilters) {
        btnClearFilters.addEventListener('click', () => {
            state.filters.date = '';
            document.getElementById('dateFilter').value = '';

            state.filters.selectedSellers = [];
            Array.from(document.querySelectorAll('#sellerDropdown input[type="checkbox"]')).forEach(cb => cb.checked = false);
            document.getElementById('selectedSellersText').textContent = "Todos os Vendedores";

            state.filters.statusTab = 'TODOS';
            document.querySelectorAll('.status-tab').forEach(t => t.classList.remove('active'));
            const defaultTab = document.querySelector('.status-tab[data-status="TODOS"]');
            if (defaultTab) defaultTab.classList.add('active');

            state.filters.tableStart = '';
            state.filters.tableEnd = '';
            document.getElementById('tableStartDate').value = '';
            document.getElementById('tableEndDate').value = '';

            if (state.charts.pie) {
                state.charts.pie.showingSellers = false;
            }

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
        if (state.filters.date) {
            tableData = tableData.filter(nc => formatExcelDate(nc.data) === state.filters.date);
        } else {
            if (state.filters.tableStart) {
                tableData = tableData.filter(nc => formatExcelDate(nc.data) >= state.filters.tableStart);
            }
            if (state.filters.tableEnd) {
                tableData = tableData.filter(nc => formatExcelDate(nc.data) <= state.filters.tableEnd);
            }
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
        updateMetric('valTotal', totalBase);
        updateMetric('valComercial', countConforme);
        updateMetric('valNaoComercial', countPendente);
        updateMetric('valDivergencia', countDivergencia);
        updateMetric('valReincidente', countReincidente);

        const getPerc = (val) => totalBase > 0 ? ((val / totalBase) * 100).toFixed(1) + "%" : "0%";
        document.getElementById('percComercial').innerText = getPerc(countConforme) + " conforme";
        document.getElementById('percNaoComercial').innerText = getPerc(countPendente) + " pendente";
        document.getElementById('percDivergencia').innerText = getPerc(countDivergencia) + " divergência";
        document.getElementById('percReincidente').innerText = getPerc(countReincidente) + " reincidente";

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
            switch (status) {
                case 'CONFORME': statusClass = 'status-done'; break;
                case 'REINCIDENTE': statusClass = 'status-reincidente'; break;
                case 'DIVERGENCIA': statusClass = 'status-warning'; break;
                default: statusClass = 'status-pending'; break;
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
        // Plugin to draw text in the center of the doughnut chart
        const centerTextPlugin = {
            id: 'centerText',
            beforeDraw: function (chart) {
                if (chart.config.type !== 'doughnut') return;
                const ctx = chart.ctx;
                const width = chart.width;
                const height = chart.height;
                ctx.restore();

                let textLines = [];
                if (state.filters.selectedSellers.length > 0) {
                    if (state.filters.selectedSellers.length <= 2) {
                        textLines = state.filters.selectedSellers;
                    } else {
                        textLines = [`${state.filters.selectedSellers.length} Vendedores`];
                    }
                } else {
                    textLines = ["Todos"];
                }

                ctx.font = "bold 14px Inter";
                ctx.textBaseline = "middle";
                ctx.fillStyle = "#94a3b8";

                const lineHeight = 18;
                const totalHeight = textLines.length * lineHeight;
                const startY = (height - totalHeight) / 2 + (lineHeight / 2);

                textLines.forEach((line, i) => {
                    const textX = Math.round((width - ctx.measureText(line).width) / 2);
                    const textY = startY + (i * lineHeight);
                    ctx.fillText(line, textX, textY);
                });
                ctx.save();
            }
        };

        const outlabelsPlugin = {
            id: 'outlabels',
            afterDraw: function (chart) {
                if (chart.config.type !== 'doughnut') return;
                const ctx = chart.ctx;
                chart.data.datasets.forEach((dataset, i) => {
                    chart.getDatasetMeta(i).data.forEach((element, index) => {
                        const dataValue = dataset.data[index];
                        if (dataValue === 0) return;

                        const total = dataset.data.reduce((acc, val) => acc + val, 0);
                        const pct = Math.round((dataValue / total) * 100) + '%';

                        const midAngle = element.startAngle + (element.endAngle - element.startAngle) / 2;
                        const outerRadius = element.outerRadius;
                        const x = chart.chartArea.left + chart.chartArea.width / 2;
                        const y = chart.chartArea.top + chart.chartArea.height / 2;

                        const x1 = x + Math.cos(midAngle) * outerRadius;
                        const y1 = y + Math.sin(midAngle) * outerRadius;

                        const lineLength = 12;
                        const x2 = x + Math.cos(midAngle) * (outerRadius + lineLength);
                        const y2 = y + Math.sin(midAngle) * (outerRadius + lineLength);

                        const isRight = x2 > x;
                        const x3 = isRight ? x2 + 8 : x2 - 8;

                        ctx.beginPath();
                        ctx.moveTo(x1, y1);
                        ctx.lineTo(x2, y2);
                        ctx.lineTo(x3, y2);
                        ctx.strokeStyle = dataset.backgroundColor[index];
                        ctx.lineWidth = 1.5;
                        ctx.stroke();

                        ctx.font = "bold 11px Inter";
                        ctx.fillStyle = dataset.backgroundColor[index];
                        ctx.textBaseline = "middle";
                        ctx.textAlign = isRight ? "left" : "right";
                        ctx.fillText(pct, isRight ? x3 + 3 : x3 - 3, y2);
                    });
                });
            }
        };

        const verticalLinePlugin = {
            id: 'verticalLine',
            afterDraw: chart => {
                if (chart.config.type !== 'line') return;
                if (chart.tooltip?._active?.length) {
                    let x = chart.tooltip._active[0].element.x;
                    let yAxis = chart.scales.y;
                    let ctx = chart.ctx;
                    ctx.save();
                    ctx.beginPath();
                    ctx.moveTo(x, yAxis.top);
                    ctx.lineTo(x, yAxis.bottom);
                    ctx.lineWidth = 1;
                    ctx.strokeStyle = 'rgba(255, 255, 255, 0.4)';
                    ctx.setLineDash([5, 5]);
                    ctx.stroke();
                    ctx.restore();
                }
            }
        };
        Chart.register(centerTextPlugin, outlabelsPlugin, verticalLinePlugin);

const ctxPie = document.getElementById('pieChart').getContext('2d');
        state.charts.pie = new Chart(ctxPie, {
            type: 'doughnut',
            data: { labels: [], datasets: [{ data: [], backgroundColor: [], borderWidth: 0 }] },
            options: { 
                cutout: '55%', 
                rotation: 0, 
                animation: {
                    animateRotate: true,
                    duration: 800, // Tempo um pouquinho maior para o olho acompanhar
                    easing: 'easeOutQuart' // A MÁGICA DA FLUIDEZ: começa rápido e freia macio no final
                },
                layout: { padding: 0 },
                plugins: { 
                    legend: { 
                        position: 'bottom', 
                        align: 'center', 
                        labels: { color: '#94a3b8', font: { size: 12 }, padding: 15, boxWidth: 12 } 
                    } 
                }, 
                maintainAspectRatio: false,
                onClick: (event, elements) => {
                    if (elements.length > 0) {
                        const index = elements[0].index;
                        let label = state.charts.pie.data.labels[index];
                        
                        if (state.charts.pie.showingSellers) {
                            label = label.split(' (')[0];
                            state.filters.selectedSellers = [label];
                            Array.from(sellerDropdown.querySelectorAll('input[type="checkbox"]')).forEach(cb => {
                                cb.checked = (cb.value === label);
                            });
                            selectedSellersText.textContent = label;
                        } else {
                            state.filters.selectedSellers = [];
                            Array.from(sellerDropdown.querySelectorAll('input[type="checkbox"]')).forEach(cb => {
                                cb.checked = false;
                            });
                            selectedSellersText.textContent = "Todos os Vendedores";
                        }

                        // Manda a rosca girar 360 graus
                        state.charts.pie.options.rotation = (state.charts.pie.options.rotation || 0) + 360;
                        
                        // Atualiza os dados imediatamente junto com o giro
                        applyFilters();
                    }
                }
            }
        });

        const ctxLine = document.getElementById('lineChart').getContext('2d');
        state.charts.line = new Chart(ctxLine, {
            type: 'line',
            data: { labels: [], datasets: [] },
            options: {
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    legend: { display: true, position: 'top', labels: { color: '#94a3b8', font: { size: 10 } } },
                    tooltip: {
                        callbacks: {
                            afterBody: function (context) {
                                const index = context[0].dataIndex;
                                const label = state.charts.line.data.labels[index];
                                const dayKey = label.split('/').reverse().join('-');
                                const sellers = state.lineSellersData[dayKey] || {};
                                let text = ['', 'Vendedores neste dia:'];
                                for (let s in sellers) {
                                    text.push(`- ${s}: ${sellers[s]}`);
                                }
                                return text;
                            }
                        }
                    }
                },
                scales: { y: { grid: { color: '#1e293b' }, ticks: { color: '#64748b' } }, x: { grid: { display: false }, ticks: { color: '#64748b' } } },
                maintainAspectRatio: false,
                onClick: (event, elements) => {
                    const chart = state.charts.line;
                    if (!chart) return;

                    const xAxis = chart.scales.x;
                    const yAxis = chart.scales.y;
                    const eY = event.y !== undefined ? event.y : event.native.offsetY;
                    const eX = event.x !== undefined ? event.x : event.native.offsetX;

                    // Se clicou na área do eixo X (ou próximo das labels) ou num ponto da linha
                    if (eY >= xAxis.top || elements.length > 0) {
                        let index;
                        if (elements.length > 0) {
                            index = elements[0].index;
                        } else {
                            const dataX = xAxis.getValueForPixel(eX);
                            if (dataX === undefined) return;
                            index = Math.round(dataX);
                        }

                        const dayKey = Object.keys(state.lineSellersData)[index];
                        if (dayKey) {
                            document.getElementById('dateFilter').value = dayKey;
                            state.filters.date = dayKey;
                            applyFilters();
                        }
                    }
                }
            }
        });

        // Gráfico de Tendência (Verso do Card)
        const ctxTrend = document.getElementById('trendChart').getContext('2d');
        state.charts.trend = new Chart(ctxTrend, {
            type: 'line',
            data: { labels: [], datasets: [] },
            options: {
                plugins: { legend: { display: false }, tooltip: { mode: 'index', intersect: false } },
                scales: {
                    y: { grid: { color: '#1e293b' }, ticks: { color: '#64748b' } },
                    x: { grid: { display: false }, ticks: { color: '#64748b' } }
                },
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false }
            }
        });

        // Listeners do Card Flip
        const btnFlipCard = document.getElementById('btnFlipCard');
        const btnFlipBack = document.getElementById('btnFlipBack');
        const evolucaoContainer = document.getElementById('evolucaoContainer');
        if (btnFlipCard && evolucaoContainer) {
            btnFlipCard.addEventListener('click', () => {
                evolucaoContainer.classList.add('flipped');
            });
            btnFlipBack.addEventListener('click', () => {
                evolucaoContainer.classList.remove('flipped');
            });
        }

        initSpark('sparkTotal', '#3b82f6');
        initSpark('sparkComercial', '#10b981');
        initSpark('sparkNaoComercial', '#ef4444');
        initSpark('sparkDivergencia', '#f97316');
        initSpark('sparkReincidente', '#ec4899');
    }

    function initSpark(id, color) {
        const ctx = document.getElementById(id).getContext('2d');
        state.charts.sparks[id] = new Chart(ctx, {
            type: 'line',
            data: { labels: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], datasets: [{ data: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0], borderColor: color, borderWidth: 2, pointRadius: 0, tension: 0.4, fill: false }] },
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
        const days = {};
        metricData.forEach(nc => {
            const day = formatExcelDate(nc.data);
            if (!day) return;
            if (!days[day]) days[day] = { CONFORME: 0, PENDENTE: 0, DIVERGENCIA: 0, REINCIDENTE: 0, sellers: {} };

            const g = state.dataGeral.find(item => item.pedido.toString() == nc.pedido.toString());
            const s = computeStatus(nc, g);
            days[day][s]++;

            const seller = nc.vendedor || 'Desc.';
            days[day].sellers[seller] = (days[day].sellers[seller] || 0) + 1;
        });

        const sortedDays = Object.keys(days).sort();
        state.lineSellersData = {};
        sortedDays.forEach(d => { state.lineSellersData[d] = days[d].sellers; });

        state.charts.line.data.labels = sortedDays.map(d => d.split('-').slice(1).reverse().join('/'));

        state.charts.line.data.datasets = [
            { label: 'Conforme', data: sortedDays.map(d => days[d].CONFORME), borderColor: '#10b981', backgroundColor: 'transparent', tension: 0.4 },
            { label: 'Pendente', data: sortedDays.map(d => days[d].PENDENTE), borderColor: '#ef4444', backgroundColor: 'transparent', tension: 0.4 },
            { label: 'Divergência', data: sortedDays.map(d => days[d].DIVERGENCIA), borderColor: '#f97316', backgroundColor: 'transparent', tension: 0.4 },
            { label: 'Reincidente', data: sortedDays.map(d => days[d].REINCIDENTE), borderColor: '#ec4899', backgroundColor: 'transparent', tension: 0.4 }
        ];

        state.charts.line.update();

        // Calculate Trend Data (Linear Regression of PENDENTE + DIVERGENCIA per day)
        const esquecimentosPerDay = sortedDays.map(d => days[d].PENDENTE + days[d].DIVERGENCIA);

        function calculateTrendLine(dataPoints) {
            let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
            const n = dataPoints.length;
            if (n === 0) return [];
            for (let i = 0; i < n; i++) {
                sumX += i;
                sumY += dataPoints[i];
                sumXY += i * dataPoints[i];
                sumX2 += i * i;
            }
            const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
            const intercept = (sumY - slope * sumX) / n;
            return dataPoints.map((_, i) => slope * i + intercept);
        }

        const trendData = calculateTrendLine(esquecimentosPerDay);

        let avg = 0;
        if (esquecimentosPerDay.length > 0) {
            avg = esquecimentosPerDay.reduce((a, b) => a + b, 0) / esquecimentosPerDay.length;
        }
        document.getElementById('trendAvg').textContent = avg.toFixed(1);

        const trendStatusEl = document.getElementById('trendStatus');
        if (trendData.length > 1) {
            const first = trendData[0];
            const last = trendData[trendData.length - 1];
            if (last > first + 0.1) {
                trendStatusEl.textContent = 'Aumentando ⚠️';
                trendStatusEl.style.color = '#ef4444';
            } else if (last < first - 0.1) {
                trendStatusEl.textContent = 'Diminuindo 📉';
                trendStatusEl.style.color = '#10b981';
            } else {
                trendStatusEl.textContent = 'Estável ➡️';
                trendStatusEl.style.color = '#94a3b8';
            }
        } else {
            trendStatusEl.textContent = 'N/A';
            trendStatusEl.style.color = '#64748b';
        }

        if (state.charts.trend) {
            state.charts.trend.data.labels = state.charts.line.data.labels;
            state.charts.trend.data.datasets = [
                {
                    label: 'Tendência',
                    data: trendData,
                    borderColor: '#0ea5e9',
                    borderWidth: 3,
                    borderDash: [5, 5],
                    backgroundColor: 'transparent',
                    pointRadius: 0,
                    tension: 0.1
                },
                {
                    label: 'Esquecimentos Reais',
                    data: esquecimentosPerDay,
                    borderColor: 'rgba(239, 68, 68, 0.4)',
                    borderWidth: 1,
                    backgroundColor: 'rgba(239, 68, 68, 0.1)',
                    fill: true,
                    tension: 0.4,
                    pointRadius: 2
                }
            ];
            state.charts.trend.update();
        }

        Object.keys(state.charts.sparks).forEach(id => {
            const chart = state.charts.sparks[id];
            chart.data.datasets[0].data = Array.from({ length: 10 }, () => Math.floor(Math.random() * 50));
            chart.update();
        });
    }
});

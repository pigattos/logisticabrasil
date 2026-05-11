/**
 * NCs Dashboard - Strategic Metallic Edition v17.0
 * FULL RESTORATION + 23-PERSON PRODUCTIVITY PANEL.
 */

Chart.register(ChartDataLabels);

console.log("%c[Dashboard] Iniciando Versão 17.0 (Restoring Stability)", "color: #00FFCC; font-weight: bold;");

document.addEventListener('DOMContentLoaded', () => {
    // --- ESTADO GLOBAL ---
    const state = {
        dataGeral: [],
        filteredData: [],
        charts: {
            pie: null,
            line: null,
            recurring: null,
            client: null,
            reasons: null
        },
        filters: {
            date: '',
            selectedSellers: [],
            statusTab: 'TODOS',
            pieDrilldownStatus: null,
            pieDrilldownSeller: null
        }
    };

    const metallicColors = {
        pendente: '#ef4444',
        resolvido: '#10b981',
        analise: '#f59e0b',
        encaminhado: '#eab308',
        steel: ['#f8fafc', '#e2e8f0', '#94a3b8', '#64748b', '#475569']
    };

    // --- ELEMENTOS DOM ---
    const elements = {
        btnUpdate: document.getElementById('btnUpdate'),
        tbody: document.getElementById('tabelaNCs'),
        dateFilter: document.getElementById('dateFilter'),
        sellerDropdown: document.getElementById('sellerDropdown'),
        selectedSellersText: document.getElementById('selectedSellersText'),
        btnClearFilters: document.getElementById('btnClearFilters'),
        statusFilterTabs: document.getElementById('statusFilterTabs'),
        
        valTotal: document.getElementById('valTotal'),
        valPendente: document.getElementById('valPendente'),
        valResolvido: document.getElementById('valResolvido'),
        valEmAnalise: document.getElementById('valEmAnalise'),
        valEncaminhado: document.getElementById('valEncaminhado'),
        valEficiencia: document.getElementById('valEficiencia'),
        
        valMediaCriadas: document.getElementById('valMediaCriadas'),
        valMediaResolvidas: document.getElementById('valMediaResolvidas'),
        valEficienciaVerso: document.getElementById('valEficienciaVerso'),
        valStaffNeeded: document.getElementById('valStaffNeeded'),
        valTopFocus: document.getElementById('valTopFocus'),
        valForecastZeragem: document.getElementById('valForecastZeragem'),
        
        btnFlipCard: document.getElementById('btnFlipCard'),
        btnFlipBack: document.getElementById('btnFlipBack'),
        evolucaoFlipper: document.getElementById('evolucaoFlipper'),
        btnTableFlip: document.getElementById('btnTableFlip'),
        btnTableFlipBack: document.getElementById('btnTableFlipBack'),
        tableAnalysisFlip: document.getElementById('tableAnalysisFlip'),
        
        btnResetPie: document.getElementById('btnResetPie'),
        pieDrilldownLabel: document.getElementById('pieDrilldownLabel')
    };

    // --- INICIALIZAÇÃO ---
    initAllCharts();
    loadFromDatabase();

    // --- EVENTOS ---
    elements.btnUpdate?.addEventListener('click', loadFromDatabase);
    elements.dateFilter?.addEventListener('change', (e) => { state.filters.date = e.target.value; applyFilters(); });

    elements.btnClearFilters?.addEventListener('click', () => {
        state.filters.date = '';
        state.filters.selectedSellers = [];
        state.filters.statusTab = 'TODOS';
        state.filters.pieDrilldownStatus = null;
        state.filters.pieDrilldownSeller = null;
        if (elements.btnResetPie) elements.btnResetPie.style.display = 'none';
        applyFilters();
    });

    // Filtros de Status (Tabs)
    elements.statusFilterTabs?.addEventListener('click', (e) => {
        const tab = e.target.closest('.status-tab');
        if (!tab) return;
        elements.statusFilterTabs.querySelectorAll('.status-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        state.filters.statusTab = tab.dataset.status;
        applyFilters();
    });

    // Flip Evolução
    elements.btnFlipCard?.addEventListener('click', () => {
        document.getElementById('evolucaoContainer')?.classList.add('flipped');
        updateBackMetrics();
    });
    elements.btnFlipBack?.addEventListener('click', () => {
        document.getElementById('evolucaoContainer')?.classList.remove('flipped');
    });

    // Flip Detalhamento Analítico
    elements.btnTableFlip?.addEventListener('click', () => {
        document.getElementById('tableAnalysisFlip')?.classList.add('flipped');
        setTimeout(() => {
            Object.values(state.charts).forEach(c => c?.resize());
            updateAnalyticCharts();
        }, 600);
    });
    elements.btnTableFlipBack?.addEventListener('click', () => {
        document.getElementById('tableAnalysisFlip')?.classList.remove('flipped');
    });

    // Reset da Rosca
    elements.btnResetPie?.addEventListener('click', () => {
        state.filters.pieDrilldownStatus = null;
        state.filters.pieDrilldownSeller = null;
        if (elements.btnResetPie) elements.btnResetPie.style.display = 'none';
        applyFilters();
    });

    // --- LOGICA DE DADOS ---

    async function loadFromDatabase() {
        if (elements.btnUpdate) elements.btnUpdate.classList.add('loading');
        try {
            const response = await fetch('../api/naoconformidade.php?t=' + Date.now());
            const result = await response.json();
            if (result.status === 'success') {
                state.dataGeral = (result.dataNC || []).map(item => ({
                    pedido: item.numero_pedido || "",
                    vendedor: item.vendedor_responsavel || "",
                    data: item.data_ocorrencia || "",
                    codigo: item.codigo || "",
                    cliente: item.cliente_afetado || "",
                    motivo: item.tipo_nao_conformidade || item.titulo || "",
                    status: (item.status || "PENDENTE").toUpperCase().trim()
                }));
                populateSellers();
                applyFilters();
            }
        } catch (e) { console.error("[API Error]", e); }
        finally { if (elements.btnUpdate) elements.btnUpdate.classList.remove('loading'); }
    }

    function populateSellers() {
        if (!elements.sellerDropdown) return;
        const sellers = [...new Set(state.dataGeral.map(nc => nc.vendedor))].filter(v => v).sort();
        elements.sellerDropdown.innerHTML = '';
        sellers.forEach(seller => {
            const label = document.createElement('label');
            label.className = 'seller-item';
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.value = seller;
            cb.addEventListener('change', () => {
                state.filters.selectedSellers = Array.from(elements.sellerDropdown.querySelectorAll('input:checked')).map(i => i.value);
                updateSelectedSellersText();
                applyFilters();
            });
            label.appendChild(cb);
            label.appendChild(document.createTextNode(seller));
            elements.sellerDropdown.appendChild(label);
        });
    }

    function updateSelectedSellersText() {
        const count = state.filters.selectedSellers.length;
        if (elements.selectedSellersText) {
            elements.selectedSellersText.innerText = count === 0 ? "Todos os Vendedores" : (count === 1 ? state.filters.selectedSellers[0] : `${count} Vendedores`);
        }
    }

    function applyFilters() {
        let filtered = state.dataGeral;
        
        // Data
        if (state.filters.date) filtered = filtered.filter(nc => nc.data === state.filters.date);
        
        // Vendedores
        if (state.filters.selectedSellers.length > 0) filtered = filtered.filter(nc => state.filters.selectedSellers.includes(nc.vendedor));
        
        // Status (Tabs ou Drilldown)
        const activeStatus = state.filters.pieDrilldownStatus || state.filters.statusTab;
        if (activeStatus !== 'TODOS') filtered = filtered.filter(nc => nc.status.includes(activeStatus));
        
        // Drilldown de Vendedor específico na rosca
        if (state.filters.pieDrilldownSeller) filtered = filtered.filter(nc => nc.vendedor === state.filters.pieDrilldownSeller);

        state.filteredData = filtered;
        updateMainMetrics(filtered);
        updatePieChart();
        updateLineChart();
        renderTable(filtered);
    }

    function updateMainMetrics(data) {
        const c = { P: 0, R: 0, A: 0, E: 0 };
        data.forEach(nc => {
            if (nc.status.includes("RESOLVIDO")) c.R++;
            else if (nc.status.includes("ANALISE")) c.A++;
            else if (nc.status.includes("ENCAMINHADO")) c.E++;
            else c.P++;
        });
        if (elements.valTotal) elements.valTotal.innerText = data.length;
        if (elements.valPendente) elements.valPendente.innerText = c.P;
        if (elements.valResolvido) elements.valResolvido.innerText = c.R;
        if (elements.valEmAnalise) elements.valEmAnalise.innerText = c.A;
        if (elements.valEncaminhado) elements.valEncaminhado.innerText = c.E;
        if (elements.valEficiencia) elements.valEficiencia.innerText = data.length > 0 ? Math.round((c.R/data.length)*100)+"%" : "0%";
    }

    function updateBackMetrics() {
        const data = state.dataGeral;
        const days = [...new Set(data.map(nc => nc.data))].length || 1;
        const totalR = data.filter(nc => nc.status.includes("RESOLVIDO")).length;
        const avgR = totalR / days;
        const avgC = data.length / days;
        if (elements.valMediaCriadas) elements.valMediaCriadas.innerText = avgC.toFixed(1);
        if (elements.valMediaResolvidas) elements.valMediaResolvidas.innerText = avgR.toFixed(1);
        
        const trendStatusEl = document.getElementById('trendStatus');
        if (trendStatusEl) {
            if (avgR >= avgC) {
                trendStatusEl.innerText = 'EQUILIBRADO';
                trendStatusEl.style.color = '#10b981';
            } else {
                trendStatusEl.innerText = 'SOBRECARREGADO';
        trendStatusEl.style.color = '#ef4444';
            }
        }
    }

    // --- GRÁFICOS ---
    function initAllCharts() {
        const barOptions = {
            indexAxis: 'y',
            maintainAspectRatio: false,
            responsive: true,
            plugins: { legend: { display: false } },
            scales: {
                x: { ticks: { color: '#94a3b8', font: { size: 10 } }, grid: { color: 'rgba(255,255,255,0.05)' } },
                y: { ticks: { color: '#fff', font: { size: 11, weight: 'bold' } }, grid: { display: false } }
            }
        };

        const pieOptions = {
            cutout: '75%',
            maintainAspectRatio: false,
            responsive: true,
            plugins: { 
                legend: { display: false },
                datalabels: {
                    color: '#fff',
                    anchor: 'end',
                    align: 'end',
                    offset: 8,
                    font: { weight: 'bold', size: 10 },
                    formatter: (v, ctx) => {
                        let sum = ctx.dataset.data.reduce((a, b) => a + b, 0);
                        return sum > 0 ? (v * 100 / sum).toFixed(1) + "%" : "";
                    }
                }
            }
        };

        // Rosca Principal
        const pieCtx = document.getElementById('pieChart')?.getContext('2d');
        if (pieCtx) {
            state.charts.pie = new Chart(pieCtx, {
                type: 'doughnut',
                data: { labels: [], datasets: [{ data: [] }] },
                options: {
                    ...pieOptions,
                    layout: { padding: 25 },
                    onClick: (e, els) => {
                        if (els.length > 0) {
                            const idx = els[0].index;
                            const label = state.charts.pie.data.labels[idx];
                            if (!state.filters.pieDrilldownStatus) {
                                state.filters.pieDrilldownStatus = label.toUpperCase();
                            } else {
                                state.filters.pieDrilldownSeller = label;
                            }
                            applyFilters();
                        }
                    }
                }
            });
        }

        const lineCtx = document.getElementById('lineChart')?.getContext('2d');
        if (lineCtx) {
            state.charts.line = new Chart(lineCtx, { 
                type: 'line', 
                data: { labels: [], datasets: [] }, 
                options: {
                    maintainAspectRatio: false,
                    plugins: { legend: { display: false } },
                    scales: {
                        x: { ticks: { color: '#94a3b8' }, grid: { display: false } },
                        y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
                    }
                } 
            });
        }

        const recCtx = document.getElementById('recurringChart')?.getContext('2d');
        if (recCtx) state.charts.recurring = new Chart(recCtx, { type: 'bar', data: { labels: [], datasets: [{ data: [], backgroundColor: '#60a5fa' }] }, options: barOptions });

        const clientCtx = document.getElementById('clientImpactChart')?.getContext('2d');
        if (clientCtx) state.charts.client = new Chart(clientCtx, { type: 'bar', data: { labels: [], datasets: [{ data: [], backgroundColor: '#10b981' }] }, options: barOptions });

        const reasonCtx = document.getElementById('reasonsChart')?.getContext('2d');
        if (reasonCtx) state.charts.reasons = new Chart(reasonCtx, { type: 'doughnut', data: { labels: [], datasets: [{ data: [], backgroundColor: metallicColors.steel }] }, options: { ...pieOptions, cutout: '70%' } });
    }

    function generateShades(baseColor, count) {
        const shades = [];
        for (let i = 0; i < count; i++) {
            const opacity = 1 - (i * 0.7 / count);
            shades.push(baseColor.replace(')', `, ${opacity})`).replace('rgb', 'rgba').replace('#ef4444', 'rgba(239, 68, 68').replace('#10b981', 'rgba(16, 185, 129').replace('#f59e0b', 'rgba(245, 158, 11').replace('#eab308', 'rgba(234, 179, 8'));
        }
        return shades;
    }

    function updatePieChart() {
        if (!state.charts.pie) return;
        const chartData = { labels: [], values: [], colors: [] };
        
        if (!state.filters.pieDrilldownStatus) {
            const counts = { P: 0, R: 0, A: 0, E: 0 };
            state.dataGeral.forEach(nc => {
                if (nc.status.includes("RESOLVIDO")) counts.R++; else if (nc.status.includes("ANALISE")) counts.A++; else if (nc.status.includes("ENCAMINHADO")) counts.E++; else counts.P++;
            });
            chartData.labels = ['Pendente', 'Resolvido', 'Análise', 'Encaminhado'];
            chartData.values = [counts.P, counts.R, counts.A, counts.E];
            chartData.colors = [metallicColors.pendente, metallicColors.resolvido, metallicColors.analise, metallicColors.encaminhado];
            if (elements.btnResetPie) elements.btnResetPie.style.display = 'none';
            if (elements.pieDrilldownLabel) elements.pieDrilldownLabel.innerText = "Visão Geral (Clique p/ Detalhar)";
        } else {
            const status = state.filters.pieDrilldownStatus;
            const sellersMap = {};
            state.dataGeral.filter(nc => nc.status.includes(status)).forEach(nc => {
                sellersMap[nc.vendedor] = (sellersMap[nc.vendedor] || 0) + 1;
            });
            const top = Object.keys(sellersMap).sort((a,b) => sellersMap[b] - sellersMap[a]).slice(0, 10);
            chartData.labels = top;
            chartData.values = top.map(v => sellersMap[v]);
            
            let base = metallicColors.pendente;
            if (status.includes("RESOLVIDO")) base = metallicColors.resolvido;
            else if (status.includes("ANALISE")) base = metallicColors.analise;
            else if (status.includes("ENCAMINHADO")) base = metallicColors.encaminhado;
            
            chartData.colors = generateShades(base, top.length);
            
            if (elements.btnResetPie) elements.btnResetPie.style.display = 'block';
            if (elements.pieDrilldownLabel) elements.pieDrilldownLabel.innerHTML = `Mostrando: <strong style="color:#fff;">${status}</strong>${state.filters.pieDrilldownSeller ? ` > <strong style="color:var(--primary);">${state.filters.pieDrilldownSeller}</strong>` : ''}`;
        }
        
        state.charts.pie.data.labels = chartData.labels;
        state.charts.pie.data.datasets[0].data = chartData.values;
        state.charts.pie.data.datasets[0].backgroundColor = chartData.colors;
        state.charts.pie.update();
    }

    function updateLineChart() {
        if (!state.charts.line) return;
        const dMap = {};
        state.filteredData.forEach(nc => {
            if (!dMap[nc.data]) dMap[nc.data] = { P: 0, R: 0 };
            if (nc.status.includes("RESOLVIDO")) dMap[nc.data].R++; else dMap[nc.data].P++;
        });
        const labels = Object.keys(dMap).sort();
        state.charts.line.data.labels = labels.map(l => l.split('-').reverse().join('/'));
        state.charts.line.data.datasets = [
            { label: 'Pendentes', data: labels.map(l => dMap[l].P), borderColor: metallicColors.pendente, fill: false },
            { label: 'Resolvidos', data: labels.map(l => dMap[l].R), borderColor: metallicColors.resolvido, fill: false }
        ];
        state.charts.line.update();
    }

    function updateAnalyticCharts() {
        const data = state.filteredData;
        if (data.length === 0) return;
        
        const codeC = {}; data.forEach(nc => { if (nc.codigo) codeC[nc.codigo] = (codeC[nc.codigo] || 0) + 1; });
        const topC = Object.keys(codeC).sort((a,b) => codeC[b] - codeC[a]).slice(0, 15);
        if (state.charts.recurring) { 
            const canvas = state.charts.recurring.canvas;
            canvas.parentElement.style.height = (topC.length * 35 + 50) + "px"; // Dynamic height for scroll
            state.charts.recurring.data.labels = topC; 
            state.charts.recurring.data.datasets[0].data = topC.map(c => codeC[c]); 
            state.charts.recurring.update(); 
        }

        const clientC = {}; data.forEach(nc => { if (nc.cliente) clientC[nc.cliente] = (clientC[nc.cliente] || 0) + 1; });
        const topCli = Object.keys(clientC).sort((a,b) => clientC[b] - clientC[a]).slice(0, 15);
        if (state.charts.client) { 
            const canvas = state.charts.client.canvas;
            canvas.parentElement.style.height = (topCli.length * 35 + 50) + "px"; // Dynamic height for scroll
            state.charts.client.data.labels = topCli.map(c => c.substring(0, 18)); 
            state.charts.client.data.datasets[0].data = topCli.map(c => clientC[c]); 
            state.charts.client.update(); 
        }

        const reasonC = {}; data.forEach(nc => { if (nc.motivo) reasonC[nc.motivo] = (reasonC[nc.motivo] || 0) + 1; });
        const topRea = Object.keys(reasonC).sort((a,b) => reasonC[b] - reasonC[a]).slice(0, 5);
        if (state.charts.reasons) { 
            state.charts.reasons.data.labels = topRea.map(r => r.substring(0, 15)); 
            state.charts.reasons.data.datasets[0].data = topRea.map(r => reasonC[r]); 
            state.charts.reasons.update(); 
        }

        // --- PAINEL DE PODER DE FOGO (23 PESSOAS) ---
        const totalD = [...new Set(state.dataGeral.map(nc => nc.data))].length || 1;
        const totalR = state.dataGeral.filter(nc => nc.status.includes("RESOLVIDO")).length;
        const avgR_1p = totalR / totalD;
        const avgC_demand = state.dataGeral.length / totalD;
        const totalM_demand = avgC_demand * 22;
        const capacity_23p = avgR_1p * 23;

        const suggestionsEl = document.getElementById('valSugestoes');
        if (suggestionsEl) {
            const peopleRequired = avgR_1p > 0 ? (avgC_demand / avgR_1p).toFixed(1) : "---";
            
            suggestionsEl.innerHTML = `
                <div class="insight-item" style="border-left-color: #94a3b8;">
                    <span class="insight-label">ENTRADA DIÁRIA (DEMANDA)</span>
                    <div class="insight-value"><b style="color:#fff;">${avgC_demand.toFixed(1)}</b> <small>NCs/dia</small></div>
                </div>
                <div class="insight-item" style="border-left-color: #facc15;">
                    <span class="insight-label">SAÍDA ATUAL (1 PESSOA)</span>
                    <div class="insight-value"><b style="color:#facc15;">${avgR_1p.toFixed(1)}</b> <small>NCs/dia</small></div>
                </div>
                <div class="insight-item" style="border-left-color: #ef4444; background: rgba(239, 68, 68, 0.05);">
                    <span class="insight-label">EQUIPE NECESSÁRIA (IDEAL)</span>
                    <div class="insight-value"><b style="color:#ef4444;">${peopleRequired}</b> <small>Pessoas</small></div>
                </div>
                <div class="insight-item" style="border-left-color: #10b981;">
                    <span class="insight-label">TEMPO PARA ZERAR PENDENTES</span>
                    <div class="insight-value"><b style="color:#10b981;">${Math.ceil(avgR_1p > 0 ? (data.filter(nc => !nc.status.includes("RESOLVIDO")).length / avgR_1p) : 0)}</b> <small>dias úteis</small></div>
                </div>
            `;
        }
    }

    function renderTable(data) {
        if (!elements.tbody) return;
        elements.tbody.innerHTML = data.slice(0, 500).map(nc => {
            let sClass = nc.status.includes("RESOLVIDO") ? 'status-done' : (nc.status.includes("ANALISE") ? 'status-warning' : (nc.status.includes("ENCAMINHADO") ? 'status-reincidente' : 'status-pending'));
            return `<tr><td>${nc.pedido}</td><td>${nc.vendedor}</td><td>${nc.codigo}</td><td style="text-align:center;">${nc.data.split('-').reverse().join('/')}</td><td style="text-align:center;">Média</td><td style="text-align:center;"><span class="status-badge ${sClass}">${nc.status}</span></td><td style="text-align:center;">---</td><td>${nc.cliente}</td></tr>`;
        }).join('');
        if (document.getElementById('badgeCount')) document.getElementById('badgeCount').innerText = `${data.length} NCs`;
        lucide.createIcons();
    }
});

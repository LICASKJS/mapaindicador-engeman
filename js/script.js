const FILES = {
  homolog: 'dados/fornecedores_homologados.xlsx',
  iqf: 'dados/atendimento controle_qualidade.xlsx'
};

const SCORE_THRESHOLD = 70;
const MAX_CHAT_HISTORY = 6;
const dateFilter = { day: null, month: null, year: null };

const RECURRENCE_IGNORE_SUBJECTS = new Set([
  'bom fornecedor',
  'bom atendimento',
  'atendimento bom',
  'atende a expectativa',
  'atende as expectativas',
  'atende expectativa',
  'atende expecativa',
  'sem comentarios',
  'sem comentario'
]);
const RECURRENCE_STOP_WORDS = new Set(['', 'de', 'da', 'do', 'das', 'dos', 'para', 'por', 'com', 'sem', 'uma', 'um', 'no', 'na', 'nos', 'nas', 'ao', 'a', 'e', 'em', 'sobre']);

const state = {
  homolog: [],
  iqf: [],
  combined: [],
  occ: [],
  recurrence: [],
  recurrenceFilter: '',
  missing: [],
  iqfFiltered: null,
  occFiltered: null,
  monthlyIqf: new Map()
};

const charts = {
  status: null,
  gauge: null,
  trend: null,
  iqf: null,
  occ: null
};

const chatState = {
  history: [],
  loading: false,
  typingNode: null
};

document.addEventListener('DOMContentLoaded', init);

async function init() {
  showLoading(true);
  try {
    if (typeof XLSX === 'undefined') {
      throw new Error('Biblioteca XLSX nao carregada');
    }
    await loadData();
    bindControls();
    renderAll();
  } catch (err) {
    console.error('[init]', err);
    alert('Falha ao carregar as planilhas. Verifique os arquivos e recarregue a pagina.');
  } finally {
    showLoading(false);
  }
}

async function loadData() {
  const [homRows, iqfRows] = await Promise.all([
    loadWorkbook(FILES.homolog),
    loadWorkbook(FILES.iqf)
  ]);
  state.homolog = homRows.map(mapHomolog).filter(Boolean);
  state.iqf = iqfRows
    .map(mapIqf)
    .filter((row) => row.code || row.name || row.iqf !== null || row.occ);
  buildState();
}

async function loadWorkbook(path) {
  const response = await fetch(path);
  if (!response.ok) {
    throw new Error('Falha ao carregar ' + path + ': ' + response.status + ' ' + response.statusText);
  }
  const buffer = await response.arrayBuffer();
  const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true });
  const rows = [];
  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    if (worksheet) {
      rows.push(...XLSX.utils.sheet_to_json(worksheet, { defval: null }));
    }
  });
  return rows;
}

function mapHomolog(row) {
  const normalized = normalizeKeys(row);
  const code = toId(normalized.codigo || normalized.codagente || normalized.codfornecedor || normalized.id);
  const name = safeString(normalized.nomefantasia || normalized.agente || normalized.fornecedor || normalized.nome);
  if (!code && !name) {
    return null;
  }
  const score = toNumber(normalized.notahomologacao || normalized.nota);
  const status = mapStatus(normalized.aprovado || normalized.status || normalized.situacao || normalized.qualifica);
  const expire = toISODate(normalized.datavencimento || normalized.validade || normalized.data);
  return { code, name, score, status, expire };
}

function mapIqf(row) {
  const normalized = normalizeKeys(row);
  return {
    code: toId(normalized.codagente || normalized.codigo || normalized.codfornecedor || normalized.id),
    name: safeString(normalized.nomeagente || normalized.fornecedor || normalized.nome),
    iqf: toNumber(normalized.nota || normalized.notaiqf || normalized.iqf),
    date: toISODate(normalized.data || normalized.datamedicao || normalized.datageracao),
    document: safeString(normalized.documento || normalized.numerodocumento || normalized.doc),
    occ: safeString(normalized.observacao || normalized.ocorrencia || normalized.comentario || normalized.notaocorrencia),
    origin: safeString(normalized.origem || normalized.tipo || normalized.classificacao)
  };
}

function buildState() {
  const index = new Map();
  const missing = [];

  const register = (key, sample) => {
    if (!key) {
      return;
    }
    if (!index.has(key)) {
      index.set(key, {
        sum: 0,
        count: 0,
        last: null,
        doc: sample.document || '',
        name: sample.name || ''
      });
    }
    const bucket = index.get(key);
    if (sample.iqf !== null) {
      bucket.sum += sample.iqf;
      bucket.count += 1;
    }
    if (sample.date && (!bucket.last || sample.date > bucket.last)) {
      bucket.last = sample.date;
    }
    if (sample.document && !bucket.doc) {
      bucket.doc = sample.document;
    }
    if (sample.name && !bucket.name) {
      bucket.name = sample.name;
    }
  };

  state.iqf.forEach((row) => {
    if (row.code) {
      register('code:' + row.code, row);
    }
    if (row.name) {
      register('name:' + normalizeText(row.name), row);
    }
  });

  state.combined = state.homolog
    .filter((h) => h.score !== null)
    .map((h) => {
      const byCode = h.code ? index.get('code:' + h.code) : null;
      const byName = h.name ? index.get('name:' + normalizeText(h.name)) : null;
      const bucket = byCode || byName;
      const iqfSamples = bucket?.count || 0;
      const iqfAverage = iqfSamples ? roundValue(bucket.sum / bucket.count) : null;
      const homologScore = roundValue(h.score);
      if (!iqfSamples) {
        missing.push({ code: h.code, name: h.name, status: h.status, score: homologScore });
      }
      const status = deriveStatus(h.status, iqfAverage, homologScore);
      return {
        code: h.code,
        name: (bucket && bucket.name) || h.name,
        status,
        statusBase: h.status,
        homolog: homologScore,
        iqf: iqfAverage,
        expire: h.expire,
        iqfSamples,
        lastIqf: (bucket && bucket.last) || null,
        document: (bucket && bucket.doc) || ''
      };
    })
    .sort((a, b) => a.name.localeCompare(b.name, 'pt-BR', { sensitivity: 'base' }));

  state.missing = missing;
  state.iqfFiltered = state.iqf.slice();
  state.occ = state.iqf
    .filter((row) => row.occ)
    .map((row) => ({
      name: row.name,
      document: row.document,
      date: row.date,
      occ: row.occ,
      severity: severityLevel(row.occ)
    }));
  state.occFiltered = state.occ.slice();
  state.monthlyIqf = aggregateMonthly(state.iqf);
}

function aggregateMonthly(rows) {
  const monthly = new Map();
  rows.forEach((row) => {
    if (!row.date || row.iqf === null) {
      return;
    }
    const key = row.date.slice(0, 7);
    if (!monthly.has(key)) {
      monthly.set(key, { sum: 0, count: 0 });
    }
    const bucket = monthly.get(key);
    bucket.sum += row.iqf;
    bucket.count += 1;
  });
  return monthly;
}

function matchesDate(iso) {
  if (!dateFilter.day && !dateFilter.month && !dateFilter.year) {
    return true;
  }
  if (!iso) {
    return false;
  }
  const [year, month, day] = iso.split('-');
  if (dateFilter.year && year !== dateFilter.year) {
    return false;
  }
  if (dateFilter.month) {
    const [filterYear, filterMonth] = dateFilter.month.split('-');
    if (filterYear !== year || filterMonth !== month) {
      return false;
    }
  }
  if (dateFilter.day && dateFilter.day !== iso) {
    return false;
  }
  return true;
}

function getFilteredCombined() {
  return state.combined.filter((row) => matchesDate(row.lastIqf));
}

function getFilteredIqfRows() {
  return state.iqf.filter((row) => matchesDate(row.date));
}

function getFilteredOccurrences() {
  return state.occ.filter((row) => matchesDate(row.date));
}

function getGaugePalette(value) {
  if (value >= 85) {
  return ['#104585f8'];
  }
  if (value >= SCORE_THRESHOLD) {
    return ['#b17919ff', '#b45309'];
  }
  return ['#dc2626', '#7f1d1d'];
}

function lightenColor(hex, alpha = 0.8) {
  const sanitized = (hex || '').replace('#', '');
  if (sanitized.length !== 7) {
    return hex;
  }
  const r = parseInt(sanitized.slice(0, 2), 16);
  const g = parseInt(sanitized.slice(2, 4), 16);
  const b = parseInt(sanitized.slice(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}


function bindControls() {
  const dayInput = document.getElementById('dateFilterDay');
  const monthInput = document.getElementById('dateFilterMonth');
  const yearInput = document.getElementById('dateFilterYear');
  const clearBtn = document.getElementById('dateFilterClear');
  const triggerDateRender = () => {
    renderAll();
  };

  if (dayInput) {
    dayInput.value = dateFilter.day || '';
    dayInput.addEventListener('change', (event) => {
      dateFilter.day = event.target.value || null;
      triggerDateRender();
    });
  }
  if (monthInput) {
    monthInput.value = dateFilter.month || '';
    monthInput.addEventListener('change', (event) => {
      dateFilter.month = event.target.value || null;
      triggerDateRender();
    });
  }
  if (yearInput) {
    yearInput.value = dateFilter.year || '';
    yearInput.addEventListener('input', (event) => {
      const value = event.target.value.trim();
      dateFilter.year = value ? value.padStart(4, '0') : null;
    });
    yearInput.addEventListener('change', (event) => {
      const value = event.target.value.trim();
      dateFilter.year = value ? value.padStart(4, '0') : null;
      triggerDateRender();
    });
  }
  if (clearBtn) {
    clearBtn.addEventListener('click', () => {
      dateFilter.day = null;
      dateFilter.month = null;
      dateFilter.year = null;
      if (dayInput) dayInput.value = '';
      if (monthInput) monthInput.value = '';
      if (yearInput) yearInput.value = '';
      triggerDateRender();
    });
  }

  const searchInput = document.getElementById('supplierSearch');
  const statusSelect = document.getElementById('statusFilter');
  if (searchInput) {
    searchInput.addEventListener('input', renderSuppliersTable);
  }
  if (statusSelect) {
    statusSelect.addEventListener('change', renderSuppliersTable);
  }

  const iqfSearch = document.getElementById('iqfSearch');
  const supplierSelect = document.getElementById('supplierFilter');
  if (iqfSearch) {
    iqfSearch.addEventListener('input', () => {
      renderIqfTable();
      renderIqfSummary();
      renderIqfChart();
    });
  }
  if (supplierSelect) {
    supplierSelect.addEventListener('change', () => {
      renderIqfTable();
      renderIqfSummary();
      renderIqfChart();
    });
  }

  const occSearch = document.getElementById('occurrenceSearch');
  if (occSearch) {
    occSearch.addEventListener('input', renderOccTable);
  }

  const chatPanel = document.getElementById('aiChatPanel');
  if (chatPanel) {
    const sendBtn = document.getElementById('aiSendBtn');
    const input = document.getElementById('aiUserInput');
    const keyInput = document.getElementById('openaiKey');
    chatPanel.querySelectorAll('[data-question]').forEach((button) => {
      button.addEventListener('click', () => {
        if (input) {
          input.value = button.getAttribute('data-question') || '';
          input.focus();
        }
      });
    });
    if (sendBtn) {
      sendBtn.addEventListener('click', sendAIMessage);
    }
    if (input) {
      input.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' && !event.shiftKey) {
          event.preventDefault();
          sendAIMessage();
        }
      });
    }
    if (keyInput) {
      const stored = localStorage.getItem('openaiKey');
      if (stored) {
        keyInput.value = stored;
      }
      keyInput.addEventListener('blur', () => {
        localStorage.setItem('openaiKey', keyInput.value.trim());
      });
    }
  }

  const recurrenceFilter = document.getElementById('recurrenceSupplierFilter');
  if (recurrenceFilter) {
    recurrenceFilter.addEventListener('change', handleRecurrenceFilterChange);
  }
}

function renderAll() {
  if (document.getElementById('suppliersTable')) {
    renderDashboard();
  }
  if (document.getElementById('iqfTable')) {
    renderIqfPage();
  }
  if (document.getElementById('occurrencesTable')) {
    renderOccurrencesPage();
  }
}

function renderDashboard() {
  const combined = getFilteredCombined();
  const approved = combined.filter((row) => row.status === 'Homologado');
  const avgIqfGlobal = avg(combined.map((row) => row.iqf));
  const avgHomologGlobal = avg(combined.map((row) => row.homolog));

  setText('totalSuppliers', String(approved.length));
  setText('avgIqf', avgIqfGlobal !== null ? formatPercent(avgIqfGlobal) : '--');
  setText('pendingOccurrences', avgHomologGlobal !== null ? formatPercent(avgHomologGlobal) : '--');

  renderSuppliersTable();
  renderStatusChart();
  renderIqfGauge();
  renderTrendChart();
}

function renderSuppliersTable() {
  const body = document.querySelector('#suppliersTable tbody');
  if (!body) {
    return;
  }
  const search = normalizeText(document.getElementById('supplierSearch')?.value || '');
  const status = document.getElementById('statusFilter')?.value || '';
  const rows = getFilteredCombined().filter((row) => {
    if (search && !normalizeText((row.name || '') + ' ' + (row.code || '')).includes(search)) {
      return false;
    }
    if (status === 'approved' && row.status !== 'Homologado') {
      return false;
    }
    if (status === 'rejected' && row.status !== 'Reprovado') {
      return false;
    }
    if (status === 'pending' && row.status !== 'Pendente') {
      return false;
    }
    return true;
  });

  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="5">Nenhum fornecedor encontrado.</td></tr>';
    updateMissingHint();
    return;
  }

  body.innerHTML = rows
    .map((row) => {
      const iqfDetails = row.iqfSamples
        ? '<small class="table-note">' + row.iqfSamples + ' amostras&nbsp;&middot;&nbsp;' + (formatDate(row.lastIqf) || 'Sem data') + '</small>'
        : '';
      const rowClass = row.status === 'Reprovado' ? ' class="row-rejected"' : '';
      return '<tr' + rowClass + '>' +
        '<td>' + escapeHtml(row.name || '---') + '</td>' +
        '<td><span class="status-badge ' + badge(row.status) + '">' + row.status + '</span></td>' +
        '<td>' + formatPercent(row.iqf) + iqfDetails + '</td>' +
        '<td>' + formatPercent(row.homolog) + '</td>' +
        '<td>' + (formatDate(row.expire) || 'Sem data') + '</td>' +
        '</tr>';
    })
    .join('');

  updateMissingHint();
}

function updateMissingHint() {
  const hint = document.getElementById('missingSuppliersHint');
  if (!hint) {
    return;
  }
  if (dateFilter.day || dateFilter.month || dateFilter.year) {
    hint.innerText = '';
    hint.style.display = 'none';
    return;
  }
  if (!state.missing.length) {
    hint.innerText = '';
    hint.style.display = 'none';
    return;
  }
  const preview = state.missing
    .slice(0, 3)
    .map((item) => escapeHtml(item.name || item.code || '---'))
    .join(', ');
  hint.innerHTML = 'Faltando IQF para <strong>' + state.missing.length + '</strong> fornecedores (ex.: ' + preview + (state.missing.length > 3 ? '...' : '') + ').';
  hint.style.display = 'block';
}

function renderStatusChart() {
  const canvas = document.getElementById('suppliersChart');
  if (!canvas || typeof Chart === 'undefined') {
    return;
  }
  const ctx = canvas.getContext('2d');
  const counts = { Homologado: 0, Reprovado: 0 };
  const rows = getFilteredCombined();
  rows.forEach((row) => {
    if (row.status === 'Homologado' || row.status === 'Reprovado') {
      counts[row.status] += 1;
    }
  });
  const total = counts.Homologado + counts.Reprovado;
  if (charts.status) {
    charts.status.destroy();
  }
  const normalizeGradient = (from, to) => {
    const gradient = ctx.createLinearGradient(0, 0, canvas.width, canvas.height);
    gradient.addColorStop(0, from);
    gradient.addColorStop(1, to);
    return gradient;
  };
  const palette = [
    normalizeGradient('#0072ff', '#2c3e50'),
    normalizeGradient('#b34700', '#ff6600')
  ];
  const centerLabelPlugin = {
    id: 'statusCenterLabel',
    afterDraw(chart) {
      const meta = chart.getDatasetMeta(0);
      if (!meta || !meta.data.length) {
        return;
      }
      const { x, y } = meta.data[0];
      const { ctx: drawCtx } = chart;
      drawCtx.save();
      drawCtx.font = '700 20px "Inter", sans-serif';
      drawCtx.fillStyle = '#0f2027';
      drawCtx.textAlign = 'center';
      drawCtx.fillText(String(total), x, y - 4);
      drawCtx.font = '500 11px "Inter", sans-serif';
      drawCtx.fillStyle = '#566c8dff';
      drawCtx.fillText('Total', x, y + 12);
      drawCtx.restore();
    }
  };
  charts.status = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['Homologado', 'Reprovado'],
      datasets: [
        {
          data: [counts.Homologado, counts.Reprovado],
          backgroundColor: palette,
          hoverBackgroundColor: palette,
          borderWidth: 2,
          borderColor: '#0f172a0d',
          hoverOffset: 12,
          spacing: 4
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '68%',
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            usePointStyle: true,
            padding: 16
          }
        },
        tooltip: {
          callbacks: {
            label: (context) => {
              const value = context.parsed;
              const pct = total ? Math.round((value / total) * 100) : 0;
              return `${context.label}: ${value} (${pct}%)`;
            }
          }
        }
      }
    },
    plugins: [centerLabelPlugin]
  });
}

function renderIqfGauge() {
  const canvas = document.getElementById('iqfGaugeChart');
  const labelEl = document.getElementById('iqfGaugeMonth');
  const valueEl = document.getElementById('iqfGaugeValue');
  if (!canvas || typeof Chart === 'undefined') {
    return;
  }
  const ctx = canvas.getContext('2d');
  const canvasHost = canvas.parentElement;
  if (canvasHost) {
    const hostHeight = canvasHost.clientHeight || 220;
    canvas.width = canvasHost.clientWidth || 260;
    canvas.height = hostHeight;
  }
  const hasDateFilter = Boolean(dateFilter.day || dateFilter.month || dateFilter.year);
  const filteredRows = getFilteredIqfRows();
  let displayLabel = '--';
  let displayValue = null;
  if (hasDateFilter) {
    if (dateFilter.day) {
      displayLabel = formatDate(dateFilter.day) || '--';
    } else if (dateFilter.month) {
      displayLabel = formatMonth(dateFilter.month);
    } else if (dateFilter.year) {
      displayLabel = dateFilter.year;
    }
  }
  if (filteredRows.length) {
    if (dateFilter.day) {
      displayValue = avg(filteredRows.map((row) => row.iqf));
    } else {
      const monthly = aggregateMonthly(filteredRows);
      const months = Array.from(monthly.keys()).sort();
      if (months.length) {
        let targetKey = months[months.length - 1];
        if (dateFilter.month && monthly.has(dateFilter.month)) {
          targetKey = dateFilter.month;
        } else if (dateFilter.year) {
          const yearMatches = months.filter((key) => key.startsWith(dateFilter.year + '-'));
          if (yearMatches.length) {
            targetKey = yearMatches[yearMatches.length - 1];
          }
        }
        const bucket = monthly.get(targetKey);
        if (bucket && bucket.count) {
          displayValue = bucket.sum / bucket.count;
          displayLabel = formatMonth(targetKey);
        }
      }
    }
  } else if (!hasDateFilter) {
    const months = Array.from(state.monthlyIqf.keys()).sort();
    if (months.length) {
      const latestKey = months[months.length - 1];
      const bucket = state.monthlyIqf.get(latestKey);
      if (bucket && bucket.count) {
        displayValue = bucket.sum / bucket.count;
        displayLabel = formatMonth(latestKey);
      }
    }
  }
  if (labelEl) {
    labelEl.innerText = displayLabel;
  }
  if (valueEl) {
    valueEl.innerText = displayValue !== null ? formatPercent(displayValue) : '--';
  }
  if (displayValue === null || Number.isNaN(displayValue)) {
    if (charts.gauge) {
      charts.gauge.destroy();
      charts.gauge = null;
    }
    return;
  }
  const safeValue = Math.max(0, Math.min(100, displayValue));
  const [primaryColor, trackColor] = getGaugePalette(safeValue);
  const secondaryColor = lightenColor(primaryColor, 0.65);
  const gradient = ctx.createLinearGradient(0, canvas.height, canvas.width, 0);
  gradient.addColorStop(0, primaryColor);
  gradient.addColorStop(1, secondaryColor);
  if (charts.gauge) {
    charts.gauge.destroy();
  }
  charts.gauge = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['IQF', ''],
      datasets: [
        {
          data: [safeValue, Math.max(0, 100 - safeValue)],
          backgroundColor: [gradient, trackColor],
          hoverBackgroundColor: [gradient, trackColor],
          borderWidth: 0,
          hoverOffset: 6,
          circumference: 180,
          rotation: -90,
          cutout: '74%',
          spacing: 2
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: () => 'Media IQF: ' + formatPercent(safeValue)
          }
        }
      }
    }
  });
}

function renderTrendChart() {
  const canvas = document.getElementById('iqfTrendChart');
  if (!canvas || typeof Chart === 'undefined') {
    return;
  }
  const rows = getFilteredIqfRows();
  const hasFilter = Boolean(dateFilter.day || dateFilter.month || dateFilter.year);
  const sourceRows = rows.length ? rows : (hasFilter ? [] : state.iqf);
  const monthly = aggregateMonthly(sourceRows);
  const labels = Array.from(monthly.keys()).sort();
  if (!labels.length) {
    if (charts.trend) {
      charts.trend.destroy();
      charts.trend = null;
    }
    return;
  }
  const ctx = canvas.getContext('2d');
  if (charts.trend) {
    charts.trend.destroy();
  }
  const gradient = ctx.createLinearGradient(0, 0, 0, canvas.height || 300);
  gradient.addColorStop(0, 'rgba(56, 189, 248, 0.35)');
  gradient.addColorStop(1, 'rgba(56, 189, 248, 0.05)');
  const values = labels.map((key) => {
    const bucket = monthly.get(key);
    return bucket && bucket.count ? roundValue(bucket.sum / bucket.count) : 0;
  });
  charts.trend = new Chart(ctx, {
    type: 'line',
    data: {
      labels: labels.map(formatMonth),
      datasets: [
        {
          label: 'Media IQF',
          data: values,
          borderColor: '#2193b0',
          borderWidth: 3,
          backgroundColor: gradient,
          tension: 0.35,
          fill: true,
          pointRadius: 3,
          pointHoverRadius: 6,
          pointBackgroundColor: '#2193b0',
          pointBorderColor: '#2193b0'
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (context) => 'Media IQF: ' + formatPercent(context.parsed.y)
          }
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(148, 163, 184, 0.15)' },
          ticks: { color: 'var(--text-secondary)' }
        },
        y: {
          beginAtZero: true,
          suggestedMax: 100,
          grid: { color: 'rgba(148, 163, 184, 0.15)' },
          ticks: { color: 'var(--text-secondary)', callback: (value) => value + '%' }
        }
      }
    }
  });
}

function renderIqfPage() {
  populateSupplierOptions();
  renderIqfTable();
  renderIqfSummary();
  renderIqfChart();
}

function populateSupplierOptions() {
  const select = document.getElementById('supplierFilter');
  if (!select) {
    return;
  }
  const current = select.value;
  const names = Array.from(new Set(state.iqf.map((row) => row.name).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'pt-BR', { sensitivity: 'base' }));
  const options = ['<option value="">Todos os Fornecedores</option>']
    .concat(names.map((name) => '<option value="' + escapeHtml(name) + '">' + escapeHtml(name) + '</option>'));
  select.innerHTML = options.join('');
  if (current) {
    select.value = current;
  }
}

function renderIqfTable() {
  const body = document.querySelector('#iqfTable tbody');
  if (!body) {
    return;
  }
  const term = normalizeText(document.getElementById('iqfSearch')?.value || '');
  const selected = document.getElementById('supplierFilter')?.value || '';
  const rows = state.iqf.filter((row) => {
    if (!matchesDate(row.date)) {
      return false;
    }
    if (selected && row.name !== selected) {
      return false;
    }
    if (term && !normalizeText((row.name || '') + ' ' + (row.document || '')).includes(term)) {
      return false;
    }
    return true;
  });
  state.iqfFiltered = rows;
  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="3">Nenhum registro encontrado.</td></tr>';
    return;
  }
  body.innerHTML = rows
    .map((row) => {
      return '<tr>' +
        '<td>' + escapeHtml(row.name || '---') + '</td>' +
        '<td>' + formatPercent(row.iqf) + '</td>' +
        '<td>' + (formatDate(row.date) || 'Sem data') + '</td>' +
        '</tr>';
    })
    .join('');
}

function renderIqfSummary() {
  const rows = Array.isArray(state.iqfFiltered) ? state.iqfFiltered : state.iqf;
  const values = rows.map((row) => row.iqf).filter((value) => value !== null && !Number.isNaN(value));
  setText('generalIqfAvg', values.length ? formatPercent(avg(values)) : '--');
  setText('bestIqf', values.length ? formatPercent(Math.max(...values)) : '--');
  setText('worstIqf', values.length ? formatPercent(Math.min(...values)) : '--');
}

function renderIqfChart() {
  const canvas = document.getElementById('iqfMonthlyChart');
  if (!canvas || typeof Chart === 'undefined') {
    return;
  }
  const rows = Array.isArray(state.iqfFiltered) ? state.iqfFiltered : getFilteredIqfRows();
  const hasFilter = Boolean(dateFilter.day || dateFilter.month || dateFilter.year);
  const sourceRows = rows.length ? rows : (hasFilter ? [] : state.iqf);
  const monthly = aggregateMonthly(sourceRows);
  const labels = Array.from(monthly.keys()).sort();
  if (!labels.length) {
    if (charts.iqf) {
      charts.iqf.destroy();
      charts.iqf = null;
    }
    return;
  }
  const ctx = canvas.getContext('2d');
  if (charts.iqf) {
    charts.iqf.destroy();
  }
  const gradient = ctx.createLinearGradient(0, 0, 0, canvas.height || 300);
  gradient.addColorStop(0, 'rgba(14, 165, 233, 0.4)');
  gradient.addColorStop(1, 'rgba(14, 165, 233, 0.05)');
  const values = labels.map((key) => {
    const bucket = monthly.get(key);
    return bucket && bucket.count ? roundValue(bucket.sum / bucket.count) : 0;
  });
  charts.iqf = new Chart(ctx, {
    type: 'line',
    data: {
      labels: labels.map(formatMonth),
      datasets: [
        {
          label: 'Media IQF',
          data: values,
          borderColor: '#0ea5e9',
          borderWidth: 3,
          backgroundColor: gradient,
          tension: 0.35,
          fill: true,
          pointRadius: 3,
          pointHoverRadius: 6,
          pointBackgroundColor: '#0ea5e9',
          pointBorderColor: '#0ea5e9'
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (context) => 'Media IQF: ' + formatPercent(context.parsed.y)
          }
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(148, 163, 184, 0.15)' },
          ticks: { color: 'var(--text-secondary)' }
        },
        y: {
          beginAtZero: true,
          suggestedMax: 100,
          grid: { color: 'rgba(148, 163, 184, 0.15)' },
          ticks: { color: 'var(--text-secondary)', callback: (value) => value + '%' }
        }
      }
    }
  });
}

function renderOccurrencesPage() {
  renderOccTable();
}

function renderOccTable() {
  const body = document.querySelector('#occurrencesTable tbody');
  if (!body) {
    return;
  }
  const term = normalizeText(document.getElementById('occurrenceSearch')?.value || '');
  const baseRows = state.occ.filter((row) => {
    if (!matchesDate(row.date)) {
      return false;
    }
    if (term && !normalizeText((row.name || '') + ' ' + (row.occ || '') + ' ' + (row.document || '')).includes(term)) {
      return false;
    }
    return true;
  });
  let recurrenceKey = state.recurrenceFilter;
  if (recurrenceKey && !baseRows.some((row) => normalizeSupplierKey(row.name) === recurrenceKey)) {
    recurrenceKey = '';
    state.recurrenceFilter = '';
    const select = document.getElementById('recurrenceSupplierFilter');
    if (select) {
      select.value = '';
    }
  }
  const rows = baseRows.filter((row) => {
    if (recurrenceKey && normalizeSupplierKey(row.name) !== recurrenceKey) {
      return false;
    }
    return true;
  });
  state.occFiltered = rows;
  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="3">Nenhuma ocorrencia encontrada.</td></tr>';
    renderRecurrenceSection();
    return;
  }
  body.innerHTML = rows
    .map((row) => {
      return '<tr class="severity-' + row.severity + '">' +
        '<td>' + escapeHtml(row.name || '---') + '</td>' +
        '<td>' + escapeHtml(row.occ) + '</td>' +
        '<td>' + escapeHtml(row.document || '') + '</td>' +
        '</tr>';
    })
    .join('');
  renderRecurrenceSection(baseRows);
}

function renderRecurrenceSection(rowsSource) {
  const rows = Array.isArray(rowsSource)
    ? rowsSource
    : (Array.isArray(state.occFiltered) ? state.occFiltered : state.occ);
  const recurrenceData = buildRecurrenceData(rows);
  state.recurrence = recurrenceData;
  populateRecurrenceFilter(recurrenceData);
  renderRecurrenceChart(recurrenceData);
  renderRecurrenceDetails();
}

function shouldIgnoreRecurrenceSubject(subjectKey) {
  if (!subjectKey) {
    return true;
  }
  for (const ignore of RECURRENCE_IGNORE_SUBJECTS) {
    if (subjectKey === ignore || subjectKey.includes(ignore)) {
      return true;
    }
  }
  return false;
}

function normalizeRecurrenceSubject(subject) {
  const normalized = normalizeText(subject).replace(/\s+/g, ' ').trim();
  if (!normalized) {
    return '';
  }
  const tokens = normalized
    .split(' ')
    .filter((token) => token.length > 2 && !RECURRENCE_STOP_WORDS.has(token));
  return tokens.join(' ') || normalized;
}

function normalizeSupplierKey(name) {
  return normalizeText(name || '').replace(/\s+/g, ' ').trim();
}

function buildRecurrenceData(rows) {
  const suppliers = new Map();
  rows.forEach((row) => {
    const supplierName = safeString(row.name);
    const subject = safeString(row.occ);
    if (!supplierName || !subject) {
      return;
    }
    const supplierKey = normalizeSupplierKey(supplierName);
    if (!supplierKey) {
      return;
    }
    const subjectKey = normalizeRecurrenceSubject(subject);
    if (!subjectKey || shouldIgnoreRecurrenceSubject(subjectKey)) {
      return;
    }
    if (!suppliers.has(supplierKey)) {
      suppliers.set(supplierKey, { name: supplierName, subjects: new Map() });
    }
    const supplier = suppliers.get(supplierKey);
    if (!supplier.subjects.has(subjectKey)) {
      supplier.subjects.set(subjectKey, { label: subject.trim(), count: 0 });
    }
    const subjectEntry = supplier.subjects.get(subjectKey);
    subjectEntry.count += 1;
    if (!subjectEntry.label) {
      subjectEntry.label = subject.trim();
    }
  });
  const recurrence = [];
  suppliers.forEach((supplier, supplierKey) => {
    const topics = [];
    supplier.subjects.forEach((subject) => {
      if (subject.count >= 2) {
        topics.push(subject.label + ' (' + subject.count + 'x)');
      }
    });
    if (topics.length) {
      recurrence.push({ supplier: supplier.name, key: supplierKey, total: topics.length, topics });
    }
  });
  return recurrence.sort((a, b) => {
    if (b.total === a.total) {
      return a.supplier.localeCompare(b.supplier, 'pt-BR', { sensitivity: 'base' });
    }
    return b.total - a.total;
  });
}

function renderRecurrenceChart(data) {
  const canvas = document.getElementById('recurrenceChart');
  const emptyState = document.getElementById('recurrenceEmptyState');
  if (!canvas || typeof Chart === 'undefined') {
    return;
  }
  if (!data.length) {
    if (charts.occ) {
      charts.occ.destroy();
      charts.occ = null;
    }
    canvas.style.display = 'none';
    if (emptyState) {
      emptyState.classList.add('active');
    }
    return;
  }
  canvas.style.display = 'block';
  if (emptyState) {
    emptyState.classList.remove('active');
  }
  const filteredData = state.recurrenceFilter ? data.filter((item) => item.key === state.recurrenceFilter) : data.slice(0, 8);
  const targetData = filteredData.length ? filteredData : data.slice(0, 8);
  const labels = targetData.map((item) => item.supplier);
  const values = targetData.map((item) => item.total);
  if (charts.occ) {
    charts.occ.destroy();
  }
  const ctx = canvas.getContext('2d');
  charts.occ = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          label: 'Reincidencias',
          data: values,
          backgroundColor: ' #e26126ff',
          borderRadius: 8,
          borderSkipped: false
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: {
          ticks: {
            color: 'var(--text-secondary)',
            autoSkip: false,
            maxRotation: 45,
            minRotation: 0
          }
        },
        y: {
          beginAtZero: true,
          ticks: {
            stepSize: 1,
            color: 'var(--text-secondary)'
          }
        }
      },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (context) => 'Reincidencias: ' + context.parsed.y,
            afterLabel: (context) => {
              const topics = targetData[context.dataIndex]?.topics || [];
              return topics.length ? 'Topicos: ' + topics.join(', ') : '';
            }
          }
        }
      }
    }
  });
}

function populateRecurrenceFilter(data) {
  const select = document.getElementById('recurrenceSupplierFilter');
  if (!select) {
    state.recurrenceFilter = '';
    return;
  }
  if (!data.length) {
    select.innerHTML = '<option value="">Nenhuma reincidencia detectada</option>';
    select.disabled = true;
    select.value = '';
    state.recurrenceFilter = '';
    return;
  }
  const options = data
    .map((item) => {
      return '<option value="' + escapeHtml(item.key) + '">' + escapeHtml(item.supplier) + ' (' + item.total + ')</option>';
    })
    .join('');
  select.innerHTML = '<option value="">Todos os reincidentes</option>' + options;
  select.disabled = false;
  if (state.recurrenceFilter && !data.some((item) => item.key === state.recurrenceFilter)) {
    state.recurrenceFilter = '';
  }
  select.value = state.recurrenceFilter;
}

function renderRecurrenceDetails() {
  const container = document.getElementById('recurrenceTopics');
  if (!container) {
    return;
  }
  if (!state.recurrence.length) {
    container.innerHTML = '<p class="recurrence-topics-empty">Nenhuma reincidencia encontrada para os filtros aplicados.</p>';
    return;
  }
  if (!state.recurrenceFilter) {
    container.innerHTML = '<p class="recurrence-topics-empty">Selecione um fornecedor reincidente para explorar os assuntos recorrentes.</p>';
    return;
  }
  const supplier = state.recurrence.find((item) => item.key === state.recurrenceFilter);
  if (!supplier) {
    container.innerHTML = '<p class="recurrence-topics-empty">Selecao invalida. Escolha outro fornecedor.</p>';
    return;
  }
  const topics = supplier.topics
    .map((topic) => '<li>' + escapeHtml(topic) + '</li>')
    .join('');
  container.innerHTML =
    '<h4>Assuntos reincidentes de ' + escapeHtml(supplier.supplier) + '</h4>' +
    '<ul class="recurrence-topics-list">' + topics + '</ul>';
}

function handleRecurrenceFilterChange(event) {
  state.recurrenceFilter = event.target.value || '';
  renderOccTable();
}

function exportIQFData() {
  const rows = Array.isArray(state.iqfFiltered) ? state.iqfFiltered : getFilteredIqfRows();
  if (!rows.length) {
    alert('Sem dados para exportar.');
    return;
  }
  const data = [
    ['Fornecedor', 'Nota IQF', 'Data'],
    ...rows.map((row) => [
      row.name || '',
      row.iqf !== null ? row.iqf.toString() : '',
      formatDate(row.date) || ''
    ])
  ];
  downloadCsv('iqf.csv', data);
}

function exportOccurrences() {
  const rows = Array.isArray(state.occFiltered) ? state.occFiltered : state.occ;
  if (!rows.length) {
    alert('Sem dados para exportar.');
    return;
  }
  const data = [
    ['Fornecedor', 'Ocorrencia', 'Documento', 'Severidade'],
    ...rows.map((row) => [row.name || '', row.occ || '', row.document || '', row.severity])
  ];
  downloadCsv('ocorrencias.csv', data);
}

function downloadCsv(filename, rows) {
  const csv = rows
    .map((row) => row.map((cell) => '"' + String(cell || '').replace(/"/g, '""') + '"').join(';'))
    .join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

function toggleAIChat() {
  const panel = document.getElementById('aiChatPanel');
  const trigger = document.querySelector('.ai-floating-btn');
  if (!panel) {
    return;
  }
  const willOpen = !panel.classList.contains('open');
  panel.classList.toggle('open', willOpen);
  panel.setAttribute('aria-hidden', willOpen ? 'false' : 'true');
  if (trigger) {
    trigger.classList.toggle('active', willOpen);
  }
  if (willOpen) {
    setTimeout(() => {
      const input = document.getElementById('aiUserInput');
      if (input) {
        input.focus();
      }
    }, 150);
  }
}

function handleAIKeyPress(event) {
  if (event.key === 'Enter' && !event.shiftKey) {
    event.preventDefault();
    sendAIMessage();
  }
}

async function sendAIMessage() {
  if (chatState.loading) {
    return;
  }
  const input = document.getElementById('aiUserInput');
  const keyInput = document.getElementById('openaiKey');
  const container = document.getElementById('aiChatMessages');
  if (!input || !container) {
    return;
  }
  const message = input.value.trim();
  if (!message) {
    return;
  }
  const apiKey = keyInput?.value?.trim();
  if (!apiKey) {
    appendChatBubble(container, 'assistant', 'Informe sua OpenAI API key para continuar.');
    if (keyInput) {
      keyInput.focus();
    }
    return;
  }
  localStorage.setItem('openaiKey', apiKey);

  appendChatBubble(container, 'user', message);
  input.value = '';
  input.focus();

  chatState.history.push({ role: 'user', content: message });
  if (chatState.history.length > MAX_CHAT_HISTORY * 2) {
    chatState.history = chatState.history.slice(-MAX_CHAT_HISTORY * 2);
  }

  chatState.loading = true;
  const typingNode = appendChatBubble(container, 'assistant', 'Digitando...', { pending: true });
  chatState.typingNode = typingNode;

  try {
    const contextMessages = [
      { role: 'system', content: CHAT_PROMPT },
      { role: 'system', content: 'Snapshot atual: ' + buildChatContext() },
      ...chatState.history.slice(-MAX_CHAT_HISTORY)
    ];
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'Bearer ' + apiKey
      },
      body: JSON.stringify({
        model: 'gpt-4o-mini',
        temperature: 0.2,
        messages: contextMessages
      })
    });
    if (!response.ok) {
      throw new Error(await response.text());
    }
    const payload = await response.json();
    const answer = payload?.choices?.[0]?.message?.content?.trim() || 'Nao encontrei informacoes relevantes agora.';
    updateChatBubble(typingNode, answer);
    chatState.history.push({ role: 'assistant', content: answer });
  } catch (error) {
    console.error('[chat]', error);
    updateChatBubble(typingNode, 'Nao foi possivel consultar a API. Verifique a chave e tente novamente.');
  } finally {
    chatState.loading = false;
    chatState.typingNode = null;
  }
}

function appendChatBubble(container, role, text, options) {
  const wrapper = document.createElement('div');
  wrapper.className = 'ai-message ' + role + (options?.pending ? ' pending' : '');
  const bubble = document.createElement('div');
  bubble.className = 'message-content';
  bubble.innerHTML = formatChatContent(text);
  wrapper.appendChild(bubble);
  container.appendChild(wrapper);
  container.scrollTop = container.scrollHeight;
  return bubble;
}

function updateChatBubble(node, text) {
  if (!node) {
    return;
  }
  const wrapper = node.parentElement;
  if (wrapper) {
    wrapper.classList.remove('pending');
  }
  node.innerHTML = formatChatContent(text);
}

const CHAT_PROMPT = [
  'Voce e um especialista em qualificacao e homologacao de fornecedores.',
  'Use linguagem objetiva, clara e organizada.',
  'Baseie as respostas nos dados do snapshot sempre que possivel e destaque riscos e oportunidades.'
].join(' ');

function buildChatContext() {
  if (!state.combined.length) {
    return 'Sem dados carregados.';
  }
  const total = state.combined.length;
  const homologados = state.combined.filter((row) => row.status === 'Homologado').length;
  const reprovados = state.combined.filter((row) => row.status === 'Reprovado').length;
  const mediaIqf = avg(state.combined.map((row) => row.iqf));
  const mediaHom = avg(state.combined.map((row) => row.homolog));
  const top = state.combined
    .slice()
    .sort((a, b) => b.iqf - a.iqf)
    .slice(0, 5)
    .map((row) => ({ fornecedor: row.name, iqf: row.iqf, homolog: row.homolog, status: row.status, vencimento: row.expire }));
  const faltantes = state.missing.slice(0, 10).map((item) => item.name || item.code);
  return JSON.stringify({ total, homologados, reprovados, mediaIqf, mediaHom, top, faltantes });
}

function showLoading(show) {
  const overlay = document.getElementById('loadingOverlay');
  if (overlay) {
    overlay.style.display = show ? 'flex' : 'none';
  }
}

function setText(id, value) {
  const element = document.getElementById(id);
  if (element) {
    element.innerText = value;
  }
}

function avg(values) {
  const filtered = values.filter((value) => value !== null && !Number.isNaN(value));
  if (!filtered.length) {
    return null;
  }
  return filtered.reduce((sum, value) => sum + value, 0) / filtered.length;
}

function normalizeKeys(row) {
  const result = {};
  Object.entries(row || {}).forEach(([key, value]) => {
    const normalized = normalizeText(key).replace(/[^a-z0-9]/g, '');
    if (normalized) {
      result[normalized] = value;
    }
  });
  return result;
}

function normalizeText(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value)
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9\s]/gi, '')
    .trim()
    .toLowerCase();
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }
  const cleaned = String(value).trim().replace(/\s/g, '').replace(/\.(?=\d{3})/g, '').replace(',', '.');
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function toISODate(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  if (value instanceof Date && !Number.isNaN(value)) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === 'number') {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(epoch.getTime() + value * 86400000).toISOString().slice(0, 10);
  }
  const raw = String(value).trim();
  const match = raw.match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})$/);
  if (match) {
    return match[3] + '-' + match[2] + '-' + match[1];
  }
  const parsed = new Date(raw);
  return Number.isNaN(parsed) ? null : parsed.toISOString().slice(0, 10);
}

function formatPercent(value) {
  if (value === null || value === undefined || Number.isNaN(value)) {
    return '--';
  }
  return formatNumber(value) + '%';
}

function formatNumber(value) {
  return new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value);
}

function formatDate(iso) {
  if (!iso) {
    return null;
  }
  const parts = iso.split('-');
  if (parts.length !== 3) {
    return null;
  }
  return parts[2] + '/' + parts[1] + '/' + parts[0];
}

function formatMonth(key) {
  if (!key) {
    return '--';
  }
  const parts = key.split('-');
  if (parts.length !== 2) {
    return key;
  }
  return parts[1] + '/' + parts[0].slice(-2);
}

function escapeHtml(value) {
  return String(value || '').replace(/[&<>"']/g, (match) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  })[match]);
}

function safeString(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function toId(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  return String(value).trim().replace(/\.0+$/, '');
}

function roundValue(value) {
  return Number.isFinite(value) ? Number(value.toFixed(2)) : value;
}

function deriveStatus(baseStatus, iqfScore, homologScore) {
  const belowThreshold = (score) => score !== null && score < SCORE_THRESHOLD;
  if (belowThreshold(iqfScore) || belowThreshold(homologScore)) {
    return 'Reprovado';
  }
  if (baseStatus === 'Homologado') {
    return 'Homologado';
  }
  if (baseStatus === 'Reprovado') {
    return 'Reprovado';
  }
  return 'Pendente';
}

function badge(status) {
  if (status === 'Homologado') {
    return 'status-homologado';
  }
  if (status === 'Reprovado') {
    return 'status-reprovado';
  }
  return 'status-pendente';
}

function severityLevel(text) {
  const normalized = normalizeText(text);
  if (!normalized) {
    return 'low';
  }
  if (/nao\s+efetuou|falha|critico|grave/.test(normalized)) {
    return 'critical';
  }
  if (/atraso|problema|irregular|bloqueio/.test(normalized)) {
    return 'high';
  }
  if (/pendente|ajuste|analise|monitor/.test(normalized)) {
    return 'medium';
  }
  return 'low';
}

function formatChatContent(text) {
  const safe = escapeHtml(text || '');
  const blocks = safe.split(/\n{2,}/);
  return blocks
    .map((block) => {
      const lines = block.split('\n');
      const listItems = lines.filter((line) => /^[-•]/.test(line.trim()));
      if (listItems.length === lines.length && listItems.length > 0) {
        const items = lines
          .map((line) => '<li>' + line.trim().replace(/^[-•]\s*/, '') + '</li>')
          .join('');
        return '<ul>' + items + '</ul>';
      }
      return '<p>' + lines.join('<br>') + '</p>';
    })
    .join('');
}

function mapStatus(value) {
  const normalized = normalizeText(value);
  if (['s', 'sim', 'homologado', 'aprovado', 'ativo', 'qualificado'].includes(normalized)) {
    return 'Homologado';
  }
  if (['n', 'nao', 'reprovado', 'bloqueado'].includes(normalized)) {
    return 'Reprovado';
  }
  return 'Pendente';
}

window.toggleAIChat = toggleAIChat;
window.sendAIMessage = sendAIMessage;
window.handleAIKeyPress = handleAIKeyPress;
window.exportIQFData = exportIQFData;
window.exportOccurrences = exportOccurrences;

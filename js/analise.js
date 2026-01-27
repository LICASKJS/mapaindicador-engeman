const DATA_FILES = {
  homolog: 'dados/fornecedores_homologados.xlsx',
  iqf: 'dados/atendimento controle_qualidade.xlsx'
};

const uploadedDataFiles = {
  homolog: null,
  iqf: null
};

const IQF_SEGMENTS = [
  { id: 'excelente', label: 'Excelentes (IQF ≥ 90)', min: 90, max: Infinity },
  { id: 'bom', label: 'Bons (80 ≤ IQF < 90)', min: 80, max: 89.99 },
  { id: 'regular', label: 'Regulares (70 ≤ IQF < 80)', min: 70, max: 79.99 },
  { id: 'critico', label: 'Críticos (IQF < 70)', min: -Infinity, max: 69.99 },
  { id: 'sem-dados', label: 'Sem medição IQF', none: true }
];

const API_KEY_STORAGE_KEY = 'analise-openai-key';
const OPENAI_API_KEY_DEFAULT =
  (typeof process !== 'undefined' && process.env && process.env.OPENAI_API_KEY) ||
  (typeof window !== 'undefined' && window.OPENAI_API_KEY) ||
  '';
const EMAIL_API_ENDPOINT = window.EMAIL_API_ENDPOINT || '';
const EMAIL_API_TOKEN = window.EMAIL_API_TOKEN || '';


const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/i;
const MAX_SUPPLIER_SUGGESTIONS = 8;
const MAX_HISTORY_POINTS = 12;
let incrementalId = 0;

const state = {
  suppliers: [],
  supplierById: new Map(),
  segments: [],
  segmentById: new Map(),
  selectedSupplier: null,
  lastFeedback: '',
  lastFeedbackHtml: '',
  lastEmailHtml: '',
  feedbackSupplierId: null,
  monthlySummary: new Map(),
  availableMonths: [],
  emailSubjectTemplate: 'Feedback IQF - {{fornecedor}}',
  openAiApiKey: '',
  inputDefaultPlaceholder: ''
};

const dom = {
  messages: null,
  mainMenu: null,
  userInput: null,
  sendBtn: null,
  attachFilesBtn: null,
  dataFilesInput: null,
  toast: null,
  settingsBtn: null,
  settingsOverlay: null,
  settingsCloseBtn: null,
  settingsSaveBtn: null,
  settingsClearBtn: null,
  settingsEmailSubject: null,
  apiKeyBar: null,
  apiKeyInput: null,
  apiKeyToggle: null,
  applyApiKeyBtn: null,
  clearApiKeyBtn: null,
  apiKeyStatus: null
};

function cacheSafePath(path) {
  if (!path) {
    return path;
  }
  const cacheBust = Date.now().toString(36);
  const base =
    typeof window !== 'undefined' && window.location
      ? window.location.origin && window.location.origin !== 'null'
        ? window.location.origin
        : window.location.href
      : '';
  try {
    const resolved = new URL(path, base || undefined);
    resolved.searchParams.set('cb', cacheBust);
    return resolved.toString();
  } catch (error) {
    const [cleanPath, ...rest] = String(path).split('?');
    const encodedPath = encodeURI(cleanPath);
    const query = rest.length ? '?' + rest.join('?') : '';
    const separator = query ? '&' : '?';
    return encodedPath + query + separator + 'cb=' + cacheBust;
  }
}

function getOpenAiApiKey() {
  return (state.openAiApiKey || OPENAI_API_KEY_DEFAULT || '').trim();
}

function setOpenAiApiKey(value, options) {
  const normalized = value ? String(value).trim() : '';
  const previous = state.openAiApiKey;
  state.openAiApiKey = normalized;
  try {
    if (normalized) {
      localStorage.setItem(API_KEY_STORAGE_KEY, normalized);
    } else {
      localStorage.removeItem(API_KEY_STORAGE_KEY);
    }
  } catch (error) {
    console.warn('[analise:setOpenAiApiKey]', error);
  }
  if (dom.apiKeyInput && dom.apiKeyInput.value !== normalized) {
    dom.apiKeyInput.value = normalized;
  }
  updateApiKeyStatus();
  const changed = previous !== normalized;
  if (options?.showToast && changed) {
    showToast(normalized ? 'Chave OpenAI salva.' : 'Chave OpenAI removida.');
  }
  return changed;
}

function updateApiKeyStatus() {
  if (!dom.apiKeyStatus) {
    return;
  }
  const key = getOpenAiApiKey();
  if (key) {
    const masked =
      key.length > 12 ? key.slice(0, 4) + '...' + key.slice(-4) : 'chave ativa';
    dom.apiKeyStatus.textContent = 'Chave ativa (' + masked + ')';
    dom.apiKeyStatus.dataset.status = 'ready';
  } else {
    dom.apiKeyStatus.textContent = 'Nenhuma chave informada';
    dom.apiKeyStatus.dataset.status = 'empty';
  }
}

function toggleApiKeyVisibility() {
  if (!dom.apiKeyInput || !dom.apiKeyToggle) {
    return;
  }
  const hidden = dom.apiKeyInput.type === 'password';
  dom.apiKeyInput.type = hidden ? 'text' : 'password';
  dom.apiKeyToggle.classList.toggle('active', hidden);
  dom.apiKeyToggle.setAttribute('aria-label', hidden ? 'Ocultar chave' : 'Mostrar chave');
}

function handleApplyApiKey() {
  if (!dom.apiKeyInput) {
    return;
  }
  const changed = setOpenAiApiKey(dom.apiKeyInput.value, { showToast: true });
  if (!changed) {
    updateApiKeyStatus();
    showToast(getOpenAiApiKey() ? 'Chave OpenAI ativa.' : 'Nenhuma chave informada.');
  }
}

function handleClearApiKey() {
  const changed = setOpenAiApiKey('', { showToast: true });
  if (!changed) {
    updateApiKeyStatus();
  }
  if (dom.apiKeyInput) {
    dom.apiKeyInput.focus();
  }
}

function handleApiKeyBlur() {
  if (!dom.apiKeyInput) {
    return;
  }
  setOpenAiApiKey(dom.apiKeyInput.value, { showToast: false });
}

function handleApiKeyKeydown(event) {
  if (event.key === 'Enter') {
    event.preventDefault();
    handleApplyApiKey();
  } else if (event.key === 'Escape') {
    event.preventDefault();
    if (dom.apiKeyInput) {
      dom.apiKeyInput.value = state.openAiApiKey;
    }
  }
}

document.addEventListener('DOMContentLoaded', init);

window.showSubmenu = handleShowSubmenu;
window.showIndicadoresMensais = handleIndicadoresMensais;
window.contactBase = handleContactBase;
window.showProcedimentoEngeman = handleProcedimento;
window.previewEmailTemplate = previewEmailTemplate;
window.openSupplierSearch = handleSupplierSearch;

async function init() {
  cacheDom();
  bindEvents();
  loadSettings();

  const loadingMessage = appendBotMessage('Carregando planilhas de fornecedores, aguarde...');
  try {
    await loadData();
    updateMessage(
      loadingMessage,
      'Bases carregadas com sucesso! Acesse "Indicadores dos Fornecedores" para explorar os agrupamentos por IQF.',
      true
    );
    showToast('Dados atualizados. Escolha um agrupamento para iniciar a analise.');
  } catch (error) {
    console.error('[analise:init]', error);
    updateMessage(
      loadingMessage,
      'Nao foi possivel carregar as planilhas. Use o botao de clipe para anexar as duas planilhas (.xlsx) e tente novamente.',
      true
    );
  }
}

function cacheDom() {
  dom.messages = document.getElementById('messagesContainer');
  dom.mainMenu = document.getElementById('main-menu');
  dom.userInput = document.getElementById('userInput');
  dom.sendBtn = document.getElementById('sendBtn');
  dom.attachFilesBtn = document.getElementById('attachFilesBtn');
  dom.dataFilesInput = document.getElementById('dataFilesInput');
  if (dom.userInput && !state.inputDefaultPlaceholder) {
    state.inputDefaultPlaceholder = dom.userInput.placeholder || '';
  }
  dom.toast = document.getElementById('toast');
  dom.settingsBtn = document.getElementById('settingsBtn');
  dom.settingsOverlay = document.getElementById('settingsOverlay');
  dom.settingsCloseBtn = document.getElementById('closeSettingsBtn');
  dom.settingsSaveBtn = document.getElementById('saveSettingsBtn');
  dom.settingsClearBtn = document.getElementById('clearSettingsBtn');
  dom.settingsEmailSubject = document.getElementById('emailSubjectInput');
  dom.apiKeyBar = document.getElementById('apiKeyBar');
  dom.apiKeyInput = document.getElementById('apiKeyInput');
  dom.apiKeyToggle = document.getElementById('toggleApiKeyBtn');
  dom.applyApiKeyBtn = document.getElementById('applyApiKeyBtn');
  dom.clearApiKeyBtn = document.getElementById('clearApiKeyBtn');
  dom.apiKeyStatus = document.getElementById('apiKeyStatus');
  updateApiKeyStatus();
}

function bindEvents() {
  if (dom.sendBtn) {
    dom.sendBtn.addEventListener('click', handleUserSend);
  }
  if (dom.userInput) {
    dom.userInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        handleUserSend();
      }
    });
  }
  if (dom.attachFilesBtn && dom.dataFilesInput) {
    dom.attachFilesBtn.addEventListener('click', () => dom.dataFilesInput.click());
    dom.dataFilesInput.addEventListener('change', handleDataFilesSelected);
  }
  if (dom.settingsBtn) {
    dom.settingsBtn.addEventListener('click', openSettings);
  }
  if (dom.settingsCloseBtn) {
    dom.settingsCloseBtn.addEventListener('click', closeSettings);
  }
  if (dom.settingsOverlay) {
    dom.settingsOverlay.addEventListener('click', (event) => {
      if (event.target === dom.settingsOverlay) {
        closeSettings();
      }
    });
  }
  if (dom.settingsSaveBtn) {
    dom.settingsSaveBtn.addEventListener('click', saveSettings);
  }
  if (dom.settingsClearBtn) {
    dom.settingsClearBtn.addEventListener('click', clearSettings);
  }
  if (dom.apiKeyInput) {
    dom.apiKeyInput.addEventListener('keydown', handleApiKeyKeydown);
    dom.apiKeyInput.addEventListener('blur', handleApiKeyBlur);
  }
  if (dom.applyApiKeyBtn) {
    dom.applyApiKeyBtn.addEventListener('click', handleApplyApiKey);
  }
  if (dom.clearApiKeyBtn) {
    dom.clearApiKeyBtn.addEventListener('click', handleClearApiKey);
  }
  if (dom.apiKeyToggle) {
    dom.apiKeyToggle.addEventListener('click', toggleApiKeyVisibility);
  }
}

function guessDataFileKind(filename) {
  const name = String(filename || '').toLowerCase();
  if (!name) {
    return null;
  }
  if (name.includes('homolog')) {
    return 'homolog';
  }
  if (name.includes('qualidade') || name.includes('controle') || name.includes('atendimento') || name.includes('iqf')) {
    return 'iqf';
  }
  return null;
}

async function handleDataFilesSelected(event) {
  const files = Array.from(event?.target?.files || []);
  if (!files.length) {
    return;
  }
  if (event?.target) {
    event.target.value = '';
  }

  const picked = { homolog: null, iqf: null };

  files.forEach((file) => {
    const kind = guessDataFileKind(file.name);
    if (kind && !picked[kind]) {
      picked[kind] = file;
    }
  });

  if (!picked.homolog && !picked.iqf && files.length >= 2) {
    picked.homolog = files[0];
    picked.iqf = files[1];
  } else if (!picked.homolog && picked.iqf && files.length >= 2) {
    picked.homolog = files.find((file) => file !== picked.iqf) || null;
  } else if (!picked.iqf && picked.homolog && files.length >= 2) {
    picked.iqf = files.find((file) => file !== picked.homolog) || null;
  }

  if (!picked.homolog || !picked.iqf) {
    showToast('Selecione as duas planilhas (.xlsx) para atualizar as bases.');
    appendBotMessage(
      'Para atualizar as bases, anexe as duas planilhas (.xlsx): <strong>fornecedores_homologados</strong> e <strong>atendimento controle_qualidade</strong>.',
      true
    );
    return;
  }

  uploadedDataFiles.homolog = picked.homolog;
  uploadedDataFiles.iqf = picked.iqf;

  const loadingMessage = appendBotMessage('Carregando planilhas anexadas...');
  showToast('Atualizando bases...');

  try {
    await loadData();
    state.selectedSupplier = null;
    state.feedbackSupplierId = null;
    state.lastFeedback = '';
    state.lastFeedbackHtml = '';
    state.lastEmailHtml = '';
    updateMessage(
      loadingMessage,
      'Bases atualizadas com as planilhas anexadas. Digite <strong>menu</strong> para voltar ao inicio.',
      true
    );
    showToast('Bases atualizadas.');
  } catch (error) {
    console.error('[analise:data-upload]', error);
    updateMessage(
      loadingMessage,
      'Falha ao carregar as planilhas anexadas. Verifique os arquivos e tente novamente.',
      true
    );
    showToast('Falha ao atualizar bases.');
  }
}

function loadSettings() {
  let persistedApiKey = '';
  try {
    const persistedSubject = localStorage.getItem('analise-email-subject');
    if (persistedSubject) {
      state.emailSubjectTemplate = persistedSubject;
      if (dom.settingsEmailSubject) {
        dom.settingsEmailSubject.value = persistedSubject;
      }
    } else if (dom.settingsEmailSubject) {
      dom.settingsEmailSubject.value = state.emailSubjectTemplate;
    }
    persistedApiKey = localStorage.getItem(API_KEY_STORAGE_KEY) || '';
  } catch (error) {
    console.warn('[analise:loadSettings]', error);
  }
  if (persistedApiKey) {
    state.openAiApiKey = persistedApiKey;
  } else if (OPENAI_API_KEY_DEFAULT) {
    state.openAiApiKey = OPENAI_API_KEY_DEFAULT;
  }
  if (dom.apiKeyInput) {
    dom.apiKeyInput.value = state.openAiApiKey;
    dom.apiKeyInput.type = 'password';
  }
  updateApiKeyStatus();
}

async function loadData() {
  if (typeof XLSX === 'undefined') {
    throw new Error('Biblioteca XLSX nao carregada');
  }
  const [homRows, iqfRows] = await Promise.all([
    loadWorkbook(uploadedDataFiles.homolog || DATA_FILES.homolog),
    loadWorkbook(uploadedDataFiles.iqf || DATA_FILES.iqf)
  ]);
  const homolog = homRows.map(mapHomolog).filter(Boolean);
  const iqf = iqfRows
    .map(mapIqf)
    .filter((row) => row.code || row.name || row.iqf !== null || row.occ);
  buildSupplierState(homolog, iqf);
}

async function loadWorkbook(source) {
  if (source && typeof source === 'object' && typeof source.arrayBuffer === 'function') {
    const buffer = await source.arrayBuffer();
    return parseWorkbookRows(buffer);
  }

  const path = String(source || '');
  const response = await fetch(cacheSafePath(path), { cache: 'no-store' });
  if (!response.ok) {
    throw new Error('Falha ao carregar ' + path + ' (' + response.status + ')');
  }
  const buffer = await response.arrayBuffer();
  return parseWorkbookRows(buffer);
}

function parseWorkbookRows(buffer) {
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

function buildSupplierState(homologRows, iqfRows) {
  const codeBuckets = new Map();
  const nameBuckets = new Map();

  iqfRows.forEach((row) => {
    if (row.code) {
      const key = String(row.code);
      if (!codeBuckets.has(key)) {
        codeBuckets.set(key, []);
      }
      codeBuckets.get(key).push(row);
    }
    if (row.name) {
      const normalized = normalizeText(row.name);
      if (normalized) {
        if (!nameBuckets.has(normalized)) {
          nameBuckets.set(normalized, []);
        }
        nameBuckets.get(normalized).push(row);
      }
    }
  });

  const suppliers = [];
  const usedIds = new Set();
  const seenKeys = new Set();

  const registerSupplier = (supplier) => {
    if (!supplier || usedIds.has(supplier.id)) {
      return;
    }
    usedIds.add(supplier.id);
    suppliers.push(supplier);
  };

  homologRows.forEach((hom) => {
    const codeKey = hom.code ? String(hom.code) : null;
    const nameKey = hom.name ? normalizeText(hom.name) : null;
    const byCode = codeKey && codeBuckets.get(codeKey) ? codeBuckets.get(codeKey) : [];
    const byName = nameKey && nameBuckets.get(nameKey) ? nameBuckets.get(nameKey) : [];
    const rows = byCode.length >= byName.length ? byCode : byName;
    const supplier = composeSupplier(hom, rows, usedIds);
    registerSupplier(supplier);
    if (codeKey) {
      seenKeys.add('code:' + codeKey);
    }
    if (nameKey) {
      seenKeys.add('name:' + nameKey);
    }
  });

  codeBuckets.forEach((rows, code) => {
    if (!seenKeys.has('code:' + code)) {
      const placeholderHom = { code, name: rows[0]?.name || null, status: 'Pendente', score: null, expire: null };
      registerSupplier(composeSupplier(placeholderHom, rows, usedIds));
      seenKeys.add('code:' + code);
    }
  });

  nameBuckets.forEach((rows, normalizedName) => {
    if (!seenKeys.has('name:' + normalizedName)) {
      const placeholderHom = { code: null, name: rows[0]?.name || null, status: 'Pendente', score: null, expire: null };
      registerSupplier(composeSupplier(placeholderHom, rows, usedIds));
      seenKeys.add('name:' + normalizedName);
    }
  });

  suppliers.sort((a, b) => a.name.localeCompare(b.name, 'pt-BR', { sensitivity: 'accent' }));
  state.suppliers = suppliers;
  state.supplierById = new Map(suppliers.map((supplier) => [supplier.id, supplier]));
  state.segments = prepareSegments(suppliers);
  state.segmentById = new Map(state.segments.map((segment) => [segment.id, segment]));
  buildMonthlySummaries(iqfRows, suppliers);
}

function composeSupplier(hom, iqfRows, usedIds) {
  const iqfSummary = summarizeIqf(iqfRows || []);
  const code = hom?.code ? String(hom.code) : iqfRows?.[0]?.code ? String(iqfRows[0].code) : null;
  const name = safeSupplierName(hom?.name, iqfRows?.[0]?.name);
  const normalizedName = normalizeText(name);
  const homologScore = roundValue(hom?.score);
  const baseStatus = hom?.status || 'Pendente';
  const status = deriveStatus(baseStatus, iqfSummary.average, homologScore);
  const id = buildSupplierId(code, normalizedName, usedIds);

  return {
    id,
    code,
    name,
    normalizedName,
    status,
    baseStatus,
    homologScore,
    averageIqf: iqfSummary.average,
    iqfSamples: iqfSummary.samples,
    lastIqfDate: iqfSummary.lastDate,
    expire: hom?.expire || null,
    document: iqfSummary.document,
    occurrences: iqfSummary.occurrences,
    iqfHistory: iqfSummary.history,
    searchText: [name || '', code || ''].join(' ').toLowerCase(),
    baseStatusRaw: hom?.status || null
  };
}

function summarizeIqf(rows) {
  const history = [];
  const occurrences = [];
  let document = '';

  rows.forEach((row) => {
    if (row.iqf !== null) {
      history.push({
        value: roundValue(row.iqf),
        date: row.date,
        formattedDate: formatDate(row.date)
      });
    }
    if (row.occ) {
      occurrences.push({
        text: row.occ,
        date: row.date,
        formattedDate: formatDate(row.date),
        document: row.document,
        severity: severityLevel(row.occ)
      });
    }
    if (!document && row.document) {
      document = row.document;
    }
  });

  history.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  occurrences.sort((a, b) => (b.date || '').localeCompare(a.date || ''));

  const trimmedHistory = history.slice(0, MAX_HISTORY_POINTS);
  const samples = history.length;
  const average = samples ? roundValue(history.reduce((sum, entry) => sum + (entry.value || 0), 0) / samples) : null;
  const lastDate = history.length ? history[0].date : null;

  return { history: trimmedHistory, occurrences, average, samples, lastDate, document };
}

function safeSupplierName(...names) {
  const options = names.filter(Boolean).map((value) => safeString(value));
  const preferred = options.find((value) => value.length > 0);
  return preferred || 'Fornecedor nao identificado';
}

function prepareSegments(suppliers) {
  return IQF_SEGMENTS.map((segment) => {
    const matches = suppliers.filter((supplier) => {
      if (segment.none) {
        return supplier.averageIqf === null;
      }
      if (supplier.averageIqf === null) {
        return false;
      }
      return supplier.averageIqf >= segment.min && supplier.averageIqf <= segment.max;
    });
    return { ...segment, suppliers: matches };
  }).filter((segment) => segment.suppliers.length);
}

function buildMonthlySummaries(iqfRows, suppliers) {
  const supplierIndex = new Map();
  suppliers.forEach((supplier) => {
    if (supplier.code) {
      supplierIndex.set('code:' + String(supplier.code).trim(), supplier);
    }
    if (supplier.normalizedName) {
      supplierIndex.set('name:' + supplier.normalizedName, supplier);
    }
  });

  const summary = new Map();

  iqfRows.forEach((row) => {
    if (!row || row.iqf === null || row.iqf === undefined || !row.date) {
      return;
    }
    const monthKey = row.date.slice(0, 7);
    if (!/^\d{4}-\d{2}$/.test(monthKey)) {
      return;
    }
    if (!summary.has(monthKey)) {
      summary.set(monthKey, { totalSum: 0, totalCount: 0, suppliers: new Map() });
    }
    const monthEntry = summary.get(monthKey);
    monthEntry.totalSum += row.iqf;
    monthEntry.totalCount += 1;

    const normalizedName = normalizeText(row.name);
    let supplierKey = null;
    if (row.code) {
      supplierKey = 'code:' + String(row.code).trim();
    } else if (normalizedName) {
      supplierKey = 'name:' + normalizedName;
    } else {
      supplierKey = 'name:' + safeString(row.name).toLowerCase();
    }

    if (!monthEntry.suppliers.has(supplierKey)) {
      const supplierRef =
        supplierIndex.get(supplierKey) ||
        (row.code ? supplierIndex.get('code:' + String(row.code).trim()) : null) ||
        (normalizedName ? supplierIndex.get('name:' + normalizedName) : null) ||
        null;
      monthEntry.suppliers.set(supplierKey, {
        key: supplierKey,
        name: safeSupplierName(row.name),
        code: supplierRef?.code || row.code || null,
        status: supplierRef?.status || 'Pendente',
        sum: 0,
        count: 0
      });
    }
    const supplierEntry = monthEntry.suppliers.get(supplierKey);
    supplierEntry.sum += row.iqf;
    supplierEntry.count += 1;
  });

  state.monthlySummary = summary;
  state.availableMonths = Array.from(summary.keys()).sort((a, b) => b.localeCompare(a));
}

function handleShowSubmenu(key) {
  if (key !== 'fornecedores') {
    appendBotMessage('Menu em construção. Em breve novas opções aqui.', true);
    return;
  }
  if (!state.suppliers.length) {
    appendBotMessage('Os dados ainda estão sendo carregados. Aguarde alguns instantes e tente novamente.', true);
    return;
  }
  renderIqfSegments();
}

function handleSupplierSearch() {
  if (!dom.userInput) {
    appendBotMessage('Campo de busca indisponivel no momento.', true);
    return;
  }
  const content = createMessage('bot');
  content.innerHTML =
    '<p><strong>Pesquisar fornecedor</strong></p>' +
    '<p>Digite o nome completo ou parte do nome no campo inferior e pressione Enter. O sistema listará as correspondências mais próximas.</p>';
  dom.userInput.placeholder = 'Pesquisar fornecedor por nome...';
  dom.userInput.value = '';
  dom.userInput.focus();
  showToast('Digite o nome do fornecedor e pressione Enter para pesquisar.');
}

function renderMainMenuOptions() {
  const content = createMessage('bot');
  const intro = document.createElement('p');
  intro.innerHTML = '<strong>Selecione uma op&ccedil;&atilde;o para continuar:</strong>';
  content.appendChild(intro);

  const actions = document.createElement('div');
  actions.className = 'message-actions';
  const options = [
    { label: 'Pesquisar Fornecedor', handler: handleSupplierSearch },
    { label: 'Indicadores dos Fornecedores', handler: () => handleShowSubmenu('fornecedores') },
    { label: 'Indicadores Mensais', handler: handleIndicadoresMensais },
    { label: 'Procedimento Engeman', handler: handleProcedimento },
    { label: 'Contato Base', handler: handleContactBase },
    { label: 'Configura&ccedil;&otilde;es', handler: openSettings }
  ];
  options.forEach((option) => {
    const button = document.createElement('button');
    button.className = 'message-action-btn';
    button.textContent = option.label;
    button.addEventListener('click', () => option.handler());
    actions.appendChild(button);
  });
  content.appendChild(actions);
}

function renderIqfSegments() {
  const content = createMessage('bot');
  const intro = document.createElement('p');
  intro.innerHTML = 'Escolha um agrupamento de fornecedores pela média IQF:';
  content.appendChild(intro);

  const actions = document.createElement('div');
  actions.className = 'message-actions';
  state.segments.forEach((segment) => {
    const button = document.createElement('button');
    button.className = 'message-action-btn';
    button.innerHTML = segment.label + ' (' + segment.suppliers.length + ')';
    button.addEventListener('click', () => showSuppliersBySegment(segment.id));
    actions.appendChild(button);
  });
  content.appendChild(actions);
}

function showSuppliersBySegment(segmentId) {
  const segment = state.segmentById.get(segmentId);
  if (!segment) {
    appendBotMessage('Segmento nao localizado. Atualize a página e tente novamente.', true);
    return;
  }
  const content = createMessage('bot');
  const title = document.createElement('p');
  title.innerHTML = '<strong>' + segment.label + '</strong> — ' + segment.suppliers.length + ' fornecedor(es). Selecione para detalhar:';
  content.appendChild(title);

  if (!segment.suppliers.length) {
    const empty = document.createElement('p');
    empty.textContent = 'Nenhum fornecedor neste agrupamento.';
    content.appendChild(empty);
    return;
  }

  const list = document.createElement('div');
  list.className = 'message-actions supplier-grid';
  const sortedSuppliers = segment.suppliers
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name, 'pt-BR', { sensitivity: 'accent' }));
  sortedSuppliers.forEach((supplier) => {
    const button = document.createElement('button');
    button.className = 'message-action-btn';
    button.textContent = supplier.name + (supplier.code ? ' (' + supplier.code + ')' : '');
    button.addEventListener('click', () => showSupplierDetails(supplier.id));
    list.appendChild(button);
  });
  content.appendChild(list);
  state.currentSegmentId = segmentId;
}

function showSupplierDetails(supplierId) {
  const supplier = state.supplierById.get(supplierId);
  if (!supplier) {
    appendBotMessage('Fornecedor nao encontrado. Utilize o menu novamente.', true);
    return;
  }
  state.selectedSupplier = supplier;
  state.feedbackSupplierId = null;
  state.lastFeedback = '';
  state.lastFeedbackHtml = '';
  state.lastEmailHtml = '';

  const content = createMessage('bot');
  const header = document.createElement('div');
  header.innerHTML =
    '<h3>' +
    escapeHtml(supplier.name) +
    '</h3>' +
    renderStatusBadge(supplier.status) +
    '<p>Código: <strong>' +
    (supplier.code ? escapeHtml(supplier.code) : 'n/d') +
    '</strong> • Nota Homologação: <strong>' +
    (supplier.homologScore !== null ? supplier.homologScore : 'n/d') +
    '</strong> • Média IQF: <strong>' +
    (supplier.averageIqf !== null ? supplier.averageIqf : 'n/d') +
    '</strong></p>' +
    '<p>Registros IQF: <strong>' +
    supplier.iqfSamples +
    '</strong> • Última medição: <strong>' +
    (formatDate(supplier.lastIqfDate) || 'n/d') +
    '</strong></p>';
  content.appendChild(header);

  if (supplier.expire) {
    const expire = document.createElement('p');
    expire.textContent = 'Validade da homologação: ' + formatDate(supplier.expire);
    content.appendChild(expire);
  }

  const feedbackId = safeDomId('feedback', supplier.id);
  const feedbackCard = document.createElement('div');
  feedbackCard.className = 'feedback-card';
  feedbackCard.id = feedbackId;
  content.appendChild(feedbackCard);

  const refreshLink = document.createElement('button');
  refreshLink.className = 'inline-link';
  refreshLink.type = 'button';
  refreshLink.textContent = 'Reprocessar analise';
  refreshLink.addEventListener('click', () => {
    generateSupplierFeedback(supplier, feedbackCard);
  });
  feedbackCard.appendChild(refreshLink);
  feedbackCard.__refreshButton = refreshLink;

  const occurrencesCard = document.createElement('div');
  occurrencesCard.className = 'occurrences-card';
  occurrencesCard.innerHTML = '<h4>Ocorrências relevantes</h4>';
  if (supplier.occurrences.length) {
    const list = document.createElement('ul');
    supplier.occurrences.forEach((occ) => {
      const item = document.createElement('li');
      item.innerHTML =
        '<strong>' +
        (occ.formattedDate || 'Data n/d') +
        ':</strong> ' +
        escapeHtml(occ.text) +
        (occ.document ? ' (Doc: ' + escapeHtml(occ.document) + ')' : '');
      list.appendChild(item);
    });
    occurrencesCard.appendChild(list);
  } else {
    const empty = document.createElement('p');
    empty.textContent = 'Sem ocorrências registradas para este fornecedor.';
    occurrencesCard.appendChild(empty);
  }
  content.appendChild(occurrencesCard);

  if (supplier.iqfHistory.length) {
    const historyCard = document.createElement('div');
    historyCard.className = 'occurrences-card';
    historyCard.innerHTML = '<h4>Histórico de IQF</h4>';
    const historyList = document.createElement('ul');
    supplier.iqfHistory.forEach((entry) => {
      const item = document.createElement('li');
      item.textContent = (entry.formattedDate || 'Data n/d') + ' \u2013 IQF: ' + entry.value;
      historyList.appendChild(item);
    });
    historyCard.appendChild(historyList);
    content.appendChild(historyCard);
  }

  const hint = document.createElement('p');
  hint.className = 'message-hint';
  hint.innerHTML = 'Digite o e-mail do fornecedor na caixa de mensagem para preparar o envio automatico do feedback.';
  content.appendChild(hint);

  generateSupplierFeedback(supplier, feedbackCard);
}

function buildSupplierClassificationNarrative(supplier) {
  const iqf = Number.isFinite(supplier.averageIqf) ? supplier.averageIqf : null;
  const formattedIqf = iqf !== null ? formatScoreValue(iqf) : 'N/D';
  const samples = supplier.iqfSamples || 0;
  const lastMeasurement = supplier.lastIqfDate ? formatDate(supplier.lastIqfDate) : 'n/d';
  const expireDate = supplier.expire ? formatDate(supplier.expire) : null;
  const criteriaAtencao = [
    'Cumprimento de prazos conforme o pedido de compra ou contrato.',
    'Comunicação, garantia e suporte pós-venda.',
    'Qualidade do material ou serviço entregue.',
    'Conformidade com os itens descritos no pedido ou contrato.'
  ];
  const criteriaCriticos = [
    'Cumprimento de prazos conforme o pedido de compra ou contrato.',
    'Conformidade com os itens descritos no pedido ou contrato.',
    'Qualidade do material ou serviço entregue.',
    'Comunicação, garantia e suporte pós-venda.',
    'Embalagem e identificação do material.',
    'Cumprimento das normas de segurança.',
    'Envio de documentos obrigatórios (boleto, notas fiscais e certificados necessarios).'
  ];

  let message = '';
  if (iqf === null) {
    message =
      'Ainda nao há medições registradas para este fornecedor. Solicitar atualização do ciclo de avaliação.';
  } else if (iqf > 75) {
    message =
      'Agradecemos pela parceria e informamos que a empresa obteve excelente performance na avaliação mais recente.' +
      '\n\n📊 Nota IQF: ' +
      formattedIqf +
      '\n🏆 Classificação: Aprovado - desempenho de excelência.' +
      '\n\nEsse resultado reforça o comprometimento com qualidade, prazos e conformidade. Mantemos confiança na continuidade desta parceria sólida.';
  } else if (iqf >= 70) {
    const temas = pickRandom(criteriaAtencao, 2);
    message =
      'Compartilhamos o resultado da avaliação periódica de desempenho.' +
      '\n\n📊 Nota IQF: ' +
      formattedIqf +
      '\n⚠️ Classificação: Em atenção - desempenho no limite mínimo.' +
      '\n\nRecomendamos foco nos seguintes aspectos para evitar regressões:\n' +
      temas.map((tema) => '- ' + tema).join('\n') +
      '\n\nA manutenção dos indicadores é essencial para preservar a parceria.';
  } else {
    const temas = pickRandom(criteriaCriticos, 3);
    message =
      'Informamos que o fornecedor foi reprovado no Índice de Qualidade (IQF) e precisa de ações imediatas.' +
      '\n\n📊 Nota IQF: ' +
      formattedIqf +
      '\n❌ Classificação: Reprovado - abaixo do padrão mínimo (70,00).' +
      '\n\nFalhas observadas:\n' +
      temas.map((tema) => '- ' + tema).join('\n') +
      '\n\nSolicitamos analise interna das nao conformidades e plano corretivo. A reincidência pode impactar futuros fornecimentos.';
  }

  const occSummary = summarizeOccurrenceTexts(supplier.occurrences);
  if (occSummary) {
    message += '\n\n🔴 Ocorrências registradas no atendimento:\n' + occSummary;
  }

  message +=
    '\n\n📑 Indicadores complementares:\n' +
    '- Avaliações consideradas: ' +
    samples +
    '\n' +
    '- Ultima avaliação IQF: ' +
    lastMeasurement;
  if (expireDate) {
    message += '\n- Validade da homologacao: ' + expireDate;
  }

  message +=
    '\n\nLegenda de notas:\n' +
    '- 0 a 69: Reprovado - nao atingiu os critérios mínimos de qualidade e conformidade.\n' +
    '- A partir de 70: Aprovado - desempenho dentro dos parâmetros estabelecidos.';

  return message;
}

function summarizeOccurrenceTexts(occurrences) {
  if (!occurrences || !occurrences.length) {
    return '';
  }
  const unique = [];
  const seen = new Set();
  occurrences.forEach((occ) => {
    const text = safeString((occ && (occ.text || occ.occ)) || '');
    if (text && !seen.has(text)) {
      seen.add(text);
      unique.push(text);
    }
  });
  if (!unique.length) {
    return '';
  }
  const lines = unique.slice(0, 5).map((entry) => '• ' + entry);
  if (unique.length > lines.length) {
    lines.push('• ... outras ocorrências foram registradas no periodo.');
  }
  return lines.join('\n');
}

function generateSupplierFeedback(supplier, container) {
  if (!container) {
    return;
  }
  const baselineText = buildSupplierClassificationNarrative(supplier);
  const baselineHtml = formatFeedback(baselineText);
  renderFeedbackCard(container, {
    title: 'Analise executiva',
    subtitle: 'Resumo gerado com dados do IQF',
    icon: '??',
    bodyHtml: baselineHtml,
    hint: 'Gerando insights adicionais com IA...'
  });
  appendRefreshControl(container);
  state.lastFeedback = baselineText;
  state.lastFeedbackHtml = baselineHtml;
  state.lastEmailHtml = '';
  state.feedbackSupplierId = supplier.id;
  const apiKey = getOpenAiApiKey();
  if (!apiKey) {
    renderFeedbackCard(container, {
      title: 'Analise executiva',
      subtitle: 'Resumo gerado com dados do IQF',
      icon: '??',
      bodyHtml: baselineHtml,
      hint: 'Informe a chave da API OpenAI no painel de configuracoes (icone de engrenagem) para complementar a analise automaticamente.'
    });
    appendRefreshControl(container);
    showToast('Abra o painel de configuracoes (icone de engrenagem) e cole a chave da API OpenAI para habilitar os insights automaticos.');
    if (dom.apiKeyInput) {
      dom.apiKeyInput.focus();
    }
    return;
  }

  fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: 'Bearer ' + apiKey
    },
    body: JSON.stringify({
      model: 'gpt-4o-mini',
      temperature: 0.2,
      messages: [
        { role: 'system', content: 'Você é um especialista em qualificação de fornecedores.' },
        { role: 'user', content: buildSupplierPrompt(supplier) }
      ]
    })
  })
    .then((response) => {
      if (!response.ok) {
        return response.text().then((text) => {
          throw new Error(text || 'Falha ao consultar a API');
        });
      }
      return response.json();
    })
    .then((payload) => {
      if (state.feedbackSupplierId !== supplier.id) {
        return;
      }
      const answer = payload?.choices?.[0]?.message?.content?.trim();
      if (!answer) {
        throw new Error('Resposta vazia da API');
      }
      const combined = baselineText + '\n\n' + answer;
      const formatted = formatFeedback(combined);
      state.lastFeedback = combined;
      state.lastFeedbackHtml = formatted;
      state.lastEmailHtml = '';
      state.feedbackSupplierId = supplier.id;
      renderFeedbackCard(container, {
        title: 'Analise executiva',
        subtitle: 'Complementada com IA (GPT-4o mini)',
        icon: '🤖',
        bodyHtml: formatted
      });
      appendRefreshControl(container);
    })
    .catch((error) => {
      console.error('[analise:gpt]', error);
      state.lastFeedback = baselineText;
      state.lastFeedbackHtml = baselineHtml;
      state.lastEmailHtml = '';
      renderFeedbackCard(container, {
        title: 'Analise executiva',
        subtitle: 'Resumo gerado com dados do IQF',
        icon: '⚠️',
        bodyHtml: baselineHtml,
        hint: 'Nao foi possivel complementar a analise com IA agora. Tente novamente em instantes.'
      });
      appendRefreshControl(container);
      showToast('Erro ao consultar a API OpenAI. Tente novamente em instantes.');
    });
}

function buildSupplierPrompt(supplier) {
  const iqfLines = supplier.iqfHistory.length
    ? supplier.iqfHistory
        .map((entry) => '- ' + (entry.formattedDate || 'Data n/d') + ': IQF ' + entry.value)
        .join('\n')
    : '- sem histórico registrado';
  const occLines = supplier.occurrences.length
    ? supplier.occurrences
        .slice(0, 6)
        .map((occ) => '- ' + (occ.formattedDate || 'Data n/d') + ': ' + occ.text)
        .join('\n')
    : '- sem ocorrências registradas';

  const statusAlert =
    'Alerta de risco: ' +
    (supplier.status === 'Reprovado'
      ? 'Fornecedor reprovado - detalhe claramente todas as falhas e acione plano corretivo imediato.'
      : supplier.status === 'Pendente' || supplier.status === 'Em atencao'
      ? 'Fornecedor em atencao - destaque impactos potenciais e orientacoes preventivas.'
      : 'Fornecedor homologado - manter tom reconhecendo desempenho, mas reforce pontos de melhoria.');

  const situationalGuidance = [
    supplier.status === 'Reprovado'
      ? 'Em "Riscos e Pontos de Atencao", explicite as consequencias de reprovação e indique reforço de governança.'
      : supplier.status === 'Pendente' || supplier.status === 'Em atencao'
      ? 'Em "Riscos e Pontos de Atencao", sinalize urgência moderada e medidas preventivas para evitar escalonamento.'
      : 'Em "Pontos Fortes", reconheça o desempenho e proponha melhorias incrementais para manter a maturidade.'
  ].join(' ');

  return [
    'Voce atua como consultor senior de supply chain da Engeman. Produza uma analise executiva em portugues do Brasil com base apenas nos dados abaixo.',
    statusAlert,
    '',
    'Perfil do fornecedor:',
    '- Nome: ' + supplier.name,
    '- Codigo interno: ' + (supplier.code || 'n/d'),
    '- Status consolidado: ' + supplier.status + ' (status de origem: ' + (supplier.baseStatusRaw || 'n/d') + ')',
    '- Nota de homologacao vigente: ' + (supplier.homologScore !== null ? supplier.homologScore : 'n/d'),
    '- Media IQF consolidada: ' + (supplier.averageIqf !== null ? supplier.averageIqf : 'sem avaliacao'),
    '- Total de avaliacoes IQF: ' + supplier.iqfSamples,
    '- Ultima avaliacao IQF registrada: ' + (formatDate(supplier.lastIqfDate) || 'n/d'),
    '- Validade da homologacao: ' + (formatDate(supplier.expire) || 'n/d'),
    '',
    'Historico recente de IQF (ordem cronologica decrescente):',
    iqfLines,
    '',
    'Ocorrencias relevantes registradas:',
    occLines,
    '',
    'Estruture a resposta seguindo exatamente o formato abaixo, usando frases curtas com metricas numericas sempre que possivel:',
    'Visao Geral:',
    '- ...',
    'Pontos Fortes:',
    '- ...',
    'Riscos e Pontos de Atencao:',
    '- ...',
    'Recomendacoes Taticas:',
    '- ...',
    'Plano de Monitoramento (30 dias):',
    '- ...',
    'Observacoes Complementares:',
    '- ...',
    situationalGuidance,
    '',
    'Nao invente informacoes. Quando algum dado nao estiver disponivel, registre essa ausencia.'
  ].join('\n');
}

function applyInlineFormatting(text) {
  return escapeHtml(text).replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
}

function formatFeedback(text) {
  const lines = String(text || '').split(/\r?\n/);
  const html = [];
  let listBuffer = [];

  const flushList = () => {
    if (listBuffer.length) {
      html.push('<ul>' + listBuffer.join('') + '</ul>');
      listBuffer = [];
    }
  };

  lines.forEach((raw) => {
    const line = raw.trim();
    if (!line) {
      flushList();
      return;
    }
    if (/^(?:[-*\u2022\u0007])\s+/.test(line)) {
      const item = applyInlineFormatting(line.replace(/^(?:[-*\u2022\u0007])\s*/, '').trim());
      listBuffer.push('<li>' + item + '</li>');
      return;
    }
    flushList();
    if (/^[A-Za-z0-9].*:\s*$/.test(line)) {
      const heading = applyInlineFormatting(line.replace(/:$/, '').trim());
      html.push('<h4>' + heading + '</h4>');
      return;
    }
    html.push('<p>' + applyInlineFormatting(line) + '</p>');
  });
  flushList();
  return html.join('');
}
function renderFeedbackCard(container, config) {
  if (!container) {
    return;
  }
  const title = config?.title ? escapeHtml(config.title) : 'Analise executiva';
  const subtitle = config?.subtitle ? escapeHtml(config.subtitle) : '';
  const icon = escapeHtml(config?.icon || '🤖');
  const bodyHtml = config?.bodyHtml || '<p>Nenhum conteudo disponível.</p>';
  const hintText = config?.hint ? escapeHtml(config.hint) : null;
  container.innerHTML =
    '<div class="ai-card-header">' +
    '<span class="ai-chip" aria-hidden="true">' +
    icon +
    '</span>' +
    '<div class="ai-card-titles">' +
    '<p class="ai-card-title">' +
    title +
    '</p>' +
    (subtitle ? '<p class="ai-card-subtitle">' + subtitle + '</p>' : '') +
    '</div>' +
    '</div>' +
    '<div class="ai-card-body">' +
    bodyHtml +
    '</div>' +
    (hintText ? '<p class="ai-hint">' + hintText + '</p>' : '');
}

function appendRefreshControl(container) {
  if (!container || !container.__refreshButton) {
    return;
  }
  container.appendChild(container.__refreshButton);
}

function buildEmailNarrative(supplier) {
  const iqf = Number.isFinite(supplier.averageIqf) ? supplier.averageIqf : null;
  const parts = [];

  if (iqf === null) {
    parts.push(
      '<p>Ainda nao identificamos medições recentes de IQF para este fornecedor. Recomendamos atualizar a avaliação para manter o controle de desempenho.</p>'
    );
  } else if (iqf > 75) {
    parts.push(
      '<p>Agradecemos pela parceria e informamos que sua empresa obteve excelente performance na nossa avaliação mais recente.</p>',
      '<p>📊 Nota IQF: <strong>' + formatScoreValue(iqf) + '</strong><br>🏆 Classificação: Aprovado - Desempenho de excelência.</p>',
      '<p>Esse resultado demonstra alto nível de comprometimento com qualidade, prazos e conformidade. Seguimos confiantes na continuidade desta parceria sólida e eficiente.</p>',
      '<p>Qualquer dúvida, estamos à disposição.</p>'
    );
  } else if (iqf >= 70) {
    const criteriosFalhos = randomSample(
      [
        'Cumprimento de prazos conforme o pedido de compra ou contrato.',
        'Comunicação, garantia e suporte pós-venda.',
        'Qualidade do material ou serviço entregue.',
        'Conformidade com os itens descritos no pedido ou contrato.'
      ],
      2
    );
    parts.push(
      '<p>Compartilhamos abaixo o resultado da avaliação periódica de desempenho:</p>',
      '<p>📊 Nota IQF: <strong>' + formatScoreValue(iqf) + '</strong><br>⚠️ Nota mínima exigida: 70,00</p>',
      '<p>Embora tecnicamente aprovado, o resultado indica performance no limite mínimo aceitável. Recomendamos atenção especial aos seguintes aspectos:</p>',
      '<ul>' + criteriosFalhos.map((c) => '<li>' + escapeHtml(c) + '</li>').join('') + '</ul>',
      '<p>A manutenção de bons indicadores é essencial para seguirmos com uma parceria de confiança e excelência.</p>'
    );
  } else {
    const criteriosFalhos = randomSample(
      [
        'Cumprimento de prazos conforme o pedido de compra ou contrato.',
        'Conformidade com os itens descritos no pedido ou contrato.',
        'Qualidade do material ou serviço entregue.',
        'Comunicação, garantia e suporte pós-venda.',
        'Embalagem e identificação do material.',
        'Cumprimento das normas de segurança.',
        'Envio de documentos obrigatórios (boleto, NF-e, certificados).'
      ],
      3
    );
    parts.push(
      '<p>Informamos que, conforme a avaliação periódica de desempenho, sua empresa foi reprovada no Índice de Qualidade do Fornecedor (IQF):</p>',
      '<p>📊 Nota IQF: <strong>' + formatScoreValue(iqf) + '</strong><br>❌ Classificação: Reprovado - Abaixo do padrão mínimo (70,00)</p>',
      '<p>A reprovação ocorreu devido a falhas nos seguintes critérios:</p>',
      '<ul>' + criteriosFalhos.map((c) => '<li>' + escapeHtml(c) + '</li>').join('') + '</ul>',
      '<p>Solicitamos analise interna das nao conformidades e implementação de medidas corretivas. A reincidência pode impactar futuros fornecimentos.</p>'
    );
  }

  const occurrencesHtml = formatOccurrencesHtml(supplier.occurrences);
  if (occurrencesHtml) {
    parts.push(occurrencesHtml);
  }

  parts.push(
    '<h4>Legendas de notas</h4>',
    '<ul>' +
      '<li><strong>0 a 69: Reprovado</strong> - indica que o fornecedor nao atingiu os critérios mínimos de qualidade e conformidade.</li>' +
      '<li><strong>A partir de 70: Aprovado</strong> - indica desempenho satisfatório conforme os critérios estabelecidos.</li>' +
      '</ul>',
    '<p>Em caso de apontamentos negativos, pedimos a analise e correção. A reincidência de problemas pode suspender o fornecedor da Engeman.</p>',
    '<p>Seguimos confiantes na continuidade desta parceria sólida e eficiente.</p>',
    '<p>Atenciosamente,<br>Equipe de Suprimentos</p>'
  );

  return { html: parts.join('') };
}

function randomSample(list, count) {
  const pool = list.slice();
  const picks = [];
  const target = Math.min(count, pool.length);
  while (picks.length < target && pool.length) {
    const index = Math.floor(Math.random() * pool.length);
    picks.push(pool.splice(index, 1)[0]);
  }
  return picks;
}

function formatOccurrencesHtml(occurrences) {
  if (!occurrences || !occurrences.length) {
    return '';
  }
  const items = occurrences.slice(0, 5).map((occ) => {
    const dateLabel = occ.formattedDate || formatDate(occ.date) || 'Data nao informada';
    const text = escapeHtml(occ.text || occ.occ || '');
    return '<li><strong>' + escapeHtml(dateLabel) + ':</strong> ' + text + '</li>';
  });
  return '<p>🔴 <strong>Ocorrências registradas no atendimento:</strong></p><ul>' + items.join('') + '</ul>';
}

function buildEmailHtml(supplier, narrativeHtml, analysisHtml) {
  const supplierDisplayName = supplier.name && supplier.name.trim() ? supplier.name : null;
  const supplierName = escapeHtml(supplierDisplayName || 'N/D');
  const greetingName = escapeHtml(supplierDisplayName || (supplier.code ? 'Fornecedor ' + supplier.code : 'Parceiro Engeman'));
  const supplierCode = supplier.code ? escapeHtml(supplier.code) : 'N/D';
  const statusLabel = escapeHtml(supplier.status || 'Pendente');
  const statusBadgeClass = getStatusBadgeClass(supplier.status);
  const iqfScore = formatScoreValue(supplier.averageIqf);
  const homologScore = formatScoreValue(supplier.homologScore);
  const expire = formatDate(supplier.expire) || 'N/D';
  const lastIqf = formatDate(supplier.lastIqfDate) || 'N/D';
  const iqfSamples = supplier.iqfSamples || 0;
  const occurrenceCount = supplier.occurrences.length || 0;

  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Feedback IQF - ${supplierName}</title>
<style>
body{margin:0;background:#f3f4f6;font-family:"Segoe UI",Arial,sans-serif;color:#111827;}
.wrapper{padding:32px 16px;}
.content{max-width:640px;margin:0 auto;background:#ffffff;border-radius:20px;overflow:hidden;box-shadow:0 18px 42px rgba(15,23,42,0.12);}
.header{background:linear-gradient(135deg,#fb923c,#ef4444);color:#ffffff;padding:32px 28px 24px;}
.header h1{margin:0;font-size:24px;font-weight:700;}
.header p{margin:8px 0 0;font-size:14px;color:rgba(255,255,255,0.85);}
.card{padding:28px 28px 24px;border-top:1px solid #f1f5f9;}
.card:first-of-type{border-top:none;}
.card h2{margin:0 0 18px;font-size:18px;color:#0f172a;}
.card h4{margin:20px 0 10px;font-size:15px;color:#0f172a;}
.card p{margin:0 0 14px;line-height:1.6;color:#1f2937;}
.card ul{margin:0 0 14px;padding-left:20px;color:#1f2937;}
.summary-table{width:100%;border-collapse:collapse;}
.summary-table td{padding:14px;border:1px solid #e2e8f0;vertical-align:top;}
.summary-table span{display:block;font-size:11px;text-transform:uppercase;letter-spacing:0.04em;color:#6b7280;}
.summary-table strong{display:block;margin-top:6px;font-size:16px;color:#0f172a;}
.summary-table small{display:block;margin-top:4px;font-size:12px;color:#64748b;}
.status-badge{display:inline-block;padding:6px 12px;border-radius:999px;font-size:12px;font-weight:600;color:#ffffff;}
.badge-homologado{background:#059669;}
.badge-reprovado{background:#dc2626;}
.badge-pendente{background:#d97706;}
.footer{background:#f9fafb;padding:20px 28px;font-size:12px;color:#64748b;text-align:center;}
.analysis-card ul{margin-bottom:12px;}
</style>
</head>
<body>
<div class="wrapper">
  <div class="content">
    <div class="header">
      <h1>Engeman | Feedback IQF</h1>
      <p>Relatorio automatico do indice de qualidade do fornecedor</p>
    </div>
    <div class="card">
      <h2>Resumo do fornecedor</h2>
      <table class="summary-table" role="presentation" cellspacing="0" cellpadding="0">
        <tr>
          <td>
            <span>Fornecedor</span>
            <strong>${supplierName}</strong>
          </td>
          <td>
            <span>Codigo</span>
            <strong>${supplierCode}</strong>
          </td>
        </tr>
        <tr>
          <td>
            <span>Status consolidado</span>
            <span class="status-badge ${statusBadgeClass}">${statusLabel}</span>
          </td>
          <td>
            <span>Notas</span>
            <strong>IQF ${iqfScore}</strong>
            <small>Homologacao ${homologScore}</small>
          </td>
        </tr>
        <tr>
          <td>
            <span>Ultima avaliacao IQF</span>
            <strong>${lastIqf}</strong>
          </td>
          <td>
            <span>Validade da homologacao</span>
            <strong>${expire}</strong>
          </td>
        </tr>
        <tr>
          <td>
            <span>Amostras IQF</span>
            <strong>${iqfSamples}</strong>
          </td>
          <td>
            <span>Ocorrencias registradas</span>
            <strong>${occurrenceCount}</strong>
          </td>
        </tr>
      </table>
    </div>
    <div class="card">
      <h2>Mensagem oficial</h2>
      <p>Ola ${greetingName},</p>
      ${narrativeHtml}
    </div>
    <div class="card analysis-card">
      <h2>Insights IA</h2>
      ${analysisHtml}
    </div>
    <div class="footer">
      Engeman Suprimentos - Este e-mail foi gerado automaticamente. Para duvidas, contate o time de Suprimentos.
    </div>
  </div>
</div>
</body>
</html>`;
}

function getStatusBadgeClass(status) {
  if (status === 'Homologado') {
    return 'badge-homologado';
  }
  if (status === 'Reprovado') {
    return 'badge-reprovado';
  }
  return 'badge-pendente';
}

function formatScoreValue(value) {
  if (!Number.isFinite(value)) {
    return 'N/D';
  }
  return value.toFixed(2);
}

function pickRandom(list, count) {
  const source = Array.isArray(list) ? list.slice() : [];
  const picks = [];
  const limit = Math.min(count, source.length);
  for (let i = 0; i < limit; i += 1) {
    const index = Math.floor(Math.random() * source.length);
    picks.push(source.splice(index, 1)[0]);
  }
  return picks;
}

function htmlToPlainText(html) {
  const cleaned = String(html || '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/?(?:p|div|section|tr|h[1-6])>/gi, '\n')
    .replace(/<li>/gi, '- ')
    .replace(/<\/li>/gi, '\n')
    .replace(/<\/?[^>]+>/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  return decodeBasicEntities(cleaned);
}

function decodeBasicEntities(text) {
  const entities = {
    ' ': ' ',
    '&': '&',
    '<': '<',
    '>': '>',
    '"': '"',
    '&#39;': "'",
    'á': 'a',
    'à': 'a',
    'ã': 'a',
    'ç': 'c',
    'é': 'e',
    'ê': 'e',
    'í': 'i',
    'ó': 'o',
    'ô': 'o',
    'õ': 'o',
    'ú': 'u',
    '&uuml;': 'u'
  };
  return String(text || '').replace(/&[a-z#0-9]+;/gi, (entity) => entities[entity.toLowerCase()] || entity);
}


function tryCopyHtml(html) {
  if (!navigator.clipboard || typeof navigator.clipboard.writeText !== 'function') {
    return Promise.resolve(false);
  }
  return navigator.clipboard.writeText(html).then(
    () => true,
    () => false
  );
}

function previewEmailTemplate() {
  if (!state.lastEmailHtml) {
    showToast('Gere o feedback e informe o e-mail para criar o layout completo.');
    return;
  }
  const encoded = 'data:text/html;charset=utf-8,' + encodeURIComponent(state.lastEmailHtml);
  window.open(encoded, '_blank', 'noopener,noreferrer');
}

function handleUserSend() {
  if (!dom.userInput) {
    return;
  }
  const message = dom.userInput.value.trim();
  if (!message) {
    dom.userInput.placeholder = state.inputDefaultPlaceholder || dom.userInput.placeholder;
    return;
  }
  dom.userInput.value = '';
  dom.userInput.placeholder = state.inputDefaultPlaceholder || dom.userInput.placeholder;
  appendUserMessage(message);

  if (EMAIL_REGEX.test(message)) {
    sendSupplierEmail(message);
    return;
  }

  handleSupplierQuery(message);
}

function handleSupplierQuery(query) {
  const normalized = normalizeText(query);
  if (!normalized) {
    appendBotMessage('Informe o nome ou código do fornecedor que deseja visualizar.', true);
    return;
  }
  if (['menu', 'opcoes', 'opcao', 'inicio', 'voltar', 'principal'].includes(normalized)) {
    renderMainMenuOptions();
    showToast('Menu principal exibido novamente.');
    return;
  }
  const exact = state.suppliers.find(
    (supplier) => supplier.code === query || supplier.normalizedName === normalized
  );
  if (exact) {
    showSupplierDetails(exact.id);
    return;
  }
  const matches = state.suppliers
    .filter((supplier) => supplier.searchText.includes(query.toLowerCase()))
    .slice(0, MAX_SUPPLIER_SUGGESTIONS);
  if (!matches.length) {
    appendBotMessage(
      'Nenhum fornecedor encontrado para "' + escapeHtml(query) + '". Utilize o menu de agrupamentos por IQF para navegar.',
      true
    );
    return;
  }
  const content = createMessage('bot');
  const intro = document.createElement('p');
  intro.innerHTML = 'Fornecedores relacionados a <strong>' + escapeHtml(query) + '</strong>:';
  content.appendChild(intro);

  const actions = document.createElement('div');
  actions.className = 'message-actions';
  matches.forEach((supplier) => {
    const button = document.createElement('button');
    button.className = 'message-action-btn';
    button.textContent = supplier.name + (supplier.code ? ' (' + supplier.code + ')' : '');
    button.addEventListener('click', () => showSupplierDetails(supplier.id));
    actions.appendChild(button);
  });
  content.appendChild(actions);
}

function sendSupplierEmail(email) {
  if (!state.selectedSupplier) {
    appendBotMessage('Selecione um fornecedor antes de informar o e-mail para envio.', true);
    return;
  }
  if (!state.lastFeedbackHtml || state.feedbackSupplierId !== state.selectedSupplier.id) {
    appendBotMessage(
      'Gere o feedback com a IA antes de enviar o e-mail. Clique em "Reprocessar analise" caso precise atualizar o conteudo.',
      true
    );
    return;
  }
  if (!email || !EMAIL_REGEX.test(email)) {
    appendBotMessage('Informe um endereco de e-mail valido para prosseguir com o envio automatico.', true);
    return;
  }
  if (!EMAIL_API_ENDPOINT) {
    appendBotMessage(
      'O envio automatico de e-mail nao esta configurado. Defina window.EMAIL_API_ENDPOINT antes de utilizar esta funcionalidade.',
      true
    );
    return;
  }

  const supplier = state.selectedSupplier;
  const subjectTemplate = state.emailSubjectTemplate || 'Feedback IQF - {{fornecedor}}';
  const displayName = supplier.name && supplier.name.trim() ? supplier.name : supplier.code ? 'Fornecedor ' + supplier.code : 'Fornecedor';
  let subject = subjectTemplate.includes('{{fornecedor}}')
    ? subjectTemplate.replace(/{{fornecedor}}/gi, displayName)
    : subjectTemplate + ' - ' + displayName;

  const narrative = buildEmailNarrative(supplier);
  const analysisHtml = state.lastFeedbackHtml || '<p>Analise da IA indisponivel.</p>';
  const emailHtml = buildEmailHtml(supplier, narrative.html, analysisHtml);
  const plainTextBody = htmlToPlainText(emailHtml);
  state.lastEmailHtml = emailHtml;

  const payload = {
    to: email,
    subject,
    html: emailHtml,
    text: plainTextBody,
    supplier: {
      id: supplier.id,
      name: supplier.name,
      code: supplier.code,
      status: supplier.status,
      averageIqf: supplier.averageIqf,
      homologScore: supplier.homologScore
    },
    generatedAt: new Date().toISOString()
  };

  showToast('Enviando e-mail automatico...');
  tryCopyHtml(emailHtml).catch(() => {});

  const headers = { 'Content-Type': 'application/json' };
  if (EMAIL_API_TOKEN) {
    headers.Authorization = 'Bearer ' + EMAIL_API_TOKEN;
  }

  fetch(EMAIL_API_ENDPOINT, {
    method: 'POST',
    headers,
    body: JSON.stringify(payload)
  })
    .then((response) => {
      if (!response.ok) {
        return response.text().then((text) => {
          throw new Error(text || 'Falha ao enviar o e-mail.');
        });
      }
      return response.json().catch(() => ({}));
    })
    .then(() => {
      showToast('E-mail enviado com sucesso!');
      appendBotMessage(
        'E-mail enviado automaticamente para <strong>' +
          escapeHtml(email) +
          '</strong>. <button type="button" class="inline-link" onclick="previewEmailTemplate()">Ver layout completo</button>',
        true
      );
    })
    .catch((error) => {
      console.error('[analise:send-email]', error);
      showToast('Falha ao enviar o e-mail. Verifique as configuracoes.');
      appendBotMessage(
        'Nao foi possivel concluir o envio automatico. Verifique a configuracao do endpoint de e-mail e tente novamente.',
        true
      );
    });
}

function handleIndicadoresMensais() {
  if (!state.availableMonths || !state.availableMonths.length) {
    appendBotMessage('Nenhum dado mensal disponível no momento. Verifique se a planilha de qualidade possui registros validos.', true);
    return;
  }
  const months = state.availableMonths.slice(0, 18);
  const content = createMessage('bot');
  const intro = document.createElement('p');
  intro.innerHTML = 'Selecione um mês para visualizar a analise consolidada dos fornecedores:';
  content.appendChild(intro);

  const actions = document.createElement('div');
  actions.className = 'message-actions';
  months.forEach((monthKey) => {
    const button = document.createElement('button');
    button.className = 'message-action-btn';
    button.textContent = formatMonthLabel(monthKey);
    button.addEventListener('click', () => renderMonthlyAnalysis(monthKey));
    actions.appendChild(button);
  });
  content.appendChild(actions);

  if (state.availableMonths.length > months.length) {
    const hint = document.createElement('p');
    hint.className = 'message-hint';
    hint.textContent = 'Exibindo os 18 meses mais recentes. Utilize os arquivos de origem para consultar periodos anteriores.';
    content.appendChild(hint);
  }
}

function renderMonthlyAnalysis(monthKey) {
  const monthEntry = state.monthlySummary && typeof state.monthlySummary.get === 'function' ? state.monthlySummary.get(monthKey) : null;
  if (!monthEntry) {
    appendBotMessage('Dados nao encontrados para o periodo selecionado. Atualize os arquivos de origem e tente novamente.', true);
    return;
  }

  if (!monthEntry.suppliers || !monthEntry.suppliers.size) {
    appendBotMessage(
      'Nenhuma medição disponível para ' + escapeHtml(formatMonthLabel(monthKey)) + '. Atualize as planilhas e tente novamente.',
      true
    );
    return;
  }

  const monthLabel = formatMonthLabel(monthKey);
  const supplierSummaries = Array.from(monthEntry.suppliers.values()).map((entry) => ({
    key: entry.key,
    name: entry.name,
    code: entry.code,
    status: entry.status,
    sum: entry.sum,
    count: entry.count,
    avg: entry.count ? roundValue(entry.sum / entry.count) : null
  }));

  supplierSummaries.sort((a, b) => a.name.localeCompare(b.name || '', 'pt-BR', { sensitivity: 'accent' }));

  const totalSuppliers = supplierSummaries.length;
  const totalSamples = monthEntry.totalCount;
  const globalAverage = totalSamples ? roundValue(monthEntry.totalSum / totalSamples) : null;
  const reprovados = supplierSummaries
    .filter((item) => item.avg !== null && item.avg < 70)
    .sort((a, b) => (a.avg ?? 101) - (b.avg ?? 101));
  const emAtencao = supplierSummaries
    .filter((item) => item.avg !== null && item.avg >= 70 && item.avg <= 75)
    .sort((a, b) => (a.avg ?? 101) - (b.avg ?? 101));
  const aprovados = supplierSummaries.filter((item) => item.avg !== null && item.avg > 75);
  const excelencia = supplierSummaries
    .filter((item) => item.avg !== null && item.avg >= 85)
    .sort((a, b) => b.avg - a.avg)
    .slice(0, 5);

  const container = createMessage('bot');
  const summaryCard = document.createElement('div');
  summaryCard.className = 'monthly-summary';
  summaryCard.innerHTML =
    '<h3>Indicadores de ' +
    escapeHtml(monthLabel) +
    '</h3>' +
    '<p><strong>Média global IQF:</strong> ' +
    (globalAverage !== null ? formatScoreValue(globalAverage) : 'N/D') +
    '</p>' +
    '<p><strong>Avaliações consideradas:</strong> ' +
    totalSamples +
    ' • <strong>Fornecedores avaliados:</strong> ' +
    totalSuppliers +
    '</p>' +
    '<p><strong>Distribuição:</strong> Aprovados ' +
    aprovados.length +
    ' | Em atenção ' +
    emAtencao.length +
    ' | Reprovados ' +
    reprovados.length +
    '</p>';
  container.appendChild(summaryCard);

  const buildListSection = (title, items, emptyMessage, cssClass) => {
    const section = document.createElement('div');
    section.className = 'monthly-subsection ' + (cssClass || '');
    section.innerHTML = '<h4>' + title + '</h4>';
    if (!items.length) {
      const empty = document.createElement('p');
      empty.className = 'empty';
      empty.textContent = emptyMessage;
      section.appendChild(empty);
      return section;
    }
    const ul = document.createElement('ul');
    items.forEach((item) => {
      const li = document.createElement('li');
      const statusTag =
        item.status === 'Reprovado'
          ? '❌'
          : item.status === 'Homologado'
          ? '🏆'
          : item.status === 'Pendente'
          ? '⏳'
          : '⚙️';
      li.textContent =
        statusTag +
        ' ' +
        item.name +
        ' — média ' +
        (item.avg !== null ? formatScoreValue(item.avg) : 'N/D') +
        ' (' +
        item.count +
        ' avaliações)';
      ul.appendChild(li);
    });
    section.appendChild(ul);
    return section;
  };

  container.appendChild(
    buildListSection(
      'Fornecedores reprovados (IQF < 70)',
      reprovados.slice(0, 10),
      'Nenhum fornecedor reprovado neste periodo.',
      'monthly-subsection-critical'
    )
  );
  if (reprovados.length > 10) {
    const note = document.createElement('p');
    note.className = 'message-hint';
    note.textContent = 'Exibindo 10 de ' + reprovados.length + ' fornecedores reprovados.';
    container.appendChild(note);
  }

  container.appendChild(
    buildListSection(
      'Fornecedores em atenção (70 ≤ IQF ≤ 75)',
      emAtencao.slice(0, 10),
      'Nenhum fornecedor em estado de atenção neste periodo.',
      'monthly-subsection-warning'
    )
  );
  if (emAtencao.length > 10) {
    const note = document.createElement('p');
    note.className = 'message-hint';
    note.textContent = 'Exibindo 10 de ' + emAtencao.length + ' fornecedores em atenção.';
    container.appendChild(note);
  }

  container.appendChild(
    buildListSection(
      'Destaques de excelência (IQF ≥ 85)',
      excelencia,
      'Nenhum fornecedor atingiu o patamar de excelência neste mês.',
      'monthly-subsection-success'
    )
  );

  const aiCard = document.createElement('div');
  aiCard.className = 'feedback-card';
  renderFeedbackCard(aiCard, {
    title: 'Resumo estratégico com IA',
    subtitle: 'Gerado a partir do IQF mensal',
    icon: '🧠',
    bodyHtml: '<p>Gerando analise detalhada, aguarde alguns segundos...</p>',
    hint: 'Aguarde enquanto consultamos o modelo de IA.'
  });
  container.appendChild(aiCard);

  generateMonthlyNarrative(monthKey, monthEntry, supplierSummaries, aiCard);
}



function generateMonthlyNarrative(monthKey, monthEntry, supplierSummaries, cardNode) {
  if (!cardNode) {
    return;
  }
  const apiKey = getOpenAiApiKey();
  if (!apiKey) {
    renderFeedbackCard(cardNode, {
      title: 'Resumo estratégico com IA',
      subtitle: 'Configure a chave da API para continuar',
      icon: '🔒',
      bodyHtml: '<p>Informe a chave da API OpenAI no painel de configuracoes (icone de engrenagem) para gerar a analise mensal automaticamente.</p>'
    });
    showToast('Abra o painel de configuracoes (icone de engrenagem) e cole a chave da API OpenAI para gerar o resumo mensal.');
    if (dom.apiKeyInput) {
      dom.apiKeyInput.focus();
    }
    return;
  }
  const prompt = buildMonthlyPrompt(monthKey, monthEntry, supplierSummaries);
  fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: 'Bearer ' + apiKey
    },
    body: JSON.stringify({
      model: 'gpt-4o-mini',
      temperature: 0.2,
      messages: [
        { role: 'system', content: 'Você é um especialista em performance de fornecedores.' },
        { role: 'user', content: prompt }
      ]
    })
  })
    .then((response) => {
      if (!response.ok) {
        return response.text().then((text) => {
          throw new Error(text || 'Falha ao consultar a API');
        });
      }
      return response.json();
    })
    .then((payload) => {
      const answer = payload?.choices?.[0]?.message?.content?.trim();
      if (!answer) {
        throw new Error('Resposta vazia da API');
      }
      renderFeedbackCard(cardNode, {
        title: 'Resumo estratégico com IA',
        subtitle: 'Complementado com IA (GPT-4o mini)',
        icon: '🤖',
        bodyHtml: formatFeedback(answer)
      });
    })
    .catch((error) => {
      console.error('[analise:monthly-gpt]', error);
      renderFeedbackCard(cardNode, {
        title: 'Resumo estratégico com IA',
        subtitle: 'Nao foi possivel atualizar agora',
        icon: '⚠️',
        bodyHtml: '<p>Nao foi possivel gerar a analise neste momento. Tente novamente em instantes ou verifique a chave da API.</p>'
      });
    });
}

function buildMonthlyPrompt(monthKey, monthEntry, supplierSummaries) {
  const monthLabel = formatMonthLabel(monthKey);
  const totalEvaluations = monthEntry.totalCount;
  const globalAverage = totalEvaluations ? roundValue(monthEntry.totalSum / totalEvaluations) : null;
  const reprovados = supplierSummaries.filter((item) => item.avg !== null && item.avg < 70);
  const emAtencao = supplierSummaries.filter((item) => item.avg !== null && item.avg >= 70 && item.avg <= 75);
  const aprovados = supplierSummaries.filter((item) => item.avg !== null && item.avg > 75);
  const excelencia = supplierSummaries
    .filter((item) => item.avg !== null && item.avg >= 85)
    .sort((a, b) => b.avg - a.avg)
    .slice(0, 10);

  const reprovLines = reprovados.length
    ? reprovados
        .slice(0, 20)
        .map((item) => '- ' + item.name + ' | média ' + formatScoreValue(item.avg) + ' | avaliações ' + item.count)
        .join('\n')
    : '- Nenhum fornecedor reprovado no periodo.';

  const atencaoLines = emAtencao.length
    ? emAtencao
        .slice(0, 20)
        .map((item) => '- ' + item.name + ' | média ' + formatScoreValue(item.avg) + ' | avaliações ' + item.count)
        .join('\n')
    : '- Nenhum fornecedor em atenção.';

  const excelenciaLines = excelencia.length
    ? excelencia
        .map((item) => '- ' + item.name + ' | média ' + formatScoreValue(item.avg) + ' | avaliações ' + item.count)
        .join('\n')
    : '- Nenhum destaque de excelência no periodo.';

  const baseSnapshot = supplierSummaries.slice(0, 40).map((item) => {
    if (item.avg === null) {
      return '- ' + item.name + ': sem medições registradas';
    }
    return '- ' + item.name + ': média ' + formatScoreValue(item.avg) + ' (' + item.count + ' avaliações)';
  });
  if (supplierSummaries.length > baseSnapshot.length) {
    baseSnapshot.push('- ... lista truncada para manter o prompt enxuto.');
  }

  return [
    'Analise mensal de ' + monthLabel + ' considerando ' + totalEvaluations + ' avaliações -válidas.',
    'IQF médio global: ' + (globalAverage !== null ? formatScoreValue(globalAverage) : 'N/D') + '.',
    'Distribuição por status: aprovados ' + aprovados.length + ', atenção ' + emAtencao.length + ', reprovados ' + reprovados.length + '.',
    'Destaques de excelência (>=85):\n' + excelenciaLines,
    'Fornecedores reprovados (média < 70):\n' + reprovLines,
    'Fornecedores em atenção (70 a 75):\n' + atencaoLines,
    'Panorama resumido por fornecedor:\n' + baseSnapshot.join('\n'),
    'Estruture a resposta exatamente nas seções: Visão Geral, Pontos de Atenção, Fornecedores Reprovados, Ações Imediatas, Conclusão, Alertas Prioritários.',
    'Na seção "Ações Imediatas" utilize exatamente o texto: "Envio de notificação via e-mail aos fornecedores reprovados no IQF mensal e, em caso de reincidência, abertura de RAC para analise, tratativas e possível suspensão do fornecedor."',
    'Se alguma seção nao tiver informações relevantes, escreva "Nenhum registro" de forma direta.',
    'Em "Fornecedores Reprovados" detalhe impactos e solicitação de plano de ação.',
    'Na seção "Alertas Prioritários" produza até três bullets iniciando com negrito destacando riscos críticos ou oportunidades urgentes.',
    'Use tom executivo, direto e orientado à decisão, evitando repetir dados literalmente sem interpretação. Deixando tudo organizado e bem feito.'
  ].join('\n');
}


function handleContactBase() {
  const message = [
      '<p><strong>Responsáveis pelas Bases:</strong></p>',
      '<p><strong>1️⃣ FILIAL MACAÉ:</strong></p>',
      '<p><strong>Compradores:</strong></p>',
      '<p>Naiara Galdino | Pryscila.</p>',
      '<p><strong>2️⃣ FILIAL PERNAMBUCO: </strong></p>',
      '<p><strong>Compradores:</strong></p>',
      '<p>Suzan Bezerril.</p>',
      '<p><strong>3️⃣ FILIAL PARACURU:</strong></p>',
      '<p><strong>Compradores:</strong></p>',
      '<p>Iran Victor.</p>',
      '<p><strong>4️⃣ FILIAL FAFEN ( BA E SE ):</strong></p>',
      '<p><strong>Compradores:</strong></p>',
      '<p>Jennyfer.</p>',
      '<p><strong>5️⃣ FILIAL SÃO PAULO:</strong></p>',
      '<p><strong>Compradores:</strong></p>',
      '<p>Gilberto Trajano.</p>'

  ].join('');
   appendBotMessage(message, true);
}

function handleProcedimento() {
  const message = [
    '<p><strong>🧠 Procedimento Engeman para Avaliação de Fornecedores</strong></p>',
    '<p><strong>1️⃣ PG.SM.01 - Aquisição:</strong> Define o fluxo de compra, cotação, analise e aprovação dos materiais e serviços.</p>',
    '<p><strong>2️⃣ PG.SM.02 - Avaliação de Fornecedores:</strong> Critérios de desempenho (IQF), RACs e homologações.</p>',
    '<p><strong>3️⃣ PG.SM.03 - Almoxarifado:</strong> Inspeções, controle e tratamento de nao conformidades.</p>',
    '<p><strong>🔎 Observações Importantes</strong></p>',
    '<p><strong>- As analises são feitas mensalmente com base nas ocorrências.</strong></p>',
    '<p><strong>- A nota IQF varia de 0 a 100.</strong></p>',
    '<p><strong>- Fornecedores com IQF abaixo de 70 são REPROVADOS.</strong></p>',
    '<p><strong>- Objetivo: garantir o padrão Engeman e um relacionamento com parceiros confiáveis.</strong></p>',
    '<p><strong>📌 Dúvidas? Consulte o Setor de Suprimentos.</strong></p>'
  ].join('');
  appendBotMessage(message, true);
}


function openSettings() {
  if (dom.settingsOverlay) {
    dom.settingsOverlay.classList.add('open');
    dom.settingsOverlay.setAttribute('aria-hidden', 'false');
  }
}

function closeSettings() {
  if (dom.settingsOverlay) {
    dom.settingsOverlay.classList.remove('open');
    dom.settingsOverlay.setAttribute('aria-hidden', 'true');
  }
}

function saveSettings() {
  if (dom.settingsEmailSubject) {
    state.emailSubjectTemplate = dom.settingsEmailSubject.value.trim() || state.emailSubjectTemplate;
    localStorage.setItem('analise-email-subject', state.emailSubjectTemplate);
  }
  showToast('Configurações salvas.');
  closeSettings();
}

function clearSettings() {
  state.emailSubjectTemplate = 'Feedback IQF - {{fornecedor}}';
  if (dom.settingsEmailSubject) {
    dom.settingsEmailSubject.value = state.emailSubjectTemplate;
  }
  localStorage.removeItem('analise-email-subject');
  showToast('Configurações removidas.');
}

function createMessage(type) {
  const wrapper = document.createElement('div');
  wrapper.className = type === 'user' ? 'message user-message' : 'message bot-message';
  const content = document.createElement('div');
  content.className = 'message-content';
  wrapper.appendChild(content);
  if (dom.messages) {
    dom.messages.appendChild(wrapper);
    dom.messages.scrollTop = dom.messages.scrollHeight;
  }
  return content;
}

function appendBotMessage(text, allowHtml) {
  const content = createMessage('bot');
  if (allowHtml) {
    content.innerHTML = text;
  } else {
    content.textContent = text;
  }
  return content;
}

function appendUserMessage(text) {
  const content = createMessage('user');
  content.textContent = text;
  return content;
}

function updateMessage(node, text, allowHtml) {
  if (!node) {
    return;
  }
  if (allowHtml) {
    node.innerHTML = text;
  } else {
    node.textContent = text;
  }
}

function renderStatusBadge(status) {
  const normalized = (status || '').toLowerCase();
  let badgeClass = 'status-em-atencao';
  if (normalized === 'homologado') {
    badgeClass = 'status-aprovado';
  } else if (normalized === 'reprovado') {
    badgeClass = 'status-reprovado';
  }
  return '<span class="status-pill ' + badgeClass + '">' + escapeHtml(status || 'Pendente') + '</span>';
}

function showToast(message) {
  if (!dom.toast) {
    return;
  }
  dom.toast.innerHTML = message;
  dom.toast.classList.add('show');
  clearTimeout(showToast.timeout);
  showToast.timeout = setTimeout(() => {
    dom.toast?.classList.remove('show');
  }, 4200);
}

function buildSupplierId(code, normalizedName, usedIds) {
  if (code) {
    const id = 'code-' + String(code).replace(/[^a-z0-9]/gi, '');
    if (!usedIds.has(id)) {
      return id;
    }
  }
  if (normalizedName) {
    const id = 'name-' + normalizedName.replace(/[^a-z0-9]/gi, '-');
    if (!usedIds.has(id)) {
      return id;
    }
  }
  let id;
  do {
    incrementalId += 1;
    id = 'supplier-' + incrementalId;
  } while (usedIds.has(id));
  return id;
}

function safeDomId(prefix, key) {
  return prefix + '-' + String(key || '').replace(/[^a-z0-9_-]/gi, '');
}

function mapHomolog(row) {
  const normalized = normalizeKeys(row);
  const code = toId(normalized.codigo || normalized.codagente || normalized.codfornecedor || normalized.id);
  const name = safeString(normalized.nomefantasia || normalized.agente || normalized.fornecedor || normalized.nome);
  if (!code && !name) {
    return null;
  }
  return {
    code,
    name,
    score: toNumber(normalized.notahomologacao || normalized.nota),
    status: mapStatus(normalized.aprovado || normalized.status || normalized.situacao || normalized.qualifica),
    expire: toISODate(normalized.datavencimento || normalized.validade || normalized.data)
  };
}

function mapIqf(row) {
  const normalized = normalizeKeys(row);
  return {
    code: toId(normalized.codagente || normalized.codigo || normalized.codfornecedor || normalized.id),
    name: safeString(normalized.nomeagente || normalized.fornecedor || normalized.nome),
    iqf: toNumber(normalized.nota || normalized.notaiqf || normalized.iqf),
    date: toISODate(normalized.data || normalized.datamedicao || normalized.datageracao),
    document: safeString(normalized.documento || normalized.numerodocumento || normalized.doc),
    occ: safeString(normalized.observacao || normalized.ocorrencia || normalized.comentario || normalized.notaocorrencia)
  };
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
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s]/gi, '')
    .trim()
    .toLowerCase();
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

function toNumber(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }
  const cleaned = String(value).trim().replace(/\s/g, '').replace(/\.(?=\d{3}\b)/g, '').replace(',', '.');
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
    const excelEpoch = Date.parse('1899-12-30');
    const millis = excelEpoch + value * 86400000;
    return new Date(millis).toISOString().slice(0, 10);
  }
  const match = String(value).trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (match) {
    const day = match[1].padStart(2, '0');
    const month = match[2].padStart(2, '0');
    const year = match[3].length === 2 ? '20' + match[3] : match[3];
    return year + '-' + month + '-' + day;
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return value;
  }
  return null;
}

function roundValue(value) {
  return Number.isFinite(value) ? Number(value.toFixed(2)) : value;
}

function formatMonthLabel(key) {
  if (!key) {
    return '--';
  }
  const parts = String(key).split('-');
  if (parts.length !== 2) {
    return key;
  }
  return parts[1] + '/' + parts[0];
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

function deriveStatus(baseStatus, iqfScore, homologScore) {
  const belowThreshold = (score) => score !== null && score < 70;
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

function escapeHtml(value) {
  return String(value || '').replace(/[&<>"']/g, (match) =>
    ({
      '&': '&',
      '<': '<',
      '>': '>',
      '"': '"',
      "'": '&#39;'
    }[match])
  );
}

/* global XLSX, BracketEngine */
'use strict';

const STORAGE_KEY = 'bracketflow_tournament_state_v7';
const BENCH_COUNT = 16;

const FILL_COLORS = [
  { key: 'yellow', label: '黄',       hex: '#fef08a', border: '#eab308' },
  { key: 'sky',    label: '水色',     hex: '#bae6fd', border: '#0284c7' },
  { key: 'green',  label: '緑',       hex: '#bbf7d0', border: '#16a34a' },
  { key: 'pink',   label: 'ピンク',   hex: '#fbcfe8', border: '#db2777' },
  { key: 'orange', label: 'オレンジ', hex: '#fed7aa', border: '#ea580c' },
  { key: 'purple', label: '紫',       hex: '#e9d5ff', border: '#9333ea' },
  { key: 'gray',   label: 'グレー',   hex: '#e2e8f0', border: '#64748b' },
  { key: 'none',   label: '解除',     hex: '#ffffff', border: '#94a3b8' }
];

function createEmptyBench(size = BENCH_COUNT) {
  return new Array(Math.max(BENCH_COUNT, size)).fill(null);
}

function normalizeBench(arr) {
  const len = Array.isArray(arr) ? Math.max(BENCH_COUNT, arr.length) : BENCH_COUNT;
  const res = new Array(len).fill(null);
  if (Array.isArray(arr)) {
    for (let i = 0; i < arr.length; i++) {
      res[i] = arr[i] || null;
    }
  }
  return res;
}

function addBenchSlots(count = 4) {
  for (let i = 0; i < count; i++) state.bench.push(null);
  saveToLocalStorage();
  renderBench();
  const container = document.getElementById('bench-slots-container');
  if (container) container.scrollTop = container.scrollHeight;
}

const RULE_DEFINITIONS = [
  {
    key: 'enforceTop4Pattern',
    label: '1〜4番手を4山＆in/out規則性で高優先枠へ配置',
    desc: '1番手基準（A-B-C-D等 × in-out-out-in / out-in-in-out、同一校複数シード特例含む）で配置します。'
  },
  {
    key: 'avoidTop4UnderSeed',
    label: '各校の1〜4番手を1〜8シードの下に入れない',
    desc: '1〜4番手は1〜8シードに1〜2回戦で当たる枠を避け、他の高優先枠へ優先配置します。'
  },
  {
    key: 'lowRankFromBottomSlots',
    label: '5番手以下を最下位番手から1〜8シードの下へ配置',
    desc: '番手が低いペアから順に、1〜8シードと1〜2回戦で当たる枠（S1下→S8下）→低優先枠へ配置します。'
  },
  {
    key: 'balanceLowerMountainAndSide',
    label: '5番手以降も全体でA〜D・in/outのバランスを保つ',
    desc: '1〜4番手と合わせた学校全体で、4つの山（A〜D）およびin/outの人数差が1以内になるよう調整します。'
  },
  {
    key: 'spreadUnder8To8Blocks',
    label: '8ペア以下の学校は8つの山（8ブロック）へ分散',
    desc: '出場組数が8組以下の学校は、A1〜B1の8ブロックに重複なく散らばるように配置します。'
  },
  {
    key: 'avoidRound1SameSchool',
    label: '1回戦の同校対決を最優先回避（8ペア以上含む）',
    desc: '隣接ペア枠（1回戦）で同じ学校が当たらないよう最優先で制御します。'
  },
  {
    key: 'avoidTripleSameCard',
    label: '1回戦の同一カード3対戦以上を回避',
    desc: '同じ学校同士の対戦が1回戦に3組以上重ならないよう分散します。'
  },
  {
    key: 'avoidRound2SameSchool',
    label: '2回戦（4枠内）の同校接近を回避',
    desc: '同じ4枠ブロック内での早期同校対決を強く抑止します。'
  }
];

const state = {
  mainView: 'registration', // 'registration' | 'tournament'
  regSubTab: 'alias',       // 'alias' | 'rules' | 'consolation'
  title: '大会トーナメント表',
  matchFormat: 7,           // 5 (3G先取) | 7 (4G先取)
  entries: [],
  schoolOrder: [],          // 学校優先順位配列（先頭が1位）
  selectedSchoolIdx: null,  // 学校優先順位のクリック入替用選択インデックス
  editingEntryId: null,     // 登録名簿テーブルでインライン編集中のペアID
  schoolAliases: {},
  rules: Object.assign({}, BracketEngine.DEFAULT_RULES),
  bracketSize: 16,
  slots: [],
  activeSlotMask: [],
  bench: createEmptyBench(),
  slotFills: {},            // { [slotIdx]: 'yellow' | 'sky' | ... }
  activePaintColor: 'yellow',
  printShowFills: true,     // 印刷時に塗りつぶしを表示するか
  results: {},
  consolation: {
    preset: 'none',
    title: '順位決定戦（代表決定戦）',
    matchCount: 2,
    slots: ['', '', '', '', '', '', '', ''],
    results: {}
  },
  oop: (typeof OopEngine !== 'undefined') ? OopEngine.createDefaultOopState() : null,
  appMode: 'edit',          // 'edit' | 'paint' | 'match'
  layoutMode: 'dual',       // 'dual' | 'single'
  bracketBlockFilter: 'ALL', // 'ALL' | 'A' | 'B' | 'C' | 'D'
  printMode: 'distribution',
  printOrientation: 'portrait',
  showRankBadge: true,
  showSlotPriority: false,
  sidebarCollapsed: false,
  benchCollapsed: false,
  zoom: 1.0,
  selectedTarget: null,
  searchQuery: '',
  searchMatches: [],
  searchCursor: 0,
  alerts: {
    round1SameSchool: [],
    tripleSameCards: [],
    top4Violations: [],
    round2SameSchool: []
  },
  history: [],
  future: [],
  activeModalMatch: null
};

// =========================================================================
// トースト通知 & localStorage 自動保存
// =========================================================================
function showToast(message, type = 'info', durationMs = 2800) {
  const container = document.getElementById('toast-container');
  if (!container) return;
  const colors = {
    info: 'bg-slate-900 text-white border-slate-700',
    success: 'bg-emerald-700 text-white border-emerald-600',
    warn: 'bg-amber-600 text-white border-amber-500',
    error: 'bg-rose-700 text-white border-rose-600'
  };
  const el = document.createElement('div');
  el.className = `toast-item px-3.5 py-2 rounded-lg border shadow-lg text-xs font-bold max-w-sm ${colors[type] || colors.info}`;
  el.textContent = message;
  container.appendChild(el);
  setTimeout(() => {
    el.style.opacity = '0';
    el.style.transition = 'opacity 0.25s ease';
    setTimeout(() => el.remove(), 260);
  }, durationMs);
}

function saveToLocalStorage() {
  try {
    const payload = {
      version: '5.1',
      title: state.title,
      matchFormat: state.matchFormat,
      entries: state.entries,
      schoolOrder: state.schoolOrder,
      schoolAliases: state.schoolAliases,
      rules: state.rules,
      bracketSize: state.bracketSize,
      slots: state.slots,
      activeSlotMask: state.activeSlotMask,
      bench: state.bench,
      slotFills: state.slotFills,
      printShowFills: state.printShowFills,
      results: state.results,
      consolation: state.consolation,
      oop: state.oop,
      layoutMode: state.layoutMode,
      printMode: state.printMode,
      printOrientation: state.printOrientation
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    if (typeof window.onAppStateSaved === 'function') {
      window.onAppStateSaved(payload);
    }
  } catch (_) {
    // 無視
  }
}

function loadFromLocalStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return false;
    const data = JSON.parse(raw);
    if (!data || !Array.isArray(data.entries)) return false;

    state.title = data.title || state.title;
    state.matchFormat = (data.matchFormat === 5 || data.matchFormat === 7) ? data.matchFormat : 7;
    state.entries = data.entries || [];
    state.schoolOrder = Array.isArray(data.schoolOrder) ? data.schoolOrder : [];
    state.schoolAliases = data.schoolAliases || {};
    state.rules = Object.assign({}, BracketEngine.DEFAULT_RULES, data.rules || {});
    state.bracketSize = data.bracketSize || 16;
    state.slots = Array.isArray(data.slots) ? data.slots : [];
    state.activeSlotMask = Array.isArray(data.activeSlotMask) ? data.activeSlotMask : [];
    state.bench = normalizeBench(data.bench);
    state.slotFills = (data.slotFills && typeof data.slotFills === 'object') ? data.slotFills : {};
    state.printShowFills = (typeof data.printShowFills === 'boolean') ? data.printShowFills : true;
    state.results = data.results || {};
    if (data.consolation) state.consolation = data.consolation;
    if (typeof OopEngine !== 'undefined') {
      state.oop = OopEngine.createDefaultOopState(data.oop);
    }
    if (data.layoutMode) state.layoutMode = data.layoutMode;
    if (data.printMode) state.printMode = data.printMode;
    if (data.printOrientation) state.printOrientation = data.printOrientation;
    return true;
  } catch (_) {
    return false;
  }
}

// =========================================================================
// 学校略称・確認用表記（例：星見一１、風見六３）ユーティリティ
// =========================================================================
function toFullWidthDigits(num) {
  return String(num ?? '').replace(/[0-9]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) + 0xFEE0));
}

function computeDefaultSchoolShortName(rawSchool) {
  if (!rawSchool) return '';
  let s = String(rawSchool).trim();
  s = s.replace(/市立|町立|村立|県立|都立|府立|道立|国立|私立/g, '');
  s = s.replace(/附属中学校|付属中学校|中等教育学校|高等学校|中学校|小学校|高校|中学|学園|学院/g, '');
  s = s.replace(/第([一二三四五六七八九十廿百]+)/g, '$1');
  return s.trim() || String(rawSchool).trim();
}

function formatSchoolShortName(rawSchool) {
  if (!rawSchool) return '';
  const key = String(rawSchool).trim();
  if (Object.prototype.hasOwnProperty.call(state.schoolAliases, key)) {
    const custom = String(state.schoolAliases[key] ?? '').trim();
    if (custom !== '') return custom;
  }
  return computeDefaultSchoolShortName(key);
}

/**
 * 確認用バッジ表記：学校略称＋全角番手数字（例：星見二１、風見六３）
 */
function formatSchoolRankBadge(entry) {
  if (!entry) return '';
  return `${formatSchoolShortName(entry.school)}${toFullWidthDigits(entry.schoolRank)}`;
}

function updateSchoolAlias(rawSchool, newAlias) {
  const key = String(rawSchool).trim();
  const trimmed = String(newAlias ?? '').trim();
  const autoVal = computeDefaultSchoolShortName(key);
  if (trimmed === '' || trimmed === autoVal) {
    delete state.schoolAliases[key];
  } else {
    state.schoolAliases[key] = trimmed;
  }
  saveToLocalStorage();
  renderRegistrationTable();
  if (state.mainView === 'tournament') {
    refreshAlerts(false);
    renderBench();
    renderTourneySidebar();
    renderSheet();
  }
}

function resetAllSchoolAliases() {
  state.schoolAliases = {};
  saveToLocalStorage();
  renderSchoolAliasEditor();
  renderRegistrationTable();
  if (state.mainView === 'tournament') {
    refreshAlerts(false);
    renderBench();
    renderTourneySidebar();
    renderSheet();
  }
  showToast('学校略称を自動変換にリセットしました');
}

function getWinGames() {
  return state.matchFormat === 5 ? 3 : 4;
}

function setMatchFormat(val) {
  const num = parseInt(val, 10);
  state.matchFormat = (num === 5) ? 5 : 7;
  const sel = document.getElementById('match-format-select');
  if (sel) sel.value = String(state.matchFormat);
  const lbl = document.getElementById('match-mode-format-label');
  if (lbl) lbl.textContent = `${state.matchFormat}ゲームマッチ`;
  saveToLocalStorage();
  if (state.mainView === 'tournament') renderSheet();
  showToast(`試合形式を ${state.matchFormat}ゲームマッチ（${getWinGames()}G先取）に設定しました`);
}

function syncSchoolOrder() {
  const { ordered } = BracketEngine.resolveSchoolPriorityMap(state.entries, state.schoolOrder);
  state.schoolOrder = ordered;
}

function swapSchoolPriority(idxA, idxB) {
  syncSchoolOrder();
  if (idxA === idxB || idxA < 0 || idxB < 0 || idxA >= state.schoolOrder.length || idxB >= state.schoolOrder.length) return;
  const tmp = state.schoolOrder[idxA];
  state.schoolOrder[idxA] = state.schoolOrder[idxB];
  state.schoolOrder[idxB] = tmp;
  state.selectedSchoolIdx = null;
  saveToLocalStorage();
  renderSchoolAliasEditor();
}

function moveSchoolPriority(index, delta) {
  swapSchoolPriority(index, index + delta);
}

let schoolDragIdx = null;
function handleSchoolDragStart(e, idx) {
  if (e.target && (e.target.tagName === 'INPUT' || e.target.tagName === 'BUTTON')) {
    e.preventDefault();
    return;
  }
  schoolDragIdx = idx;
  e.dataTransfer.effectAllowed = 'move';
  try { e.dataTransfer.setData('text/plain', String(idx)); } catch (_) {}
}
function handleSchoolDragOver(e) {
  e.preventDefault();
  e.currentTarget.classList.add('drag-over-target');
}
function handleSchoolDragLeave(e) {
  e.currentTarget.classList.remove('drag-over-target');
}
function handleSchoolDrop(e, targetIdx) {
  e.preventDefault();
  e.currentTarget.classList.remove('drag-over-target');
  if (schoolDragIdx === null || schoolDragIdx === targetIdx) {
    schoolDragIdx = null;
    return;
  }
  swapSchoolPriority(schoolDragIdx, targetIdx);
  schoolDragIdx = null;
}

function handleSchoolRowClick(e, idx) {
  if (e.target && (e.target.tagName === 'INPUT' || e.target.tagName === 'BUTTON')) return;
  if (state.selectedSchoolIdx === null) {
    state.selectedSchoolIdx = idx;
    renderSchoolAliasEditor();
    return;
  }
  if (state.selectedSchoolIdx === idx) {
    state.selectedSchoolIdx = null;
    renderSchoolAliasEditor();
    return;
  }
  swapSchoolPriority(state.selectedSchoolIdx, idx);
}

function shuffleSchoolOrder(notify = true) {
  syncSchoolOrder();
  if (state.schoolOrder.length <= 1) return;
  state.schoolOrder = BracketEngine.shuffleArray(state.schoolOrder);
  state.selectedSchoolIdx = null;
  saveToLocalStorage();
  renderSchoolAliasEditor();
  if (notify) showToast('学校の優先順位をランダム抽選しました', 'success');
}

function sortSchoolOrderDefault(notify = true) {
  state.schoolOrder = [];
  state.selectedSchoolIdx = null;
  syncSchoolOrder();
  saveToLocalStorage();
  renderSchoolAliasEditor();
  if (notify) showToast('学校優先順位をシード・組数順に整列しました');
}

function renderSchoolAliasEditor() {
  const container = document.getElementById('school-alias-editor-list');
  if (!container) return;

  syncSchoolOrder();
  const schools = state.schoolOrder;
  if (schools.length === 0) {
    container.innerHTML = `<div class="p-4 text-center text-slate-400 text-[11px]">
      生徒を登録すると、学校ごとの優先順位・略称設定欄が表示されます。
    </div>`;
    return;
  }

  const schoolCounts = {};
  state.entries.forEach((e) => {
    const sc = e.school || '無所属';
    schoolCounts[sc] = (schoolCounts[sc] || 0) + 1;
  });

  container.innerHTML = schools.map((sc, idx) => {
    const currentShort = formatSchoolShortName(sc);
    const isCustom = Object.prototype.hasOwnProperty.call(state.schoolAliases, sc);
    const isSelected = state.selectedSchoolIdx === idx;
    const cnt = schoolCounts[sc] || 0;
    return `<div draggable="true"
                 onclick="handleSchoolRowClick(event, ${idx})"
                 ondragstart="handleSchoolDragStart(event, ${idx})"
                 ondragover="handleSchoolDragOver(event)"
                 ondragleave="handleSchoolDragLeave(event)"
                 ondrop="handleSchoolDrop(event, ${idx})"
                 class="px-2 py-1.5 flex items-center justify-between gap-1.5 bg-white hover:bg-slate-50 cursor-grab active:cursor-grabbing transition ${isSelected ? 'selected-slot' : ''}"
                 title="ドラッグまたは2回クリックで優先順位を入れ替え">
      <div class="flex items-center gap-1 shrink-0">
        <span class="text-slate-400 text-xs select-none leading-none" aria-hidden="true">⋮⋮</span>
        <span class="w-7 text-center text-[10px] font-bold px-1 py-0.5 rounded bg-blue-50 text-blue-800 border border-blue-200">${idx + 1}位</span>
        <div class="flex flex-col">
          <button onclick="event.stopPropagation(); moveSchoolPriority(${idx}, -1)" ${idx === 0 ? 'disabled' : ''} class="text-[9px] leading-none px-1 py-0.5 text-slate-500 hover:text-slate-900 disabled:opacity-25">▲</button>
          <button onclick="event.stopPropagation(); moveSchoolPriority(${idx}, 1)" ${idx === schools.length - 1 ? 'disabled' : ''} class="text-[9px] leading-none px-1 py-0.5 text-slate-500 hover:text-slate-900 disabled:opacity-25">▼</button>
        </div>
      </div>
      <div class="min-w-0 flex-1">
        <div class="text-slate-800 font-semibold truncate text-[11px]" title="${escapeHtml(sc)}">${escapeHtml(sc)}</div>
        <div class="text-[10px] text-slate-400">${cnt}組</div>
      </div>
      <input type="text"
             value="${escapeHtml(currentShort)}"
             onclick="event.stopPropagation()"
             onmousedown="event.stopPropagation()"
             onchange="updateSchoolAlias('${escapeJs(sc)}', this.value)"
             class="w-16 px-1.5 py-0.5 text-xs font-bold border rounded text-center focus:outline-none focus:border-teal-500 cursor-text ${isCustom ? 'bg-teal-50 border-teal-400 text-teal-900' : 'border-slate-300 text-slate-800'}"
             title="表に表示する学校略称" />
    </div>`;
  }).join('');
}

function formatPairLabel(entry, includeSchool = true) {
  if (!entry) return '';
  const names = entry.p2 ? `${entry.p1}・${entry.p2}` : entry.p1;
  if (!includeSchool || !entry.school) return names;
  return `${names}（${formatSchoolShortName(entry.school)}）`;
}

// =========================================================================
// 抽選ルール設定UI
// =========================================================================
function renderRuleCheckboxes() {
  const html = RULE_DEFINITIONS.map((def) => {
    const checked = Boolean(state.rules[def.key]);
    return `<label class="flex items-start gap-2.5 p-2 rounded-lg border cursor-pointer transition ${checked ? 'bg-slate-50 border-slate-300' : 'bg-white border-slate-200 opacity-75'}">
      <input type="checkbox" ${checked ? 'checked' : ''} onchange="toggleRule('${def.key}', this.checked)"
             class="mt-0.5 rounded text-emerald-600 focus:ring-emerald-500" />
      <div class="min-w-0">
        <div class="text-xs font-bold text-slate-800 leading-tight">${escapeHtml(def.label)}</div>
        <div class="text-[10px] text-slate-500 leading-snug mt-0.5">${escapeHtml(def.desc)}</div>
      </div>
    </label>`;
  }).join('');

  const regBox = document.getElementById('reg-rules-checkboxes');
  const popBox = document.getElementById('popover-rules-checkboxes');
  if (regBox) regBox.innerHTML = html;
  if (popBox) popBox.innerHTML = html;
}

function toggleRule(key, value) {
  state.rules[key] = Boolean(value);
  saveToLocalStorage();
  renderRuleCheckboxes();
}

function toggleRulesPopover() {
  const pop = document.getElementById('rules-popover');
  if (pop) pop.classList.toggle('hidden');
}

// =========================================================================
// 履歴管理（Undo / Redo）
// =========================================================================
function pushHistory() {
  const snapshot = JSON.stringify({
    bracketSize: state.bracketSize,
    slots: state.slots,
    activeSlotMask: state.activeSlotMask,
    bench: state.bench,
    slotFills: state.slotFills,
    results: state.results,
    consolation: state.consolation
  });
  if (state.history.length > 0 && state.history[state.history.length - 1] === snapshot) return;
  state.history.push(snapshot);
  if (state.history.length > 50) state.history.shift();
  state.future = [];
  updateUndoRedoButtons();
}

function undo() {
  if (state.history.length === 0) return;
  const current = JSON.stringify({
    bracketSize: state.bracketSize,
    slots: state.slots,
    activeSlotMask: state.activeSlotMask,
    bench: state.bench,
    slotFills: state.slotFills,
    results: state.results,
    consolation: state.consolation
  });
  state.future.push(current);
  restoreSnapshot(JSON.parse(state.history.pop()));
  updateUndoRedoButtons();
  showToast('1つ前の状態に戻しました');
}

function redo() {
  if (state.future.length === 0) return;
  const current = JSON.stringify({
    bracketSize: state.bracketSize,
    slots: state.slots,
    activeSlotMask: state.activeSlotMask,
    bench: state.bench,
    slotFills: state.slotFills,
    results: state.results,
    consolation: state.consolation
  });
  state.history.push(current);
  restoreSnapshot(JSON.parse(state.future.pop()));
  updateUndoRedoButtons();
  showToast('やり直しました');
}

function restoreSnapshot(snap) {
  state.bracketSize = snap.bracketSize;
  state.slots = snap.slots;
  state.activeSlotMask = snap.activeSlotMask;
  state.bench = normalizeBench(snap.bench);
  state.slotFills = (snap.slotFills && typeof snap.slotFills === 'object') ? snap.slotFills : {};
  state.results = snap.results;
  state.consolation = snap.consolation;
  state.selectedTarget = null;
  refreshAlerts(false);
  saveToLocalStorage();
  renderBench();
  renderTourneySidebar();
  renderSheet();
}

function updateUndoRedoButtons() {
  const btnUndo = document.getElementById('btn-undo');
  const btnRedo = document.getElementById('btn-redo');
  if (btnUndo) btnUndo.disabled = state.history.length === 0;
  if (btnRedo) btnRedo.disabled = state.future.length === 0;
}

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    if (state.selectedTarget) {
      state.selectedTarget = null;
      renderBench();
      renderSheet();
    }
    if (state.selectedSchoolIdx !== null) {
      state.selectedSchoolIdx = null;
      renderSchoolAliasEditor();
    }
    if (state.editingEntryId !== null) {
      state.editingEntryId = null;
      renderRegistrationTable();
    }
    const pop = document.getElementById('rules-popover');
    if (pop && !pop.classList.contains('hidden')) pop.classList.add('hidden');
    closeExportMenu();
    closeScoreModal();
    return;
  }
  if (state.mainView !== 'tournament') return;
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z' && !e.shiftKey) {
    e.preventDefault();
    undo();
  } else if ((e.ctrlKey || e.metaKey) && (e.key.toLowerCase() === 'y' || (e.key.toLowerCase() === 'z' && e.shiftKey))) {
    e.preventDefault();
    redo();
  }
});

// =========================================================================
// 画面・タブ・左右サイドバー・印刷切替
// =========================================================================
function switchRegSubTab(tab) {
  const tabs = ['alias', 'rules', 'consolation'];
  const targetTab = tabs.includes(tab) ? tab : 'alias';
  state.regSubTab = targetTab;
  tabs.forEach((t) => {
    const btn = document.getElementById(`left-tab-${t}`);
    const panel = document.getElementById(`reg-panel-${t}`);
    if (!btn || !panel) return;
    if (t === targetTab) {
      btn.className = 'py-1.5 rounded-lg bg-white text-blue-700 shadow-sm border border-slate-200/80 transition';
      panel.classList.remove('hidden');
    } else {
      btn.className = 'py-1.5 rounded-lg text-slate-600 hover:text-slate-900 transition';
      panel.classList.add('hidden');
    }
  });
}

function switchMainView(view) {
  state.mainView = view;
  const homeView = document.getElementById('view-home');
  const regView = document.getElementById('view-registration');
  const tourneyView = document.getElementById('view-tournament');
  const oopView = document.getElementById('view-oop');
  const spectatorView = document.getElementById('view-spectator');
  const eventBar = document.getElementById('event-switcher-bar');

  const tabHome = document.getElementById('main-tab-home');
  const tabReg = document.getElementById('main-tab-reg');
  const tabTourney = document.getElementById('main-tab-tourney');
  const tabOop = document.getElementById('main-tab-oop');
  const tabSpectator = document.getElementById('main-tab-spectator');

  const activeCls = 'px-2 sm:px-2.5 py-1 text-[11px] sm:text-xs font-bold rounded-md bg-blue-600 text-white shadow transition flex items-center gap-1 whitespace-nowrap';
  const idleCls = 'px-2 sm:px-2.5 py-1 text-[11px] sm:text-xs font-bold rounded-md text-slate-300 hover:text-white transition flex items-center gap-1 whitespace-nowrap';
  const activeSpecCls = 'px-2 sm:px-2.5 py-1 text-[11px] sm:text-xs font-bold rounded-md bg-emerald-600 text-white shadow transition flex items-center gap-1 whitespace-nowrap';
  const idleSpecCls = 'px-2 sm:px-2.5 py-1 text-[11px] sm:text-xs font-bold rounded-md text-emerald-300 hover:text-white transition flex items-center gap-1 whitespace-nowrap';

  [homeView, regView, tourneyView, oopView, spectatorView].forEach((el) => {
    if (el) el.classList.add('hidden');
  });
  if (tabHome) tabHome.className = (view === 'home') ? activeCls : idleCls;
  if (tabReg) tabReg.className = (view === 'registration') ? activeCls : idleCls;
  if (tabTourney) tabTourney.className = (view === 'tournament') ? activeCls : idleCls;
  if (tabOop) tabOop.className = (view === 'oop') ? activeCls : idleCls;
  if (tabSpectator) tabSpectator.className = (view === 'spectator') ? activeSpecCls : idleSpecCls;
  if (eventBar) eventBar.classList.toggle('hidden', view === 'spectator' || view === 'home');

  const ensureBracketReady = () => {
    if (state.slots.length === 0) {
      if (state.entries.length >= 2) {
        generateTournamentFromEntries(false, false);
      } else {
        loadRealisticSampleData(false);
        generateTournamentFromEntries(false, false);
      }
    }
  };

  if (view === 'home') {
    if (homeView) homeView.classList.remove('hidden');
    if (typeof window.onEnterHomeView === 'function') window.onEnterHomeView();
  } else if (view === 'registration') {
    if (regView) regView.classList.remove('hidden');
    renderSchoolAliasEditor();
    renderRegistrationTable();
  } else if (view === 'tournament') {
    if (tourneyView) tourneyView.classList.remove('hidden');
    ensureBracketReady();
    refreshAlerts(false);
    renderBench();
    renderTourneySidebar();
    renderSheet();
    requestAnimationFrame(() => {
      const viewport = document.getElementById('canvas-viewport');
      const sheet = document.getElementById('tournament-sheet');
      if (viewport && sheet && sheet.offsetWidth > viewport.clientWidth - 24) {
        fitZoomToScreen();
      } else {
        resetZoom100();
      }
    });
  } else if (view === 'oop') {
    if (oopView) oopView.classList.remove('hidden');
    ensureBracketReady();
    if (typeof window.onEnterOopView === 'function') window.onEnterOopView();
  } else if (view === 'spectator') {
    if (spectatorView) spectatorView.classList.remove('hidden');
    ensureBracketReady();
    if (typeof window.onEnterSpectatorView === 'function') window.onEnterSpectatorView();
  }
}

function toggleRankBadgeDisplay(forceVal) {
  state.showRankBadge = (typeof forceVal === 'boolean') ? forceVal : !state.showRankBadge;
  const btn = document.getElementById('btn-toggle-rank-badge');
  const pill = document.getElementById('rank-badge-toggle-state');
  if (btn && pill) {
    if (state.showRankBadge) {
      btn.className = 'px-2 py-0.5 text-[11px] font-bold bg-teal-600 hover:bg-teal-500 text-white border border-teal-600 rounded-md transition flex items-center gap-1';
      pill.className = 'text-[10px] px-1 rounded bg-teal-800/70 text-teal-100';
      pill.textContent = 'ON';
    } else {
      btn.className = 'px-2 py-0.5 text-[11px] font-bold bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 rounded-md transition flex items-center gap-1';
      pill.className = 'text-[10px] px-1 rounded bg-slate-200 text-slate-600';
      pill.textContent = 'OFF';
    }
  }
  if (state.mainView === 'tournament') renderSheet();
}

function toggleSlotPriorityDisplay(forceVal) {
  state.showSlotPriority = (typeof forceVal === 'boolean') ? forceVal : !state.showSlotPriority;
  const btn = document.getElementById('btn-toggle-priority');
  const pill = document.getElementById('priority-toggle-state');
  if (btn && pill) {
    if (state.showSlotPriority) {
      btn.className = 'px-2 py-0.5 text-[11px] font-bold bg-indigo-600 hover:bg-indigo-500 text-white border border-indigo-600 rounded-md transition flex items-center gap-1 shadow-sm';
      pill.className = 'text-[10px] px-1 rounded bg-indigo-800/80 text-indigo-100';
      pill.textContent = 'ON';
    } else {
      btn.className = 'px-2 py-0.5 text-[11px] font-bold bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 rounded-md transition flex items-center gap-1';
      pill.className = 'text-[10px] px-1 rounded bg-slate-200 text-slate-600';
      pill.textContent = 'OFF';
    }
  }
  if (state.mainView === 'tournament') renderSheet();
}

function updatePanelToggleButtons() {
  const btnLeft = document.getElementById('btn-sidebar-open');
  if (btnLeft) {
    btnLeft.className = state.sidebarCollapsed
      ? 'px-2 py-1 text-xs font-bold bg-slate-100 hover:bg-slate-200 text-slate-600 border border-slate-300 rounded-md transition flex items-center gap-1'
      : 'px-2 py-1 text-xs font-bold bg-blue-50 hover:bg-blue-100 text-blue-800 border border-blue-200 rounded-md transition flex items-center gap-1';
  }
  const btnRight = document.getElementById('btn-bench-toggle');
  const rightOpen = state.appMode === 'edit' && !state.benchCollapsed;
  if (btnRight) {
    btnRight.className = rightOpen
      ? 'px-2 py-1 text-xs font-bold bg-amber-100 hover:bg-amber-200 text-amber-900 border border-amber-300 rounded-md transition flex items-center gap-1'
      : 'px-2 py-1 text-xs font-bold bg-slate-100 hover:bg-slate-200 text-slate-600 border border-slate-300 rounded-md transition flex items-center gap-1';
  }
}

function toggleSidebar() {
  state.sidebarCollapsed = !state.sidebarCollapsed;
  const aside = document.getElementById('tourney-sidebar');
  if (aside) aside.classList.toggle('collapsed', state.sidebarCollapsed);
  updatePanelToggleButtons();
}

function toggleBenchSidebar() {
  if (state.appMode !== 'edit') {
    state.benchCollapsed = false;
    setAppMode('edit');
    return;
  }
  state.benchCollapsed = !state.benchCollapsed;
  const aside = document.getElementById('bench-sidebar');
  if (aside) aside.classList.toggle('collapsed', state.benchCollapsed);
  updatePanelToggleButtons();
}

function syncTitle(val) {
  state.title = val || '大会トーナメント表';
  saveToLocalStorage();
  if (state.mainView === 'tournament') renderSheet();
}

function setPrintMode(mode) {
  state.printMode = mode;
  const btnDist = document.getElementById('btn-print-dist');
  const btnCheck = document.getElementById('btn-print-check');
  if (btnDist && btnCheck) {
    if (mode === 'distribution') {
      btnDist.className = 'px-2 py-0.5 text-[11px] rounded font-bold bg-emerald-600 text-white transition';
      btnCheck.className = 'px-2 py-0.5 text-[11px] rounded font-bold text-slate-300 hover:text-white transition';
      document.body.classList.add('print-hide-ranks');
    } else {
      btnCheck.className = 'px-2 py-0.5 text-[11px] rounded font-bold bg-amber-500 text-slate-950 transition';
      btnDist.className = 'px-2 py-0.5 text-[11px] rounded font-bold text-slate-300 hover:text-white transition';
      document.body.classList.remove('print-hide-ranks');
    }
  }
  saveToLocalStorage();
}

function setPrintOrientation(orientation) {
  state.printOrientation = orientation;
  const btnP = document.getElementById('btn-orient-portrait');
  const btnL = document.getElementById('btn-orient-landscape');
  const styleEl = document.getElementById('dynamic-print-page-style');
  if (styleEl) styleEl.textContent = `@page { size: A4 ${orientation}; margin: 7mm; }`;
  if (btnP && btnL) {
    if (orientation === 'portrait') {
      btnP.className = 'px-2 py-0.5 text-[11px] rounded font-bold bg-slate-600 text-white transition';
      btnL.className = 'px-2 py-0.5 text-[11px] rounded font-bold text-slate-300 hover:text-white transition';
    } else {
      btnL.className = 'px-2 py-0.5 text-[11px] rounded font-bold bg-slate-600 text-white transition';
      btnP.className = 'px-2 py-0.5 text-[11px] rounded font-bold text-slate-300 hover:text-white transition';
    }
  }
  saveToLocalStorage();
}

function executePrint() {
  if (state.mainView !== 'tournament') switchMainView('tournament');
  setTimeout(() => window.print(), 160);
}

// =========================================================================
// 保存メニュー制御 ＆ JSON / Excel / PDF エクスポート・インポート
// =========================================================================
function toggleExportMenu() {
  const menu = document.getElementById('export-menu-popover');
  if (menu) menu.classList.toggle('hidden');
}

function closeExportMenu() {
  const menu = document.getElementById('export-menu-popover');
  if (menu) menu.classList.add('hidden');
}

document.addEventListener('click', (e) => {
  const menu = document.getElementById('export-menu-popover');
  const btn = document.getElementById('btn-export-menu');
  if (!menu || menu.classList.contains('hidden')) return;
  if (!menu.contains(e.target) && (!btn || !btn.contains(e.target))) {
    menu.classList.add('hidden');
  }
});

function getExportBaseFileName() {
  const safeTitle = (state.title || 'tournament').replace(/[\\/:*?"<>|]/g, '_');
  const dateStr = new Date().toISOString().slice(0, 10);
  return `${safeTitle}_${dateStr}`;
}

function exportJsonState() {
  if (typeof window.exportBundleJsonState === 'function') {
    window.exportBundleJsonState();
    return;
  }
  try {
    const payload = {
      version: '5.1',
      exportedAt: new Date().toISOString(),
      title: state.title,
      matchFormat: state.matchFormat,
      entries: state.entries,
      schoolOrder: state.schoolOrder,
      schoolAliases: state.schoolAliases,
      rules: state.rules,
      bracketSize: state.bracketSize,
      slots: state.slots,
      activeSlotMask: state.activeSlotMask,
      bench: state.bench,
      slotFills: state.slotFills,
      printShowFills: state.printShowFills,
      results: state.results,
      consolation: state.consolation,
      layoutMode: state.layoutMode
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${getExportBaseFileName()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    showToast('JSONファイルを保存しました', 'success');
  } catch (err) {
    alert('JSON保存中にエラーが発生しました: ' + err.message);
  }
}

function getRoundDisplayName(roundIdx, totalRounds) {
  const remaining = Math.pow(2, totalRounds - roundIdx + 1);
  if (remaining === 2) return '決勝';
  if (remaining === 4) return '準決勝';
  if (remaining === 8) return '準々決勝';
  return `${roundIdx}回戦`;
}

function exportExcelState() {
  if (typeof XLSX === 'undefined') {
    alert('Excel出力ライブラリの読み込みに失敗しました。通信環境をご確認ください。');
    return;
  }
  try {
    syncSchoolOrder();
    const schoolPriorityMap = {};
    state.schoolOrder.forEach((sc, idx) => { schoolPriorityMap[sc] = idx + 1; });
    const fillLabelMap = {};
    FILL_COLORS.forEach((c) => { fillLabelMap[c.key] = c.label; });

    const wb = XLSX.utils.book_new();

    // シート1：出場者名簿（先頭4列は本アプリ4項目読込と完全互換）
    const rosterRows = [
      ['番手番号', '生徒A', '生徒B', '学校名', '学校略称', '確認表記', 'シード', '学校優先順位']
    ];
    state.entries.forEach((e) => {
      rosterRows.push([
        e.schoolRank,
        e.p1,
        e.p2 || '',
        e.school,
        formatSchoolShortName(e.school),
        formatSchoolRankBadge(e),
        e.seed ? `第${e.seed}シード` : '',
        schoolPriorityMap[e.school] || ''
      ]);
    });
    const wsRoster = XLSX.utils.aoa_to_sheet(rosterRows);
    wsRoster['!cols'] = [
      { wch: 10 }, { wch: 16 }, { wch: 16 }, { wch: 18 },
      { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 13 }
    ];
    XLSX.utils.book_append_sheet(wb, wsRoster, '出場者名簿');

    // シート2：トーナメント配置表
    const size = state.bracketSize;
    const halfSize = size / 2;
    const adjPriority = BracketEngine.buildAdjustedSlotPriorityMap(size, state.activeSlotMask);
    const bracketRows = [
      ['枠番号', '左右', 'ブロック', '枠優先順位', '1回戦', '生徒A', '生徒B', '学校名', '略称', '校内番手', 'シード', '塗りつぶし']
    ];
    let dispNum = 1;
    for (let i = 0; i < size; i++) {
      if (!state.activeSlotMask[i]) continue;
      const entry = state.slots[i];
      const bInfo = BracketEngine.getBlockInfo(i, size);
      const isBye = BracketEngine.isByeSlot(i, state.activeSlotMask);
      const fillKey = state.slotFills[String(i)];
      bracketRows.push([
        dispNum++,
        i < halfSize ? '左山' : '右山',
        `${bInfo.code}(${bInfo.side === 'in' ? '内' : '外'})`,
        adjPriority[i] || '',
        isBye ? '足長(BYE)' : '1回戦',
        entry ? entry.p1 : '（空き枠）',
        entry ? (entry.p2 || '') : '',
        entry ? entry.school : '',
        entry ? formatSchoolShortName(entry.school) : '',
        entry ? entry.schoolRank : '',
        entry && entry.seed ? `S${entry.seed}` : '',
        fillKey ? (fillLabelMap[fillKey] || fillKey) : ''
      ]);
    }
    const wsBracket = XLSX.utils.aoa_to_sheet(bracketRows);
    wsBracket['!cols'] = [
      { wch: 8 }, { wch: 8 }, { wch: 11 }, { wch: 11 }, { wch: 11 },
      { wch: 16 }, { wch: 16 }, { wch: 18 }, { wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 10 }
    ];
    XLSX.utils.book_append_sheet(wb, wsBracket, 'トーナメント配置表');

    // シート3：試合結果一覧（本戦＋順位決定戦）
    const totalRounds = Math.log2(size);
    const tree = [];
    tree[0] = [];
    for (let i = 0; i < size; i++) {
      tree[0][i] = { active: Boolean(state.activeSlotMask[i]), entry: state.slots[i] || null };
    }
    const resultRows = [
      ['種別', 'ラウンド', '試合', '上/左ペア', 'G1', 'G2', '下/右ペア', '勝者']
    ];
    for (let r = 1; r <= totalRounds; r++) {
      const mCount = size / Math.pow(2, r);
      tree[r] = [];
      for (let m = 0; m < mCount; m++) {
        const c0 = tree[r - 1][m * 2];
        const c1 = tree[r - 1][m * 2 + 1];
        const active = c0.active || c1.active;
        const isBye = (c0.active && !c1.active) || (!c0.active && c1.active);
        const res = state.results[`${r}_${m}`] || null;
        let adv = null;
        if (isBye) adv = c0.active ? c0.entry : c1.entry;
        else if (res) adv = (res.winner === 0) ? c0.entry : c1.entry;
        tree[r][m] = { active, isBye, c0, c1, result: res, entry: adv };

        if (active && !isBye && (c0.entry || c1.entry)) {
          resultRows.push([
            '本戦',
            getRoundDisplayName(r, totalRounds),
            m + 1,
            c0.entry ? formatPairLabel(c0.entry, true) : '未定',
            res ? res.g1 : '',
            res ? res.g2 : '',
            c1.entry ? formatPairLabel(c1.entry, true) : '未定',
            adv ? formatPairLabel(adv, true) : ''
          ]);
        }
      }
    }
    if (state.consolation.preset !== 'none') {
      for (let m = 0; m < state.consolation.matchCount; m++) {
        const p1 = state.consolation.slots[m * 2] || '未定';
        const p2 = state.consolation.slots[m * 2 + 1] || '未定';
        const res = state.consolation.results[m];
        const winnerName = res ? (res.winner === 0 ? p1 : p2) : '';
        resultRows.push([
          state.consolation.title || '順位決定戦',
          '順位戦',
          m + 1,
          p1,
          res ? res.g1 : '',
          res ? res.g2 : '',
          p2,
          winnerName
        ]);
      }
    }
    const wsResults = XLSX.utils.aoa_to_sheet(resultRows);
    wsResults['!cols'] = [
      { wch: 14 }, { wch: 10 }, { wch: 6 }, { wch: 28 },
      { wch: 6 }, { wch: 6 }, { wch: 28 }, { wch: 28 }
    ];
    XLSX.utils.book_append_sheet(wb, wsResults, '試合結果一覧');

    XLSX.writeFile(wb, `${getExportBaseFileName()}.xlsx`);
    showToast('Excelファイル（3シート）を保存しました', 'success');
  } catch (err) {
    alert('Excel保存中にエラーが発生しました: ' + err.message);
  }
}

async function exportPdfState() {
  if (typeof window.html2canvas === 'undefined' || !window.jspdf || !window.jspdf.jsPDF) {
    showToast('PDFライブラリ未読込のため、ブラウザの印刷（PDF保存）画面を開きます', 'warn', 3600);
    executePrint();
    return;
  }

  const wasRegView = state.mainView !== 'tournament';
  if (wasRegView) {
    switchMainView('tournament');
    await new Promise((resolve) => setTimeout(resolve, 220));
  }

  const sheet = document.getElementById('tournament-sheet');
  if (!sheet) return;

  showToast('PDFを生成しています...', 'info', 2200);
  const prevTransform = sheet.style.transform;

  try {
    document.body.classList.add('pdf-capturing');
    sheet.style.transform = 'scale(1)';
    if (document.fonts && document.fonts.ready) {
      await document.fonts.ready;
    }
    autoFitSlotTexts(sheet);
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));

    const canvas = await window.html2canvas(sheet, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#ffffff',
      logging: false,
      windowWidth: sheet.scrollWidth + 64,
      windowHeight: sheet.scrollHeight + 64
    });

    const orientation = state.printOrientation === 'landscape' ? 'landscape' : 'portrait';
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation, unit: 'mm', format: 'a4' });
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();
    const margin = 8;
    const usableW = pageW - margin * 2;
    const usableH = pageH - margin * 2;

    const imgW = canvas.width;
    const imgH = canvas.height;
    const ratio = Math.min(usableW / imgW, usableH / imgH);
    const drawW = imgW * ratio;
    const drawH = imgH * ratio;
    const offsetX = (pageW - drawW) / 2;
    const offsetY = margin;

    const imgData = canvas.toDataURL('image/png');
    pdf.addImage(imgData, 'PNG', offsetX, offsetY, drawW, drawH, undefined, 'FAST');
    pdf.save(`${getExportBaseFileName()}.pdf`);
    showToast(`PDFファイル（A4${orientation === 'landscape' ? '横' : '縦'}）を保存しました`, 'success');
  } catch (err) {
    alert('PDF出力中にエラーが発生しました: ' + err.message);
  } finally {
    document.body.classList.remove('pdf-capturing');
    sheet.style.transform = prevTransform;
  }
}

function importJsonState(event) {
  if (typeof window.importBundleJsonState === 'function') {
    window.importBundleJsonState(event);
    return;
  }
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = JSON.parse(e.target.result);
      if (!data || !Array.isArray(data.entries)) {
        throw new Error('有効なトーナメントJSONデータではありません。');
      }
      state.title = data.title || '大会トーナメント表';
      document.getElementById('tournament-title-input').value = state.title;
      state.matchFormat = (data.matchFormat === 5 || data.matchFormat === 7) ? data.matchFormat : 7;
      const fmtSel = document.getElementById('match-format-select');
      if (fmtSel) fmtSel.value = String(state.matchFormat);
      state.entries = data.entries || [];
      state.schoolOrder = Array.isArray(data.schoolOrder) ? data.schoolOrder : [];
      state.schoolAliases = data.schoolAliases || {};
      state.rules = Object.assign({}, BracketEngine.DEFAULT_RULES, data.rules || {});
      state.bracketSize = data.bracketSize || BracketEngine.nextPowerOfTwo(Math.max(4, state.entries.length));
      state.slots = Array.isArray(data.slots) ? data.slots : [];
      state.activeSlotMask = Array.isArray(data.activeSlotMask) ? data.activeSlotMask : state.slots.map((s) => Boolean(s));
      state.bench = normalizeBench(data.bench);
      state.slotFills = (data.slotFills && typeof data.slotFills === 'object') ? data.slotFills : {};
      if (typeof data.printShowFills === 'boolean') togglePrintFills(data.printShowFills, false);
      state.results = data.results || {};
      if (data.consolation) state.consolation = data.consolation;
      if (data.layoutMode) setLayoutMode(data.layoutMode);
      state.history = [];
      state.future = [];
      updateUndoRedoButtons();
      updateRegCountPill();
      renderRuleCheckboxes();
      refreshAlerts(true);
      saveToLocalStorage();
      renderSchoolAliasEditor();
      renderRegistrationTable();
      if (state.slots.length > 0) switchMainView('tournament');
      else switchMainView('registration');
      showToast('JSONデータを読み込みました', 'success');
    } catch (err) {
      alert('JSONファイルの読み込みに失敗しました: ' + err.message);
    } finally {
      event.target.value = '';
    }
  };
  reader.readAsText(file, 'UTF-8');
}

// =========================================================================
// サンプルデータ・Excel/CSV・テキスト・個別追加
// =========================================================================
function loadRealisticSampleData(notify = true) {
  const sample = [
    { schoolRank: 1, p1: '佐藤 翔太', p2: '鈴木 健太', school: '星見第二中学校', seed: 1 },
    { schoolRank: 2, p1: '高橋 翼',   p2: '田中 蓮',   school: '星見第二中学校', seed: null },
    { schoolRank: 3, p1: '伊藤 陸',   p2: '渡辺 颯',   school: '星見第二中学校', seed: null },
    { schoolRank: 4, p1: '山本 大輝', p2: '中村 悠真', school: '星見第二中学校', seed: null },
    { schoolRank: 5, p1: '斉藤 律',   p2: '藤田 湊斗', school: '星見第二中学校', seed: null },
    { schoolRank: 1, p1: '小林 陽向', p2: '加藤 樹',   school: '風見第六中学校', seed: 2 },
    { schoolRank: 2, p1: '吉田 湊',   p2: '山田 蒼空', school: '風見第六中学校', seed: null },
    { schoolRank: 3, p1: '佐々木 朝陽', p2: '山口 律', school: '風見第六中学校', seed: null },
    { schoolRank: 1, p1: '松本 暖',   p2: '井上 新',   school: '星見第三中学校', seed: 3 },
    { schoolRank: 2, p1: '木村 結翔', p2: '林 颯真',   school: '星見第三中学校', seed: null },
    { schoolRank: 1, p1: '斎藤 碧',   p2: '清水 奏汰', school: '風見第四中学校', seed: 4 },
    { schoolRank: 2, p1: '山崎 瑛太', p2: '森 陽太',   school: '風見第四中学校', seed: null },
    { schoolRank: 1, p1: '池田 悠人', p2: '橋本 大和', school: '星見第五中学校', seed: null },
    { schoolRank: 2, p1: '岡田 健',   p2: '後藤 蓮司', school: '星見第五中学校', seed: null },
    { schoolRank: 1, p1: '阿部 蒼',   p2: '石川 怜',   school: '風見第一中学校', seed: null }
  ];

  state.entries = sample.map((item, idx) => ({
    id: `sample_${idx + 1}`,
    serial: idx + 1,
    schoolRank: item.schoolRank,
    p1: item.p1,
    p2: item.p2,
    school: item.school,
    seed: item.seed
  }));
  state.schoolOrder = [];
  state.title = '令和8年度 春季大会 個人戦トーナメント表';
  document.getElementById('tournament-title-input').value = state.title;
  document.getElementById('entry-count-input').value = '15';
  updateRegCountPill();
  renderSchoolAliasEditor();
  renderRegistrationTable();
  saveToLocalStorage();
  if (notify) showToast('サンプル15組（6校・第1〜4シード付）を読み込みました', 'success');
}

function downloadSampleExcel() {
  if (typeof XLSX === 'undefined') {
    alert('Excelライブラリの読み込みに失敗しました。通信環境をご確認ください。');
    return;
  }
  const rows = [
    ['番手番号', '生徒A', '生徒B', '学校名'],
    [1, '佐藤 翔太', '鈴木 健太', '星見第二中学校'],
    [2, '高橋 翼', '田中 蓮', '星見第二中学校'],
    [3, '伊藤 陸', '渡辺 颯', '星見第二中学校'],
    [4, '山本 大輝', '中村 悠真', '星見第二中学校'],
    [5, '斉藤 律', '藤田 湊斗', '星見第二中学校'],
    [1, '小林 陽向', '加藤 樹', '風見第六中学校'],
    [2, '吉田 湊', '山田 蒼空', '風見第六中学校'],
    [3, '佐々木 朝陽', '山口 律', '風見第六中学校'],
    [1, '松本 暖', '井上 新', '星見第三中学校'],
    [2, '木村 結翔', '林 颯真', '星見第三中学校'],
    [1, '斎藤 碧', '清水 奏汰', '風見第四中学校'],
    [2, '山崎 瑛太', '森 陽太', '風見第四中学校'],
    [1, '池田 悠人', '橋本 大和', '星見第五中学校'],
    [2, '岡田 健', '後藤 蓮司', '星見第五中学校'],
    [1, '阿部 蒼', '石川 怜', '風見第一中学校']
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{ wch: 10 }, { wch: 16 }, { wch: 16 }, { wch: 18 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '出場者名簿');
  XLSX.writeFile(wb, '出場者名簿サンプル.xlsx');
}

function handleExcelDragOver(e) {
  e.preventDefault();
  e.currentTarget.classList.add('border-emerald-500', 'bg-emerald-50/60');
}

function handleExcelDragLeave(e) {
  e.currentTarget.classList.remove('border-emerald-500', 'bg-emerald-50/60');
}

function handleExcelDrop(e) {
  e.preventDefault();
  e.currentTarget.classList.remove('border-emerald-500', 'bg-emerald-50/60');
  const file = e.dataTransfer?.files?.[0];
  if (file) processSpreadsheetFile(file);
}

function handleSpreadsheetUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  processSpreadsheetFile(file, () => { event.target.value = ''; });
}

function processSpreadsheetFile(file, onDone) {
  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });

      const parsed = [];
      const schoolAutoCounter = {};

      // 4項目（番手 / 生徒A / 生徒B / 学校名）か 旧5項目（通し / 番手 / 生徒A / 生徒B / 学校名）かを自動判別
      const firstRow = rows[0] || [];
      const h0 = String(firstRow[0] ?? '').trim();
      const h1 = String(firstRow[1] ?? '').trim();
      const h2 = String(firstRow[2] ?? '').trim();
      const h3 = String(firstRow[3] ?? '').trim();
      const h4 = String(firstRow[4] ?? '').trim();
      const hasHeader = h0.includes('番手') || h0.includes('通し') || h1.includes('生徒') || h2.includes('生徒') || h3.includes('学校') || h4.includes('学校');
      let isLegacy5Col = h0.includes('通し') || h4.includes('学校');
      if (!hasHeader && firstRow.length >= 5) {
        if (/^\d+$/.test(h0) && /^\d+$/.test(h1) && h4 !== '') isLegacy5Col = true;
      }

      rows.forEach((row, idx) => {
        if (!row || row.length < 2) return;
        if (idx === 0 && hasHeader) return;

        let rankStr, p1, p2, schoolStr;
        if (isLegacy5Col) {
          rankStr = String(row[1] ?? '').trim();
          p1 = String(row[2] ?? '').trim();
          p2 = String(row[3] ?? '').trim();
          schoolStr = String(row[4] ?? '').trim();
        } else {
          rankStr = String(row[0] ?? '').trim();
          p1 = String(row[1] ?? '').trim();
          p2 = String(row[2] ?? '').trim();
          schoolStr = String(row[3] ?? '').trim();
        }

        if (!p1 && !p2) return;

        const school = schoolStr || '無所属';
        schoolAutoCounter[school] = (schoolAutoCounter[school] || 0) + 1;
        const rankNum = parseInt(rankStr, 10) || schoolAutoCounter[school];

        parsed.push({
          id: 'ent_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7),
          serial: parsed.length + 1,
          schoolRank: rankNum,
          p1,
          p2,
          school,
          seed: null
        });
      });

      if (parsed.length === 0) {
        alert('有効な選手データが見つかりませんでした。列構成（番手番号 / 生徒A / 生徒B / 学校名）をご確認ください。');
        return;
      }
      if (parsed.length > 256) {
        alert('最大256ペアまでのため、先頭256ペアを読み込みます。');
        parsed.length = 256;
      }

      state.entries = parsed;
      state.schoolOrder = [];
      state.editingEntryId = null;
      normalizeSerials();
      updateRegCountPill();
      renderSchoolAliasEditor();
      renderRegistrationTable();
      saveToLocalStorage();
      showToast(`${parsed.length}組の名簿を読み込みました`, 'success');
    } catch (err) {
      alert('ファイルの解析に失敗しました: ' + err.message);
    } finally {
      if (typeof onDone === 'function') onDone();
    }
  };
  reader.readAsArrayBuffer(file);
}

function handleAddKeydown(e) {
  if (e.key === 'Enter') {
    e.preventDefault();
    addSingleEntry();
  }
}

function addSingleEntry() {
  const rank = parseInt(document.getElementById('add-rank').value, 10) || 1;
  const p1 = document.getElementById('add-p1').value.trim();
  const p2 = document.getElementById('add-p2').value.trim();
  const school = document.getElementById('add-school').value.trim() || '無所属';

  if (!p1) {
    alert('生徒Aの名前を入力してください。');
    return;
  }
  if (state.entries.length >= 256) {
    alert('最大256ペアまで登録可能です。');
    return;
  }

  state.entries.push({
    id: 'ent_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7),
    serial: state.entries.length + 1,
    schoolRank: rank,
    p1,
    p2,
    school,
    seed: null
  });

  document.getElementById('add-p1').value = '';
  document.getElementById('add-p2').value = '';
  document.getElementById('add-rank').value = String(rank + 1);
  document.getElementById('add-p1').focus();
  updateRegCountPill();
  renderSchoolAliasEditor();
  renderRegistrationTable();
  saveToLocalStorage();
}

function removeEntry(id) {
  state.entries = state.entries.filter((e) => e.id !== id);
  normalizeSerials();
  updateRegCountPill();
  renderSchoolAliasEditor();
  renderRegistrationTable();
  saveToLocalStorage();
}

function clearAllEntries() {
  if (state.entries.length === 0) return;
  if (!confirm('登録されているすべての生徒名簿を消去しますか？')) return;
  state.entries = [];
  state.schoolAliases = {};
  state.slots = [];
  state.activeSlotMask = [];
  state.slotFills = {};
  state.results = {};
  state.bench = createEmptyBench();
  updateRegCountPill();
  renderSchoolAliasEditor();
  renderRegistrationTable();
  saveToLocalStorage();
  showToast('名簿をすべて消去しました');
}

function syncEntryToBracket(ent) {
  if (!ent) return;
  for (let i = 0; i < state.slots.length; i++) {
    if (state.slots[i] && state.slots[i].id === ent.id) {
      state.slots[i] = Object.assign({}, ent);
    }
  }
  for (let i = 0; i < state.bench.length; i++) {
    if (state.bench[i] && state.bench[i].id === ent.id) {
      state.bench[i] = Object.assign({}, ent);
    }
  }
}

function startEditEntry(id) {
  state.editingEntryId = id;
  renderRegistrationTable();
  const p1Input = document.getElementById('edit-entry-p1');
  if (p1Input) {
    p1Input.focus();
    p1Input.select();
  }
}

function cancelEditEntry() {
  state.editingEntryId = null;
  renderRegistrationTable();
}

function handleEditEntryKeydown(e, id) {
  if (e.key === 'Enter') {
    e.preventDefault();
    saveEditEntry(id);
  } else if (e.key === 'Escape') {
    e.preventDefault();
    cancelEditEntry();
  }
}

function saveEditEntry(id) {
  const ent = state.entries.find((e) => e.id === id);
  if (!ent) return;
  const p1 = (document.getElementById('edit-entry-p1')?.value || '').trim();
  const p2 = (document.getElementById('edit-entry-p2')?.value || '').trim();
  const school = (document.getElementById('edit-entry-school')?.value || '').trim() || '無所属';
  const rankVal = Math.max(1, parseInt(document.getElementById('edit-entry-rank')?.value, 10) || 1);
  const seedRaw = document.getElementById('edit-entry-seed')?.value ?? '';
  const seedVal = seedRaw === '' ? null : parseInt(seedRaw, 10);

  if (!p1) {
    alert('生徒Aの名前を入力してください。');
    return;
  }

  if (seedVal !== null) {
    state.entries.forEach((other) => {
      if (other.id !== id && other.seed === seedVal) {
        other.seed = null;
        syncEntryToBracket(other);
      }
    });
  }

  ent.p1 = p1;
  ent.p2 = p2;
  ent.school = school;
  ent.schoolRank = rankVal;
  ent.seed = seedVal;

  syncEntryToBracket(ent);
  state.editingEntryId = null;
  syncSchoolOrder();
  renderSchoolAliasEditor();
  renderRegistrationTable();
  refreshAlerts(false);
  saveToLocalStorage();
  showToast('ペア情報を更新しました（配置済みトーナメントにも反映）', 'success');
}

function updateEntryField(id, field, value) {
  const ent = state.entries.find((e) => e.id === id);
  if (!ent) return;
  if (field === 'schoolRank') {
    ent.schoolRank = Math.max(1, parseInt(value, 10) || 1);
    syncEntryToBracket(ent);
    refreshAlerts(false);
    renderRegistrationTable();
  } else if (field === 'seed') {
    const sVal = value === '' ? null : parseInt(value, 10);
    if (sVal !== null) {
      state.entries.forEach((other) => {
        if (other.id !== id && other.seed === sVal) {
          other.seed = null;
          syncEntryToBracket(other);
        }
      });
    }
    ent.seed = sVal;
    syncEntryToBracket(ent);
    refreshAlerts(false);
    renderRegistrationTable();
  }
  saveToLocalStorage();
}

function normalizeSerials() {
  state.entries.forEach((e, idx) => { e.serial = idx + 1; });
}

function updateRegCountPill() {
  const pill = document.getElementById('reg-count-pill');
  if (pill) pill.textContent = `${state.entries.length}組`;
}

/**
 * 生成前クイック診断バーの描画（ドロー規模・番手重複・学校表記ゆれ・複数シード特例予告）
 */
function renderPrecheckSummary() {
  const box = document.getElementById('precheck-summary-bar');
  if (!box) return;

  const n = state.entries.length;
  if (n < 2) {
    box.innerHTML = `<div class="text-slate-500">2組以上登録すると、想定ドロー規模と名簿チェック結果がここに表示されます。</div>`;
    return;
  }

  const size = BracketEngine.nextPowerOfTwo(n);
  const mask = BracketEngine.createBalancedActiveMask(n, size);
  const leftCnt = mask.slice(0, size / 2).filter(Boolean).length;
  const rightCnt = mask.slice(size / 2).filter(Boolean).length;
  const byeCnt = size - n;

  // 1. 同校内の番手重複チェック
  const rankKeyCounts = {};
  const dupRankSchools = new Set();
  // 2. 学校名の表記ゆれチェック（同じ略称に複数の元学校名が紐づくケース）
  const shortToRawSchools = {};
  // 3. 同一校複数シード特例（上位4番手に2組以上シード）
  const schoolTop4Seeds = {};
  // 4. 8組超の学校
  const schoolTotals = {};

  state.entries.forEach((e) => {
    const sc = e.school || '無所属';
    schoolTotals[sc] = (schoolTotals[sc] || 0) + 1;

    const rkKey = `${sc}||${e.schoolRank}`;
    rankKeyCounts[rkKey] = (rankKeyCounts[rkKey] || 0) + 1;
    if (rankKeyCounts[rkKey] > 1) dupRankSchools.add(sc);

    const shortSc = formatSchoolShortName(sc);
    if (!shortToRawSchools[shortSc]) shortToRawSchools[shortSc] = new Set();
    shortToRawSchools[shortSc].add(sc);

    if (e.seed && e.schoolRank >= 1 && e.schoolRank <= 4) {
      schoolTop4Seeds[sc] = (schoolTop4Seeds[sc] || 0) + 1;
    }
  });

  const aliasCollisionGroups = Object.entries(shortToRawSchools)
    .filter(([, set]) => set.size > 1)
    .map(([shortSc, set]) => ({ shortSc, raws: Array.from(set) }));
  const multiSeedSchools = Object.keys(schoolTop4Seeds).filter((sc) => schoolTop4Seeds[sc] >= 2);
  const over8Schools = Object.keys(schoolTotals).filter((sc) => schoolTotals[sc] > 8);

  const badges = [];
  if (dupRankSchools.size > 0) {
    const firstSc = Array.from(dupRankSchools)[0];
    badges.push(`<button onclick="toggleRegSchoolFilter('${escapeJs(firstSc)}')" class="px-1.5 py-0.5 rounded bg-rose-100 hover:bg-rose-200 text-rose-800 border border-rose-300 font-bold transition" title="クリックで該当校を絞り込み">⚠ 番手重複: ${dupRankSchools.size}校</button>`);
  }
  if (aliasCollisionGroups.length > 0) {
    const tip = aliasCollisionGroups.map((g) => `${g.shortSc}: ${g.raws.join(' / ')}`).join(', ');
    badges.push(`<span class="px-1.5 py-0.5 rounded bg-amber-100 text-amber-900 border border-amber-300 font-bold" title="${escapeHtml(tip)}">⚠ 校名表記ゆれ疑い: ${aliasCollisionGroups.length}件</span>`);
  }
  if (multiSeedSchools.length > 0) {
    const names = multiSeedSchools.map((s) => formatSchoolShortName(s)).join('・');
    badges.push(`<span class="px-1.5 py-0.5 rounded bg-indigo-100 text-indigo-900 border border-indigo-200 font-bold" title="同一校に複数シードがあるため残山＋in/out特例を適用">複数シード特例: ${escapeHtml(names)}</span>`);
  }
  if (over8Schools.length > 0) {
    badges.push(`<span class="px-1.5 py-0.5 rounded bg-slate-200 text-slate-700 font-bold">8組超: ${over8Schools.length}校</span>`);
  }
  if (dupRankSchools.size === 0 && aliasCollisionGroups.length === 0) {
    badges.push(`<span class="px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 border border-emerald-200 font-bold">✓ 番手重複・表記ゆれなし</span>`);
  }

  box.innerHTML = `
    <div class="flex items-center justify-between gap-2 text-slate-700 font-bold">
      <span>想定ドロー：<b class="text-slate-900">${n}組 → ${size}枠</b></span>
      <span class="text-[10px] text-slate-500">BYE ${byeCnt} ／ 左${leftCnt}・右${rightCnt}</span>
    </div>
    <div class="flex flex-wrap gap-1 pt-0.5">${badges.join('')}</div>
  `;
}

function renderRegistrationTable() {
  renderPrecheckSummary();
  const tbody = document.getElementById('reg-table-body');
  const summary = document.getElementById('reg-table-summary');
  const seedSummary = document.getElementById('reg-seed-summary');
  const schoolFilter = document.getElementById('reg-school-filter');
  const chipsBox = document.getElementById('reg-school-chips');
  const schoolDatalist = document.getElementById('school-datalist');
  const searchVal = (document.getElementById('reg-search-input')?.value || '').trim().toLowerCase();

  const schoolCounts = {};
  let seedCount = 0;
  state.entries.forEach((e) => {
    schoolCounts[e.school] = (schoolCounts[e.school] || 0) + 1;
    if (e.seed) seedCount++;
  });
  const schools = Object.keys(schoolCounts);
  summary.textContent = `${state.entries.length}組 / ${schools.length}校`;

  if (seedSummary) {
    if (seedCount > 0) {
      seedSummary.classList.remove('hidden');
      seedSummary.textContent = `シード ${seedCount}組`;
    } else {
      seedSummary.classList.add('hidden');
    }
  }

  if (schoolDatalist) {
    schoolDatalist.innerHTML = schools.map((sc) => `<option value="${escapeHtml(sc)}"></option>`).join('');
  }

  const prevSchool = schoolFilter.value;
  schoolFilter.innerHTML = `<option value="">すべての学校 (${schools.length}校)</option>` +
    schools.map((sc) => `<option value="${escapeHtml(sc)}">${escapeHtml(sc)} (${schoolCounts[sc]}組)</option>`).join('');
  if (schools.includes(prevSchool)) schoolFilter.value = prevSchool;

  chipsBox.innerHTML = schools.map((sc) => {
    const active = schoolFilter.value === sc;
    return `<button onclick="toggleRegSchoolFilter('${escapeJs(sc)}')"
                    class="px-2 py-0.5 text-[11px] rounded-full border font-semibold transition ${active ? 'bg-blue-600 text-white border-blue-600' : 'bg-slate-50 hover:bg-slate-100 text-slate-700 border-slate-200'}">
      ${escapeHtml(formatSchoolShortName(sc))} <span class="opacity-75">(${schoolCounts[sc]})</span>
    </button>`;
  }).join('');

  const activeSchool = schoolFilter.value;
  const filtered = state.entries.filter((e) => {
    if (activeSchool && e.school !== activeSchool) return false;
    if (searchVal) {
      const hay = `${e.p1} ${e.p2} ${e.school} ${formatSchoolShortName(e.school)} ${formatSchoolRankBadge(e)}`.toLowerCase();
      if (!hay.includes(searchVal)) return false;
    }
    return true;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="py-12 text-center text-slate-400 text-xs">
      ${state.entries.length === 0
        ? '出場者がまだ登録されていません。左パネルの「Excel読込」や上部の「サンプル15組読込」から始められます。'
        : '条件に一致する出場ペアがありません。'}
    </td></tr>`;
    return;
  }

  tbody.innerHTML = filtered.map((e) => {
    const rankTag = formatSchoolRankBadge(e);
    const seedOptions = ['<option value="">なし</option>'];
    for (let s = 1; s <= 16; s++) {
      seedOptions.push(`<option value="${s}" ${e.seed === s ? 'selected' : ''}>第${s}シード</option>`);
    }

    if (state.editingEntryId === e.id) {
      return `<tr class="bg-blue-50/80">
        <td class="py-1.5 px-2 text-center text-blue-600 font-mono font-bold">${e.serial}</td>
        <td class="py-1.5 px-2 text-center">
          <input id="edit-entry-rank" type="number" min="1" max="50" value="${e.schoolRank}"
                 onkeydown="handleEditEntryKeydown(event, '${e.id}')"
                 class="w-12 px-1 py-1 text-xs border border-blue-400 bg-white rounded text-center font-bold focus:outline-none" />
        </td>
        <td class="py-1.5 px-2">
          <div class="flex items-center gap-1">
            <input id="edit-entry-p1" type="text" value="${escapeHtml(e.p1)}" placeholder="生徒A"
                   onkeydown="handleEditEntryKeydown(event, '${e.id}')"
                   class="w-full min-w-[78px] px-2 py-1 text-xs font-bold border border-blue-400 bg-white rounded focus:outline-none" />
            <span class="text-slate-400 font-bold">・</span>
            <input id="edit-entry-p2" type="text" value="${escapeHtml(e.p2)}" placeholder="生徒B"
                   onkeydown="handleEditEntryKeydown(event, '${e.id}')"
                   class="w-full min-w-[78px] px-2 py-1 text-xs font-bold border border-blue-400 bg-white rounded focus:outline-none" />
          </div>
        </td>
        <td class="py-1.5 px-2">
          <input id="edit-entry-school" type="text" value="${escapeHtml(e.school)}" list="school-datalist" placeholder="学校名"
                 onkeydown="handleEditEntryKeydown(event, '${e.id}')"
                 class="w-full min-w-[100px] px-2 py-1 text-xs border border-blue-400 bg-white rounded focus:outline-none" />
        </td>
        <td class="py-1.5 px-2">
          <select id="edit-entry-seed" onkeydown="handleEditEntryKeydown(event, '${e.id}')"
                  class="px-1.5 py-1 text-xs border border-blue-400 bg-white rounded font-bold">
            ${seedOptions.join('')}
          </select>
        </td>
        <td class="py-1.5 px-2 text-center whitespace-nowrap">
          <button onclick="saveEditEntry('${e.id}')" class="px-2 py-1 bg-blue-600 hover:bg-blue-500 text-white text-[11px] font-bold rounded shadow-sm transition">保存</button>
          <button onclick="cancelEditEntry()" class="px-1.5 py-1 text-slate-500 hover:text-slate-900 text-[11px] font-bold rounded">戻す</button>
        </td>
      </tr>`;
    }

    return `<tr class="hover:bg-slate-50/90">
      <td class="py-1.5 px-2.5 text-center text-slate-400 font-mono">${e.serial}</td>
      <td class="py-1.5 px-2.5 text-center">
        <input type="number" min="1" max="50" value="${e.schoolRank}" onchange="updateEntryField('${e.id}', 'schoolRank', this.value)"
               class="w-12 px-1 py-0.5 text-xs border border-slate-300 rounded text-center font-bold" />
      </td>
      <td ondblclick="startEditEntry('${e.id}')" class="py-1.5 px-2.5 font-bold text-slate-800 cursor-pointer" title="ダブルクリックまたは✎ボタンで修正">
        ${escapeHtml(e.p1)}${e.p2 ? '・' + escapeHtml(e.p2) : ''}
      </td>
      <td ondblclick="startEditEntry('${e.id}')" class="py-1.5 px-2.5 text-slate-700 cursor-pointer" title="ダブルクリックまたは✎ボタンで修正">
        <span>${escapeHtml(e.school)}</span>
        <span class="text-[10px] text-teal-800 bg-teal-50 border border-teal-200 px-1.5 py-0.2 rounded ml-1 font-bold">${escapeHtml(rankTag)}</span>
      </td>
      <td class="py-1.5 px-2.5">
        <select onchange="updateEntryField('${e.id}', 'seed', this.value)"
                class="px-1.5 py-0.5 text-xs border rounded ${e.seed ? 'bg-amber-50 border-amber-400 text-amber-900 font-bold' : 'border-slate-300 text-slate-600'}">
          ${seedOptions.join('')}
        </select>
      </td>
      <td class="py-1.5 px-2 text-center whitespace-nowrap">
        <button onclick="startEditEntry('${e.id}')" class="text-blue-600 hover:bg-blue-50 font-bold px-1.5 py-0.5 rounded transition" title="このペアを修正">✎</button>
        <button onclick="removeEntry('${e.id}')" class="text-slate-400 hover:text-rose-600 hover:bg-rose-50 font-bold px-1.5 py-0.5 rounded transition" title="削除">✕</button>
      </td>
    </tr>`;
  }).join('');
}

function toggleRegSchoolFilter(sc) {
  const sel = document.getElementById('reg-school-filter');
  sel.value = (sel.value === sc) ? '' : sc;
  renderRegistrationTable();
}

// =========================================================================
// トーナメント自動生成
// =========================================================================
function generateDummySchoolNames(schoolCount) {
  // 実在しない架空の都市名 ＋ 第一〜第六中学校（最大66校分）
  const fictionalCities = ['星見', '風見', '蒼丘', '楓野', '白雲', '翠川', '月見丘', '霞野', '朝霧', '陽向', '雲海'];
  const nums = ['第一', '第二', '第三', '第四', '第五', '第六'];
  const names = [];
  for (let i = 0; i < schoolCount; i++) {
    const city = fictionalCities[Math.floor(i / nums.length) % fictionalCities.length];
    const num = nums[i % nums.length];
    names.push(`${city}${num}中学校`);
  }
  return names;
}

function buildDummyFrame(recordHistory = true) {
  const count = Math.max(2, Math.min(256, parseInt(document.getElementById('entry-count-input').value, 10) || 15));
  const rawSchools = parseInt(document.getElementById('dummy-school-count-input')?.value, 10) || 6;
  const schoolCount = Math.max(1, Math.min(64, Math.min(count, rawSchools)));
  if (recordHistory) pushHistory();

  const schools = generateDummySchoolNames(schoolCount);
  const dummies = [];
  for (let i = 1; i <= count; i++) {
    const sc = schools[(i - 1) % schools.length];
    dummies.push({
      id: 'dummy_' + i,
      serial: i,
      schoolRank: Math.ceil(i / schools.length),
      p1: `Player ${i}`,
      p2: '',
      school: sc,
      seed: null
    });
  }
  state.entries = dummies;
  state.schoolOrder = [...schools];
  state.selectedSchoolIdx = null;
  normalizeSerials();
  updateRegCountPill();
  renderSchoolAliasEditor();
  renderRegistrationTable();
  generateTournamentFromEntries(false, true);
  switchMainView('tournament');
  showToast(`${count}組・${schoolCount}校の仮組トーナメントを作成しました`, 'success');
}

function generateAndOpenTournament() {
  if (state.entries.length < 2) {
    alert('2ペア以上の生徒を登録してからトーナメントを作成してください（「サンプル15組読込」もご利用いただけます）。');
    return;
  }
  generateTournamentFromEntries(true, true);
  switchMainView('tournament');
}

function regenerateCurrentBracket(shuffleSchools = true) {
  if (state.entries.length < 2) return;
  if (Object.keys(state.results).length > 0) {
    if (!confirm('入力済みの試合スコアはリセットされます。配置を再生成しますか？')) return;
  }
  if (shuffleSchools) {
    shuffleSchoolOrder(false);
  }
  generateTournamentFromEntries(true, true);
  renderBench();
  renderTourneySidebar();
  renderSheet();
}

function generateTournamentFromEntries(recordHistory = true, notifyUser = false) {
  if (state.entries.length < 2) return;
  if (recordHistory) pushHistory();

  syncSchoolOrder();
  const generated = BracketEngine.generateBracket(state.entries, state.rules, {
    schoolOrder: state.schoolOrder
  });
  state.bracketSize = generated.bracketSize;
  state.activeSlotMask = generated.activeSlotMask;
  state.slots = generated.slots;
  state.schoolOrder = generated.schoolOrder || state.schoolOrder;
  state.results = {};
  state.slotFills = {};
  state.bench = createEmptyBench();
  state.selectedTarget = null;
  if (state.oop) {
    state.oop.courts = {};
    state.oop.statuses = {};
    state.oop.umpires = {};
  }

  refreshAlerts(notifyUser);
  saveToLocalStorage();
}

// =========================================================================
// 3種通知チェック（1回戦同校 / 1回戦同一カード3対戦以上 / 上位4番手規則性外）
// =========================================================================
function refreshAlerts(notifyOnIssue = false) {
  state.alerts = BracketEngine.analyzeBracketAlerts(state.slots, state.activeSlotMask, state.bracketSize);
  const { round1SameSchool, tripleSameCards, top4Violations, round2SameSchool } = state.alerts;

  const criticalCount = round1SameSchool.length + tripleSameCards.length + top4Violations.length;
  const badge = document.getElementById('btn-conflict-badge');
  const listItems = document.getElementById('conflict-list-items');
  const drawer = document.getElementById('conflict-drawer');

  if (badge) {
    if (criticalCount > 0) {
      const hasR1 = round1SameSchool.length > 0;
      badge.className = hasR1
        ? 'px-2.5 py-1 text-xs rounded-md font-bold inline-flex items-center gap-1.5 border bg-rose-100 hover:bg-rose-200 text-rose-800 border-rose-300 transition'
        : 'px-2.5 py-1 text-xs rounded-md font-bold inline-flex items-center gap-1.5 border bg-orange-100 hover:bg-orange-200 text-orange-900 border-orange-300 transition';
      const parts = [];
      if (round1SameSchool.length > 0) parts.push(`1回戦同校 ${round1SameSchool.length}`);
      if (tripleSameCards.length > 0) parts.push(`同一カード3戦+ ${tripleSameCards.length}`);
      if (top4Violations.length > 0) parts.push(`上位4番手外 ${top4Violations.length}`);
      badge.innerHTML = `<span>⚠ 通知: ${parts.join(' / ')}</span>`;
    } else if (round2SameSchool.length > 0) {
      badge.className = 'px-2.5 py-1 text-xs rounded-md font-bold inline-flex items-center gap-1.5 border bg-amber-50 hover:bg-amber-100 text-amber-800 border-amber-300 transition';
      badge.innerHTML = `<span>✓ 規則OK</span><span class="text-amber-700 font-semibold">(2回戦接近 ${round2SameSchool.length}件)</span>`;
    } else {
      badge.className = 'px-2.5 py-1 text-xs rounded-md font-bold inline-flex items-center gap-1 border bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border-emerald-200 transition';
      badge.innerHTML = `<span>✓ 通知なし・規則OK</span>`;
    }
  }

  // ドロワー内詳細レンダリング
  if (listItems) {
    if (criticalCount === 0 && round2SameSchool.length === 0) {
      listItems.innerHTML = `<div class="text-emerald-700 font-semibold">1回戦同校対決・同一カード3対戦以上・上位4番手規則性違反はいずれも発生していません。</div>`;
    } else {
      let html = '';

      // 1. 1回戦同校対決
      if (round1SameSchool.length > 0) {
        html += `<div class="p-2 rounded-lg bg-rose-50 border border-rose-200">
          <div class="text-[11px] font-bold text-rose-800 mb-1">【警告】1回戦の同校対決（${round1SameSchool.length}件）</div>
          <div class="flex flex-wrap gap-1.5">
            ${round1SameSchool.map((c) => `
              <button onclick="jumpToSlot(${c.slotA})" class="px-2 py-1 bg-white hover:bg-rose-100 border border-rose-300 rounded text-rose-900 font-semibold transition">
                ${escapeHtml(formatSchoolRankBadge(c.p1))} ${escapeHtml(formatPairLabel(c.p1, false))} vs ${escapeHtml(formatSchoolRankBadge(c.p2))} ${escapeHtml(formatPairLabel(c.p2, false))}
              </button>`).join('')}
          </div>
        </div>`;
      }

      // 2. 1回戦同一カード3対戦以上
      if (tripleSameCards.length > 0) {
        html += `<div class="p-2 rounded-lg bg-purple-50 border border-purple-200">
          <div class="text-[11px] font-bold text-purple-900 mb-1">【警告】1回戦 同一カード対決が3対戦以上（${tripleSameCards.length}カード）</div>
          <div class="space-y-1.5">
            ${tripleSameCards.map((card) => `
              <div class="flex flex-wrap items-center gap-1.5">
                <span class="font-bold text-purple-900 bg-purple-100 px-2 py-0.5 rounded">${escapeHtml(formatSchoolShortName(card.schoolA))} vs ${escapeHtml(formatSchoolShortName(card.schoolB))}（${card.matches.length}対戦）</span>
                ${card.matches.map((m, idx) => `
                  <button onclick="jumpToSlot(${m.slotA})" class="px-2 py-0.5 bg-white hover:bg-purple-100 border border-purple-300 rounded text-purple-900 font-medium transition">
                    対戦${idx + 1}: ${escapeHtml(formatSchoolRankBadge(m.p1))} vs ${escapeHtml(formatSchoolRankBadge(m.p2))}
                  </button>`).join('')}
              </div>`).join('')}
          </div>
        </div>`;
      }

      // 3. 上位4番手規則性ルール外（シード起因含む）
      if (top4Violations.length > 0) {
        html += `<div class="p-2 rounded-lg bg-orange-50 border border-orange-200">
          <div class="text-[11px] font-bold text-orange-900 mb-1">【通知】上位4番手の規則性ルール外（${top4Violations.length}件）</div>
          <div class="flex flex-wrap gap-1.5">
            ${top4Violations.map((v) => {
              const sideJa = (s) => (s === 'in' ? 'in(内)' : 'out(外)');
              const seedTag = v.isMultiSeedException && v.entry.seed
                ? `<span class="ml-1 px-1 py-0.2 bg-indigo-200 text-indigo-950 rounded text-[10px] font-bold">複数シード特例適用（残山・in/out調整済）</span>`
                : v.causedBySeed
                  ? `<span class="ml-1 px-1 py-0.2 bg-amber-300 text-slate-950 rounded text-[10px] font-bold">シード配置による例外</span>`
                  : '';
              return `<button onclick="jumpToSlot(${v.slotIdx})" class="px-2 py-1 bg-white hover:bg-orange-100 border border-orange-300 rounded text-orange-950 font-semibold text-left transition">
                <span>${escapeHtml(formatSchoolRankBadge(v.entry))}（${escapeHtml(formatPairLabel(v.entry, false))}）</span>
                <span class="text-[10px] text-orange-800 ml-1">現在: <b>${v.actualCode}(${sideJa(v.actualSide)})</b> ／ 標準目標: <b>${v.expectedCode}(${sideJa(v.expectedSide)})</b> [1番手=${v.anchorCode}]</span>
                ${seedTag}
              </button>`;
            }).join('')}
          </div>
        </div>`;
      }

      // 4. 参考：2回戦同校接近
      if (round2SameSchool.length > 0) {
        html += `<div class="p-2 rounded-lg bg-amber-50/70 border border-amber-200">
          <div class="text-[11px] font-bold text-amber-800 mb-1">【参考】2回戦（4枠内）での同校接近（${round2SameSchool.length}件）</div>
          <div class="flex flex-wrap gap-1.5">
            ${round2SameSchool.map((c) => `
              <button onclick="jumpToSlot(${c.slotA})" class="px-2 py-0.5 bg-white hover:bg-amber-100 border border-amber-300 rounded text-amber-900 font-medium transition">
                ${escapeHtml(formatSchoolRankBadge(c.p1))} / ${escapeHtml(formatSchoolRankBadge(c.p2))}${c.directMatch ? ' [初戦直対]' : ''}
              </button>`).join('')}
          </div>
        </div>`;
      }

      listItems.innerHTML = html;
    }
  }

  // 重要な通知がある場合は自動でトースト表示＆ドロワー展開
  if (notifyOnIssue) {
    if (criticalCount > 0) {
      if (drawer) drawer.classList.remove('hidden');
      const msgs = [];
      if (round1SameSchool.length > 0) msgs.push(`1回戦同校対決 ${round1SameSchool.length}件`);
      if (tripleSameCards.length > 0) msgs.push(`1回戦同一カード3対戦以上 ${tripleSameCards.length}件`);
      if (top4Violations.length > 0) {
        const seedCaused = top4Violations.some((v) => v.causedBySeed);
        msgs.push(`上位4番手ルール外 ${top4Violations.length}件${seedCaused ? '（シード要因含む）' : ''}`);
      }
      showToast(`通知：${msgs.join(' / ')}`, round1SameSchool.length > 0 ? 'error' : 'warn', 4200);
    } else {
      if (drawer) drawer.classList.add('hidden');
      showToast('規則通りに配置しました（通知・違反なし）', 'success');
    }
  }
}

function toggleConflictDrawer() {
  document.getElementById('conflict-drawer').classList.toggle('hidden');
}

// =========================================================================
// 右サイド控え枠（三角トレード用）＆ 枠入替・D&D
// =========================================================================
function renderBench() {
  const container = document.getElementById('bench-slots-container');
  const usedPill = document.getElementById('bench-used-pill');
  const totalLabel = document.getElementById('bench-total-label');
  if (!container) return;

  const usedCount = state.bench.filter(Boolean).length;
  if (totalLabel) totalLabel.textContent = `${state.bench.length}枠`;
  if (usedPill) {
    if (usedCount > 0) {
      usedPill.classList.remove('hidden');
      usedPill.textContent = String(usedCount);
    } else {
      usedPill.classList.add('hidden');
    }
  }

  container.innerHTML = state.bench.map((item, idx) => {
    const isSelected = state.selectedTarget && state.selectedTarget.type === 'bench' && state.selectedTarget.index === idx;
    if (item) {
      return `<div onclick="handleTargetClick('bench', ${idx})"
                   draggable="true"
                   ondragstart="handleDragStart(event, 'bench', ${idx})"
                   ondragover="handleDragOver(event)"
                   ondragleave="handleDragLeave(event)"
                   ondrop="handleDrop(event, 'bench', ${idx})"
                   class="bench-slot ${isSelected ? 'selected' : ''} flex-1 min-h-[44px] px-2 py-1.5 bg-white border border-amber-400 rounded-lg text-xs cursor-pointer shadow-sm flex flex-col justify-center gap-0.5">
        <div class="flex items-center justify-between gap-1">
          <span class="text-[10px] font-bold text-amber-700 bg-amber-100 px-1.5 py-0.2 rounded">控え${idx + 1}</span>
          <span class="text-[10px] font-bold text-slate-600 bg-slate-100 border border-slate-200 px-1.5 py-0.2 rounded">${escapeHtml(formatSchoolRankBadge(item))}</span>
        </div>
        <div class="font-bold text-slate-900 leading-tight truncate" title="${escapeHtml(formatPairLabel(item, true))}">
          ${escapeHtml(formatPairLabel(item, false))}
        </div>
      </div>`;
    }
    return `<div onclick="handleTargetClick('bench', ${idx})"
                 ondragover="handleDragOver(event)"
                 ondragleave="handleDragLeave(event)"
                 ondrop="handleDrop(event, 'bench', ${idx})"
                 class="bench-slot ${isSelected ? 'selected' : ''} flex-1 min-h-[36px] px-2.5 py-1.5 bg-white/65 border border-dashed border-amber-300 hover:border-amber-500 rounded-lg text-[11px] text-amber-800/75 cursor-pointer flex items-center justify-between">
      <span class="font-bold">控え ${idx + 1}</span>
      <span class="text-[10px] opacity-75">空き</span>
    </div>`;
  }).join('');
}

function handleTargetClick(type, index) {
  if (state.appMode === 'paint') {
    if (type === 'slot' && state.activeSlotMask[index]) {
      applySlotFill(index);
    }
    return;
  }
  if (state.appMode !== 'edit') return;
  if (type === 'slot' && !state.activeSlotMask[index]) return;

  if (!state.selectedTarget) {
    const item = (type === 'slot') ? state.slots[index] : state.bench[index];
    if (!item) return;
    state.selectedTarget = { type, index };
    renderBench();
    renderSheet();
    return;
  }

  if (state.selectedTarget.type === type && state.selectedTarget.index === index) {
    state.selectedTarget = null;
    renderBench();
    renderSheet();
    return;
  }

  swapTargets(state.selectedTarget, { type, index });
  state.selectedTarget = null;
}

let dragSource = null;
function handleDragStart(e, type, index) {
  if (state.appMode !== 'edit') {
    e.preventDefault();
    return;
  }
  dragSource = { type, index };
  e.dataTransfer.effectAllowed = 'move';
}
function handleDragOver(e) {
  if (state.appMode !== 'edit') return;
  e.preventDefault();
  e.currentTarget.classList.add('drag-over-target');
}
function handleDragLeave(e) {
  e.currentTarget.classList.remove('drag-over-target');
}
function handleDrop(e, targetType, targetIndex) {
  e.preventDefault();
  e.currentTarget.classList.remove('drag-over-target');
  if (!dragSource) return;
  if (targetType === 'slot' && !state.activeSlotMask[targetIndex]) return;
  if (dragSource.type === targetType && dragSource.index === targetIndex) return;

  swapTargets(dragSource, { type: targetType, index: targetIndex });
  dragSource = null;
  state.selectedTarget = null;
}

function swapTargets(src, dst) {
  if (Object.keys(state.results).length > 0) {
    if (!confirm('枠を入れ替えると入力済みの勝敗スコアがクリアされます。続けますか？')) return;
  }
  pushHistory();
  const getVal = (t) => (t.type === 'slot') ? state.slots[t.index] : state.bench[t.index];
  const setVal = (t, val) => {
    if (t.type === 'slot') state.slots[t.index] = val;
    else state.bench[t.index] = val;
  };

  const v1 = getVal(src);
  const v2 = getVal(dst);
  setVal(src, v2);
  setVal(dst, v1);

  state.results = {};
  const prevCritical =
    state.alerts.round1SameSchool.length +
    state.alerts.tripleSameCards.length +
    state.alerts.top4Violations.length;
  refreshAlerts(false);
  const nextCritical =
    state.alerts.round1SameSchool.length +
    state.alerts.tripleSameCards.length +
    state.alerts.top4Violations.length;
  if (nextCritical > prevCritical) {
    refreshAlerts(true);
  }
  saveToLocalStorage();
  renderBench();
  renderTourneySidebar();
  renderSheet();
}

// =========================================================================
// 左サイドバー（配置済み生徒一覧）
// =========================================================================
function renderTourneySidebar() {
  const listEl = document.getElementById('sidebar-roster-list');
  const selectEl = document.getElementById('sidebar-school-select');
  const badgeEl = document.getElementById('sidebar-stats-badge');
  const filterText = (document.getElementById('sidebar-filter-input')?.value || '').trim().toLowerCase();
  if (!listEl) return;

  const displayNumMap = {};
  let dNum = 1;
  for (let i = 0; i < state.bracketSize; i++) {
    if (state.activeSlotMask[i]) displayNumMap[i] = dNum++;
  }

  const placedMap = new Map();
  state.slots.forEach((s, idx) => {
    if (s) {
      const bInfo = BracketEngine.getBlockInfo(idx, state.bracketSize);
      placedMap.set(s.id, {
        location: `No.${displayNumMap[idx] || (idx + 1)}`,
        slotIdx: idx,
        block: bInfo.code,
        sideJa: bInfo.side === 'in' ? '内' : '外',
        isBye: BracketEngine.isByeSlot(idx, state.activeSlotMask)
      });
    }
  });
  state.bench.forEach((b, idx) => {
    if (b) placedMap.set(b.id, { location: `控え${idx + 1}`, slotIdx: -1, block: '控え', sideJa: '', isBye: false });
  });

  const schools = [...new Set(state.entries.map((e) => e.school))];
  const prevSc = selectEl.value;
  selectEl.innerHTML = `<option value="">すべての学校 (${schools.length}校)</option>` +
    schools.map((sc) => `<option value="${escapeHtml(sc)}">${escapeHtml(formatSchoolShortName(sc))}（${escapeHtml(sc)}）</option>`).join('');
  if (schools.includes(prevSc)) selectEl.value = prevSc;

  const selectedSchool = selectEl.value;
  const filtered = state.entries.filter((e) => {
    if (selectedSchool && e.school !== selectedSchool) return false;
    if (filterText) {
      const info = placedMap.get(e.id);
      const hay = `${e.p1} ${e.p2} ${e.school} ${formatSchoolShortName(e.school)} ${formatSchoolRankBadge(e)} ${info ? info.location : ''}`.toLowerCase();
      if (!hay.includes(filterText)) return false;
    }
    return true;
  });

  badgeEl.textContent = `${filtered.length}/${state.entries.length}組`;

  if (filtered.length === 0) {
    listEl.innerHTML = `<div class="py-8 text-center text-slate-400">該当する出場者がいません</div>`;
    return;
  }

  listEl.innerHTML = filtered.map((e) => {
    const info = placedMap.get(e.id);
    const locLabel = info ? info.location : '未配置';
    const blockLabel = info ? (info.sideJa ? `${info.block}・${info.sideJa}` : info.block) : '-';
    const canJump = info && info.slotIdx >= 0;

    return `<div onclick="${canJump ? `jumpToSlot(${info.slotIdx})` : ''}"
                 class="p-2 rounded-lg border border-slate-200 hover:border-blue-400 bg-slate-50/70 hover:bg-blue-50/60 transition ${canJump ? 'cursor-pointer' : ''}">
      <div class="flex items-center justify-between gap-1 mb-0.5">
        <div class="flex items-center gap-1">
          <span class="text-[10px] font-bold px-1.5 py-0.2 rounded ${info?.slotIdx >= 0 ? 'bg-blue-100 text-blue-800' : 'bg-amber-100 text-amber-800'}">
            ${locLabel}
          </span>
          <span class="text-[9px] font-semibold text-slate-500">[${blockLabel}]</span>
          ${info?.isBye ? `<span class="text-[9px] font-bold px-1 rounded bg-emerald-100 text-emerald-800">BYE</span>` : ''}
        </div>
        <span class="text-[10px] font-bold text-teal-800 bg-teal-50 border border-teal-200 px-1 rounded">${escapeHtml(formatSchoolRankBadge(e))}</span>
      </div>
      <div class="font-bold text-slate-800 truncate flex items-center gap-1">
        <span class="truncate">${escapeHtml(e.p1)}${e.p2 ? '・' + escapeHtml(e.p2) : ''}</span>
        ${e.seed ? `<span class="text-[9px] bg-amber-400 text-slate-950 px-1 rounded font-bold shrink-0">S${e.seed}</span>` : ''}
      </div>
    </div>`;
  }).join('');
}

function jumpToSlot(slotIdx, temporaryFlash = true) {
  const el = document.querySelector(`[data-slot-idx="${slotIdx}"]`);
  if (el) {
    el.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
    if (temporaryFlash) {
      el.classList.add('jump-flash-slot');
      setTimeout(() => el.classList.remove('jump-flash-slot'), 1800);
    }
  }
}

// =========================================================================
// 盤面検索・ズーム
// =========================================================================
function handleBoardSearch(q) {
  state.searchQuery = q.trim().toLowerCase();
  state.searchMatches = [];
  state.searchCursor = 0;

  if (state.searchQuery) {
    state.slots.forEach((s, idx) => {
      if (!s) return;
      const hay = `${idx + 1} ${s.p1} ${s.p2} ${s.school} ${formatSchoolShortName(s.school)} ${formatSchoolRankBadge(s)}`.toLowerCase();
      if (hay.includes(state.searchQuery)) state.searchMatches.push(idx);
    });
  }

  const countEl = document.getElementById('board-search-count');
  if (countEl) countEl.textContent = state.searchQuery ? `${state.searchMatches.length}件` : '';
  renderSheet();
  if (state.searchMatches.length > 0) jumpToSlot(state.searchMatches[0], false);
}

function handleBoardSearchKey(e) {
  if (e.key === 'Enter') {
    e.preventDefault();
    stepBoardSearch(e.shiftKey ? -1 : 1);
  }
}

function stepBoardSearch(delta) {
  if (state.searchMatches.length === 0) return;
  state.searchCursor = (state.searchCursor + delta + state.searchMatches.length) % state.searchMatches.length;
  renderSheet();
  jumpToSlot(state.searchMatches[state.searchCursor], false);
}

function changeZoom(delta) {
  state.zoom = Math.max(0.4, Math.min(1.5, Math.round((state.zoom + delta) * 100) / 100));
  applyZoomTransform();
}

function resetZoom100() {
  state.zoom = 1.0;
  applyZoomTransform();
}

function fitZoomToScreen() {
  const viewport = document.getElementById('canvas-viewport');
  const sheet = document.getElementById('tournament-sheet');
  if (!viewport || !sheet) return;
  const vw = viewport.clientWidth - 40;
  const sw = sheet.offsetWidth || 1000;
  state.zoom = Math.max(0.4, Math.min(1.0, Math.floor((vw / sw) * 100) / 100));
  applyZoomTransform();
}

function applyZoomTransform() {
  const sheet = document.getElementById('tournament-sheet');
  const label = document.getElementById('zoom-label');
  if (sheet) sheet.style.transform = `scale(${state.zoom})`;
  if (label) label.textContent = `${Math.round(state.zoom * 100)}%`;
}

// =========================================================================
// モード切替・任意塗りつぶし・印刷時塗りつぶしトグル・敗者復活戦設定
// =========================================================================
function renderPaintPalette() {
  const box = document.getElementById('paint-palette-buttons');
  if (!box) return;
  box.innerHTML = FILL_COLORS.map((c) => {
    const active = state.activePaintColor === c.key;
    const isEraser = c.key === 'none';
    return `<button onclick="selectPaintColor('${c.key}')"
                    class="px-2 py-0.5 rounded-md text-[11px] font-bold border transition flex items-center gap-1 ${active ? 'ring-2 ring-slate-900 ring-offset-1 shadow-sm' : 'opacity-85 hover:opacity-100'}"
                    style="background-color:${c.hex}; border-color:${c.border}; color:#1e293b;"
                    title="${c.label}">
      ${isEraser ? '✕ 解除' : c.label}
    </button>`;
  }).join('');
}

function selectPaintColor(colorKey) {
  state.activePaintColor = colorKey;
  renderPaintPalette();
}

function applySlotFill(slotIdx) {
  pushHistory();
  const key = String(slotIdx);
  const current = state.slotFills[key];
  if (state.activePaintColor === 'none' || current === state.activePaintColor) {
    delete state.slotFills[key];
  } else {
    state.slotFills[key] = state.activePaintColor;
  }
  saveToLocalStorage();
  renderSheet();
}

function clearAllSlotFills() {
  if (Object.keys(state.slotFills).length === 0) return;
  pushHistory();
  state.slotFills = {};
  saveToLocalStorage();
  renderSheet();
  showToast('塗りつぶしをすべて解除しました');
}

function togglePrintFills(forceVal, notify = true) {
  state.printShowFills = (typeof forceVal === 'boolean') ? forceVal : !state.printShowFills;
  document.body.classList.toggle('print-hide-fills', !state.printShowFills);

  const chk = document.getElementById('chk-print-fills');
  if (chk) chk.checked = state.printShowFills;

  const btn = document.getElementById('btn-print-fills-toggle');
  const stEl = document.getElementById('print-fills-header-state');
  if (btn && stEl) {
    stEl.textContent = state.printShowFills ? 'ON' : 'OFF';
    btn.className = state.printShowFills
      ? 'hidden md:inline-flex items-center gap-1 px-2 py-1 text-[11px] rounded-md font-bold bg-amber-500/20 hover:bg-amber-500/30 text-amber-200 border border-amber-400/50 transition'
      : 'hidden md:inline-flex items-center gap-1 px-2 py-1 text-[11px] rounded-md font-bold bg-slate-800 hover:bg-slate-700 text-slate-400 border border-slate-700 transition';
  }
  saveToLocalStorage();
  if (notify) {
    showToast(`印刷時の塗りつぶし表示を ${state.printShowFills ? 'ON' : 'OFF'} にしました`);
  }
}

function setAppMode(mode) {
  state.appMode = mode;
  state.selectedTarget = null;
  const btnEdit = document.getElementById('btn-mode-edit');
  const btnPaint = document.getElementById('btn-mode-paint');
  const btnMatch = document.getElementById('btn-mode-match');
  const benchSidebar = document.getElementById('bench-sidebar');
  const paintBar = document.getElementById('paint-mode-bar');
  const matchBar = document.getElementById('match-mode-bar');

  const inactiveCls = 'px-2.5 py-1 text-xs rounded font-bold text-slate-600 hover:text-slate-900 transition';
  if (btnEdit) btnEdit.className = mode === 'edit' ? 'px-2.5 py-1 text-xs rounded font-bold bg-blue-600 text-white shadow-sm transition' : inactiveCls;
  if (btnPaint) btnPaint.className = mode === 'paint' ? 'px-2.5 py-1 text-xs rounded font-bold bg-amber-500 text-slate-950 shadow-sm transition' : inactiveCls;
  if (btnMatch) btnMatch.className = mode === 'match' ? 'px-2.5 py-1 text-xs rounded font-bold bg-rose-600 text-white shadow-sm transition' : inactiveCls;

  document.body.classList.toggle('paint-mode-active', mode === 'paint');
  paintBar?.classList.toggle('hidden', mode !== 'paint');
  matchBar?.classList.toggle('hidden', mode !== 'match');

  if (benchSidebar) {
    if (mode === 'edit') benchSidebar.classList.toggle('collapsed', state.benchCollapsed);
    else benchSidebar.classList.add('collapsed');
  }
  if (mode === 'paint') renderPaintPalette();
  const fmtLbl = document.getElementById('match-mode-format-label');
  if (fmtLbl) fmtLbl.textContent = `${state.matchFormat}ゲームマッチ`;
  updatePanelToggleButtons();
  renderBench();
  renderSheet();
}

function setLayoutMode(mode) {
  state.layoutMode = mode;
  const btnDual = document.getElementById('btn-layout-dual');
  const btnSingle = document.getElementById('btn-layout-single');
  if (btnDual && btnSingle) {
    if (mode === 'dual') {
      btnDual.className = 'px-2 py-0.5 rounded font-bold bg-white text-slate-900 shadow-sm';
      btnSingle.className = 'px-2 py-0.5 rounded font-bold text-slate-600';
    } else {
      btnSingle.className = 'px-2 py-0.5 rounded font-bold bg-white text-slate-900 shadow-sm';
      btnDual.className = 'px-2 py-0.5 rounded font-bold text-slate-600';
    }
  }
  saveToLocalStorage();
  if (state.mainView === 'tournament') renderSheet();
}

function applyConsolationPreset(preset) {
  state.consolation.preset = preset;
  const customBox = document.getElementById('consolation-custom-box');
  if (preset === 'none') {
    customBox.classList.add('hidden');
  } else {
    customBox.classList.remove('hidden');
    if (preset === '3rd') {
      state.consolation.title = '3位決定戦';
      state.consolation.matchCount = 1;
      state.consolation.slots = ['準決勝① 敗者', '準決勝② 敗者', '', '', '', '', '', ''];
    } else if (preset === '5to7') {
      state.consolation.title = '代表決定戦（5〜7位）';
      state.consolation.matchCount = 2;
      state.consolation.slots = ['準々決勝① 敗者', '準々決勝② 敗者', '準々決勝③ 敗者', '準々決勝④ 敗者', '', '', '', ''];
    } else if (preset === '9to16') {
      state.consolation.title = '9〜16位決定戦';
      state.consolation.matchCount = 4;
      state.consolation.slots = [
        'ベスト16① 敗者', 'ベスト16② 敗者', 'ベスト16③ 敗者', 'ベスト16④ 敗者',
        'ベスト16⑤ 敗者', 'ベスト16⑥ 敗者', 'ベスト16⑦ 敗者', 'ベスト16⑧ 敗者'
      ];
    }
    document.getElementById('consolation-title-input').value = state.consolation.title;
    document.getElementById('consolation-match-count').value = String(state.consolation.matchCount);
    renderConsolationEditor();
  }
  saveToLocalStorage();
  if (state.mainView === 'tournament') renderSheet();
}

function updateConsolationConfig() {
  state.consolation.title = document.getElementById('consolation-title-input').value;
  state.consolation.matchCount = parseInt(document.getElementById('consolation-match-count').value, 10) || 1;
  renderConsolationEditor();
  saveToLocalStorage();
  if (state.mainView === 'tournament') renderSheet();
}

function renderConsolationEditor() {
  const box = document.getElementById('consolation-slots-editor');
  const count = state.consolation.matchCount * 2;
  let html = '';
  for (let i = 0; i < count; i++) {
    html += `<input type="text" value="${escapeHtml(state.consolation.slots[i] || '')}"
                    oninput="state.consolation.slots[${i}] = this.value; saveToLocalStorage(); if(state.mainView==='tournament') renderSheet();"
                    placeholder="枠${i + 1} 出場者名"
                    class="w-full px-2.5 py-1 text-xs border border-slate-300 rounded-md" />`;
  }
  box.innerHTML = html;
}

// =========================================================================
// スコア入力モーダル
// =========================================================================
function populateScoreGameSelects() {
  const winG = getWinGames();
  let opts = '';
  for (let g = 0; g <= winG; g++) {
    opts += `<option value="${g}">${g}</option>`;
  }
  const sel1 = document.getElementById('score-p1-games');
  const sel2 = document.getElementById('score-p2-games');
  const btn1 = document.getElementById('btn-quick-win-1');
  const btn2 = document.getElementById('btn-quick-win-2');
  const titleEl = document.getElementById('score-modal-title');
  if (sel1) sel1.innerHTML = opts;
  if (sel2) sel2.innerHTML = opts;
  if (btn1) btn1.textContent = `勝(${winG})`;
  if (btn2) btn2.textContent = `勝(${winG})`;
  if (titleEl) titleEl.textContent = `試合結果・ゲームカウント入力（${state.matchFormat}ゲーム / ${winG}G先取）`;
}

function openScoreModal(matchKey, p1Name, p2Name, isConsolation = false, allowFromOop = false) {
  if (state.appMode !== 'match' && !allowFromOop && state.mainView !== 'oop') return;
  state.activeModalMatch = { matchKey, isConsolation };
  populateScoreGameSelects();
  document.getElementById('score-p1-name').textContent = p1Name || '未定';
  document.getElementById('score-p2-name').textContent = p2Name || '未定';

  const winG = getWinGames();
  const existing = isConsolation ? state.consolation.results[matchKey] : state.results[matchKey];
  const g1Val = existing && !existing.wo ? Math.min(winG, parseInt(existing.g1, 10) || 0) : winG;
  const g2Val = existing && !existing.wo ? Math.min(winG, parseInt(existing.g2, 10) || 0) : 0;
  document.getElementById('score-p1-games').value = String(g1Val);
  document.getElementById('score-p2-games').value = String(g2Val);
  document.getElementById('score-modal').classList.remove('hidden');
}

function closeScoreModal() {
  const modal = document.getElementById('score-modal');
  if (modal) modal.classList.add('hidden');
  state.activeModalMatch = null;
}

function setQuickWinner(side) {
  const winG = getWinGames();
  const sel1 = document.getElementById('score-p1-games');
  const sel2 = document.getElementById('score-p2-games');
  if (side === 1) {
    sel1.value = String(winG);
    if (parseInt(sel2.value, 10) >= winG) {
      sel2.value = String(Math.max(0, winG - 1));
    }
  } else {
    sel2.value = String(winG);
    if (parseInt(sel1.value, 10) >= winG) {
      sel1.value = String(Math.max(0, winG - 1));
    }
  }
  saveMatchScore();
}

function saveMatchScore() {
  if (!state.activeModalMatch) return;
  const g1 = parseInt(document.getElementById('score-p1-games').value, 10);
  const g2 = parseInt(document.getElementById('score-p2-games').value, 10);
  if (g1 === g2) {
    alert('同点では勝敗を確定できません。勝者側のゲーム数が多くなるよう選択してください。');
    return;
  }
  pushHistory();
  const { matchKey, isConsolation } = state.activeModalMatch;
  const winner = (g1 > g2) ? 0 : 1;

  if (isConsolation) {
    state.consolation.results[matchKey] = { winner, g1: String(g1), g2: String(g2) };
  } else {
    state.results[matchKey] = { winner, g1: String(g1), g2: String(g2) };
  }
  closeScoreModal();
  saveToLocalStorage();
  renderSheet();
  if (typeof window.onBracketResultsChanged === 'function') window.onBracketResultsChanged();
}

function clearMatchScore() {
  if (!state.activeModalMatch) return;
  pushHistory();
  const { matchKey, isConsolation } = state.activeModalMatch;
  if (isConsolation) delete state.consolation.results[matchKey];
  else delete state.results[matchKey];
  closeScoreModal();
  saveToLocalStorage();
  renderSheet();
  if (typeof window.onBracketResultsChanged === 'function') window.onBracketResultsChanged();
}

// =========================================================================
// 公式足長トーナメント描画エンジン（決勝の完全対称山・略称＋番手表記）
// =========================================================================
function renderSheet() {
  const sheet = document.getElementById('tournament-sheet');
  if (!sheet) return;

  const size = state.bracketSize;
  const totalRounds = Math.log2(size);
  const showRankBadge = state.showRankBadge;

  const alertSets = {
    r1: new Set(),
    triple: new Set(),
    top4: new Set(),
    r2: new Set()
  };
  state.alerts.round1SameSchool.forEach((c) => { alertSets.r1.add(c.slotA); alertSets.r1.add(c.slotB); });
  state.alerts.tripleSameCards.forEach((card) => {
    card.matches.forEach((m) => { alertSets.triple.add(m.slotA); alertSets.triple.add(m.slotB); });
  });
  state.alerts.top4Violations.forEach((v) => { alertSets.top4.add(v.slotIdx); });
  state.alerts.round2SameSchool.forEach((c) => { alertSets.r2.add(c.slotA); alertSets.r2.add(c.slotB); });

  const tree = [];
  tree[0] = [];
  for (let i = 0; i < size; i++) {
    tree[0][i] = {
      slotIdx: i,
      active: Boolean(state.activeSlotMask[i]),
      entry: state.slots[i] || null
    };
  }

  for (let r = 1; r <= totalRounds; r++) {
    const count = size / Math.pow(2, r);
    tree[r] = [];
    for (let m = 0; m < count; m++) {
      const c0 = tree[r - 1][m * 2];
      const c1 = tree[r - 1][m * 2 + 1];
      const active = c0.active || c1.active;
      const isBye = (c0.active && !c1.active) || (!c0.active && c1.active);
      const matchKey = `${r}_${m}`;
      const res = state.results[matchKey] || null;

      let advEntry = null;
      if (isBye) {
        advEntry = c0.active ? c0.entry : c1.entry;
      } else if (res) {
        advEntry = (res.winner === 0) ? c0.entry : c1.entry;
      }

      tree[r][m] = { round: r, match: m, active, isBye, c0, c1, result: res, entry: advEntry };
    }
  }

  const rowHeight = size >= 64 ? 26 : 32;
  const labelWidth = state.showSlotPriority
    ? (showRankBadge ? 330 : 292)
    : (showRankBadge ? 292 : 254);
  const stepWidth = Math.max(28, Math.min(48, Math.floor(210 / Math.max(1, totalRounds - 1))));

  if (state.bracketBlockFilter && state.bracketBlockFilter !== 'ALL' && size >= 8) {
    renderSingleBlockBracket(sheet, tree, size, totalRounds, rowHeight, labelWidth, stepWidth, showRankBadge, alertSets, state.bracketBlockFilter);
  } else if (state.layoutMode === 'dual' && size >= 4) {
    renderDualWingBracket(sheet, tree, size, totalRounds, rowHeight, labelWidth, stepWidth, showRankBadge, alertSets);
  } else {
    renderSingleSideBracket(sheet, tree, size, totalRounds, rowHeight, labelWidth, stepWidth, showRankBadge, alertSets);
  }
  autoFitSlotTexts(sheet);
  applyZoomTransform();
}

function setBracketBlockFilter(code) {
  state.bracketBlockFilter = ['A', 'B', 'C', 'D'].includes(code) ? code : 'ALL';
  ['ALL', 'A', 'B', 'C', 'D'].forEach((k) => {
    const btn = document.getElementById(`btn-block-filter-${k}`);
    if (!btn) return;
    btn.className = state.bracketBlockFilter === k
      ? 'px-2 py-0.5 text-[11px] font-bold rounded bg-slate-900 text-white shadow-sm transition'
      : 'px-2 py-0.5 text-[11px] font-bold rounded text-slate-600 hover:text-slate-900 transition';
  });
  renderSheet();
}

function renderSingleBlockBracket(sheet, tree, size, totalRounds, rowHeight, labelWidth, stepW, showRankBadge, alertSets, blockCode) {
  const quarterIndexMap = { A: 0, D: 1, C: 2, B: 3 };
  const qIdx = quarterIndexMap[blockCode] ?? 0;
  const quarterSize = size / 4;
  const startSlot = qIdx * quarterSize;
  const endSlot = startSlot + quarterSize;
  const blockRounds = Math.max(1, totalRounds - 2);
  const priorityMap = BracketEngine.buildAdjustedSlotPriorityMap(state.activeSlotMask, size);

  const activeIndices = [];
  for (let i = startSlot; i < endSlot; i++) {
    if (state.activeSlotMask[i]) activeIndices.push(i);
  }
  const slotToDisplayNum = {};
  let dispCounter = 1;
  for (let i = 0; i < size; i++) {
    if (state.activeSlotMask[i]) slotToDisplayNum[i] = dispCounter++;
  }

  const nodeY = {};
  nodeY[0] = {};
  activeIndices.forEach((slotIdx, visualRow) => {
    nodeY[0][slotIdx] = visualRow * rowHeight + rowHeight / 2 + 12;
  });
  for (let r = 1; r <= blockRounds; r++) {
    nodeY[r] = {};
    const mStart = startSlot / Math.pow(2, r);
    const mEnd = endSlot / Math.pow(2, r);
    for (let m = mStart; m < mEnd; m++) {
      const node = tree[r][m];
      const y0 = nodeY[r - 1][m * 2];
      const y1 = nodeY[r - 1][m * 2 + 1];
      if (!node.active) nodeY[r][m] = 0;
      else if (node.isBye) nodeY[r][m] = node.c0.active ? y0 : y1;
      else nodeY[r][m] = (y0 + y1) / 2;
    }
  }

  const totalWidth = labelWidth + blockRounds * stepW + 170;
  const canvasHeight = Math.max(180, activeIndices.length * rowHeight + 28);
  let svgBlackLines = '';
  let svgRedLines = '';
  let svgScoreTexts = '';
  let htmlLabels = '';

  const pushSvgLine = (x1, y1, x2, y2, isRed, width) => {
    const line = `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${isRed ? '#dc2626' : '#0f172a'}" stroke-width="${width || (isRed ? 2.6 : 1.4)}" stroke-linecap="square" />`;
    if (isRed) svgRedLines += line;
    else svgBlackLines += line;
  };

  activeIndices.forEach((slotIdx) => {
    const leaf = tree[0][slotIdx];
    const y = nodeY[0][slotIdx];
    const isSel = state.selectedTarget && state.selectedTarget.type === 'slot' && state.selectedTarget.index === slotIdx;
    const isSearchHit = state.searchMatches.includes(slotIdx);
    const isSearchCurr = state.searchMatches.length > 0 && state.searchMatches[state.searchCursor] === slotIdx;
    const stateCls = getSlotAlertClass(slotIdx, isSel, isSearchCurr, isSearchHit, alertSets, Boolean(leaf.entry));
    const entry = leaf.entry;
    const pRank = priorityMap[slotIdx];
    const pBadge = (state.showSlotPriority && pRank)
      ? `<span class="priority-rank-el text-[9px] font-bold px-1 py-0.2 rounded bg-indigo-50 text-indigo-700 border border-indigo-200 shrink-0">${pRank}</span>`
      : '';
    const sBadge = (state.showSlotPriority && entry && entry.seed)
      ? `<span class="priority-rank-el text-[9px] font-bold px-1 py-0.2 rounded bg-amber-400 text-slate-950 shrink-0">S${entry.seed}</span>`
      : '';
    const rBadge = (showRankBadge && entry)
      ? `<span class="rank-badge-el text-[9px] font-bold px-1 py-0.2 rounded bg-teal-50 text-teal-800 border border-teal-200 shrink-0">${escapeHtml(formatSchoolRankBadge(entry))}</span>`
      : '';

    const pairHtml = entry
      ? `<span class="pair-name-wrap flex-1 min-w-0 text-left pr-1"><span class="pair-name-autofit block whitespace-nowrap leading-snug text-xs font-bold text-slate-900">${escapeHtml(entry.p1)}${entry.p2 ? '・' + escapeHtml(entry.p2) : ''}</span></span>
         <span class="text-[11px] font-semibold text-slate-600 shrink-0">（${escapeHtml(formatSchoolShortName(entry.school))}）</span>`
      : `<span class="flex-1 text-left text-[11px] text-slate-400 italic">空き枠</span>`;

    htmlLabels += `<div data-slot-idx="${slotIdx}"
                        onclick="handleTargetClick('slot', ${slotIdx})"
                        draggable="${state.appMode === 'edit' && Boolean(entry)}"
                        ondragstart="handleDragStart(event, 'slot', ${slotIdx})"
                        ondragover="handleDragOver(event)"
                        ondragleave="handleDragLeave(event)"
                        ondrop="handleDrop(event, 'slot', ${slotIdx})"
                        style="left:0; top:${y - 13}px; width:${labelWidth - 4}px; height:26px;"
                        class="slot-card ${stateCls} absolute flex items-center gap-1 px-1 cursor-pointer">
      ${pBadge}${sBadge}${rBadge}
      <span class="w-6 text-right text-[11px] font-bold text-slate-500 shrink-0">${slotToDisplayNum[slotIdx]}</span>
      ${pairHtml}
    </div>`;
  });

  for (let r = 1; r <= blockRounds; r++) {
    const xCurr = labelWidth + r * stepW;
    const mStart = startSlot / Math.pow(2, r);
    const mEnd = endSlot / Math.pow(2, r);
    for (let m = mStart; m < mEnd; m++) {
      const node = tree[r][m];
      if (!node.active || node.isBye) continue;
      const y0 = nodeY[r - 1][m * 2];
      const y1 = nodeY[r - 1][m * 2 + 1];
      const yMid = nodeY[r][m];
      const xPrev0 = getEffectiveStartX(tree[r - 1][m * 2], r - 1, labelWidth, 1, stepW);
      const xPrev1 = getEffectiveStartX(tree[r - 1][m * 2 + 1], r - 1, labelWidth, 1, stepW);
      const res = node.result;
      const win0 = Boolean(res && res.winner === 0);
      const win1 = Boolean(res && res.winner === 1);
      const reach0 = win0 || hasReachedByWin(node.c0);
      const reach1 = win1 || hasReachedByWin(node.c1);

      pushSvgLine(xPrev0, y0, xCurr, y0, reach0);
      pushSvgLine(xCurr, y0, xCurr, yMid, win0);
      pushSvgLine(xPrev1, y1, xCurr, y1, reach1);
      pushSvgLine(xCurr, yMid, xCurr, y1, win1);

      if (res) {
        svgScoreTexts += `<text x="${xCurr - 9}" y="${y0 - 4}" text-anchor="middle" fill="#dc2626" font-size="11" font-weight="bold">${escapeHtml(res.g1)}</text>`;
        svgScoreTexts += `<text x="${xCurr - 9}" y="${y1 + 12}" text-anchor="middle" fill="#dc2626" font-size="11" font-weight="bold">${escapeHtml(res.g2)}</text>`;
      }
      if (state.appMode === 'match' && node.c0.entry && node.c1.entry) {
        const n0 = formatPairLabel(node.c0.entry, true);
        const n1 = formatPairLabel(node.c1.entry, true);
        htmlLabels += `<button onclick="openScoreModal('${r}_${m}', '${escapeJs(n0)}', '${escapeJs(n1)}', false)"
                               style="left:${xCurr - 10}px; top:${yMid - 10}px; width:20px; height:20px;"
                               class="no-print absolute rounded-full bg-rose-600 text-white text-[10px] font-bold shadow flex items-center justify-center z-10">✎</button>`;
      }
    }
  }

  const blockFinalM = startSlot / Math.pow(2, blockRounds);
  const blockFinalNode = tree[blockRounds][blockFinalM];
  const finalX = labelWidth + blockRounds * stepW;
  const finalY = nodeY[blockRounds][blockFinalM];
  if (blockFinalNode && blockFinalNode.active) {
    const won = Boolean(blockFinalNode.result || blockFinalNode.isBye);
    pushSvgLine(finalX, finalY, finalX + 24, finalY, Boolean(blockFinalNode.result), blockFinalNode.result ? 2.6 : 1.4);
    if (won && blockFinalNode.entry) {
      htmlLabels += `<div style="left:${finalX + 28}px; top:${finalY - 12}px;" class="absolute text-xs font-bold text-rose-700 bg-rose-50 border border-rose-200 px-2 py-0.5 rounded whitespace-nowrap">${blockCode}ブロック勝者: ${escapeHtml(formatPairLabel(blockFinalNode.entry, true))}</div>`;
    }
  }

  sheet.style.width = `${totalWidth + 48}px`;
  sheet.innerHTML = `
    <div class="mb-3 pb-2.5 border-b-2 border-slate-800 flex items-baseline justify-between gap-4">
      <div class="flex items-center gap-2">
        <span class="px-2 py-0.5 rounded bg-slate-900 text-white text-xs font-extrabold">${blockCode}ブロック</span>
        <h2 class="text-lg font-bold tracking-wide text-slate-900">${escapeHtml(state.title)}</h2>
      </div>
      <div class="flex items-center gap-3 text-xs text-slate-600 font-semibold shrink-0">
        <span>${state.matchFormat}ゲームマッチ</span>
        <span>${blockCode}山：${activeIndices.length}ペア</span>
      </div>
    </div>
    <div class="relative" style="width:${totalWidth}px; height:${canvasHeight}px;">
      <svg width="${totalWidth}" height="${canvasHeight}" class="absolute inset-0 pointer-events-none">${svgBlackLines}${svgRedLines}${svgScoreTexts}</svg>
      ${htmlLabels}
    </div>
  `;
}

/**
 * 生徒氏名欄に収まらない長い氏名は、見切れず全文が見えるようフォントサイズを自動縮小
 */
function autoFitSlotTexts(sheet) {
  if (!sheet) return;
  const nodes = sheet.querySelectorAll('.pair-name-autofit');
  nodes.forEach((el) => {
    const parent = el.parentElement;
    if (!parent) return;
    el.style.fontSize = '';
    const maxW = parent.clientWidth;
    const curW = el.scrollWidth;
    if (maxW > 0 && curW > maxW) {
      const baseSize = parseFloat(window.getComputedStyle(el).fontSize) || 12;
      let fitted = Math.max(7.5, Math.floor((baseSize * (maxW / curW)) * 10) / 10);
      el.style.fontSize = `${fitted}px`;
      let guard = 0;
      while (el.scrollWidth > maxW && fitted > 7 && guard < 8) {
        fitted = Math.max(7, Math.round((fitted - 0.5) * 10) / 10);
        el.style.fontSize = `${fitted}px`;
        guard++;
      }
    }
  });
}

function getEffectiveStartX(node, roundIdx, baseX, dir, stepWidth) {
  if (roundIdx === 0) return baseX;
  if (!node.isBye) return baseX + dir * (roundIdx * stepWidth);
  const activeChild = node.c0.active ? node.c0 : node.c1;
  return getEffectiveStartX(activeChild, roundIdx - 1, baseX, dir, stepWidth);
}

/**
 * 子ノードが前ラウンドの勝利によってこの試合まで勝ち上がってきているか
 * （敗退ペアであっても、前ラウンドに勝っていれば次ラウンドのT字接続点まで赤線を伸ばす）
 */
function hasReachedByWin(node) {
  if (!node || node.round === undefined) return false;
  if (node.isBye) {
    const activeChild = node.c0.active ? node.c0 : node.c1;
    return hasReachedByWin(activeChild);
  }
  return Boolean(node.result);
}

function getSlotAlertClass(slotIdx, isSelected, isSearchCurr, isSearchHit, alertSets, hasEntry) {
  const fillKey = state.slotFills[String(slotIdx)];
  const fillCls = fillKey ? `slot-custom-fill fill-${fillKey} ` : '';
  if (isSelected) return `${fillCls}selected-slot `;
  if (isSearchCurr) return `${fillCls}search-current-slot `;
  if (isSearchHit) return `${fillCls}search-hit-slot `;
  if (alertSets.r1.has(slotIdx)) return `${fillCls}same-school-r1-alert `;
  if (alertSets.triple.has(slotIdx)) return `${fillCls}triple-card-alert `;
  if (alertSets.top4.has(slotIdx)) return `${fillCls}top4-rule-alert `;
  if (alertSets.r2.has(slotIdx)) return `${fillCls}same-school-r2-alert `;
  if (!hasEntry) return `${fillCls}empty-slot-placeholder `;
  return fillCls;
}

function renderDualWingBracket(sheet, tree, size, totalRounds, rowHeight, labelWidth, stepWidth, showRankBadge, alertSets) {
  const halfSize = size / 2;
  const halfRounds = totalRounds - 1;

  // Step 1: 左右ウィングそれぞれの生Y座標と決勝接続点（rawFinalY）を事前計算
  const computeWingLayout = (isLeft) => {
    const startSlot = isLeft ? 0 : halfSize;
    const endSlot = isLeft ? halfSize : size;
    const activeIndices = [];
    for (let i = startSlot; i < endSlot; i++) {
      if (state.activeSlotMask[i]) activeIndices.push(i);
    }
    const rawNodeY = Array.from({ length: halfRounds + 1 }, () => ({}));
    activeIndices.forEach((slotIdx, visualRow) => {
      rawNodeY[0][slotIdx] = visualRow * rowHeight + rowHeight / 2;
    });

    for (let r = 1; r <= halfRounds; r++) {
      const mStart = isLeft ? 0 : (size / Math.pow(2, r + 1));
      const mEnd = isLeft ? (size / Math.pow(2, r + 1)) : (size / Math.pow(2, r));
      for (let m = mStart; m < mEnd; m++) {
        const node = tree[r][m];
        if (!node.active) continue;
        if (node.isBye) {
          const activeChildIdx = node.c0.active ? (m * 2) : (m * 2 + 1);
          rawNodeY[r][m] = rawNodeY[r - 1][activeChildIdx];
        } else {
          rawNodeY[r][m] = (rawNodeY[r - 1][m * 2] + rawNodeY[r - 1][m * 2 + 1]) / 2;
        }
      }
    }

    const finalSideM = isLeft ? 0 : 1;
    const rawFinalY = rawNodeY[halfRounds][finalSideM] ?? (rowHeight / 2);
    const rawTopY = activeIndices.length > 0 ? rawNodeY[0][activeIndices[0]] - rowHeight / 2 : 0;
    const rawBottomY = activeIndices.length > 0 ? rawNodeY[0][activeIndices[activeIndices.length - 1]] + rowHeight / 2 : rowHeight;
    return { activeIndices, rawNodeY, rawFinalY, rawTopY, rawBottomY, count: activeIndices.length };
  };

  const leftLayout = computeWingLayout(true);
  const rightLayout = computeWingLayout(false);

  // Step 2: 左右の決勝接続高さ（targetFinalY）を完全一致させ、真っ直ぐな逆T字決勝線にする
  const stemHeight = 40;
  const minFinalY = stemHeight + 54; // 優勝表記分の十分な上余白
  let targetFinalY = Math.max(leftLayout.rawFinalY, rightLayout.rawFinalY, minFinalY);
  let leftShiftY = targetFinalY - leftLayout.rawFinalY;
  let rightShiftY = targetFinalY - rightLayout.rawFinalY;

  const minTopY = Math.min(leftLayout.rawTopY + leftShiftY, rightLayout.rawTopY + rightShiftY);
  if (minTopY < 16) {
    const extraShift = 16 - minTopY;
    leftShiftY += extraShift;
    rightShiftY += extraShift;
    targetFinalY += extraShift;
  }

  const canvasHeight = Math.max(
    leftLayout.rawBottomY + leftShiftY,
    rightLayout.rawBottomY + rightShiftY,
    targetFinalY + 32
  ) + 18;

  const wingLineSpan = halfRounds * stepWidth;
  const centerBridgeWidth = 128; // 中央決勝線＆優勝カードが綺麗に収まる広さ
  const totalWidth = (labelWidth + wingLineSpan) * 2 + centerBridgeWidth;
  const cx = totalWidth / 2;
  const rawPriorityMap = BracketEngine.buildSlotPriorityMap(size);
  const underSeedMap = BracketEngine.buildUnderTop8SeedMap(size, state.activeSlotMask, rawPriorityMap);
  const slotPriorityMap = BracketEngine.buildAdjustedSlotPriorityMap(size, state.activeSlotMask);

  let svgBlackLines = '';
  let svgRedLines = '';
  let svgScoreTexts = '';
  let htmlLabels = '';
  let displayNum = 1;

  const pushSvgLine = (x1, y1, x2, y2, isRed, widthOverride) => {
    const stroke = isRed ? '#dc2626' : '#334155';
    const sw = widthOverride || (isRed ? 2.6 : 1.3);
    const line = `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${stroke}" stroke-width="${sw}" stroke-linecap="square" />`;
    if (isRed) svgRedLines += line;
    else svgBlackLines += line;
  };

  const drawWing = (isLeft, layout, shiftY) => {
    const baseX = isLeft ? labelWidth : (totalWidth - labelWidth);
    const dir = isLeft ? 1 : -1;

    // 選手ラベル描画
    // 左山：番号 / 生徒A・生徒B / （学校略称） [+確認バッジ]
    // 右山：[確認バッジ+] 生徒A・生徒B / （学校略称） / 番号
    layout.activeIndices.forEach((i) => {
      const y = layout.rawNodeY[0][i] + shiftY;
      const entry = state.slots[i];
      const num = displayNum++;
      const isSelected = state.selectedTarget && state.selectedTarget.type === 'slot' && state.selectedTarget.index === i;
      const isSearchHit = state.searchMatches.includes(i);
      const isSearchCurr = state.searchMatches[state.searchCursor] === i;
      const blockInfo = BracketEngine.getBlockInfo(i, size);

      let cls = 'slot-card absolute flex items-center gap-1.5 px-1.5 text-[12px] cursor-pointer ';
      cls += getSlotAlertClass(i, isSelected, isSearchCurr, isSearchHit, alertSets, Boolean(entry));

      const leftPos = isLeft ? 0 : (baseX + 6);
      const widthPx = labelWidth - 8;

      let innerHtml = '';
      if (entry) {
        const pairName = entry.p2 ? `${escapeHtml(entry.p1)}・${escapeHtml(entry.p2)}` : escapeHtml(entry.p1);
        const shortSc = formatSchoolShortName(entry.school);
        const rankText = formatSchoolRankBadge(entry);
        const slotPRank = slotPriorityMap[i];
        const underInfo = underSeedMap[i];
        const underLabel = underInfo ? ` [S${underInfo.seedRank}下・${underInfo.type === 'R1' ? '1回戦' : '2回戦'}]` : '';
        const blockTip = `確認用: ${rankText} ／ 枠優先順位: ${slotPRank}${underLabel} ／ ブロック: ${blockInfo.code} (${blockInfo.side === 'in' ? '内' : '外'})`;
        const priorityTag = state.showSlotPriority
          ? `<span class="priority-rank-el text-[9px] font-bold px-1 py-0.5 leading-tight rounded bg-indigo-50 text-indigo-800 border border-indigo-300 shrink-0 font-mono" title="枠優先順位">${slotPRank}</span>`
          : '';
        const rankBadge = showRankBadge
          ? `<span class="rank-badge-el text-[10px] font-bold px-1.5 py-0.5 leading-tight rounded bg-teal-50 text-teal-900 border border-teal-300 shrink-0" title="${escapeHtml(blockTip)}">${escapeHtml(rankText)}</span>`
          : '';
        // シード表記（S1等）は優先順位表示トグル・配布用印刷非表示と完全同一扱い
        const seedTag = (state.showSlotPriority && entry.seed)
          ? `<span class="priority-rank-el text-[9px] font-bold px-1 py-0.5 leading-tight rounded bg-amber-100 text-amber-900 border border-amber-300 shrink-0">S${entry.seed}</span>`
          : '';

        if (isLeft) {
          innerHtml = `
            <span class="font-mono text-[11px] font-bold text-slate-500 w-6 text-right shrink-0">${num}</span>
            <span class="pair-name-wrap flex-1 min-w-0 text-left leading-snug"><span class="pair-name-autofit block whitespace-nowrap font-bold text-slate-900">${pairName}</span></span>
            <span class="text-[11px] font-semibold text-slate-600 shrink-0 leading-snug">（${escapeHtml(shortSc)}）</span>
            ${seedTag}${rankBadge}${priorityTag}
          `;
        } else {
          innerHtml = `
            ${priorityTag}${rankBadge}${seedTag}
            <span class="pair-name-wrap flex-1 min-w-0 text-left leading-snug"><span class="pair-name-autofit block whitespace-nowrap font-bold text-slate-900">${pairName}</span></span>
            <span class="text-[11px] font-semibold text-slate-600 shrink-0 leading-snug">（${escapeHtml(shortSc)}）</span>
            <span class="font-mono text-[11px] font-bold text-slate-500 w-6 text-left shrink-0">${num}</span>
          `;
        }
      } else {
        const slotPRank = slotPriorityMap[i];
        const priorityTag = state.showSlotPriority
          ? `<span class="priority-rank-el text-[9px] font-bold px-1 py-0.2 rounded bg-indigo-50 text-indigo-800 border border-indigo-300 shrink-0 font-mono">${slotPRank}</span>`
          : '';
        innerHtml = isLeft
          ? `<span class="font-mono text-[11px] text-slate-400 w-6 text-right shrink-0">${num}</span><span class="text-[11px] text-slate-400 italic flex-1 text-left">（空き枠）</span>${priorityTag}`
          : `${priorityTag}<span class="text-[11px] text-slate-400 italic flex-1 text-left">（空き枠）</span><span class="font-mono text-[11px] text-slate-400 w-6 text-left shrink-0">${num}</span>`;
      }

      htmlLabels += `<div data-slot-idx="${i}"
                          onclick="handleTargetClick('slot', ${i})"
                          draggable="true"
                          ondragstart="handleDragStart(event, 'slot', ${i})"
                          ondragover="handleDragOver(event)"
                          ondragleave="handleDragLeave(event)"
                          ondrop="handleDrop(event, 'slot', ${i})"
                          style="left:${leftPos}px; top:${y - 13}px; width:${widthPx}px; height:26px;"
                          class="${cls}">${innerHtml}</div>`;
    });

    // ウィング内トーナメント線描画
    for (let r = 1; r <= halfRounds; r++) {
      const mStart = isLeft ? 0 : (size / Math.pow(2, r + 1));
      const mEnd = isLeft ? (size / Math.pow(2, r + 1)) : (size / Math.pow(2, r));
      const xCurr = baseX + dir * (r * stepWidth);

      for (let m = mStart; m < mEnd; m++) {
        const node = tree[r][m];
        if (!node.active || node.isBye) continue;

        const y0 = layout.rawNodeY[r - 1][m * 2] + shiftY;
        const y1 = layout.rawNodeY[r - 1][m * 2 + 1] + shiftY;
        const yMid = layout.rawNodeY[r][m] + shiftY;

        const xPrev0 = getEffectiveStartX(tree[r - 1][m * 2], r - 1, baseX, dir, stepWidth);
        const xPrev1 = getEffectiveStartX(tree[r - 1][m * 2 + 1], r - 1, baseX, dir, stepWidth);

        const res = node.result;
        const win0 = Boolean(res && res.winner === 0);
        const win1 = Boolean(res && res.winner === 1);
        // 敗退ペアでも前試合に勝っていれば、この試合の接続点（縦線・T字）まで赤線を伸ばす
        const reach0 = win0 || hasReachedByWin(node.c0);
        const reach1 = win1 || hasReachedByWin(node.c1);

        pushSvgLine(xPrev0, y0, xCurr, y0, reach0);
        pushSvgLine(xCurr, y0, xCurr, yMid, win0);
        pushSvgLine(xPrev1, y1, xCurr, y1, reach1);
        pushSvgLine(xCurr, yMid, xCurr, y1, win1);

        if (res) {
          const scoreOffsetX = isLeft ? -9 : 9;
          svgScoreTexts += `<text x="${xCurr + scoreOffsetX}" y="${y0 - 4}" text-anchor="middle" fill="#dc2626" font-size="11" font-weight="bold">${escapeHtml(res.g1)}</text>`;
          svgScoreTexts += `<text x="${xCurr + scoreOffsetX}" y="${y1 + 12}" text-anchor="middle" fill="#dc2626" font-size="11" font-weight="bold">${escapeHtml(res.g2)}</text>`;
        }

        if (state.appMode === 'match' && node.c0.entry && node.c1.entry) {
          const n0 = formatPairLabel(node.c0.entry, true);
          const n1 = formatPairLabel(node.c1.entry, true);
          htmlLabels += `<button onclick="openScoreModal('${r}_${m}', '${escapeJs(n0)}', '${escapeJs(n1)}', false)"
                                 style="left:${xCurr - 10}px; top:${yMid - 10}px; width:20px; height:20px;"
                                 class="no-print absolute rounded-full bg-rose-600 hover:bg-rose-500 text-white text-[10px] font-bold shadow flex items-center justify-center z-10"
                                 title="スコア入力">✎</button>`;
        }
      }
    }

    // 決勝中央のT字（cx）まで1本の真っ直ぐな水平線で接続
    // 準決勝を勝ち上がったペア（決勝敗退ペア含む）は中央のT字まで赤線を伸ばす
    const finalSideM = isLeft ? 0 : 1;
    const xWingEnd = getEffectiveStartX(tree[halfRounds][finalSideM], halfRounds, baseX, dir, stepWidth);
    const finalRes = tree[totalRounds][0].result;
    const isWingWinner = Boolean(finalRes && ((isLeft && finalRes.winner === 0) || (!isLeft && finalRes.winner === 1)));
    const wingReachedFinal = isWingWinner || hasReachedByWin(tree[halfRounds][finalSideM]);

    pushSvgLine(xWingEnd, targetFinalY, cx, targetFinalY, wingReachedFinal);
  };

  drawWing(true, leftLayout, leftShiftY);
  drawWing(false, rightLayout, rightShiftY);

  // Step 3: 公式トーナメント標準の左右完全対称「逆T字」決勝線（────────┴────────）
  const finalNode = tree[totalRounds][0];
  const finalRes = finalNode.result;
  const stemTopY = targetFinalY - stemHeight;

  pushSvgLine(cx, targetFinalY, cx, stemTopY, Boolean(finalRes), finalRes ? 2.8 : 1.4);

  if (finalRes) {
    svgScoreTexts += `<text x="${cx - 22}" y="${targetFinalY - 6}" text-anchor="middle" fill="#dc2626" font-size="11" font-weight="bold">${escapeHtml(finalRes.g1)}</text>`;
    svgScoreTexts += `<text x="${cx + 22}" y="${targetFinalY - 6}" text-anchor="middle" fill="#dc2626" font-size="11" font-weight="bold">${escapeHtml(finalRes.g2)}</text>`;
  }

  // 「優勝」の文字は削除し、決勝確定時のみ勝者ペア名を表示
  if (finalNode.entry) {
    const winNames = finalNode.entry.p2
      ? `${escapeHtml(finalNode.entry.p1)}・${escapeHtml(finalNode.entry.p2)}`
      : escapeHtml(finalNode.entry.p1);
    const winSchool = escapeHtml(formatSchoolShortName(finalNode.entry.school));
    htmlLabels += `<div style="left:${cx - 60}px; top:${stemTopY - 36}px; width:120px;"
                        class="absolute text-center bg-rose-50 border border-rose-300 rounded-md px-1.5 py-1 shadow-sm leading-tight">
      <div class="text-[11px] font-bold text-rose-900 truncate">${winNames}</div>
      <div class="text-[10px] font-semibold text-rose-700">（${winSchool}）</div>
    </div>`;
  }

  if (state.appMode === 'match' && finalNode.c0.entry && finalNode.c1.entry) {
    const n0 = formatPairLabel(finalNode.c0.entry, true);
    const n1 = formatPairLabel(finalNode.c1.entry, true);
    htmlLabels += `<button onclick="openScoreModal('${totalRounds}_0', '${escapeJs(n0)}', '${escapeJs(n1)}', false)"
                           style="left:${cx - 11}px; top:${targetFinalY - 11}px; width:22px; height:22px;"
                           class="no-print absolute rounded-full bg-rose-600 hover:bg-rose-500 text-white text-[10px] font-bold shadow flex items-center justify-center z-10"
                           title="決勝スコア入力">✎</button>`;
  }

  const consolationHtml = renderConsolationSectionHtml(totalWidth);

  sheet.style.width = `${totalWidth + 48}px`;
  sheet.innerHTML = `
    <div class="mb-3 pb-2.5 border-b-2 border-slate-800 flex items-baseline justify-between gap-4">
      <h2 class="text-lg font-bold tracking-wide text-slate-900">${escapeHtml(state.title)}</h2>
      <div class="flex items-center gap-3 text-xs text-slate-600 font-semibold shrink-0">
        <span>${state.matchFormat}ゲームマッチ</span>
        <span>出場：${state.entries.length}ペア（左 ${leftLayout.count} / 右 ${rightLayout.count}）</span>
      </div>
    </div>
    <div class="relative" style="width:${totalWidth}px; height:${canvasHeight}px;">
      <svg width="${totalWidth}" height="${canvasHeight}" class="absolute inset-0 pointer-events-none">
        ${svgBlackLines}${svgRedLines}${svgScoreTexts}
      </svg>
      ${htmlLabels}
    </div>
    ${consolationHtml}
  `;
}

function renderSingleSideBracket(sheet, tree, size, totalRounds, rowHeight, labelWidth, stepWidth, showRankBadge, alertSets) {
  const activeIndices = [];
  for (let i = 0; i < size; i++) {
    if (state.activeSlotMask[i]) activeIndices.push(i);
  }
  const count = Math.max(2, activeIndices.length);
  const canvasHeight = count * rowHeight + 28;
  const stepW = stepWidth + 14;
  const totalWidth = labelWidth + totalRounds * stepW + 190;
  const rawPriorityMap = BracketEngine.buildSlotPriorityMap(size);
  const underSeedMap = BracketEngine.buildUnderTop8SeedMap(size, state.activeSlotMask, rawPriorityMap);
  const slotPriorityMap = BracketEngine.buildAdjustedSlotPriorityMap(size, state.activeSlotMask);

  let svgBlackLines = '';
  let svgRedLines = '';
  let svgScoreTexts = '';
  let htmlLabels = '';
  const nodeY = Array.from({ length: totalRounds + 1 }, () => ({}));

  const pushSvgLine = (x1, y1, x2, y2, isRed, widthOverride) => {
    const stroke = isRed ? '#dc2626' : '#334155';
    const sw = widthOverride || (isRed ? 2.6 : 1.3);
    const line = `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${stroke}" stroke-width="${sw}" stroke-linecap="square" />`;
    if (isRed) svgRedLines += line;
    else svgBlackLines += line;
  };

  activeIndices.forEach((slotIdx, visualRow) => {
    const y = visualRow * rowHeight + rowHeight / 2 + 12;
    nodeY[0][slotIdx] = y;
    const entry = state.slots[slotIdx];
    const num = visualRow + 1;
    const isSelected = state.selectedTarget && state.selectedTarget.type === 'slot' && state.selectedTarget.index === slotIdx;
    const isSearchHit = state.searchMatches.includes(slotIdx);
    const isSearchCurr = state.searchMatches[state.searchCursor] === slotIdx;
    const blockInfo = BracketEngine.getBlockInfo(slotIdx, size);

    let cls = 'slot-card absolute flex items-center gap-1.5 px-1.5 text-[12px] cursor-pointer ';
    cls += getSlotAlertClass(slotIdx, isSelected, isSearchCurr, isSearchHit, alertSets, Boolean(entry));

    let innerHtml = '';
    if (entry) {
      const pairName = entry.p2 ? `${escapeHtml(entry.p1)}・${escapeHtml(entry.p2)}` : escapeHtml(entry.p1);
      const shortSc = formatSchoolShortName(entry.school);
      const rankText = formatSchoolRankBadge(entry);
      const slotPRank = slotPriorityMap[slotIdx];
      const underInfo = underSeedMap[slotIdx];
      const underLabel = underInfo ? ` [S${underInfo.seedRank}下・${underInfo.type === 'R1' ? '1回戦' : '2回戦'}]` : '';
      const blockTip = `確認用: ${rankText} ／ 枠優先順位: ${slotPRank}${underLabel} ／ ブロック: ${blockInfo.code} (${blockInfo.side === 'in' ? '内' : '外'})`;
      const priorityTag = state.showSlotPriority
        ? `<span class="priority-rank-el text-[9px] font-bold px-1 py-0.5 leading-tight rounded bg-indigo-50 text-indigo-800 border border-indigo-300 shrink-0 font-mono" title="枠優先順位">${slotPRank}</span>`
        : '';
      const rankBadge = showRankBadge
        ? `<span class="rank-badge-el text-[10px] font-bold px-1.5 py-0.5 leading-tight rounded bg-teal-50 text-teal-900 border border-teal-300 shrink-0" title="${escapeHtml(blockTip)}">${escapeHtml(rankText)}</span>`
        : '';
      const seedTag = (state.showSlotPriority && entry.seed)
        ? `<span class="priority-rank-el text-[9px] font-bold px-1 py-0.5 leading-tight rounded bg-amber-100 text-amber-900 border border-amber-300 shrink-0">S${entry.seed}</span>`
        : '';
      innerHtml = `<span class="font-mono text-[11px] font-bold text-slate-500 w-6 text-right shrink-0">${num}</span><span class="pair-name-wrap flex-1 min-w-0 text-left leading-snug"><span class="pair-name-autofit block whitespace-nowrap font-bold text-slate-900">${pairName}</span></span><span class="text-[11px] font-semibold text-slate-600 shrink-0 leading-snug">（${escapeHtml(shortSc)}）</span>${seedTag}${rankBadge}${priorityTag}`;
    } else {
      const slotPRank = slotPriorityMap[slotIdx];
      const priorityTag = state.showSlotPriority
        ? `<span class="priority-rank-el text-[9px] font-bold px-1 py-0.2 rounded bg-indigo-50 text-indigo-800 border border-indigo-300 shrink-0 font-mono">${slotPRank}</span>`
        : '';
      innerHtml = `<span class="font-mono text-[11px] text-slate-400 w-6 text-right shrink-0">${num}</span><span class="text-[11px] text-slate-400 italic flex-1 text-left">（空き枠）</span>${priorityTag}`;
    }

    htmlLabels += `<div data-slot-idx="${slotIdx}" onclick="handleTargetClick('slot', ${slotIdx})"
                        draggable="true" ondragstart="handleDragStart(event, 'slot', ${slotIdx})"
                        ondragover="handleDragOver(event)" ondragleave="handleDragLeave(event)" ondrop="handleDrop(event, 'slot', ${slotIdx})"
                        style="left:0px; top:${y - 13}px; width:${labelWidth - 8}px; height:26px;" class="${cls}">${innerHtml}</div>`;
  });

  for (let r = 1; r <= totalRounds; r++) {
    const mCount = size / Math.pow(2, r);
    const xCurr = labelWidth + r * stepW;
    for (let m = 0; m < mCount; m++) {
      const node = tree[r][m];
      if (!node.active) continue;
      if (node.isBye) {
        const activeChildIdx = node.c0.active ? (m * 2) : (m * 2 + 1);
        nodeY[r][m] = nodeY[r - 1][activeChildIdx];
        continue;
      }
      const y0 = nodeY[r - 1][m * 2];
      const y1 = nodeY[r - 1][m * 2 + 1];
      const yMid = (y0 + y1) / 2;
      nodeY[r][m] = yMid;

      const xPrev0 = getEffectiveStartX(tree[r - 1][m * 2], r - 1, labelWidth, 1, stepW);
      const xPrev1 = getEffectiveStartX(tree[r - 1][m * 2 + 1], r - 1, labelWidth, 1, stepW);
      const res = node.result;
      const win0 = Boolean(res && res.winner === 0);
      const win1 = Boolean(res && res.winner === 1);
      const reach0 = win0 || hasReachedByWin(node.c0);
      const reach1 = win1 || hasReachedByWin(node.c1);

      pushSvgLine(xPrev0, y0, xCurr, y0, reach0);
      pushSvgLine(xCurr, y0, xCurr, yMid, win0);
      pushSvgLine(xPrev1, y1, xCurr, y1, reach1);
      pushSvgLine(xCurr, yMid, xCurr, y1, win1);

      if (res) {
        svgScoreTexts += `<text x="${xCurr - 9}" y="${y0 - 4}" text-anchor="middle" fill="#dc2626" font-size="11" font-weight="bold">${escapeHtml(res.g1)}</text>`;
        svgScoreTexts += `<text x="${xCurr - 9}" y="${y1 + 12}" text-anchor="middle" fill="#dc2626" font-size="11" font-weight="bold">${escapeHtml(res.g2)}</text>`;
      }
      if (state.appMode === 'match' && node.c0.entry && node.c1.entry) {
        const n0 = formatPairLabel(node.c0.entry, true);
        const n1 = formatPairLabel(node.c1.entry, true);
        htmlLabels += `<button onclick="openScoreModal('${r}_${m}', '${escapeJs(n0)}', '${escapeJs(n1)}', false)"
                               style="left:${xCurr - 10}px; top:${yMid - 10}px; width:20px; height:20px;"
                               class="no-print absolute rounded-full bg-rose-600 text-white text-[10px] font-bold shadow flex items-center justify-center z-10">✎</button>`;
      }
    }
  }

  const finalY = nodeY[totalRounds][0];
  const finalX = labelWidth + totalRounds * stepW;
  const finalHasWinner = Boolean(tree[totalRounds][0].result);
  pushSvgLine(finalX, finalY, finalX + 26, finalY, finalHasWinner, finalHasWinner ? 2.6 : 1.4);
  if (tree[totalRounds][0].entry) {
    htmlLabels += `<div style="left:${finalX + 32}px; top:${finalY - 12}px;" class="absolute text-xs font-bold text-rose-700 bg-rose-50 border border-rose-200 px-2 py-0.5 rounded">${escapeHtml(formatPairLabel(tree[totalRounds][0].entry, true))}</div>`;
  }

  const consolationHtml = renderConsolationSectionHtml(totalWidth);
  sheet.style.width = `${totalWidth + 48}px`;
  sheet.innerHTML = `
    <div class="mb-3 pb-2.5 border-b-2 border-slate-800 flex items-baseline justify-between gap-4">
      <h2 class="text-lg font-bold tracking-wide text-slate-900">${escapeHtml(state.title)}</h2>
      <div class="flex items-center gap-3 text-xs text-slate-600 font-semibold shrink-0">
        <span>${state.matchFormat}ゲームマッチ</span>
        <span>出場：${state.entries.length}ペア</span>
      </div>
    </div>
    <div class="relative" style="width:${totalWidth}px; height:${canvasHeight}px;">
      <svg width="${totalWidth}" height="${canvasHeight}" class="absolute inset-0 pointer-events-none">${svgBlackLines}${svgRedLines}${svgScoreTexts}</svg>
      ${htmlLabels}
    </div>
    ${consolationHtml}
  `;
}

function renderConsolationSectionHtml(totalWidth) {
  if (state.consolation.preset === 'none') return '';
  const count = state.consolation.matchCount;
  let cards = '';
  for (let m = 0; m < count; m++) {
    const p1 = state.consolation.slots[m * 2] || '未定';
    const p2 = state.consolation.slots[m * 2 + 1] || '未定';
    const res = state.consolation.results[m];
    const win0 = res && res.winner === 0;
    const win1 = res && res.winner === 1;

    cards += `<div class="border border-slate-300 rounded-md p-2.5 bg-white flex items-center justify-between gap-3">
      <div class="space-y-1.5 text-xs flex-1 min-w-0">
        <div class="flex items-center justify-between gap-2">
          <span class="font-bold truncate ${win0 ? 'text-rose-700' : 'text-slate-800'}">${escapeHtml(p1)}</span>
          ${res ? `<span class="font-bold text-rose-600">${escapeHtml(res.g1)}</span>` : ''}
        </div>
        <div class="flex items-center justify-between gap-2">
          <span class="font-bold truncate ${win1 ? 'text-rose-700' : 'text-slate-800'}">${escapeHtml(p2)}</span>
          ${res ? `<span class="font-bold text-rose-600">${escapeHtml(res.g2)}</span>` : ''}
        </div>
      </div>
      ${state.appMode === 'match'
        ? `<button onclick="openScoreModal('${m}', '${escapeJs(p1)}', '${escapeJs(p2)}', true)" class="no-print px-2 py-1 bg-rose-600 text-white text-[10px] font-bold rounded shrink-0">結果入力</button>`
        : ''}
    </div>`;
  }

  return `<div class="mt-5 pt-4 border-t border-slate-300" style="max-width:${totalWidth}px;">
    <h3 class="text-sm font-bold text-slate-900 mb-2.5">${escapeHtml(state.consolation.title)}</h3>
    <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">${cards}</div>
  </div>`;
}

// =========================================================================
// 文字列エスケープ ＆ 初期化
// =========================================================================
function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escapeJs(str) {
  return String(str ?? '')
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/\n/g, ' ');
}

(function init() {
  const restored = loadFromLocalStorage();
  if (!restored || state.entries.length === 0) {
    loadRealisticSampleData(false);
    generateTournamentFromEntries(false, false);
  }
  document.getElementById('tournament-title-input').value = state.title;
  const fmtSel = document.getElementById('match-format-select');
  if (fmtSel) fmtSel.value = String(state.matchFormat);
  const fmtLbl = document.getElementById('match-mode-format-label');
  if (fmtLbl) fmtLbl.textContent = `${state.matchFormat}ゲームマッチ`;
  populateScoreGameSelects();
  updateRegCountPill();
  renderRuleCheckboxes();
  renderPaintPalette();
  setPrintMode(state.printMode);
  togglePrintFills(state.printShowFills, false);
  setPrintOrientation(state.printOrientation);
  setLayoutMode(state.layoutMode);
  updatePanelToggleButtons();
  renderSchoolAliasEditor();
  renderRegistrationTable();
  refreshAlerts(false);
})();

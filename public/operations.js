/* global state, BracketEngine, OopEngine, XLSX, html2canvas, jspdf, qrcode */
/* global saveToLocalStorage, renderSheet, renderBench, renderTourneySidebar, renderSchoolAliasEditor, renderRegistrationTable, renderRuleCheckboxes, refreshAlerts, updateRegCountPill, updateUndoRedoButtons, populateScoreGameSelects, setLayoutMode, togglePrintFills, formatSchoolShortName, formatPairLabel, escapeHtml, escapeJs, showToast, openScoreModal, switchMainView, pushHistory, generateTournamentFromEntries, normalizeBench, createEmptyBench */
'use strict';

/**
 * BracketFlow — Phase 1 Operations Module
 * 1. 親「大会」＞ 複数「種目（会場別）」マルチイベント管理（JSON完全互換）
 * 2. 進行表・コート運営（ハイブリッドOOP・ドラッグ移動・敗者審判・順位決定戦統合・印刷/PDF/Excel）
 * 3. 生徒・保護者向け 公開ビュー（スマホ最適化3タブ・個人情報制御・QRコード・スタッフPIN）
 */
(function () {
  const BUNDLE_STORAGE_KEY = 'bracketflow_tournament_bundle_v1';
  const SPECTATOR_FAVS_KEY = 'bracketflow_spectator_favs_v1';

  const opsUi = {
    bundle: null,
    suppressBundleSync: false,
    oopSelectedMoveKey: null,
    oopDragKey: null,
    oopOpenWoKey: null,
    oopHideFinished: false,   // 終了済みをたたんで現在・次試合を左へ寄せる
    oopQueueDrawerOpen: false,
    oopSplitBlock: 'ALL',     // 'ALL' | 'A' | 'B' | 'C' | 'D' | 'FINALS'
    oopConflictsOpen: false,
    homeSearchQuery: '',
    oopUndoStack: [],         // 進行表のコート移動・再配分Undoスタック
    oopFinishedCollapsed: {}, // { [courtKey]: boolean }
    standalonePublic: null,   // { meta, liveEvent, verifyPasscodeFn, unlocked }
    umpireModalMatchKey: null,
    eventModalMode: 'add',    // 'add' | 'edit'
    spectator: {
      frameMode: 'mobile',    // 'mobile' | 'full'
      tab: 'courts',          // 'courts' | 'bracket' | 'search'
      eventId: null,
      searchQuery: '',
      schoolFilter: 'ALL',
      courtFilter: 'ALL',     // 'ALL' | 'ACTIVE' | 'CHECKED'
      expandedCourts: {},     // { [courtKey]: boolean }
      bracketBlock: 'ALL',    // 'ALL' | 'A' | 'B' | 'C' | 'D'
      bracketStage: 'ALL',    // 'ALL' | '32' | '16' | '8'
      bracketZoom: 0.78,
      myCheckCollapsed: false,
      unlockedPasscode: '',
      organizerBypassClosed: false,
      favorites: new Set()
    }
  };

  function formatLocalClock(isoStr) {
    const d = isoStr ? new Date(isoStr) : new Date();
    if (Number.isNaN(d.getTime())) return '';
    return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
  }

  function formatLocalDateTime(isoStr) {
    if (!isoStr) return '-';
    const d = new Date(isoStr);
    if (Number.isNaN(d.getTime())) return String(isoStr).replace('T', ' ').slice(0, 16);
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    const hh = String(d.getHours()).padStart(2, '0');
    const mi = String(d.getMinutes()).padStart(2, '0');
    return `${yyyy}/${mm}/${dd} ${hh}:${mi}`;
  }

  function randomId(prefix) {
    return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 7)}`;
  }

  function generateStaffPin() {
    return String(Math.floor(1000 + Math.random() * 9000));
  }

  function cloneDeep(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  function captureActiveStateSnapshot() {
    return {
      version: '5.1',
      title: state.title,
      matchFormat: state.matchFormat,
      entries: cloneDeep(state.entries || []),
      schoolOrder: cloneDeep(state.schoolOrder || []),
      schoolAliases: cloneDeep(state.schoolAliases || {}),
      rules: cloneDeep(state.rules || BracketEngine.DEFAULT_RULES),
      bracketSize: state.bracketSize || 16,
      slots: cloneDeep(state.slots || []),
      activeSlotMask: cloneDeep(state.activeSlotMask || []),
      bench: cloneDeep(state.bench || createEmptyBench()),
      slotFills: cloneDeep(state.slotFills || {}),
      printShowFills: Boolean(state.printShowFills),
      results: cloneDeep(state.results || {}),
      consolation: cloneDeep(state.consolation || {
        preset: 'none',
        title: '順位決定戦（代表決定戦）',
        matchCount: 2,
        slots: ['', '', '', '', '', '', '', ''],
        results: {}
      }),
      oop: OopEngine.createDefaultOopState(state.oop),
      ignoredAlerts: cloneDeep(state.ignoredAlerts || {
        round1SameSchool: false,
        tripleSameCards: false,
        top4Violations: false,
        duplicateStudents: false,
        round2SameSchool: false
      }),
      layoutMode: state.layoutMode || 'dual',
      printMode: state.printMode || 'distribution',
      printOrientation: state.printOrientation || 'portrait'
    };
  }

  function hydrateStateFromSnapshot(snap) {
    const data = snap || {};
    state.title = data.title || '大会トーナメント表';
    state.matchFormat = (data.matchFormat === 5 || data.matchFormat === 7) ? data.matchFormat : 7;
    state.entries = Array.isArray(data.entries) ? cloneDeep(data.entries) : [];
    state.schoolOrder = Array.isArray(data.schoolOrder) ? cloneDeep(data.schoolOrder) : [];
    state.schoolAliases = data.schoolAliases ? cloneDeep(data.schoolAliases) : {};
    state.rules = Object.assign({}, BracketEngine.DEFAULT_RULES, data.rules || {});
    state.bracketSize = data.bracketSize || BracketEngine.nextPowerOfTwo(Math.max(4, state.entries.length || 4));
    state.slots = Array.isArray(data.slots) ? cloneDeep(data.slots) : [];
    state.activeSlotMask = Array.isArray(data.activeSlotMask) ? cloneDeep(data.activeSlotMask) : state.slots.map((s) => Boolean(s));
    state.bench = normalizeBench(data.bench);
    state.slotFills = (data.slotFills && typeof data.slotFills === 'object') ? cloneDeep(data.slotFills) : {};
    state.printShowFills = (typeof data.printShowFills === 'boolean') ? data.printShowFills : true;
    state.results = data.results ? cloneDeep(data.results) : {};
    state.consolation = data.consolation ? cloneDeep(data.consolation) : {
      preset: 'none',
      title: '順位決定戦（代表決定戦）',
      matchCount: 2,
      slots: ['', '', '', '', '', '', '', ''],
      results: {}
    };
    state.oop = OopEngine.createDefaultOopState(data.oop);
    state.ignoredAlerts = Object.assign({
      round1SameSchool: false,
      tripleSameCards: false,
      top4Violations: false,
      duplicateStudents: false,
      round2SameSchool: false
    }, data.ignoredAlerts || {});
    state.layoutMode = data.layoutMode || 'dual';
    state.printMode = data.printMode || 'distribution';
    state.printOrientation = data.printOrientation || 'portrait';
    state.history = [];
    state.future = [];
    state.selectedTarget = null;
    state.editingEntryId = null;

    // UI入力要素の同期
    const titleInput = document.getElementById('tournament-title-input');
    if (titleInput) titleInput.value = state.title;
    const fmtSel = document.getElementById('match-format-select');
    if (fmtSel) fmtSel.value = String(state.matchFormat);
    const fmtLbl = document.getElementById('match-mode-format-label');
    if (fmtLbl) fmtLbl.textContent = `${state.matchFormat}ゲームマッチ`;
    populateScoreGameSelects();
    togglePrintFills(state.printShowFills, false);
    setLayoutMode(state.layoutMode);
    updateUndoRedoButtons();
    updateRegCountPill();
    renderRuleCheckboxes();
    refreshAlerts(false);
  }

  function createEmptyEventSnapshot(eventName, venueName, courtCount) {
    return {
      version: '5.1',
      title: `${opsUi.bundle ? opsUi.bundle.name + ' ' : ''}${eventName}`,
      matchFormat: 7,
      entries: [],
      schoolOrder: [],
      schoolAliases: {},
      rules: Object.assign({}, BracketEngine.DEFAULT_RULES),
      bracketSize: 16,
      slots: [],
      activeSlotMask: [],
      bench: createEmptyBench(),
      slotFills: {},
      printShowFills: true,
      results: {},
      consolation: {
        preset: 'none',
        title: '順位決定戦（代表決定戦）',
        matchCount: 2,
        slots: ['', '', '', '', '', '', '', ''],
        results: {}
      },
      oop: OopEngine.createDefaultOopState({
        venue: venueName || '第1会場',
        courtCount: parseInt(courtCount, 10) || 4
      }),
      layoutMode: 'dual',
      printMode: 'distribution',
      printOrientation: 'portrait'
    };
  }

  function getActiveEvent() {
    if (!opsUi.bundle || !Array.isArray(opsUi.bundle.events) || opsUi.bundle.events.length === 0) return null;
    return opsUi.bundle.events.find((ev) => ev.id === opsUi.bundle.activeEventId) || opsUi.bundle.events[0];
  }

  function saveBundleToLocalStorage() {
    if (!opsUi.bundle) return;
    try {
      localStorage.setItem(BUNDLE_STORAGE_KEY, JSON.stringify(opsUi.bundle));
    } catch (_) {
      // 無視
    }
    if (!opsUi.suppressBundleSync && window.BracketCloud && typeof window.BracketCloud.scheduleCloudSync === 'function') {
      window.BracketCloud.scheduleCloudSync(false);
    }
  }

  function loadSpectatorFavorites() {
    try {
      const raw = localStorage.getItem(SPECTATOR_FAVS_KEY);
      const arr = raw ? JSON.parse(raw) : [];
      opsUi.spectator.favorites = new Set(Array.isArray(arr) ? arr : []);
    } catch (_) {
      opsUi.spectator.favorites = new Set();
    }
  }

  function saveSpectatorFavorites() {
    try {
      localStorage.setItem(SPECTATOR_FAVS_KEY, JSON.stringify(Array.from(opsUi.spectator.favorites)));
    } catch (_) {
      // 無視
    }
  }

  function initBundle() {
    loadSpectatorFavorites();
    try {
      const raw = localStorage.getItem(BUNDLE_STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && Array.isArray(parsed.events) && parsed.events.length > 0) {
          opsUi.bundle = parsed;
          if (!opsUi.bundle.publish) {
            opsUi.bundle.publish = { enabled: true, passcode: '', nameDisplayMode: 'full' };
          }
          opsUi.bundle.events.forEach((ev, idx) => {
            if (!ev.staffPin) ev.staffPin = generateStaffPin();
            if (!ev.venue) ev.venue = (ev.data && ev.data.oop && ev.data.oop.venue) || `第${idx + 1}会場`;
          });
          const activeEv = getActiveEvent();
          if (activeEv && activeEv.data) {
            opsUi.suppressBundleSync = true;
            hydrateStateFromSnapshot(activeEv.data);
            opsUi.suppressBundleSync = false;
            renderSchoolAliasEditor();
            renderRegistrationTable();
          }
          renderEventSwitcherBar();
          updateOopHeaderBadge();
          updatePublishIndicator();
          return;
        }
      }
    } catch (_) {
      // フォールバックへ
    }

    // 初回マイグレーション：現行の単一種目 state を親大会＞第1種目として取り込む
    const initialSnap = captureActiveStateSnapshot();
    const evId = randomId('ev');
    opsUi.bundle = {
      version: '6.0-bundle',
      id: randomId('tour'),
      name: '令和8年度 春季地区大会',
      publish: {
        enabled: true,
        passcode: '',
        nameDisplayMode: 'full'
      },
      activeEventId: evId,
      events: [
        {
          id: evId,
          name: '男子個人',
          venue: (initialSnap.oop && initialSnap.oop.venue) || '第1会場',
          staffPin: generateStaffPin(),
          data: initialSnap
        }
      ]
    };
    saveBundleToLocalStorage();
    renderEventSwitcherBar();
    updateOopHeaderBadge();
    updatePublishIndicator();
  }

  // =========================================================================
  // 1. 親大会・複数種目（会場別）バー ＆ モーダル
  // =========================================================================
  function renderEventSwitcherBar() {
    if (!opsUi.bundle) return;
    const parentInput = document.getElementById('parent-tournament-name-input');
    if (parentInput && document.activeElement !== parentInput) {
      parentInput.value = opsUi.bundle.name || '大会';
    }
    const tabsEl = document.getElementById('event-tabs-container');
    const activeEv = getActiveEvent();
    if (tabsEl) {
      tabsEl.innerHTML = opsUi.bundle.events.map((ev) => {
        const isActive = activeEv && ev.id === activeEv.id;
        const count = (ev.data && Array.isArray(ev.data.entries)) ? ev.data.entries.length : 0;
        return `<button onclick="switchActiveEvent('${escapeJs(ev.id)}')"
                        class="px-2.5 py-0.5 text-[11px] font-bold rounded transition flex items-center gap-1.5 border ${isActive
                          ? 'bg-emerald-600 text-white border-emerald-500 shadow-sm'
                          : 'bg-slate-900/60 text-slate-300 border-slate-700 hover:bg-slate-700 hover:text-white'}">
                  <span>${escapeHtml(ev.name)}</span>
                  <span class="text-[9px] px-1 rounded ${isActive ? 'bg-slate-950/30 text-emerald-100' : 'bg-slate-800 text-slate-400'}">${count}組</span>
                </button>`;
      }).join('');
    }
    const venuePill = document.getElementById('active-event-venue-pill');
    if (venuePill && activeEv) {
      const cCount = (state.oop && state.oop.courtCount) || 4;
      venuePill.textContent = `会場: ${activeEv.venue || '第1会場'} (${cCount}面)`;
    }
  }

  function updateParentTournamentName(val) {
    if (!opsUi.bundle) return;
    opsUi.bundle.name = String(val || '').trim() || '大会';
    saveBundleToLocalStorage();
    if (state.mainView === 'oop') renderOopView();
    if (state.mainView === 'spectator') renderSpectatorView();
  }

  function switchActiveEvent(eventId) {
    if (!opsUi.bundle) return;
    const currentEv = getActiveEvent();
    if (currentEv && currentEv.id === eventId) return;
    const targetEv = opsUi.bundle.events.find((ev) => ev.id === eventId);
    if (!targetEv) return;

    // 現在の種目を保存
    if (currentEv) {
      currentEv.data = captureActiveStateSnapshot();
      currentEv.venue = (currentEv.data.oop && currentEv.data.oop.venue) || currentEv.venue;
    }

    opsUi.bundle.activeEventId = targetEv.id;
    opsUi.oopSelectedMoveKey = null;
    opsUi.oopOpenWoKey = null;
    opsUi.oopUndoStack = [];

    opsUi.suppressBundleSync = true;
    hydrateStateFromSnapshot(targetEv.data);
    opsUi.suppressBundleSync = false;

    saveBundleToLocalStorage();
    renderEventSwitcherBar();
    updateOopHeaderBadge();
    if (window.BracketCloud && typeof window.BracketCloud.attachAdminLiveListener === 'function') {
      window.BracketCloud.attachAdminLiveListener();
    }

    if (state.mainView === 'registration') {
      renderSchoolAliasEditor();
      renderRegistrationTable();
    } else if (state.mainView === 'tournament') {
      if (state.slots.length === 0 && state.entries.length >= 2) {
        generateTournamentFromEntries(false, false);
      }
      renderBench();
      renderTourneySidebar();
      renderSheet();
    } else if (state.mainView === 'oop') {
      renderOopView();
    } else if (state.mainView === 'spectator') {
      opsUi.spectator.eventId = targetEv.id;
      renderSpectatorView();
    }
    showToast(`種目「${targetEv.name}」に切り替えました`, 'info', 1800);
  }

  function openEventModal(mode) {
    opsUi.eventModalMode = mode === 'edit' ? 'edit' : 'add';
    const modal = document.getElementById('event-modal');
    const titleEl = document.getElementById('event-modal-title');
    const nameInput = document.getElementById('event-modal-name');
    const venueInput = document.getElementById('event-modal-venue');
    const courtRow = document.getElementById('event-modal-court-row');
    const extraActions = document.getElementById('event-modal-extra-actions');
    const courtsSel = document.getElementById('event-modal-courts');

    if (opsUi.eventModalMode === 'add') {
      const nextNum = (opsUi.bundle ? opsUi.bundle.events.length : 1) + 1;
      const presetNames = ['男子個人', '女子個人', '男子団体', '女子団体'];
      const suggested = presetNames[nextNum - 1] || `種目${nextNum}`;
      titleEl.textContent = '新しい種目（別会場・別進行表）を追加';
      nameInput.value = suggested;
      venueInput.value = `第${nextNum}会場`;
      courtsSel.value = '4';
      courtRow.classList.remove('hidden');
      extraActions.classList.add('hidden');
    } else {
      const ev = getActiveEvent();
      if (!ev) return;
      titleEl.textContent = `種目設定：${ev.name}`;
      nameInput.value = ev.name || '';
      venueInput.value = (state.oop && state.oop.venue) || ev.venue || '第1会場';
      courtRow.classList.add('hidden');
      extraActions.classList.remove('hidden');
    }
    modal.classList.remove('hidden');
    setTimeout(() => nameInput.focus(), 40);
  }

  function closeEventModal() {
    const modal = document.getElementById('event-modal');
    if (modal) modal.classList.add('hidden');
  }

  function submitEventModal() {
    if (!opsUi.bundle) return;
    const name = String(document.getElementById('event-modal-name').value || '').trim();
    const venue = String(document.getElementById('event-modal-venue').value || '').trim() || '第1会場';
    if (!name) {
      alert('種目名を入力してください。');
      return;
    }

    if (opsUi.eventModalMode === 'add') {
      const courts = parseInt(document.getElementById('event-modal-courts').value, 10) || 4;
      const curEv = getActiveEvent();
      if (curEv) curEv.data = captureActiveStateSnapshot();

      const newId = randomId('ev');
      const newEv = {
        id: newId,
        name,
        venue,
        staffPin: generateStaffPin(),
        data: createEmptyEventSnapshot(name, venue, courts)
      };
      opsUi.bundle.events.push(newEv);
      closeEventModal();
      switchActiveEvent(newId);
      switchMainView('registration');
      showToast(`種目「${name}（${venue}）」を追加しました`, 'success');
    } else {
      const ev = getActiveEvent();
      if (!ev) return;
      ev.name = name;
      ev.venue = venue;
      if (state.oop) state.oop.venue = venue;
      ev.data = captureActiveStateSnapshot();
      saveToLocalStorage();
      saveBundleToLocalStorage();
      closeEventModal();
      renderEventSwitcherBar();
      if (state.mainView === 'oop') renderOopView();
      showToast(`種目設定を更新しました`, 'success');
    }
  }

  function duplicateActiveEvent() {
    if (!opsUi.bundle) return;
    const ev = getActiveEvent();
    if (!ev) return;
    ev.data = captureActiveStateSnapshot();
    const copyData = cloneDeep(ev.data);
    const copyId = randomId('ev');
    const copyName = `${ev.name}（コピー）`;
    copyData.title = `${opsUi.bundle.name} ${copyName}`;
    opsUi.bundle.events.push({
      id: copyId,
      name: copyName,
      venue: ev.venue,
      staffPin: generateStaffPin(),
      data: copyData
    });
    closeEventModal();
    switchActiveEvent(copyId);
    showToast(`種目「${copyName}」を複製しました`, 'success');
  }

  function deleteActiveEvent() {
    if (!opsUi.bundle) return;
    if (opsUi.bundle.events.length <= 1) {
      alert('最低1つの種目が必要なため削除できません。');
      return;
    }
    const ev = getActiveEvent();
    if (!ev) return;
    if (!confirm(`種目「${ev.name}」を削除しますか？この操作は取り消せません。`)) return;
    opsUi.bundle.events = opsUi.bundle.events.filter((item) => item.id !== ev.id);
    const nextEv = opsUi.bundle.events[0];
    opsUi.bundle.activeEventId = nextEv.id;
    opsUi.suppressBundleSync = true;
    hydrateStateFromSnapshot(nextEv.data);
    opsUi.suppressBundleSync = false;
    saveBundleToLocalStorage();
    closeEventModal();
    renderEventSwitcherBar();
    switchMainView(state.mainView);
    showToast(`種目「${ev.name}」を削除しました`, 'warn');
  }

  // =========================================================================
  // 2. マルチイベント完全対応 JSON 保存・読込（旧単一種目JSONも互換読込）
  // =========================================================================
  function exportBundleJsonState() {
    try {
      const curEv = getActiveEvent();
      const activeSnap = captureActiveStateSnapshot();
      if (curEv) curEv.data = activeSnap;
      saveBundleToLocalStorage();

      // 旧バージョン互換のためアクティブ種目のトップレベル項目も併記
      const payload = Object.assign({}, activeSnap, {
        version: '6.0-bundle',
        exportedAt: new Date().toISOString(),
        bundle: cloneDeep(opsUi.bundle)
      });

      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const safeName = (opsUi.bundle.name || state.title || 'tournament').replace(/[\\/:*?"<>|]/g, '_');
      const dateStr = new Date().toISOString().slice(0, 10);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${safeName}_${dateStr}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      showToast(`大会データ（全${opsUi.bundle.events.length}種目）をJSON保存しました`, 'success');
    } catch (err) {
      alert('JSON保存中にエラーが発生しました: ' + err.message);
    }
  }

  function importBundleJsonState(event) {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = JSON.parse(e.target.result);
        if (data && data.bundle && Array.isArray(data.bundle.events) && data.bundle.events.length > 0) {
          opsUi.bundle = data.bundle;
          if (!opsUi.bundle.publish) opsUi.bundle.publish = { enabled: true, passcode: '', nameDisplayMode: 'full' };
          const activeEv = getActiveEvent();
          opsUi.suppressBundleSync = true;
          hydrateStateFromSnapshot(activeEv.data);
          opsUi.suppressBundleSync = false;
          saveBundleToLocalStorage();
          saveToLocalStorage();
          renderEventSwitcherBar();
          updateOopHeaderBadge();
          updatePublishIndicator();
          if (state.slots.length > 0) switchMainView('tournament');
          else switchMainView('registration');
          showToast(`大会データ（${opsUi.bundle.events.length}種目）を読み込みました`, 'success');
        } else if (data && Array.isArray(data.entries)) {
          // 旧単一種目JSONを現在の種目へ読込
          opsUi.suppressBundleSync = true;
          hydrateStateFromSnapshot(data);
          opsUi.suppressBundleSync = false;
          const activeEv = getActiveEvent();
          if (activeEv) activeEv.data = captureActiveStateSnapshot();
          saveBundleToLocalStorage();
          saveToLocalStorage();
          renderEventSwitcherBar();
          updateOopHeaderBadge();
          if (state.slots.length > 0) switchMainView('tournament');
          else switchMainView('registration');
          showToast('単一種目JSONデータを現在の種目へ読み込みました', 'success');
        } else {
          throw new Error('有効なBracketFlow JSONデータではありません。');
        }
      } catch (err) {
        alert('JSONファイルの読み込みに失敗しました: ' + err.message);
      } finally {
        event.target.value = '';
      }
    };
    reader.readAsText(file, 'UTF-8');
  }

  // =========================================================================
  // 3. 進行表・コート運営（OOP）エンジン連携 ＆ UI
  // =========================================================================
  function getActiveMatchCatalog(nameDisplayMode) {
    const mode = nameDisplayMode || 'full';
    const formatter = (entry) => OopEngine.formatSpectatorPairName(entry, mode, formatSchoolShortName);
    return OopEngine.buildMatchCatalog({
      bracketSize: state.bracketSize,
      slots: state.slots,
      activeSlotMask: state.activeSlotMask,
      results: state.results,
      consolation: state.consolation,
      labelFormatter: formatter
    });
  }

  function syncOopWithBracket() {
    const catalog = getActiveMatchCatalog('full');
    const reconciled = OopEngine.reconcileOopState(state.oop, catalog);
    state.oop = reconciled.oop;
    return { catalog, unassigned: reconciled.unassigned };
  }

  function updateOopHeaderBadge() {
    const pill = document.getElementById('oop-live-pill');
    if (!pill || !state.oop || !state.oop.statuses) return;
    let activeCount = 0;
    Object.values(state.oop.statuses).forEach((st) => {
      if (st === 'playing' || st === 'calling') activeCount++;
    });
    if (activeCount > 0) {
      pill.textContent = String(activeCount);
      pill.classList.remove('hidden');
    } else {
      pill.classList.add('hidden');
    }
  }

  function renderOopView() {
    const { catalog, unassigned } = syncOopWithBracket();
    updateOopHeaderBadge();
    renderEventSwitcherBar();

    // ツールバー値の同期
    const venueInput = document.getElementById('oop-venue-input');
    if (venueInput && document.activeElement !== venueInput) {
      venueInput.value = state.oop.venue || '第1会場';
    }
    const courtSel = document.getElementById('oop-court-count-select');
    if (courtSel) courtSel.value = String(state.oop.courtCount || 4);
    const settingsSummary = document.getElementById('oop-settings-summary');
    if (settingsSummary) {
      settingsSummary.textContent = `${state.oop.venue || '第1会場'}・${state.oop.courtCount || 4}面`;
    }

    const btnBlock = document.getElementById('btn-oop-mode-block');
    const btnRound = document.getElementById('btn-oop-mode-round');
    if (btnBlock && btnRound) {
      const onCls = 'px-2 py-1 text-[11px] font-bold rounded-md bg-white text-slate-900 shadow-sm transition';
      const offCls = 'px-2 py-1 text-[11px] font-bold rounded-md text-slate-600 hover:text-slate-900 transition';
      btnBlock.className = state.oop.assignMode === 'block' ? onCls : offCls;
      btnRound.className = state.oop.assignMode === 'round' ? onCls : offCls;
    }

    const umpireBtn = document.getElementById('btn-oop-umpire-toggle');
    if (umpireBtn) {
      const isOn = state.oop.umpireRule !== 'none';
      umpireBtn.textContent = `敗者審判: ${isOn ? 'ON' : 'OFF'}`;
      umpireBtn.className = isOn
        ? 'px-2.5 py-1 text-[11px] font-bold rounded-lg border border-emerald-300 bg-emerald-50 text-emerald-800 transition'
        : 'px-2.5 py-1 text-[11px] font-bold rounded-lg border border-slate-300 bg-white text-slate-600 transition';
    }

    const btnBoard = document.getElementById('btn-oop-view-board');
    const btnSplit = document.getElementById('btn-oop-view-split');
    const btnTable = document.getElementById('btn-oop-view-table');
    const boardWs = document.getElementById('oop-board-workspace');
    const tableWs = document.getElementById('oop-table-workspace');
    const splitPane = document.getElementById('oop-split-bracket-pane');
    const viewStyle = state.oop.viewStyle || 'board';
    const isTable = viewStyle === 'table';
    const isSplit = viewStyle === 'split';
    if (btnBoard && btnTable) {
      const onCls = 'px-2.5 py-1 text-[11px] font-bold rounded-md bg-slate-900 text-white shadow-sm transition';
      const offCls = 'px-2.5 py-1 text-[11px] font-bold rounded-md text-slate-600 hover:text-slate-900 transition';
      btnBoard.className = viewStyle === 'board' ? onCls : offCls;
      if (btnSplit) btnSplit.className = isSplit ? onCls : offCls;
      btnTable.className = isTable ? onCls : offCls;
    }
    if (boardWs) boardWs.classList.toggle('hidden', isTable);
    if (tableWs) tableWs.classList.toggle('hidden', !isTable);
    if (splitPane) splitPane.classList.toggle('hidden', !isSplit);

    // アナウンス稼働インジケータ
    const annDot = document.getElementById('oop-announce-dot');
    const hasAnn = Boolean((opsUi.bundle && opsUi.bundle.announcement && opsUi.bundle.announcement.text) || (state.oop && state.oop.announcement && state.oop.announcement.text));
    if (annDot) annDot.classList.toggle('hidden', !hasAnn);

    const blankBtn = document.getElementById('btn-oop-add-blank');
    if (blankBtn) {
      const pickingBlank = opsUi.oopSelectedMoveKey === '__NEW_BLANK__';
      blankBtn.className = pickingBlank
        ? 'px-2.5 py-1 text-[11px] font-bold rounded-lg border border-amber-600 bg-amber-500 text-white shadow-sm transition'
        : 'px-2.5 py-1 text-[11px] font-bold rounded-lg border border-amber-300 bg-amber-50 hover:bg-amber-100 text-amber-900 transition';
    }

    const cancelMoveBtn = document.getElementById('btn-oop-cancel-move');
    if (cancelMoveBtn) cancelMoveBtn.classList.toggle('hidden', !opsUi.oopSelectedMoveKey);

    const hideFinBtn = document.getElementById('btn-oop-hide-finished');
    if (hideFinBtn) {
      hideFinBtn.textContent = `終了済をたたむ: ${opsUi.oopHideFinished ? 'ON' : 'OFF'}`;
      hideFinBtn.className = opsUi.oopHideFinished
        ? 'px-2.5 py-1 text-[11px] font-bold rounded-lg border border-blue-400 bg-blue-50 text-blue-800 transition'
        : 'px-2.5 py-1 text-[11px] font-bold rounded-lg border border-slate-300 bg-white text-slate-600 hover:bg-slate-50 transition';
    }

    const undoOopBtn = document.getElementById('btn-oop-undo');
    if (undoOopBtn) {
      undoOopBtn.disabled = opsUi.oopUndoStack.length === 0;
    }

    document.body.classList.toggle('oop-move-mode-active', Boolean(opsUi.oopSelectedMoveKey));
    const moveHint = document.getElementById('oop-move-hint');
    if (moveHint) {
      if (opsUi.oopSelectedMoveKey === '__NEW_BLANK__') {
        moveHint.innerHTML = `<span class="text-amber-700 font-bold">空き枠を連続挿入中：複数の「＋」をクリックして何個でも設置できます（もう一度「＋ 空き枠」またはEscで終了）</span>`;
      } else if (opsUi.oopSelectedMoveKey && OopEngine.isBlankKey(opsUi.oopSelectedMoveKey)) {
        moveHint.innerHTML = `<span class="text-blue-700 font-bold">空き枠の移動先（「＋」）をクリックしてください</span>`;
      } else if (opsUi.oopSelectedMoveKey && catalog.matchMap[opsUi.oopSelectedMoveKey]) {
        const sm = catalog.matchMap[opsUi.oopSelectedMoveKey];
        moveHint.innerHTML = `<span class="text-blue-700 font-bold">移動先を選択中（${escapeHtml(sm.matchNo)}）：挿入したい試合の間の「＋」をクリックしてください</span>`;
      } else {
        moveHint.textContent = isSplit
          ? '左のトーナメントT字アイコンを右のコートへドラッグ（またはクリック→＋）して進行表を作成できます'
          : 'カードをドラッグまたは「移動」で、任意の試合と試合の間に挿入できます';
      }
    }

    // 進捗サマリー ＆ 時系列矛盾チェック
    let playing = 0;
    let calling = 0;
    let finished = 0;
    let readyWaiting = 0;
    catalog.matches.forEach((m) => {
      const st = state.oop.statuses[m.key] || 'waiting';
      if (st === 'playing') playing++;
      else if (st === 'calling') calling++;
      else if (st === 'finished' || st === 'wo') finished++;
      else if (m.isReady) readyWaiting++;
    });
    const total = catalog.matches.length;
    const pct = total > 0 ? Math.round((finished / total) * 100) : 0;
    const conflictInfo = OopEngine.detectTimelineConflicts(state.oop, catalog);
    const conflictCount = conflictInfo.conflicts.length;

    const statsBar = document.getElementById('oop-stats-bar');
    if (statsBar) {
      const conflictBadge = conflictCount > 0
        ? `<button onclick="toggleOopConflictsDrawer()" class="px-2 py-0.5 rounded-full bg-rose-600 hover:bg-rose-500 text-white font-bold shadow-sm transition">⚠ 時系列矛盾 ${conflictCount}件 ▾</button>`
        : `<span class="px-2 py-0.5 rounded-full bg-emerald-50 border border-emerald-200 font-bold text-emerald-700">✓ 時系列OK</span>`;
      statsBar.innerHTML = `
        <span class="px-2 py-0.5 rounded-full bg-slate-100 border border-slate-200 font-bold text-slate-700">全${total}試合（進捗 ${pct}%）</span>
        <span class="px-2 py-0.5 rounded-full bg-rose-100 border border-rose-200 font-bold text-rose-800">試合中 ${playing}</span>
        <span class="px-2 py-0.5 rounded-full bg-amber-100 border border-amber-200 font-bold text-amber-800">コール中 ${calling}</span>
        <span class="px-2 py-0.5 rounded-full bg-blue-50 border border-blue-200 font-bold text-blue-800">実施可能待機 ${readyWaiting}</span>
        <span class="px-2 py-0.5 rounded-full bg-emerald-100 border border-emerald-200 font-bold text-emerald-800">終了 ${finished}</span>
        ${conflictBadge}
      `;
    }

    const conflictsDrawer = document.getElementById('oop-conflicts-drawer');
    if (conflictsDrawer) {
      if (conflictCount === 0) {
        opsUi.oopConflictsOpen = false;
        conflictsDrawer.classList.add('hidden');
      } else {
        conflictsDrawer.classList.toggle('hidden', !opsUi.oopConflictsOpen);
        conflictsDrawer.innerHTML = `
          <div class="flex items-center justify-between font-bold text-rose-900">
            <span>⚠ 時系列の矛盾が ${conflictCount} 件あります（空き枠の挿入または順序入れ替えで解消できます）</span>
            <button onclick="toggleOopConflictsDrawer()" class="text-rose-600 hover:text-rose-900 px-1">✕</button>
          </div>
          <div class="flex flex-wrap gap-1.5 pt-1">
            ${conflictInfo.conflicts.map((cf) => `
              <button onclick="jumpToOopCard('${escapeJs(cf.matchKey)}')"
                      class="px-2 py-1 rounded bg-white hover:bg-rose-100 border border-rose-300 text-rose-900 text-left font-semibold transition">
                <b>${escapeHtml(cf.matchNo)}</b>（C${cf.courtKey}-第${cf.colIdx + 1}）: ${escapeHtml(cf.message)}
              </button>
            `).join('')}
          </div>
        `;
      }
    }

    if (!isTable) {
      if (isSplit) renderOopSplitBracket(catalog);
      renderOopBoard(catalog, unassigned, conflictInfo.conflictMap);
    } else {
      renderOopPrintSheet(catalog);
    }
  }

  function renderOopBoard(catalog, unassigned, conflictMap) {
    const grid = document.getElementById('oop-courts-grid');
    if (!grid) return;
    const cMap = conflictMap || {};

    const nCourts = state.oop.courtCount || 4;
    const formatter = (entry) => OopEngine.formatSpectatorPairName(entry, 'full', formatSchoolShortName);

    // 未割当サイドバーの表示制御（2分割モード中は左トーナメントから直接ドラッグできるため自動非表示、手動展開時のみ表示）
    const sidebar = document.getElementById('oop-queue-sidebar');
    if (sidebar) {
      const showSidebar = opsUi.oopQueueDrawerOpen || (state.oop.viewStyle !== 'split' && unassigned.length > 0);
      sidebar.classList.toggle('hidden', !showSidebar);
    }

    // 全コートの最大表示試合数を算出（ヘッダー列「第1試合・第2試合…」用）
    let maxVisibleCols = 1;
    const courtRowsData = [];
    for (let c = 1; c <= nCourts; c++) {
      const cKey = String(c);
      const courtKeys = Array.isArray(state.oop.courts[cKey]) ? state.oop.courts[cKey] : [];
      const realCount = courtKeys.filter((k) => !OopEngine.isBlankKey(k)).length;
      const live = OopEngine.getCourtLiveSlots(courtKeys, catalog.matchMap, state.oop.statuses);
      const items = [];
      let hiddenFinishedCount = 0;

      // 各試合のロール（現在/次/次々）を事前マップ
      const roleMap = {};
      if (live.current) roleMap[live.current.key] = 'CURRENT';
      if (live.next) roleMap[live.next.key] = live.current ? 'NEXT' : 'READY';
      if (live.onDeck) roleMap[live.onDeck.key] = 'ONDECK';

      courtKeys.forEach((mKey, rawIdx) => {
        if (OopEngine.isBlankKey(mKey)) {
          items.push({ key: mKey, isBlank: true, indexInCourt: rawIdx });
          return;
        }
        const m = catalog.matchMap[mKey];
        if (!m) return;
        const st = state.oop.statuses[mKey] || 'waiting';
        if (opsUi.oopHideFinished && (st === 'finished' || st === 'wo')) {
          hiddenFinishedCount++;
          return;
        }
        items.push({ key: mKey, isBlank: false, match: m, status: st, indexInCourt: rawIdx, role: roleMap[mKey] || '' });
      });

      if (items.length > maxVisibleCols) maxVisibleCols = items.length;
      courtRowsData.push({ c, cKey, courtKeys, realCount, live, items, hiddenFinishedCount });
    }

    const renderInsertZone = (courtKey, insertIdx, labelText) => `
      <div class="oop-insert-zone shrink-0"
           title="${escapeHtml(labelText || 'ここに挿入')}"
           onclick="handleOopInsertZoneClick('${courtKey}', ${insertIdx})"
           ondragover="handleOopInsertDragOver(event)"
           ondragleave="handleOopInsertDragLeave(event)"
           ondrop="handleOopInsertDrop(event, '${courtKey}', ${insertIdx})">
        <span class="oop-insert-badge text-[11px] font-extrabold text-blue-700 select-none">＋</span>
      </div>
    `;

    // 全コート横タイムライン一覧テーブル構造
    let rowsHtml = '';
    courtRowsData.forEach(({ c, cKey, courtKeys, realCount, live, items, hiddenFinishedCount }) => {
      let statusBadge = `<span class="px-1.5 py-0.5 rounded text-[10px] font-bold bg-slate-200 text-slate-700">空き</span>`;
      let rowAccent = 'border-l-slate-300';
      if (live.current && live.current.status === 'playing') {
        statusBadge = `<span class="px-1.5 py-0.5 rounded text-[10px] font-bold bg-rose-600 text-white inline-flex items-center gap-1"><span class="w-1.5 h-1.5 rounded-full bg-white live-dot"></span>試合中</span>`;
        rowAccent = 'border-l-rose-500';
      } else if (live.current && live.current.status === 'calling') {
        statusBadge = `<span class="px-1.5 py-0.5 rounded text-[10px] font-bold bg-amber-500 text-white">コール中</span>`;
        rowAccent = 'border-l-amber-500';
      } else if (realCount > 0 && live.finished.length === realCount) {
        statusBadge = `<span class="px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800">完了</span>`;
        rowAccent = 'border-l-emerald-500';
      }

      // 横並びの試合カード＋空き枠カード＋カード間挿入スロット
      let laneHtml = '';
      if (hiddenFinishedCount > 0) {
        laneHtml += `
          <button onclick="toggleOopHideFinished()"
                  class="shrink-0 w-16 mr-1 rounded-lg bg-slate-100 hover:bg-slate-200 border border-slate-300 text-[10px] font-bold text-slate-600 flex flex-col items-center justify-center p-1 transition"
                  title="クリックして終了済み試合を再表示">
            <span class="text-emerald-700">済 ${hiddenFinishedCount}</span>
            <span class="text-[9px] text-slate-400">表示▸</span>
          </button>
        `;
      }

      if (items.length === 0) {
        laneHtml += `
          <div class="flex items-center flex-1 min-h-[96px]">
            ${renderInsertZone(cKey, courtKeys.length, `コート${c}の先頭へ挿入`)}
            <div class="flex-1 min-w-[200px] text-center text-[11px] text-slate-400 py-4 px-3 border border-dashed border-slate-200 rounded-lg mx-1"
                 onclick="handleOopInsertZoneClick('${cKey}', ${courtKeys.length})"
                 ondragover="handleOopDragOver(event)"
                 ondragleave="handleOopDragLeave(event)"
                 ondrop="handleOopInsertDrop(event, '${cKey}', ${courtKeys.length})">
              ${courtKeys.length === 0 ? '割り当て試合なし（ここへドラッグ＆ドロップ）' : '未終了の試合はありません'}
            </div>
          </div>
        `;
      } else {
        // 先頭の挿入スロット
        laneHtml += renderInsertZone(cKey, items[0].indexInCourt, `コート${c} 第${items[0].indexInCourt + 1}試合の前に挿入`);
        items.forEach((item, visIdx) => {
          if (item.isBlank) {
            laneHtml += renderOopBlankCardHtml(item, cKey);
          } else {
            laneHtml += renderOopTimelineCardHtml(item, cKey, courtKeys, catalog.matchMap, formatter, cMap[item.key] || []);
          }
          const nextInsertIdx = item.indexInCourt + 1;
          const curLabel = item.isBlank ? '空き枠' : item.match.matchNo;
          const nextItem = items[visIdx + 1];
          const label = nextItem
            ? `コート${c} ${curLabel}と${nextItem.isBlank ? '空き枠' : nextItem.match.matchNo}の間に挿入`
            : `コート${c} の末尾へ追加`;
          laneHtml += renderInsertZone(cKey, nextInsertIdx, label);
        });
      }

      rowsHtml += `
        <div class="flex items-stretch bg-white border border-slate-200 border-l-4 ${rowAccent} rounded-xl shadow-sm overflow-visible min-w-max">
          <!-- 固定左列：コート情報 -->
          <div class="sticky left-0 z-10 w-28 shrink-0 bg-slate-50/95 backdrop-blur-sm border-r border-slate-200 px-2.5 py-2 flex flex-col justify-between rounded-l-lg">
            <div class="space-y-1">
              <div class="text-xs font-extrabold text-slate-900">コート ${c}</div>
              <div>${statusBadge}</div>
            </div>
            <div class="pt-1.5 border-t border-slate-200/80 text-[10px] font-bold text-slate-500 flex items-center justify-between">
              <span>進捗</span>
              <span>${live.finished.length}/${realCount}</span>
            </div>
          </div>

          <!-- 横スクロール試合レーン -->
          <div class="flex items-stretch py-2 px-1.5">
            ${laneHtml}
          </div>
        </div>
      `;
    });

    // 列見出し（第1試合・第2試合…の目安バー）
    let colHeaderCells = '';
    for (let i = 0; i < Math.max(4, maxVisibleCols); i++) {
      colHeaderCells += `<div class="w-[228px] shrink-0 text-center text-[10px] font-bold text-slate-500">${opsUi.oopHideFinished ? `待機順 ${i + 1}` : `第 ${i + 1} 試合`}</div>`;
    }

    grid.innerHTML = `
      <div class="space-y-2 min-w-max pb-4">
        <div class="flex items-center min-w-max sticky top-0 z-20 bg-slate-100/95 backdrop-blur-sm py-1 border-b border-slate-200">
          <div class="sticky left-0 z-30 w-28 shrink-0 bg-slate-100 px-3 text-[10px] font-extrabold text-slate-600 uppercase tracking-wider">全${nCourts}コート一覧</div>
          <div class="flex items-center px-3 gap-0">
            ${colHeaderCells}
          </div>
        </div>
        ${rowsHtml}
      </div>
    `;

    // 未割当キュー描画
    const unassignedCountEl = document.getElementById('oop-unassigned-count');
    const unassignedListEl = document.getElementById('oop-unassigned-list');
    if (unassignedCountEl) {
      unassignedCountEl.textContent = String(unassigned.length);
      unassignedCountEl.className = unassigned.length > 0
        ? 'px-1.5 rounded-full bg-amber-500 text-white text-[10px]'
        : 'px-1.5 rounded-full bg-slate-300 text-slate-800 text-[10px]';
    }
    if (unassignedListEl) {
      if (unassigned.length === 0) {
        unassignedListEl.innerHTML = `<div class="text-center py-6 text-[11px] text-slate-400">未割当の試合はありません<br/>全試合がコートに配分済です</div>`;
      } else {
        unassignedListEl.innerHTML = unassigned.map((mKey) => {
          const m = catalog.matchMap[mKey];
          if (!m) return '';
          const isSelected = opsUi.oopSelectedMoveKey === mKey;
          return `<div draggable="true"
                       ondragstart="handleOopDragStart(event, '${escapeJs(mKey)}')"
                       ondragend="handleOopDragEnd(event)"
                       class="oop-match-card ${isSelected ? 'oop-selected-move' : ''} p-2 rounded-lg border border-slate-200 bg-slate-50 text-[11px] space-y-1 cursor-grab">
            <div class="flex items-center justify-between gap-1">
              <span class="font-bold text-slate-800">${escapeHtml(m.matchNo)} <span class="text-[10px] text-slate-500">${escapeHtml(m.roundLabel)}</span></span>
              <button onclick="selectOopMatchForMove('${escapeJs(mKey)}')" class="px-1.5 py-0.5 text-[10px] font-bold bg-white border border-slate-300 rounded hover:bg-blue-50 text-blue-700">挿入移動</button>
            </div>
            <div class="text-[10px] text-slate-700 truncate">${escapeHtml(m.p1Label)}</div>
            <div class="text-[10px] text-slate-700 truncate">vs ${escapeHtml(m.p2Label)}</div>
          </div>`;
        }).join('');
      }
    }
  }

  function renderOopBlankCardHtml(item, courtKey) {
    const bKey = item.key;
    const isSelected = opsUi.oopSelectedMoveKey === bKey;
    const label = (state.oop.blankLabels && state.oop.blankLabels[bKey]) || '空き（進行調整）';
    return `
      <div draggable="true"
           ondragstart="handleOopDragStart(event, '${escapeJs(bKey)}')"
           ondragend="handleOopDragEnd(event)"
           ondragover="handleOopCardDragOver(event)"
           ondragleave="handleOopCardDragLeave(event)"
           ondrop="handleOopCardDrop(event, '${courtKey}', ${item.indexInCourt})"
           class="oop-match-card ${isSelected ? 'oop-selected-move' : ''} w-[216px] shrink-0 p-2 rounded-lg border-2 border-dashed border-amber-300 bg-amber-50/50 text-[11px] flex flex-col justify-between gap-1.5 cursor-grab active:cursor-grabbing">
        <div class="flex items-center justify-between gap-1">
          <span class="px-1.5 py-0.5 rounded bg-amber-200/80 text-amber-950 text-[9.5px] font-extrabold">空き枠・進行調整</span>
          <div class="flex items-center gap-1">
            <button onclick="selectOopMatchForMove('${escapeJs(bKey)}')" class="px-1.5 py-0.2 text-[9px] font-bold rounded border bg-white text-slate-600 border-slate-300 hover:bg-blue-50">移動</button>
            <button onclick="removeOopItem('${escapeJs(bKey)}')" class="px-1 text-[11px] font-bold text-slate-400 hover:text-rose-600" title="空き枠を削除">×</button>
          </div>
        </div>
        <div class="text-center py-2 px-1">
          <div class="text-xs font-extrabold text-amber-900 truncate" title="${escapeHtml(label)}">${escapeHtml(label)}</div>
          <div class="text-[9.5px] text-amber-700/80 mt-0.5">この列は意図的に空けて進行を揃えます</div>
        </div>
        <div class="flex items-center justify-between gap-1">
          <button onclick="duplicateOopBlankRight('${courtKey}', ${item.indexInCourt})" class="px-2 py-0.5 text-[9.5px] font-bold bg-amber-100 hover:bg-amber-200 border border-amber-300 rounded text-amber-950" title="すぐ右隣に同じ空き枠をもう1つ追加">＋1枠</button>
          <button onclick="editOopBlankLabel('${escapeJs(bKey)}')" class="px-2 py-0.5 text-[9.5px] font-bold bg-white hover:bg-amber-100 border border-amber-300 rounded text-amber-900">ラベル変更</button>
        </div>
      </div>
    `;
  }

  function renderOopTimelineCardHtml(item, courtKey, courtKeys, matchMap, formatter, cardConflicts) {
    const m = item.match;
    const st = item.status;
    const role = item.role;
    const conflicts = Array.isArray(cardConflicts) ? cardConflicts : [];
    const hasConflict = conflicts.length > 0;
    const isSelected = opsUi.oopSelectedMoveKey === m.key;
    const isWoOpen = opsUi.oopOpenWoKey === m.key;
    const isFinished = st === 'finished' || st === 'wo';
    const umpire = OopEngine.getUmpireLabel(courtKeys, item.indexInCourt, matchMap, state.oop, formatter);

    let cardStyle = 'bg-white border-slate-200';
    let rolePill = '';
    if (st === 'playing') {
      cardStyle = 'bg-rose-50/85 border-rose-400 ring-1 ring-rose-300';
      rolePill = `<span class="px-1.5 py-0.2 rounded text-[9px] font-bold bg-rose-600 text-white">試合中</span>`;
    } else if (st === 'calling') {
      cardStyle = 'bg-amber-50/90 border-amber-400 ring-1 ring-amber-300';
      rolePill = `<span class="px-1.5 py-0.2 rounded text-[9px] font-bold bg-amber-500 text-white">コール</span>`;
    } else if (isFinished) {
      cardStyle = 'bg-slate-50/90 border-slate-200 opacity-85';
      const sc = m.result ? `${m.result.g1}-${m.result.g2}` : '済';
      rolePill = `<span class="px-1.5 py-0.2 rounded text-[9px] font-bold bg-emerald-100 text-emerald-800">${escapeHtml(sc)}</span>`;
    } else if (role === 'NEXT' || role === 'READY') {
      cardStyle = 'bg-blue-50/50 border-blue-300';
      rolePill = `<span class="px-1.5 py-0.2 rounded text-[9px] font-bold bg-blue-600 text-white">${role === 'NEXT' ? '次試合' : '次予定'}</span>`;
    } else if (role === 'ONDECK') {
      cardStyle = 'bg-sky-50/30 border-slate-300';
      rolePill = `<span class="px-1.5 py-0.2 rounded text-[9px] font-bold bg-slate-200 text-slate-700">次々</span>`;
    }
    if (hasConflict) {
      cardStyle = 'bg-rose-50 border-rose-500 ring-2 ring-rose-400/70';
    }

    const mtBadge = m.isConsolation
      ? `<span class="px-1 rounded text-[9px] font-bold bg-purple-100 text-purple-800">順位</span>`
      : `<span class="px-1 rounded text-[9px] font-bold bg-slate-100 text-slate-600">${escapeHtml(m.mountain)}</span>`;

    const readyDisabled = !m.isReady ? 'opacity-40 cursor-not-allowed' : '';
    const canMove = !isFinished;
    const dragAttrs = canMove
      ? `draggable="true" ondragstart="handleOopDragStart(event, '${escapeJs(m.key)}')" ondragend="handleOopDragEnd(event)"`
      : '';

    const win0 = Boolean(m.result && m.result.winner === 0);
    const win1 = Boolean(m.result && m.result.winner === 1);

    let actionBtns = '';
    if (isFinished) {
      actionBtns = `
        <button onclick="openOopScoreInput('${escapeJs(m.key)}')"
                class="px-2 py-0.5 text-[10px] font-bold rounded bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 transition">結果修正</button>
      `;
    } else if (st === 'waiting') {
      actionBtns = `
        <button ${!m.isReady ? 'disabled' : ''} onclick="setOopMatchStatus('${escapeJs(m.key)}', 'calling')"
                class="px-1.5 py-0.5 text-[10px] font-bold rounded bg-amber-500 hover:bg-amber-400 text-white transition ${readyDisabled}">コール</button>
        <button ${!m.isReady ? 'disabled' : ''} onclick="setOopMatchStatus('${escapeJs(m.key)}', 'playing')"
                class="px-1.5 py-0.5 text-[10px] font-bold rounded bg-rose-600 hover:bg-rose-500 text-white transition ${readyDisabled}">開始</button>
        <button ${!m.isReady ? 'disabled' : ''} onclick="openOopScoreInput('${escapeJs(m.key)}')"
                class="px-1.5 py-0.5 text-[10px] font-bold rounded bg-blue-600 hover:bg-blue-500 text-white transition ${readyDisabled}">結果</button>
        <button ${!m.isReady ? 'disabled' : ''} onclick="toggleOopWoChooser('${escapeJs(m.key)}')"
                class="px-1.5 py-0.5 text-[10px] font-bold rounded bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300 transition ${readyDisabled}">棄権</button>
      `;
    } else if (st === 'calling') {
      actionBtns = `
        <button onclick="setOopMatchStatus('${escapeJs(m.key)}', 'playing')"
                class="px-2 py-0.5 text-[10px] font-bold rounded bg-rose-600 hover:bg-rose-500 text-white transition">開始</button>
        <button onclick="openOopScoreInput('${escapeJs(m.key)}')"
                class="px-1.5 py-0.5 text-[10px] font-bold rounded bg-blue-600 hover:bg-blue-500 text-white transition">結果</button>
        <button onclick="toggleOopWoChooser('${escapeJs(m.key)}')"
                class="px-1.5 py-0.5 text-[10px] font-bold rounded bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 transition">棄権</button>
        <button onclick="setOopMatchStatus('${escapeJs(m.key)}', 'waiting')"
                class="px-1 py-0.5 text-[10px] font-bold text-slate-500 hover:text-slate-800">戻す</button>
      `;
    } else if (st === 'playing') {
      actionBtns = `
        <button onclick="openOopScoreInput('${escapeJs(m.key)}')"
                class="px-2 py-0.5 text-[10px] font-bold rounded bg-blue-600 hover:bg-blue-500 text-white shadow-sm transition">スコア入力</button>
        <button onclick="toggleOopWoChooser('${escapeJs(m.key)}')"
                class="px-1.5 py-0.5 text-[10px] font-bold rounded bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 transition">棄権</button>
        <button onclick="setOopMatchStatus('${escapeJs(m.key)}', 'waiting')"
                class="px-1 py-0.5 text-[10px] font-bold text-slate-500 hover:text-slate-800">戻す</button>
      `;
    }

    const woPanel = isWoOpen ? `
      <div class="mt-1 p-1.5 rounded bg-amber-50 border border-amber-300 space-y-1 text-[10px]">
        <div class="font-bold text-amber-950">不戦勝ペアを選択：</div>
        <button onclick="recordOopWalkover('${escapeJs(m.key)}', 0)" class="w-full text-left px-1.5 py-0.5 rounded bg-white hover:bg-emerald-50 border border-emerald-300 font-bold text-emerald-900 truncate">○ ${escapeHtml(m.p1Label)}</button>
        <button onclick="recordOopWalkover('${escapeJs(m.key)}', 1)" class="w-full text-left px-1.5 py-0.5 rounded bg-white hover:bg-emerald-50 border border-emerald-300 font-bold text-emerald-900 truncate">○ ${escapeHtml(m.p2Label)}</button>
      </div>` : '';

    return `
      <div data-oop-card-key="${escapeHtml(m.key)}"
           ${dragAttrs}
           ondragover="handleOopCardDragOver(event)"
           ondragleave="handleOopCardDragLeave(event)"
           ondrop="handleOopCardDrop(event, '${courtKey}', ${item.indexInCourt})"
           class="oop-match-card ${isSelected ? 'oop-selected-move' : ''} w-[216px] shrink-0 p-2 rounded-lg border ${cardStyle} text-[11px] flex flex-col justify-between gap-1 ${canMove ? 'cursor-grab active:cursor-grabbing' : ''}">
        <div class="space-y-1">
          <!-- 1行目：試合番号・ラウンド・状態・移動・外す -->
          <div class="flex items-center justify-between gap-1">
            <div class="flex items-center gap-1 min-w-0">
              <span class="font-extrabold text-slate-900 text-[11px]">${escapeHtml(m.matchNo)}</span>
              <span class="text-[10px] font-bold text-slate-500 truncate">${escapeHtml(m.roundLabel)}</span>
              ${mtBadge}
            </div>
            <div class="flex items-center gap-0.5 shrink-0">
              ${rolePill}
              ${canMove ? `<button onclick="selectOopMatchForMove('${escapeJs(m.key)}')" class="px-1.5 py-0.2 text-[9px] font-bold rounded border ${isSelected ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-300 hover:bg-blue-50'}" title="別の場所・試合間へ挿入移動">移動</button>` : ''}
              ${st === 'waiting' ? `<button onclick="removeOopItem('${escapeJs(m.key)}')" class="px-1 text-[11px] font-bold text-slate-400 hover:text-rose-600" title="進行表から外して未配置に戻す">×</button>` : ''}
            </div>
          </div>

          ${hasConflict ? `
            <div class="px-1.5 py-0.5 rounded bg-rose-600 text-white text-[9.5px] font-bold leading-tight" title="${escapeHtml(conflicts.map((x) => x.message).join(' / '))}">
              ⚠ 時系列矛盾: ${escapeHtml(conflicts[0].message)}
            </div>` : ''}

          <!-- 2行目：対戦カード -->
          <div class="pl-1.5 border-l-2 ${m.isReady ? 'border-slate-300' : 'border-dashed border-slate-200'} space-y-0.5 leading-tight">
            <div class="text-[11px] truncate ${win0 ? 'font-extrabold text-rose-700' : (m.p1Entry ? 'font-bold text-slate-900' : 'text-slate-400 italic')}" title="${escapeHtml(m.p1Label)}">${escapeHtml(m.p1Label)}</div>
            <div class="text-[11px] truncate ${win1 ? 'font-extrabold text-rose-700' : (m.p2Entry ? 'font-bold text-slate-900' : 'text-slate-400 italic')}" title="${escapeHtml(m.p2Label)}">${escapeHtml(m.p2Label)}</div>
          </div>

          <!-- 3行目：審判メモ -->
          ${umpire.text || state.oop.umpireRule !== 'none' ? `
            <div class="flex items-center justify-between gap-1 text-[9.5px] ${umpire.isCustom ? 'bg-indigo-50 text-indigo-800 border-indigo-200' : 'bg-slate-50 text-slate-600 border-slate-200'} border rounded px-1.5 py-0.2">
              <span class="truncate" title="${escapeHtml(umpire.text || '審判未設定')}">審: ${escapeHtml(umpire.text || '未設定')}</span>
              <button onclick="openUmpireModal('${escapeJs(m.key)}')" class="text-[9px] font-bold text-blue-600 hover:underline shrink-0">✎</button>
            </div>` : ''}
        </div>

        <!-- 4行目：ステータス操作ボタン -->
        <div class="flex items-center gap-1 pt-0.5 flex-wrap">
          ${actionBtns}
        </div>
        ${woPanel}
      </div>
    `;
  }

  function renderOopFinishedMiniCardHtml(item) {
    const m = item.match;
    const res = m.result;
    const scoreStr = res ? `${res.g1}-${res.g2}` : '終了';
    const win0 = res && res.winner === 0;
    const win1 = res && res.winner === 1;
    return `
      <div class="p-1.5 rounded bg-white border border-slate-200 text-[10px] flex items-center justify-between gap-2">
        <div class="min-w-0 flex-1">
          <div class="flex items-center gap-1">
            <span class="font-bold text-slate-500">${escapeHtml(m.matchNo)}</span>
            <span class="px-1 rounded bg-emerald-100 text-emerald-800 font-bold">${escapeHtml(scoreStr)}</span>
            ${m.isWo ? `<span class="px-1 rounded bg-amber-100 text-amber-800 font-bold">WO</span>` : ''}
          </div>
          <div class="truncate ${win0 ? 'font-bold text-rose-700' : 'text-slate-500'}">${escapeHtml(m.p1Label)}</div>
          <div class="truncate ${win1 ? 'font-bold text-rose-700' : 'text-slate-500'}">${escapeHtml(m.p2Label)}</div>
        </div>
        <button onclick="openOopScoreInput('${escapeJs(m.key)}')"
                class="px-1.5 py-0.5 text-[9px] font-bold bg-slate-100 hover:bg-slate-200 text-slate-700 rounded shrink-0">修正</button>
      </div>
    `;
  }

  function renderOopPrintSheet(catalog) {
    const sheet = document.getElementById('oop-print-sheet');
    if (!sheet) return;
    const ev = getActiveEvent();
    const nCourts = state.oop.courtCount || 4;
    const formatter = (entry) => OopEngine.formatSpectatorPairName(entry, 'full', formatSchoolShortName);

    // 8面を超える場合は読みやすさのため8面ずつ段分け
    const chunkSize = nCourts > 8 ? Math.ceil(nCourts / 2) : nCourts;
    const chunks = [];
    for (let start = 1; start <= nCourts; start += chunkSize) {
      const end = Math.min(nCourts, start + chunkSize - 1);
      const list = [];
      for (let c = start; c <= end; c++) list.push(String(c));
      chunks.push(list);
    }

    const tablesHtml = chunks.map((courtList) => {
      let maxRows = 0;
      courtList.forEach((cKey) => {
        const len = (state.oop.courts[cKey] || []).length;
        if (len > maxRows) maxRows = len;
      });
      maxRows = Math.max(1, maxRows);

      let rowsHtml = '';
      for (let r = 0; r < maxRows; r++) {
        rowsHtml += `<tr class="align-top">
          <td class="border border-slate-400 bg-slate-50 text-center font-bold text-[11px] px-1.5 py-2 w-10">第${r + 1}<br/>試合</td>
          ${courtList.map((cKey) => {
            const cKeys = state.oop.courts[cKey] || [];
            const mKey = cKeys[r];
            if (OopEngine.isBlankKey(mKey)) {
              const bLabel = (state.oop.blankLabels && state.oop.blankLabels[mKey]) || '空き（進行調整）';
              return `<td class="border border-slate-400 bg-amber-50/40 p-2 text-center text-[10px] font-bold text-amber-900 align-middle">${escapeHtml(bLabel)}</td>`;
            }
            const m = mKey ? catalog.matchMap[mKey] : null;
            if (!m) {
              return `<td class="border border-slate-300 bg-slate-50/40 p-1.5 text-center text-[10px] text-slate-300">-</td>`;
            }
            const umpire = OopEngine.getUmpireLabel(cKeys, r, catalog.matchMap, state.oop, formatter);
            const res = m.result;
            const st = state.oop.statuses[mKey] || 'waiting';
            const scoreBox = res
              ? `<span class="px-1.5 py-0.5 rounded bg-rose-50 border border-rose-300 text-rose-700 font-extrabold text-[10px]">${escapeHtml(res.g1)} - ${escapeHtml(res.g2)}</span>`
              : (st === 'playing'
                ? `<span class="px-1.5 py-0.5 rounded bg-rose-600 text-white font-bold text-[9px]">試合中</span>`
                : (st === 'calling'
                  ? `<span class="px-1.5 py-0.5 rounded bg-amber-500 text-white font-bold text-[9px]">コール中</span>`
                  : `<span class="inline-block px-2 py-0.5 border border-slate-300 rounded text-[9px] text-slate-400">　 - 　</span>`));
            return `<td class="border border-slate-400 p-1.5 text-[10px] leading-snug">
              <div class="flex items-center justify-between gap-1 mb-1 border-b border-slate-200 pb-0.5">
                <span class="font-extrabold text-slate-900">${escapeHtml(m.matchNo)} <span class="font-semibold text-slate-600">${escapeHtml(m.roundLabel)}</span></span>
                ${scoreBox}
              </div>
              <div class="font-bold text-slate-900 break-words">${escapeHtml(m.p1Label)}</div>
              <div class="text-[9px] text-slate-400 font-bold">vs</div>
              <div class="font-bold text-slate-900 break-words">${escapeHtml(m.p2Label)}</div>
              ${umpire.text ? `<div class="mt-1 pt-0.5 border-t border-dashed border-slate-200 text-[9px] text-slate-600">審判: ${escapeHtml(umpire.text)}</div>` : ''}
            </td>`;
          }).join('')}
        </tr>`;
      }

      return `
        <div class="overflow-x-auto mb-5">
          <table class="w-full border-collapse table-fixed">
            <thead>
              <tr class="bg-slate-800 text-white text-xs">
                <th class="border border-slate-600 py-1.5 px-1 w-10">順</th>
                ${courtList.map((cKey) => `<th class="border border-slate-600 py-1.5 px-2">コート ${cKey}</th>`).join('')}
              </tr>
            </thead>
            <tbody>${rowsHtml}</tbody>
          </table>
        </div>
      `;
    }).join('');

    sheet.innerHTML = `
      <div class="mb-4 pb-2.5 border-b-2 border-slate-800 flex flex-wrap items-baseline justify-between gap-3">
        <div>
          <div class="text-[11px] font-bold text-slate-500">${escapeHtml(opsUi.bundle ? opsUi.bundle.name : '')}</div>
          <h2 class="text-lg font-extrabold text-slate-900">${escapeHtml(ev ? ev.name : state.title)} — コート別進行表（OOP）</h2>
        </div>
        <div class="flex flex-wrap items-center gap-3 text-xs font-bold text-slate-700">
          <span>会場: ${escapeHtml(state.oop.venue || '第1会場')}</span>
          <span>コート数: ${nCourts}面</span>
          <span>${state.matchFormat}ゲームマッチ</span>
          <span>全${catalog.matches.length}試合</span>
        </div>
      </div>
      ${tablesHtml}
    `;
  }

  // OOP操作アクション
  function updateOopVenue(val) {
    if (!state.oop) state.oop = OopEngine.createDefaultOopState();
    state.oop.venue = String(val || '').trim() || '第1会場';
    const ev = getActiveEvent();
    if (ev) ev.venue = state.oop.venue;
    saveToLocalStorage();
    renderEventSwitcherBar();
  }

  function pushOopUndo() {
    if (!state.oop) return;
    opsUi.oopUndoStack.push({
      courtCount: state.oop.courtCount,
      courts: cloneDeep(state.oop.courts || {}),
      statuses: cloneDeep(state.oop.statuses || {}),
      blankLabels: cloneDeep(state.oop.blankLabels || {}),
      manualEmpty: Boolean(state.oop.manualEmpty)
    });
    if (opsUi.oopUndoStack.length > 30) opsUi.oopUndoStack.shift();
  }

  function undoOopAction() {
    if (!opsUi.oopUndoStack.length || !state.oop) return;
    const prev = opsUi.oopUndoStack.pop();
    state.oop.courtCount = prev.courtCount;
    state.oop.courts = prev.courts;
    state.oop.statuses = prev.statuses;
    state.oop.blankLabels = prev.blankLabels || {};
    state.oop.manualEmpty = Boolean(prev.manualEmpty);
    state.oop.updatedAt = new Date().toISOString();
    saveToLocalStorage();
    renderOopView();
    showToast('直前の進行表変更を元に戻しました', 'info', 1800);
  }

  function changeOopCourtCount(val) {
    const nextCount = Math.max(1, Math.min(24, parseInt(val, 10) || 4));
    pushOopUndo();
    const catalog = getActiveMatchCatalog('full');
    state.oop.courtCount = nextCount;
    state.oop.courts = OopEngine.generateCourtAssignments(catalog.matches, nextCount, state.oop.assignMode, state.oop, true);
    state.oop.updatedAt = new Date().toISOString();
    saveToLocalStorage();
    renderOopView();
    showToast(`コート数を${nextCount}面に変更し、待機試合を再配分しました`, 'success');
  }

  function setOopAssignMode(mode) {
    state.oop.assignMode = mode === 'round' ? 'round' : 'block';
    saveToLocalStorage();
    renderOopView();
  }

  function regenerateOopSchedule(preserveActive) {
    const catalog = getActiveMatchCatalog('full');
    if (!preserveActive && !confirm('試合中・終了済みの位置も含めて全試合のコート割当を初期ルールで再作成しますか？（スコア結果は消えません）')) {
      return;
    }
    pushOopUndo();
    state.oop.courts = OopEngine.generateCourtAssignments(
      catalog.matches,
      state.oop.courtCount,
      state.oop.assignMode,
      state.oop,
      Boolean(preserveActive)
    );
    state.oop.updatedAt = new Date().toISOString();
    saveToLocalStorage();
    renderOopView();
    showToast(preserveActive ? '進行中・終了済みを維持して待機試合を再配分しました' : '進行表を初期ルールで再生成しました', 'success');
  }

  function toggleOopUmpireRule() {
    state.oop.umpireRule = (state.oop.umpireRule === 'none') ? 'loser' : 'none';
    saveToLocalStorage();
    renderOopView();
  }

  function setOopViewStyle(style) {
    state.oop.viewStyle = ['board', 'split', 'table'].includes(style) ? style : 'board';
    saveToLocalStorage();
    renderOopView();
  }

  function setOopSplitBlock(blk) {
    opsUi.oopSplitBlock = ['ALL', 'A', 'B', 'C', 'D', 'FINALS'].includes(blk) ? blk : 'ALL';
    renderOopView();
  }

  function toggleOopConflictsDrawer() {
    opsUi.oopConflictsOpen = !opsUi.oopConflictsOpen;
    renderOopView();
  }

  function highlightOopCard(matchKey) {
    document.querySelectorAll('[data-oop-card-key]').forEach((el) => {
      el.classList.toggle('oop-card-hover-link', Boolean(matchKey) && el.dataset.oopCardKey === matchKey);
    });
  }

  function jumpToOopCard(matchKey) {
    const el = document.querySelector(`[data-oop-card-key="${matchKey}"]`);
    if (el) {
      el.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
      el.classList.add('oop-card-hover-link');
      setTimeout(() => el.classList.remove('oop-card-hover-link'), 2200);
    }
  }

  function selectNewOopBlankForInsert() {
    opsUi.oopSelectedMoveKey = (opsUi.oopSelectedMoveKey === '__NEW_BLANK__') ? null : '__NEW_BLANK__';
    renderOopView();
  }

  function editOopBlankLabel(blankKey) {
    const cur = (state.oop.blankLabels && state.oop.blankLabels[blankKey]) || '空き（進行調整）';
    const next = prompt('空き枠のラベルを入力してください（例：ベスト8揃え待機、昼休憩、コート整備）', cur);
    if (next === null) return;
    if (!state.oop.blankLabels) state.oop.blankLabels = {};
    state.oop.blankLabels[blankKey] = String(next).trim() || '空き（進行調整）';
    state.oop.updatedAt = new Date().toISOString();
    saveToLocalStorage();
    renderOopView();
  }

  function removeOopItem(itemKey) {
    if (!state.oop || !state.oop.courts) return;
    pushOopUndo();
    Object.keys(state.oop.courts).forEach((cKey) => {
      state.oop.courts[cKey] = (state.oop.courts[cKey] || []).filter((k) => k !== itemKey);
    });
    if (OopEngine.isBlankKey(itemKey) && state.oop.blankLabels) {
      delete state.oop.blankLabels[itemKey];
    }
    state.oop.manualEmpty = true;
    state.oop.updatedAt = new Date().toISOString();
    saveToLocalStorage();
    renderOopView();
    showToast(OopEngine.isBlankKey(itemKey) ? '空き枠を削除しました' : '試合を進行表から外して未配置に戻しました', 'info', 1800);
  }

  function alignOopQuarterfinals() {
    const catalog = getActiveMatchCatalog('full');
    const hasQF = catalog.matches.some((m) => m.roundLabel === '準々決勝');
    const targetLabel = hasQF ? '準々決勝' : '準決勝';
    pushOopUndo();
    const res = OopEngine.alignRoundWithBlanks(state.oop, catalog, targetLabel, `${targetLabel}揃え待機`);
    state.oop.courts = res.courts;
    state.oop.blankLabels = res.blankLabels;
    state.oop.updatedAt = new Date().toISOString();
    saveToLocalStorage();
    renderOopView();
    if (res.insertedCount > 0) {
      showToast(`${targetLabel}を第${res.alignedColumn}試合列へ揃えるため空き枠を${res.insertedCount}個挿入しました`, 'success');
    } else {
      showToast(`${targetLabel}の開始列はすでに揃っています`, 'info');
    }
  }

  function clearAllOopAssignments() {
    if (!confirm('待機中の試合をすべて未配置にして、左のトーナメントから手動で組み直しますか？（試合中・終了済みは保持されます）')) return;
    pushOopUndo();
    Object.keys(state.oop.courts || {}).forEach((cKey) => {
      state.oop.courts[cKey] = (state.oop.courts[cKey] || []).filter((k) => {
        if (OopEngine.isBlankKey(k)) return false;
        const st = state.oop.statuses[k];
        return st === 'playing' || st === 'calling' || st === 'finished' || st === 'wo';
      });
    });
    state.oop.blankLabels = {};
    state.oop.manualEmpty = true;
    state.oop.updatedAt = new Date().toISOString();
    saveToLocalStorage();
    renderOopView();
    showToast('待機試合を未配置に戻しました。左のT字アイコンから右へドラッグして配置できます。', 'info', 3000);
  }

  /**
   * 2分割モード左ペイン：トーナメント表＋各対戦のT字結合点にドラッグ可能なアイコンを描画
   */
  function renderOopSplitBracket(catalog) {
    const box = document.getElementById('oop-split-bracket-scroll');
    if (!box) return;

    // ブロック切替ボタンのスタイル同期
    ['ALL', 'A', 'B', 'C', 'D', 'FINALS'].forEach((b) => {
      const btn = document.getElementById(`btn-oop-blk-${b}`);
      if (!btn) return;
      btn.className = opsUi.oopSplitBlock === b
        ? 'px-2 py-0.5 text-[10px] font-bold rounded bg-slate-900 text-white shadow-sm'
        : 'px-2 py-0.5 text-[10px] font-bold rounded bg-slate-200 hover:bg-slate-300 text-slate-700';
    });

    // 各試合の現在のコート位置をマップ
    const locMap = {};
    Object.keys(state.oop.courts || {}).forEach((cKey) => {
      (state.oop.courts[cKey] || []).forEach((k, idx) => {
        if (!OopEngine.isBlankKey(k)) locMap[k] = { court: cKey, order: idx + 1 };
      });
    });

    const blk = opsUi.oopSplitBlock;
    const size = state.bracketSize || 16;
    const totalRounds = Math.round(Math.log2(size));
    const tree = catalog.tree;

    // 準決・決勝＋順位決定戦リスト表示モード
    if (blk === 'FINALS') {
      const finalMatches = catalog.matches.filter((m) => m.isConsolation || m.round >= Math.max(1, totalRounds - 1));
      if (!finalMatches.length) {
        box.innerHTML = `<div class="py-10 text-center text-xs text-slate-400">準決勝・決勝・順位戦がありません</div>`;
        return;
      }
      box.innerHTML = `<div class="space-y-2">${finalMatches.map((m) => {
        const loc = locMap[m.key];
        const isPicked = opsUi.oopSelectedMoveKey === m.key;
        const badgeCls = isPicked
          ? 'bg-amber-500 text-white border-amber-600'
          : (loc ? 'bg-emerald-50 text-emerald-700 border-emerald-400' : 'bg-blue-600 text-white border-blue-700');
        return `
          <div draggable="true"
               ondragstart="handleOopDragStart(event, '${escapeJs(m.key)}')"
               ondragend="handleOopDragEnd(event)"
               onmouseenter="highlightOopCard('${escapeJs(m.key)}')"
               onmouseleave="highlightOopCard(null)"
               class="p-2.5 rounded-lg border border-slate-200 bg-slate-50 flex items-center justify-between gap-2 cursor-grab">
            <div class="min-w-0 flex-1 text-xs">
              <div class="font-extrabold text-slate-800">${escapeHtml(m.matchNo)} <span class="text-[10px] text-slate-500">${escapeHtml(m.roundLabel)}</span></div>
              <div class="truncate font-semibold text-slate-700 mt-0.5">${escapeHtml(m.p1Label)} vs ${escapeHtml(m.p2Label)}</div>
            </div>
            <button type="button" onclick="selectOopMatchForMove('${escapeJs(m.key)}')"
                    class="px-2 py-1 rounded-full border text-[10px] font-extrabold shrink-0 ${badgeCls}">
              ${loc ? `C${loc.court}-${loc.order}` : `⠿ ${escapeHtml(m.matchNo)}`}
            </button>
          </div>
        `;
      }).join('')}</div>`;
      return;
    }

    // 全体 または A/B/C/D ブロックの部分トーナメント描画
    let startSlot = 0;
    let endSlot = size;
    let maxRound = totalRounds;
    if (blk !== 'ALL' && size >= 8) {
      const qMap = { A: 0, D: 1, C: 2, B: 3 };
      const qIdx = qMap[blk] ?? 0;
      const qSize = size / 4;
      startSlot = qIdx * qSize;
      endSlot = startSlot + qSize;
      maxRound = Math.max(1, totalRounds - 2);
    }

    const activeIndices = [];
    for (let i = startSlot; i < endSlot; i++) {
      if (state.activeSlotMask[i]) activeIndices.push(i);
    }
    const slotToNum = {};
    let dNo = 1;
    for (let i = 0; i < size; i++) {
      if (state.activeSlotMask[i]) slotToNum[i] = dNo++;
    }

    const rowH = activeIndices.length >= 24 ? 26 : 30;
    const labelW = 165;
    const stepW = Math.max(48, Math.min(62, Math.floor(240 / Math.max(1, maxRound))));
    const totalW = labelW + maxRound * stepW + 64;
    const totalH = Math.max(180, activeIndices.length * rowH + 28);

    const nodeY = { 0: {} };
    activeIndices.forEach((slotIdx, vRow) => {
      nodeY[0][slotIdx] = vRow * rowH + rowH / 2 + 12;
    });
    for (let r = 1; r <= maxRound; r++) {
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

    const effX = (node, round) => {
      if (round === 0) return labelW;
      if (node.isBye) return effX(node.c0.active ? node.c0 : node.c1, round - 1);
      return labelW + round * stepW;
    };

    let svgLines = '';
    let htmlItems = '';

    activeIndices.forEach((slotIdx) => {
      const leaf = tree[0][slotIdx];
      const y = nodeY[0][slotIdx];
      const nameStr = leaf.entry ? OopEngine.formatSpectatorPairName(leaf.entry, 'full', formatSchoolShortName) : '空き';
      htmlItems += `<div style="left:0; top:${y - 10}px; width:${labelW - 6}px; height:20px;"
                         class="absolute flex items-center gap-1 text-[11px] leading-tight">
        <span class="w-5 text-right font-bold text-slate-400 shrink-0">${slotToNum[slotIdx]}</span>
        <span class="font-bold text-slate-800 truncate" title="${escapeHtml(nameStr)}">${escapeHtml(nameStr)}</span>
      </div>`;
    });

    for (let r = 1; r <= maxRound; r++) {
      const xCurr = labelW + r * stepW;
      const mStart = startSlot / Math.pow(2, r);
      const mEnd = endSlot / Math.pow(2, r);
      for (let m = mStart; m < mEnd; m++) {
        const node = tree[r][m];
        if (!node.active || node.isBye) continue;
        const y0 = nodeY[r - 1][m * 2];
        const y1 = nodeY[r - 1][m * 2 + 1];
        const yMid = nodeY[r][m];
        const x0 = effX(node.c0, r - 1);
        const x1 = effX(node.c1, r - 1);
        const isDone = Boolean(node.result);
        const col = isDone ? '#dc2626' : '#64748b';
        svgLines += `<line x1="${x0}" y1="${y0}" x2="${xCurr}" y2="${y0}" stroke="${col}" stroke-width="1.4" />`;
        svgLines += `<line x1="${x1}" y1="${y1}" x2="${xCurr}" y2="${y1}" stroke="${col}" stroke-width="1.4" />`;
        svgLines += `<line x1="${xCurr}" y1="${y0}" x2="${xCurr}" y2="${y1}" stroke="${col}" stroke-width="1.4" />`;

        const mKey = `${r}_${m}`;
        const mObj = catalog.matchMap[mKey];
        if (mObj) {
          const loc = locMap[mKey];
          const isPicked = opsUi.oopSelectedMoveKey === mKey;
          const cls = isPicked
            ? 'bg-amber-500 text-white border-amber-600 ring-2 ring-amber-300'
            : (mObj.isFinished
              ? 'bg-slate-200 text-slate-600 border-slate-300'
              : (loc
                ? 'bg-emerald-50 text-emerald-800 border-emerald-500 shadow-sm'
                : 'bg-blue-600 text-white border-blue-700 shadow-sm'));
          const text = mObj.isFinished ? `済` : (loc ? `C${loc.court}-${loc.order}` : `⠿ ${mObj.matchNo}`);
          htmlItems += `
            <button type="button"
                    draggable="${!mObj.isFinished}"
                    ondragstart="handleOopDragStart(event, '${escapeJs(mKey)}')"
                    ondragend="handleOopDragEnd(event)"
                    onclick="selectOopMatchForMove('${escapeJs(mKey)}')"
                    onmouseenter="highlightOopCard('${escapeJs(mKey)}')"
                    onmouseleave="highlightOopCard(null)"
                    style="left:${xCurr}px; top:${yMid}px; transform:translate(-50%,-50%);"
                    title="${escapeHtml(mObj.matchNo)}（${escapeHtml(mObj.roundLabel)}）${escapeHtml(mObj.p1Label)} vs ${escapeHtml(mObj.p2Label)}"
                    class="oop-t-handle absolute px-1.5 py-0.5 rounded-full border text-[10px] font-extrabold whitespace-nowrap leading-none z-10 ${cls}">
              ${escapeHtml(text)}
            </button>
          `;
        }
      }
    }

    box.innerHTML = `
      <div class="relative" style="width:${totalW}px; height:${totalH}px;">
        <svg width="${totalW}" height="${totalH}" class="absolute inset-0 pointer-events-none">${svgLines}</svg>
        ${htmlItems}
      </div>
    `;
  }

  function toggleOopCourtFinished(courtKey) {
    opsUi.oopFinishedCollapsed[courtKey] = (opsUi.oopFinishedCollapsed[courtKey] === false);
    renderOopView();
  }

  function setOopMatchStatus(matchKey, newStatus) {
    if (!state.oop) return;
    if (newStatus === 'playing' || newStatus === 'calling') {
      // 同一コートに別の試合中がある場合、確認なしでスムーズに切り替えられるよう注意トーストのみ
      const courtOfKey = Object.keys(state.oop.courts).find((c) => (state.oop.courts[c] || []).includes(matchKey));
      if (courtOfKey && newStatus === 'playing') {
        const anotherPlaying = (state.oop.courts[courtOfKey] || []).find(
          (k) => k !== matchKey && state.oop.statuses[k] === 'playing'
        );
        if (anotherPlaying) {
          showToast(`コート${courtOfKey}には別の試合中カードがあります。必要に応じて終了処理してください。`, 'warn', 3200);
        }
      }
    }
    state.oop.statuses[matchKey] = newStatus;
    state.oop.updatedAt = new Date().toISOString();
    opsUi.oopOpenWoKey = null;
    saveToLocalStorage();
    renderOopView();
  }

  function openOopScoreInput(matchKey) {
    const catalog = getActiveMatchCatalog('full');
    const m = catalog.matchMap[matchKey];
    if (!m) return;
    opsUi.oopOpenWoKey = null;
    const rawKey = m.isConsolation ? String(m.consolationIndex) : m.key;
    openScoreModal(rawKey, m.p1Label, m.p2Label, m.isConsolation, true);
  }

  function toggleOopWoChooser(matchKey) {
    opsUi.oopOpenWoKey = (opsUi.oopOpenWoKey === matchKey) ? null : matchKey;
    renderOopView();
  }

  function recordOopWalkover(matchKey, winnerSide) {
    const catalog = getActiveMatchCatalog('full');
    const m = catalog.matchMap[matchKey];
    if (!m) return;
    pushHistory();
    const resObj = {
      winner: winnerSide === 0 ? 0 : 1,
      g1: winnerSide === 0 ? '○' : '棄',
      g2: winnerSide === 1 ? '○' : '棄',
      wo: true
    };
    if (m.isConsolation) {
      state.consolation.results[String(m.consolationIndex)] = resObj;
    } else {
      state.results[m.key] = resObj;
    }
    state.oop.statuses[matchKey] = 'wo';
    state.oop.updatedAt = new Date().toISOString();
    opsUi.oopOpenWoKey = null;
    saveToLocalStorage();
    renderSheet();
    renderOopView();
    showToast(`${m.matchNo} を棄権（WO）で確定し、トーナメント表へ反映しました`, 'success');
  }

  function selectOopMatchForMove(matchKey) {
    opsUi.oopSelectedMoveKey = (opsUi.oopSelectedMoveKey === matchKey) ? null : matchKey;
    renderOopView();
  }

  function cancelOopMoveSelection() {
    opsUi.oopSelectedMoveKey = null;
    renderOopView();
  }

  function moveOopMatchToCourt(matchKey, targetCourtKey, insertIndex) {
    if (!matchKey || !state.oop || !state.oop.courts[targetCourtKey]) return;
    pushOopUndo();
    const isContinuousBlank = (matchKey === '__NEW_BLANK__' && opsUi.oopSelectedMoveKey === '__NEW_BLANK__');
    let keyToInsert = matchKey;
    if (matchKey === '__NEW_BLANK__') {
      keyToInsert = `__blank_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 6)}`;
      if (!state.oop.blankLabels) state.oop.blankLabels = {};
      state.oop.blankLabels[keyToInsert] = '空き（進行調整）';
    }
    const moved = OopEngine.moveMatchInCourts(state.oop.courts, keyToInsert, targetCourtKey, insertIndex);
    state.oop.courts = moved.courts;
    state.oop.updatedAt = new Date().toISOString();
    // 空き枠クリック連続モード時は解除せず続けて複数設置できるようにする
    if (!isContinuousBlank) {
      opsUi.oopSelectedMoveKey = null;
      document.body.classList.remove('oop-move-mode-active');
    }
    document.body.classList.remove('oop-dragging-active');
    saveToLocalStorage();
    renderOopView();
    showToast(
      OopEngine.isBlankKey(keyToInsert)
        ? `コート${targetCourtKey}の第${moved.insertedIndex + 1}試合位置へ空き枠を挿入しました${isContinuousBlank ? '（続けて挿入可・Esc終了）' : ''}`
        : `コート${targetCourtKey}の第${moved.insertedIndex + 1}試合位置へ移動しました`,
      'info',
      2000
    );
  }

  function duplicateOopBlankRight(courtKey, idx) {
    if (!state.oop || !state.oop.courts[courtKey]) return;
    pushOopUndo();
    const origKey = state.oop.courts[courtKey][idx];
    const origLabel = (state.oop.blankLabels && state.oop.blankLabels[origKey]) || '空き（進行調整）';
    const newKey = `__blank_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 6)}`;
    if (!state.oop.blankLabels) state.oop.blankLabels = {};
    state.oop.blankLabels[newKey] = origLabel;
    state.oop.courts[courtKey].splice(idx + 1, 0, newKey);
    state.oop.updatedAt = new Date().toISOString();
    saveToLocalStorage();
    renderOopView();
    showToast(`コート${courtKey}に空き枠を追加しました`, 'info', 1600);
  }

  function toggleOopHideFinished() {
    opsUi.oopHideFinished = !opsUi.oopHideFinished;
    renderOopView();
  }

  function toggleOopQueueDrawer() {
    opsUi.oopQueueDrawerOpen = !opsUi.oopQueueDrawerOpen;
    renderOopView();
  }

  function reorderOopMatch(courtKey, matchKey, delta) {
    const list = state.oop.courts[courtKey];
    if (!Array.isArray(list)) return;
    const idx = list.indexOf(matchKey);
    if (idx < 0) return;
    const nextIdx = idx + delta;
    if (nextIdx < 0 || nextIdx >= list.length) return;
    const tmp = list[idx];
    list[idx] = list[nextIdx];
    list[nextIdx] = tmp;
    state.oop.updatedAt = new Date().toISOString();
    saveToLocalStorage();
    renderOopView();
  }

  function handleOopDragStart(event, matchKey) {
    opsUi.oopDragKey = matchKey;
    document.body.classList.add('oop-dragging-active');
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = 'move';
      event.dataTransfer.setData('text/plain', matchKey);
    }
  }

  function handleOopDragEnd() {
    opsUi.oopDragKey = null;
    document.body.classList.remove('oop-dragging-active');
    document.querySelectorAll('.oop-insert-hover,.oop-drop-left,.oop-drop-right,.oop-court-drop-active').forEach((el) => {
      el.classList.remove('oop-insert-hover', 'oop-drop-left', 'oop-drop-right', 'oop-court-drop-active');
    });
  }

  function handleOopDragOver(event) {
    event.preventDefault();
    if (event.currentTarget) event.currentTarget.classList.add('oop-court-drop-active');
  }

  function handleOopDragLeave(event) {
    if (event.currentTarget) event.currentTarget.classList.remove('oop-court-drop-active');
  }

  function handleOopDrop(event, targetCourtKey) {
    event.preventDefault();
    if (event.currentTarget) event.currentTarget.classList.remove('oop-court-drop-active');
    const key = opsUi.oopDragKey || (event.dataTransfer && event.dataTransfer.getData('text/plain'));
    handleOopDragEnd();
    if (key) moveOopMatchToCourt(key, targetCourtKey);
  }

  // 試合と試合の間の挿入スロット用ハンドラ
  function handleOopInsertZoneClick(targetCourtKey, insertIndex) {
    if (!opsUi.oopSelectedMoveKey) return;
    moveOopMatchToCourt(opsUi.oopSelectedMoveKey, targetCourtKey, insertIndex);
  }

  function handleOopInsertDragOver(event) {
    event.preventDefault();
    event.stopPropagation();
    if (event.currentTarget) event.currentTarget.classList.add('oop-insert-hover');
  }

  function handleOopInsertDragLeave(event) {
    if (event.currentTarget) event.currentTarget.classList.remove('oop-insert-hover');
  }

  function handleOopInsertDrop(event, targetCourtKey, insertIndex) {
    event.preventDefault();
    event.stopPropagation();
    const key = opsUi.oopDragKey || (event.dataTransfer && event.dataTransfer.getData('text/plain'));
    handleOopDragEnd();
    if (key) moveOopMatchToCourt(key, targetCourtKey, insertIndex);
  }

  // カード本体上にドロップされた場合も左半分なら前・右半分なら後ろへ正確に挿入
  function handleOopCardDragOver(event) {
    event.preventDefault();
    event.stopPropagation();
    const el = event.currentTarget;
    if (!el) return;
    const rect = el.getBoundingClientRect();
    const isLeft = event.clientX < rect.left + rect.width / 2;
    el.classList.toggle('oop-drop-left', isLeft);
    el.classList.toggle('oop-drop-right', !isLeft);
  }

  function handleOopCardDragLeave(event) {
    const el = event.currentTarget;
    if (el) el.classList.remove('oop-drop-left', 'oop-drop-right');
  }

  function handleOopCardDrop(event, targetCourtKey, cardIndex) {
    event.preventDefault();
    event.stopPropagation();
    const el = event.currentTarget;
    let insertIdx = cardIndex;
    if (el) {
      const rect = el.getBoundingClientRect();
      insertIdx = (event.clientX < rect.left + rect.width / 2) ? cardIndex : cardIndex + 1;
    }
    const key = opsUi.oopDragKey || (event.dataTransfer && event.dataTransfer.getData('text/plain'));
    handleOopDragEnd();
    if (key) moveOopMatchToCourt(key, targetCourtKey, insertIdx);
  }

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && opsUi.oopSelectedMoveKey) {
      cancelOopMoveSelection();
    }
  });

  function autoAssignUnassignedMatches() {
    const { unassigned } = syncOopWithBracket();
    if (unassigned.length === 0) {
      showToast('未割当の試合はありません', 'info');
      return;
    }
    const nCourts = state.oop.courtCount || 4;
    unassigned.forEach((mKey) => {
      let best = '1';
      for (let c = 2; c <= nCourts; c++) {
        if ((state.oop.courts[String(c)] || []).length < (state.oop.courts[best] || []).length) {
          best = String(c);
        }
      }
      state.oop.courts[best].push(mKey);
    });
    state.oop.updatedAt = new Date().toISOString();
    saveToLocalStorage();
    renderOopView();
    showToast(`未割当${unassigned.length}試合を空きコートへ投入しました`, 'success');
  }

  // 選手登録番号（トーナメント表示番号 1..N）からペア一覧を構築
  function buildActiveSlotNumberList() {
    const list = [];
    let num = 1;
    const size = state.bracketSize || (state.slots ? state.slots.length : 0);
    for (let i = 0; i < size; i++) {
      if (!state.activeSlotMask[i]) continue;
      const entry = state.slots[i];
      if (entry) {
        const pairStr = OopEngine.formatSpectatorPairName(entry, 'full', formatSchoolShortName);
        list.push({ num, slotIdx: i, entry, label: `${num}番 ${pairStr}` });
      }
      num++;
    }
    return list;
  }

  // 審判メモ編集モーダル（登録番号から引用指定対応）
  function openUmpireModal(matchKey) {
    const catalog = getActiveMatchCatalog('full');
    const m = catalog.matchMap[matchKey];
    if (!m) return;
    opsUi.umpireModalMatchKey = matchKey;
    document.getElementById('umpire-modal-match-label').textContent = `${m.matchNo}（${m.roundLabel}）：${m.p1Label} vs ${m.p2Label}`;
    document.getElementById('umpire-modal-input').value = (state.oop.umpires && state.oop.umpires[matchKey]) || '';

    const numInput = document.getElementById('umpire-slot-num-input');
    const previewEl = document.getElementById('umpire-slot-num-preview');
    const selEl = document.getElementById('umpire-pair-select');
    if (numInput) numInput.value = '';
    if (previewEl) previewEl.textContent = 'トーナメントの番号を入力すると該当ペアが表示されます';
    if (selEl) {
      const pairs = buildActiveSlotNumberList();
      selEl.innerHTML = `<option value="">一覧から登録番号・ペアを選択（全${pairs.length}組）...</option>` +
        pairs.map((p) => `<option value="${escapeHtml(p.label)}">${escapeHtml(p.label)}</option>`).join('');
    }

    document.getElementById('umpire-modal').classList.remove('hidden');
    setTimeout(() => {
      if (numInput) numInput.focus();
    }, 40);
  }

  function previewUmpireByNumber(val) {
    const previewEl = document.getElementById('umpire-slot-num-preview');
    if (!previewEl) return;
    const n = parseInt(val, 10);
    if (!n || n < 1) {
      previewEl.textContent = 'トーナメントの番号を入力すると該当ペアが表示されます';
      previewEl.className = 'text-[11px] font-semibold text-blue-800 min-h-[16px]';
      return;
    }
    const found = buildActiveSlotNumberList().find((p) => p.num === n);
    if (found) {
      previewEl.textContent = `一致: ${found.label}`;
      previewEl.className = 'text-[11px] font-bold text-emerald-700 min-h-[16px]';
    } else {
      previewEl.textContent = `番号 ${n} に該当する出場ペアがありません`;
      previewEl.className = 'text-[11px] font-bold text-rose-600 min-h-[16px]';
    }
  }

  function applyUmpireByNumber(appendMode) {
    const numInput = document.getElementById('umpire-slot-num-input');
    const textInput = document.getElementById('umpire-modal-input');
    if (!numInput || !textInput) return;
    const n = parseInt(numInput.value, 10);
    const found = buildActiveSlotNumberList().find((p) => p.num === n);
    if (!found) {
      showToast('該当する選手登録番号が見つかりません', 'warn');
      return;
    }
    if (appendMode && textInput.value.trim()) {
      textInput.value = `${textInput.value.trim()} ／ ${found.label}`;
    } else {
      textInput.value = found.label;
    }
    numInput.value = '';
    previewUmpireByNumber('');
    showToast(`${found.label} を審判に反映しました`, 'info', 1600);
  }

  function applyUmpireFromSelect(val) {
    if (!val) return;
    const textInput = document.getElementById('umpire-modal-input');
    if (textInput) textInput.value = val;
    const selEl = document.getElementById('umpire-pair-select');
    if (selEl) selEl.value = '';
  }

  function closeUmpireModal() {
    const m = document.getElementById('umpire-modal');
    if (m) m.classList.add('hidden');
    opsUi.umpireModalMatchKey = null;
  }

  function saveUmpireOverride() {
    const key = opsUi.umpireModalMatchKey;
    if (!key) return;
    const val = String(document.getElementById('umpire-modal-input').value || '').trim();
    if (!state.oop.umpires) state.oop.umpires = {};
    if (val) state.oop.umpires[key] = val;
    else delete state.oop.umpires[key];
    saveToLocalStorage();
    closeUmpireModal();
    renderOopView();
  }

  function clearUmpireOverride() {
    const key = opsUi.umpireModalMatchKey;
    if (!key) return;
    if (state.oop.umpires) delete state.oop.umpires[key];
    saveToLocalStorage();
    closeUmpireModal();
    renderOopView();
  }

  // 進行表 印刷・PDF・Excel 出力
  function printOopSheet() {
    const { catalog } = syncOopWithBracket();
    renderOopPrintSheet(catalog);
    const cleanup = () => {
      document.body.classList.remove('printing-oop');
      window.removeEventListener('afterprint', cleanup);
    };
    window.addEventListener('afterprint', cleanup);
    document.body.classList.add('printing-oop');
    setTimeout(() => window.print(), 120);
  }

  async function exportOopPdf() {
    if (typeof html2canvas === 'undefined' || typeof jspdf === 'undefined') {
      alert('PDF出力ライブラリの読み込みが完了していません。印刷メニューの「PDFに保存」をご利用ください。');
      return;
    }
    const prevStyle = state.oop.viewStyle;
    const { catalog } = syncOopWithBracket();
    state.oop.viewStyle = 'table';
    renderOopView();
    renderOopPrintSheet(catalog);

    const sheet = document.getElementById('oop-print-sheet');
    if (!sheet) return;
    showToast('進行表PDFを作成しています...', 'info', 2200);
    try {
      if (document.fonts && document.fonts.ready) await document.fonts.ready;
      document.body.classList.add('oop-pdf-capturing');
      await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));

      const canvas = await html2canvas(sheet, {
        scale: 2,
        backgroundColor: '#ffffff',
        useCORS: true,
        logging: false
      });
      const { jsPDF } = jspdf;
      const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();
      const margin = 8;
      const availW = pageW - margin * 2;
      const availH = pageH - margin * 2;
      const ratio = Math.min(availW / canvas.width, availH / canvas.height);
      const drawW = canvas.width * ratio;
      const drawH = canvas.height * ratio;
      const x = (pageW - drawW) / 2;
      pdf.addImage(canvas.toDataURL('image/png'), 'PNG', x, margin, drawW, drawH, undefined, 'FAST');

      const ev = getActiveEvent();
      const safeName = `${(opsUi.bundle && opsUi.bundle.name) || '大会'}_${(ev && ev.name) || '進行表'}_進行表`.replace(/[\\/:*?"<>|]/g, '_');
      pdf.save(`${safeName}_${new Date().toISOString().slice(0, 10)}.pdf`);
      showToast('進行表PDFを保存しました', 'success');
    } catch (err) {
      alert('進行表PDFの作成中にエラーが発生しました: ' + err.message);
    } finally {
      document.body.classList.remove('oop-pdf-capturing');
      state.oop.viewStyle = prevStyle;
      renderOopView();
    }
  }

  function exportOopExcel() {
    if (typeof XLSX === 'undefined') {
      alert('Excel出力ライブラリが読み込まれていません。');
      return;
    }
    try {
      const { catalog } = syncOopWithBracket();
      const nCourts = state.oop.courtCount || 4;
      const formatter = (entry) => OopEngine.formatSpectatorPairName(entry, 'full', formatSchoolShortName);
      const wb = XLSX.utils.book_new();

      // Sheet 1: コート別進行表マトリクス
      let maxRows = 0;
      for (let c = 1; c <= nCourts; c++) {
        maxRows = Math.max(maxRows, (state.oop.courts[String(c)] || []).length);
      }
      const header = ['試合順'];
      for (let c = 1; c <= nCourts; c++) header.push(`コート${c}`);
      const matrix = [header];
      for (let r = 0; r < maxRows; r++) {
        const row = [`第${r + 1}試合`];
        for (let c = 1; c <= nCourts; c++) {
          const cKeys = state.oop.courts[String(c)] || [];
          const mKey = cKeys[r];
          if (OopEngine.isBlankKey(mKey)) {
            row.push((state.oop.blankLabels && state.oop.blankLabels[mKey]) || '空き（進行調整）');
            continue;
          }
          const m = mKey ? catalog.matchMap[mKey] : null;
          if (!m) {
            row.push('');
          } else {
            const ump = OopEngine.getUmpireLabel(cKeys, r, catalog.matchMap, state.oop, formatter);
            const score = m.result ? ` [${m.result.g1}-${m.result.g2}]` : '';
            row.push(`${m.matchNo}(${m.roundLabel})${score}\n${m.p1Label} vs ${m.p2Label}${ump.text ? `\n審判: ${ump.text}` : ''}`);
          }
        }
        matrix.push(row);
      }
      const ws1 = XLSX.utils.aoa_to_sheet(matrix);
      ws1['!cols'] = [{ wch: 10 }, ...Array.from({ length: nCourts }, () => ({ wch: 30 }))];
      XLSX.utils.book_append_sheet(wb, ws1, 'コート別進行表');

      // Sheet 2: 全試合一覧
      const statusJa = { waiting: '待機', calling: 'コール中', playing: '試合中', finished: '終了', wo: '棄権(WO)' };
      const listRows = [['コート', 'コート順', '試合番号', 'ラウンド', '山', 'ペア1', 'ペア2', '審判', 'ステータス', 'スコア']];
      for (let c = 1; c <= nCourts; c++) {
        const cKeys = state.oop.courts[String(c)] || [];
        cKeys.forEach((mKey, idx) => {
          const m = catalog.matchMap[mKey];
          if (!m) return;
          const ump = OopEngine.getUmpireLabel(cKeys, idx, catalog.matchMap, state.oop, formatter);
          const st = state.oop.statuses[mKey] || 'waiting';
          listRows.push([
            `コート${c}`,
            idx + 1,
            m.matchNo,
            m.roundLabel,
            m.mountain,
            m.p1Label,
            m.p2Label,
            ump.text || '',
            statusJa[st] || st,
            m.result ? `${m.result.g1}-${m.result.g2}` : ''
          ]);
        });
      }
      const ws2 = XLSX.utils.aoa_to_sheet(listRows);
      ws2['!cols'] = [
        { wch: 10 }, { wch: 9 }, { wch: 10 }, { wch: 12 }, { wch: 8 },
        { wch: 24 }, { wch: 24 }, { wch: 22 }, { wch: 10 }, { wch: 10 }
      ];
      XLSX.utils.book_append_sheet(wb, ws2, '全試合一覧');

      const ev = getActiveEvent();
      const safeName = `${(opsUi.bundle && opsUi.bundle.name) || '大会'}_${(ev && ev.name) || '進行表'}_進行表`.replace(/[\\/:*?"<>|]/g, '_');
      XLSX.writeFile(wb, `${safeName}_${new Date().toISOString().slice(0, 10)}.xlsx`);
      showToast('進行表Excelを保存しました', 'success');
    } catch (err) {
      alert('進行表Excelの保存に失敗しました: ' + err.message);
    }
  }

  // =========================================================================
  // 4. 公開設定・QRコード・会場スタッフPIN モーダル
  // =========================================================================
  function getPortalPreviewUrl() {
    const base = `${window.location.origin}${window.location.pathname}`;
    const tid = (opsUi.bundle && opsUi.bundle.id) || 'demo';
    return `${base}?public=${encodeURIComponent(tid)}`;
  }

  function updatePublishIndicator() {
    if (!opsUi.bundle || !opsUi.bundle.publish) return;
    const dot = document.getElementById('header-publish-dot');
    if (dot) {
      dot.className = opsUi.bundle.publish.enabled
        ? 'w-1.5 h-1.5 rounded-full bg-emerald-400'
        : 'w-1.5 h-1.5 rounded-full bg-slate-500';
    }
  }

  function refreshPublishCloudBanner() {
    const banner = document.getElementById('pub-cloud-status-banner');
    if (!banner) return;
    const isLogged = window.BracketCloud && window.BracketCloud.isLoggedIn();
    const isSynced = window.BracketCloud && window.BracketCloud.isSynced();
    if (isLogged && isSynced) {
      banner.className = 'p-3 rounded-xl border border-emerald-200 bg-emerald-50 text-emerald-950 text-xs flex flex-wrap items-center justify-between gap-2';
      banner.innerHTML = `
        <div class="flex items-center gap-2">
          <span class="w-2 h-2 rounded-full bg-emerald-500 shrink-0"></span>
          <span><b>クラウド公開同期中：</b>このQRコード・URLは別端末のスマホからリアルタイム閲覧できます。</span>
        </div>
        <button onclick="syncNowToCloud()" class="px-2.5 py-1 text-[11px] font-bold bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg shrink-0">今すぐ最新を同期</button>
      `;
    } else if (isLogged) {
      banner.className = 'p-3 rounded-xl border border-blue-200 bg-blue-50 text-blue-950 text-xs flex flex-wrap items-center justify-between gap-2';
      banner.innerHTML = `
        <span><b>ログイン済（未クラウド保存）：</b>別端末へQR配布する前に一度クラウド保存してください。</span>
        <button onclick="syncNowToCloud()" class="px-2.5 py-1 text-[11px] font-bold bg-blue-600 hover:bg-blue-500 text-white rounded-lg shrink-0">クラウドへ公開・保存</button>
      `;
    } else {
      banner.className = 'p-3 rounded-xl border border-amber-300 bg-amber-50 text-amber-950 text-xs flex flex-wrap items-center justify-between gap-2';
      banner.innerHTML = `
        <span><b>ローカルプレビュー専用（未ログイン）：</b>別端末のスマホからQR閲覧・スタッフ同期するにはGoogleログインが必要です。</span>
        <button onclick="signInWithGoogle()" class="px-2.5 py-1 text-[11px] font-bold bg-slate-900 hover:bg-slate-800 text-white rounded-lg shrink-0">Googleログイン</button>
      `;
    }
  }

  function getStaffEventUrl(eventId) {
    const base = `${window.location.origin}${window.location.pathname}`;
    const tid = (opsUi.bundle && opsUi.bundle.id) || 'demo';
    return `${base}?staff=${encodeURIComponent(tid)}&e=${encodeURIComponent(eventId)}`;
  }

  function renderPublishModalContents() {
    if (!opsUi.bundle) return;
    refreshPublishCloudBanner();
    const pub = opsUi.bundle.publish;
    const chk = document.getElementById('pub-enabled-checkbox');
    const txt = document.getElementById('pub-enabled-text');
    const passInput = document.getElementById('pub-passcode-input');
    if (chk) chk.checked = Boolean(pub.enabled);
    if (txt) {
      txt.textContent = pub.enabled ? '公開中' : '非公開';
      txt.className = pub.enabled ? 'ml-1.5 text-xs font-bold text-emerald-700' : 'ml-1.5 text-xs font-bold text-slate-500';
    }
    if (passInput && document.activeElement !== passInput) passInput.value = pub.passcode || '';
    document.querySelectorAll('input[name="pub-name-mode"]').forEach((r) => {
      r.checked = (r.value === (pub.nameDisplayMode || 'full'));
    });

    // 種目別スタッフPIN ＆ スタッフ専用URL一覧
    const pinsEl = document.getElementById('pub-staff-pins-list');
    if (pinsEl) {
      pinsEl.innerHTML = opsUi.bundle.events.map((ev) => `
        <div class="bg-white px-2.5 py-2 rounded-lg border border-amber-200 text-[11px] space-y-1">
          <div class="flex items-center justify-between gap-2">
            <div class="font-bold text-slate-900 truncate">${escapeHtml(ev.name)} <span class="text-[10px] text-slate-500 font-normal">(${escapeHtml(ev.venue || '第1会場')})</span></div>
            <span class="font-mono font-extrabold text-amber-900 bg-amber-100 px-2 py-0.5 rounded tracking-widest shrink-0">PIN: ${escapeHtml(ev.staffPin || '0000')}</span>
          </div>
          <div class="flex items-center justify-end gap-1.5">
            <button onclick="copyStaffUrl('${escapeJs(ev.id)}')" class="px-2 py-0.5 text-[10px] font-bold bg-blue-50 hover:bg-blue-100 border border-blue-200 rounded text-blue-800">スタッフURL＋PINコピー</button>
            <button onclick="copyStaffPin('${escapeJs(ev.id)}')" class="px-1.5 py-0.5 text-[10px] font-bold bg-slate-100 hover:bg-slate-200 rounded text-slate-700">PINのみ</button>
            <button onclick="regenerateStaffPin('${escapeJs(ev.id)}')" class="px-1.5 py-0.5 text-[10px] font-bold bg-amber-100 hover:bg-amber-200 rounded text-amber-900">再発行</button>
          </div>
        </div>
      `).join('');
    }

    // URL ＆ QRコード生成
    const urlStr = getPortalPreviewUrl();
    const urlInput = document.getElementById('pub-portal-url-input');
    if (urlInput) urlInput.value = urlStr;
    const qrBox = document.getElementById('pub-qr-box');
    if (qrBox) {
      if (typeof qrcode === 'function') {
        try {
          const qr = qrcode(0, 'M');
          qr.addData(urlStr);
          qr.make();
          qrBox.innerHTML = qr.createSvgTag({ cellSize: 4, margin: 2 });
          const svg = qrBox.querySelector('svg');
          if (svg) {
            svg.setAttribute('width', '152');
            svg.setAttribute('height', '152');
          }
        } catch (_) {
          qrBox.innerHTML = `<div class="text-[11px] text-slate-400 p-4">QR生成に失敗しました</div>`;
        }
      } else {
        qrBox.innerHTML = `<div class="text-[11px] text-slate-500 p-3">オフライン環境のためQRは<br/>オンライン接続時に自動生成されます</div>`;
      }
    }
  }

  function openPublishModal() {
    const curEv = getActiveEvent();
    if (curEv) curEv.data = captureActiveStateSnapshot();
    renderPublishModalContents();
    document.getElementById('publish-modal').classList.remove('hidden');
  }

  function closePublishModal() {
    document.getElementById('publish-modal').classList.add('hidden');
  }

  function updatePublishSettingsFromModal() {
    if (!opsUi.bundle) return;
    const enabled = Boolean(document.getElementById('pub-enabled-checkbox').checked);
    const passcode = String(document.getElementById('pub-passcode-input').value || '').trim();
    const checkedRadio = document.querySelector('input[name="pub-name-mode"]:checked');
    const nameDisplayMode = checkedRadio ? checkedRadio.value : 'full';
    opsUi.bundle.publish = { enabled, passcode, nameDisplayMode };
    saveBundleToLocalStorage();
    updatePublishIndicator();
    renderPublishModalContents();
    if (state.mainView === 'spectator') renderSpectatorView();
  }

  function copyPortalUrl() {
    const url = getPortalPreviewUrl();
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(url).then(() => showToast('公開ポータルURLをコピーしました', 'success'));
    } else {
      const inp = document.getElementById('pub-portal-url-input');
      if (inp) { inp.select(); document.execCommand('copy'); }
      showToast('公開ポータルURLをコピーしました', 'success');
    }
  }

  function copyStaffPin(eventId) {
    const ev = opsUi.bundle && opsUi.bundle.events.find((e) => e.id === eventId);
    if (!ev) return;
    const text = `${opsUi.bundle.name} / ${ev.name}（${ev.venue}） 運営スタッフPIN: ${ev.staffPin}`;
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(() => showToast(`${ev.name}のスタッフPINをコピーしました`, 'success'));
    } else {
      showToast(`PIN: ${ev.staffPin}`, 'info');
    }
  }

  function copyStaffUrl(eventId) {
    const ev = opsUi.bundle && opsUi.bundle.events.find((e) => e.id === eventId);
    if (!ev) return;
    const url = getStaffEventUrl(eventId);
    const text = `【${opsUi.bundle.name}】${ev.name}（${ev.venue || '第1会場'}）会場スタッフ用\nURL: ${url}\nPIN: ${ev.staffPin}`;
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(() => showToast(`${ev.name}のスタッフURL＋PINをコピーしました`, 'success'));
    } else {
      showToast(url, 'info');
    }
  }

  function regenerateStaffPin(eventId) {
    const ev = opsUi.bundle && opsUi.bundle.events.find((e) => e.id === eventId);
    if (!ev) return;
    const oldPin = ev.staffPin;
    ev.staffPin = generateStaffPin();
    ev.pinVersion = (ev.pinVersion || 1) + 1;
    saveBundleToLocalStorage();
    if (window.BracketCloud && typeof window.BracketCloud.deleteStaffKeyFromCloud === 'function') {
      window.BracketCloud.deleteStaffKeyFromCloud(opsUi.bundle.id, ev.id, oldPin);
      window.BracketCloud.scheduleCloudSync(true);
    }
    renderPublishModalContents();
    showToast(`${ev.name}のスタッフPINを再発行しました（旧PINは無効化）`, 'success');
  }

  // =========================================================================
  // 5. 生徒・保護者向け 公開ビュー（スマホ最適化3タブ）
  // =========================================================================
  function getSpectatorEvent() {
    if (!opsUi.bundle || !opsUi.bundle.events.length) return null;
    const activeEv = getActiveEvent();
    if (activeEv) activeEv.data = captureActiveStateSnapshot();
    const target = opsUi.bundle.events.find((ev) => ev.id === opsUi.spectator.eventId) || activeEv;
    opsUi.spectator.eventId = target.id;
    return target;
  }

  function setSpectatorFrameMode(mode) {
    opsUi.spectator.frameMode = mode === 'full' ? 'full' : 'mobile';
    const frame = document.getElementById('spectator-portal-frame');
    const btnM = document.getElementById('btn-spec-frame-mobile');
    const btnF = document.getElementById('btn-spec-frame-full');
    if (frame) {
      frame.className = opsUi.spectator.frameMode === 'full'
        ? 'w-full max-w-6xl h-full min-h-0 bg-slate-50 text-slate-900 rounded-2xl shadow-2xl border border-slate-700 overflow-hidden flex flex-col'
        : 'w-full max-w-[400px] h-full max-h-[820px] min-h-0 bg-slate-50 text-slate-900 rounded-2xl shadow-2xl border border-slate-700 overflow-hidden flex flex-col';
    }
    if (btnM && btnF) {
      const on = 'px-2 py-0.5 text-[11px] font-bold rounded bg-blue-600 text-white';
      const off = 'px-2 py-0.5 text-[11px] font-bold rounded text-slate-300 hover:text-white';
      btnM.className = opsUi.spectator.frameMode === 'mobile' ? on : off;
      btnF.className = opsUi.spectator.frameMode === 'full' ? on : off;
    }
  }

  function setSpectatorTab(tab) {
    opsUi.spectator.tab = tab;
    renderSpectatorView();
  }

  function selectSpectatorEvent(eventId) {
    opsUi.spectator.eventId = eventId;
    opsUi.spectator.schoolFilter = 'ALL';
    renderSpectatorView();
  }

  function submitSpectatorPasscode() {
    const inp = document.getElementById('spec-passcode-try');
    const val = inp ? String(inp.value || '').trim() : '';
    const expected = (opsUi.bundle && opsUi.bundle.publish && opsUi.bundle.publish.passcode) || '';
    if (val === expected) {
      opsUi.spectator.unlockedPasscode = val;
      renderSpectatorView();
    } else {
      showToast('パスコードが一致しません', 'error');
    }
  }

  function bypassSpectatorGate(kind) {
    if (kind === 'closed') opsUi.spectator.organizerBypassClosed = true;
    if (kind === 'passcode') opsUi.spectator.unlockedPasscode = (opsUi.bundle && opsUi.bundle.publish && opsUi.bundle.publish.passcode) || '';
    renderSpectatorView();
  }

  function toggleSpectatorFavorite(favId) {
    if (opsUi.spectator.favorites.has(favId)) opsUi.spectator.favorites.delete(favId);
    else opsUi.spectator.favorites.add(favId);
    saveSpectatorFavorites();
    if (opsUi.standalonePublic) renderStandalonePublicFrame();
    else renderSpectatorView();
  }

  function setSpectatorSearchQuery(q) {
    opsUi.spectator.searchQuery = String(q || '');
    if (opsUi.standalonePublic) renderStandalonePublicFrame();
    else renderSpectatorView();
    const inp = document.getElementById('spec-search-input');
    if (inp) {
      inp.focus();
      inp.setSelectionRange(inp.value.length, inp.value.length);
    }
  }

  function setSpectatorSchoolFilter(school) {
    opsUi.spectator.schoolFilter = school;
    if (opsUi.standalonePublic) renderStandalonePublicFrame();
    else renderSpectatorView();
  }

  function adjustSpectatorBracketZoom(delta) {
    opsUi.spectator.bracketZoom = Math.max(0.45, Math.min(1.4, Math.round((opsUi.spectator.bracketZoom + delta) * 100) / 100));
    if (opsUi.standalonePublic) renderStandalonePublicFrame();
    else renderSpectatorView();
  }

  function refreshActiveSpectatorPortal() {
    if (opsUi.standalonePublic) renderStandalonePublicFrame();
    else renderSpectatorView();
  }

  function setSpectatorCourtFilter(filterVal) {
    opsUi.spectator.courtFilter = ['ALL', 'ACTIVE', 'CHECKED'].includes(filterVal) ? filterVal : 'ALL';
    refreshActiveSpectatorPortal();
  }

  function toggleSpectatorCourtExpand(courtKey) {
    opsUi.spectator.expandedCourts[courtKey] = !opsUi.spectator.expandedCourts[courtKey];
    refreshActiveSpectatorPortal();
  }

  function setSpectatorBracketBlock(blk) {
    opsUi.spectator.bracketBlock = ['ALL', 'A', 'B', 'C', 'D'].includes(blk) ? blk : 'ALL';
    if (opsUi.spectator.bracketBlock !== 'ALL') opsUi.spectator.bracketStage = 'ALL';
    refreshActiveSpectatorPortal();
  }

  function setSpectatorBracketStage(stage) {
    opsUi.spectator.bracketStage = ['ALL', '32', '16', '8'].includes(stage) ? stage : 'ALL';
    if (opsUi.spectator.bracketStage !== 'ALL') opsUi.spectator.bracketBlock = 'ALL';
    refreshActiveSpectatorPortal();
  }

  function fitSpectatorBracketZoom() {
    const box = document.getElementById('spec-bracket-scroll');
    const canvas = document.getElementById('spec-bracket-inner');
    if (!box || !canvas) return;
    const rawW = parseFloat(canvas.dataset.rawWidth || '500');
    const availW = Math.max(260, box.clientWidth - 24);
    opsUi.spectator.bracketZoom = Math.max(0.42, Math.min(1.15, Math.floor((availW / rawW) * 100) / 100));
    refreshActiveSpectatorPortal();
  }

  function jumpSpectatorToFirstChecked() {
    const el = document.querySelector('#spec-bracket-scroll [data-spec-checked="1"]');
    if (el) {
      el.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
    } else {
      showToast('現在の表示範囲に★チェック中のペアがありません', 'info', 2000);
    }
  }

  function toggleSpectatorMyCheckCollapsed() {
    opsUi.spectator.myCheckCollapsed = !opsUi.spectator.myCheckCollapsed;
    refreshActiveSpectatorPortal();
  }

  function focusSpectatorCourt(courtKey) {
    opsUi.spectator.tab = 'courts';
    opsUi.spectator.courtFilter = 'ALL';
    opsUi.spectator.expandedCourts[String(courtKey)] = true;
    refreshActiveSpectatorPortal();
    setTimeout(() => {
      const el = document.querySelector(`[data-spec-court="${courtKey}"]`);
      if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }, 50);
  }

  function buildSpectatorEntryMetrics(evId, evData, catalog, oop) {
    const matchCourtMap = {};
    const courtAheadMap = {};
    Object.keys(oop.courts || {}).forEach((cKey) => {
      const list = oop.courts[cKey] || [];
      let unplayedBefore = 0;
      list.forEach((k, idx) => {
        if (OopEngine.isBlankKey(k)) return;
        const st = oop.statuses[k] || 'waiting';
        matchCourtMap[k] = { court: cKey, order: idx + 1 };
        courtAheadMap[k] = unplayedBefore;
        if (st !== 'finished' && st !== 'wo') unplayedBefore++;
      });
    });

    const getEntryStatus = (entry) => {
      const related = catalog.matches.filter(
        (m) => (m.p1Entry && m.p1Entry.id === entry.id) || (m.p2Entry && m.p2Entry.id === entry.id)
      );
      let stateKind = 'waiting';
      let badge = { text: '待機中', cls: 'bg-slate-100 text-slate-600' };
      let targetCourt = null;
      let waitCount = null;

      const activeMatch = related.find((m) => {
        const st = oop.statuses[m.key];
        return st === 'playing' || st === 'calling';
      });
      if (activeMatch) {
        const st = oop.statuses[activeMatch.key];
        const loc = matchCourtMap[activeMatch.key];
        targetCourt = loc ? loc.court : null;
        stateKind = st;
        badge = st === 'playing'
          ? { text: `C${loc ? loc.court : '?'} 試合中`, cls: 'bg-rose-600 text-white' }
          : { text: `C${loc ? loc.court : '?'} コール中`, cls: 'bg-amber-500 text-white' };
      } else {
        const nextMatch = related.find((m) => !m.isFinished);
        if (nextMatch) {
          const loc = matchCourtMap[nextMatch.key];
          targetCourt = loc ? loc.court : null;
          waitCount = typeof courtAheadMap[nextMatch.key] === 'number' ? courtAheadMap[nextMatch.key] : null;
          stateKind = waitCount === 0 ? 'next' : 'upcoming';
          const waitLabel = waitCount === 0 ? '次試合！' : (waitCount !== null ? `あと${waitCount}試合` : nextMatch.matchNo);
          badge = waitCount === 0
            ? { text: `C${loc ? loc.court : '?'} ${waitLabel}`, cls: 'bg-blue-600 text-white' }
            : { text: `C${loc ? loc.court : '?'} ${waitLabel}`, cls: 'bg-sky-100 text-sky-900 border border-sky-200' };
        } else if (related.length > 0) {
          const last = related[related.length - 1];
          const wonLast = last.winnerEntry && last.winnerEntry.id === entry.id;
          stateKind = wonLast ? 'advanced' : 'eliminated';
          badge = wonLast
            ? { text: '次戦相手待ち', cls: 'bg-emerald-100 text-emerald-800' }
            : { text: '試合終了', cls: 'bg-slate-200 text-slate-600' };
        }
      }
      return { stateKind, badge, related, targetCourt, waitCount };
    };

    return { matchCourtMap, courtAheadMap, getEntryStatus };
  }

  /**
   * プレビュー・別端末公開ポータル共通の BracketFlow オリジナルUIレンダラ
   */
  function renderUnifiedSpectatorPortal(frame, opts) {
    const {
      tournamentName,
      eventsList,
      activeEvent,
      evData,
      nameMode,
      tourAnnouncement,
      isStandalone
    } = opts;

    const aliases = evData.schoolAliases || {};
    const shortSchool = (raw) => {
      const k = String(raw || '').trim();
      if (Object.prototype.hasOwnProperty.call(aliases, k) && String(aliases[k] || '').trim()) return String(aliases[k]).trim();
      return formatSchoolShortName(k);
    };
    const pairFmt = (entry) => OopEngine.formatSpectatorPairName(entry, nameMode, shortSchool);

    const catalog = OopEngine.buildMatchCatalog({
      bracketSize: evData.bracketSize || 16,
      slots: evData.slots || [],
      activeSlotMask: evData.activeSlotMask || [],
      results: evData.results || {},
      consolation: evData.consolation || null,
      labelFormatter: pairFmt
    });
    const oop = OopEngine.reconcileOopState(evData.oop, catalog).oop;
    const metrics = buildSpectatorEntryMetrics(activeEvent.id, evData, catalog, oop);

    // 進捗率・ライブ稼働数
    let liveCourtCnt = 0;
    let finishedCnt = 0;
    catalog.matches.forEach((m) => {
      const st = oop.statuses[m.key] || 'waiting';
      if (st === 'playing' || st === 'calling') liveCourtCnt++;
      else if (st === 'finished' || st === 'wo') finishedCnt++;
    });
    const totalMatches = catalog.matches.length;
    const progressPct = totalMatches > 0 ? Math.round((finishedCnt / totalMatches) * 100) : 0;
    const candidateTimestamps = [
      oop.updatedAt,
      activeEvent && activeEvent.updatedAt,
      tourAnnouncement && tourAnnouncement.updatedAt,
      oop.announcement && oop.announcement.updatedAt
    ].filter((t) => typeof t === 'string' && !Number.isNaN(new Date(t).getTime()));
    candidateTimestamps.sort((a, b) => new Date(b).getTime() - new Date(a).getTime());
    const updatedClock = formatLocalClock(candidateTimestamps[0] || new Date().toISOString());

    // ① マイチェック（★登録ペア）省スペース1行チップバー（開閉対応）
    const entries = Array.isArray(evData.entries) ? evData.entries : [];
    const checkedEntries = entries.filter((e) => opsUi.spectator.favorites.has(`${activeEvent.id}::${e.id}`));
    const checkedEntryIdSet = new Set(checkedEntries.map((e) => e.id));
    const searchTabCmd = isStandalone ? "setStandaloneSpectatorTab('search')" : "setSpectatorTab('search')";

    let myCheckBarHtml = '';
    if (checkedEntries.length > 0) {
      if (opsUi.spectator.myCheckCollapsed) {
        const urgentCnt = checkedEntries.filter((e) => {
          const k = metrics.getEntryStatus(e).stateKind;
          return k === 'playing' || k === 'calling' || k === 'next';
        }).length;
        myCheckBarHtml = `
          <div class="bg-slate-900 border-b border-teal-800/60 px-3 py-1 flex items-center justify-between gap-2 shrink-0 text-[10.5px]">
            <button onclick="toggleSpectatorMyCheckCollapsed()" class="flex items-center gap-1.5 text-amber-300 font-extrabold min-w-0">
              <span>★ マイチェック（${checkedEntries.length}組）</span>
              ${urgentCnt > 0 ? `<span class="px-1.5 py-0.2 rounded-full bg-rose-600 text-white text-[9.5px] font-extrabold">${urgentCnt}組まもなく/試合中</span>` : ''}
            </button>
            <div class="flex items-center gap-2 shrink-0">
              <button onclick="${searchTabCmd}" class="text-teal-300 hover:text-white font-bold text-[10px]">編集</button>
              <button onclick="toggleSpectatorMyCheckCollapsed()" class="px-1.5 py-0.5 rounded bg-slate-800 text-slate-200 font-bold text-[10px]">表示 ▾</button>
            </div>
          </div>
        `;
      } else {
        myCheckBarHtml = `
          <div class="bg-slate-900 border-b border-teal-800/60 px-2.5 py-1.5 flex items-center gap-1.5 shrink-0">
            <button onclick="toggleSpectatorMyCheckCollapsed()"
                    title="マイチェックをたたむ"
                    class="shrink-0 px-1.5 py-1 rounded-lg bg-amber-400/15 border border-amber-400/40 text-amber-300 text-[10px] font-extrabold flex items-center gap-0.5">
              <span>★${checkedEntries.length}</span>
              <span class="text-[9px] opacity-80">▴</span>
            </button>
            <div class="flex-1 min-w-0 flex items-center gap-1.5 overflow-x-auto no-scrollbar">
              ${checkedEntries.map((e) => {
                const st = metrics.getEntryStatus(e);
                const tapAction = st.targetCourt
                  ? `focusSpectatorCourt('${escapeJs(st.targetCourt)}')`
                  : searchTabCmd;
                const shortName = OopEngine.formatSpectatorPairName(e, nameMode, null);
                return `
                  <button onclick="${tapAction}"
                          class="shrink-0 inline-flex items-center gap-1.5 px-2 py-1 rounded-full bg-slate-800 hover:bg-slate-700 border border-slate-700 transition">
                    <span class="text-[11px] font-extrabold text-white max-w-[108px] truncate">${escapeHtml(shortName)}</span>
                    <span class="px-1.5 py-0.2 rounded-full text-[9.5px] font-extrabold shrink-0 ${st.badge.cls}">${escapeHtml(st.badge.text)}</span>
                  </button>
                `;
              }).join('')}
            </div>
            <button onclick="${searchTabCmd}"
                    title="マイチェックを追加・変更"
                    class="shrink-0 px-1.5 py-1 rounded-lg bg-slate-800 hover:bg-slate-700 text-teal-300 text-[10px] font-extrabold">＋</button>
          </div>
        `;
      }
    } else {
      myCheckBarHtml = `
        <button onclick="${searchTabCmd}"
                class="w-full text-left bg-teal-950/95 hover:bg-teal-900/90 text-teal-100 border-b border-teal-800/70 px-3 py-1 text-[10.5px] flex items-center justify-between gap-2 shrink-0 transition">
          <span class="truncate"><b class="text-amber-300 mr-1">☆ マイチェック:</b>選手に★を付けるとここに待ち順を固定表示</span>
          <span class="font-bold text-teal-300 shrink-0">登録 ▸</span>
        </button>
      `;
    }

    const tab = opsUi.spectator.tab;
    let bodyHtml = '';
    if (tab === 'courts') {
      bodyHtml = renderSpectatorCourtsTab(catalog, oop, pairFmt, checkedEntryIdSet);
    } else if (tab === 'bracket') {
      bodyHtml = renderSpectatorBracketTab(activeEvent.id, evData, catalog, nameMode, shortSchool, checkedEntryIdSet);
    } else {
      bodyHtml = renderSpectatorSearchTab(activeEvent, evData, catalog, oop, nameMode, shortSchool, pairFmt, metrics);
    }

    const evSelectFn = isStandalone ? 'selectStandalonePublicEvent' : 'selectSpectatorEvent';
    const tabSelectFn = isStandalone ? 'setStandaloneSpectatorTab' : 'setSpectatorTab';

    frame.innerHTML = `
      <!-- BracketFlow オリジナル ヘッダー（ディープティール・進行ゲージ付） -->
      <div class="${isStandalone ? 'spectator-standalone-header ' : ''}bg-gradient-to-br from-slate-950 via-slate-900 to-teal-950 text-white px-3 pt-2 pb-2 space-y-1.5 shrink-0">
        <div class="flex items-center justify-between gap-2">
          <div class="flex items-center gap-2 min-w-0">
            <span class="w-6 h-6 rounded-lg bg-gradient-to-br from-emerald-400 to-teal-500 text-slate-950 flex items-center justify-center shrink-0 shadow-sm">
              <svg class="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round">
                <path d="M4 6h5v5H4" /><path d="M4 18h5v-5H4" /><path d="M9 8.5h5v7H9" /><path d="M14 12h6" />
              </svg>
            </span>
            <div class="min-w-0">
              <div class="flex items-center gap-1.5 text-[9.5px] font-bold text-emerald-300 leading-none">
                <span class="w-1.5 h-1.5 rounded-full bg-emerald-400 live-dot"></span>
                <span>LIVE</span>
                <span class="text-slate-300 font-semibold">${escapeHtml(updatedClock)} 更新</span>
                ${nameMode === 'surname' ? `<span class="px-1 rounded bg-slate-800/90 text-slate-300 border border-slate-700">名字表示</span>` : ''}
              </div>
              <h2 class="text-sm font-extrabold leading-tight truncate mt-0.5">${escapeHtml(tournamentName)}</h2>
            </div>
          </div>
          <span class="text-[10px] font-bold px-2 py-0.5 rounded-lg bg-white/10 text-teal-100 border border-white/15 shrink-0">${escapeHtml(oop.venue || activeEvent.venue || '第1会場')}</span>
        </div>

        <!-- 種目切替ピル ＆ 進行ゲージ（コンパクト配置） -->
        <div class="flex items-center gap-2">
          <div class="flex-1 min-w-0 flex items-center gap-1 overflow-x-auto no-scrollbar">
            ${(eventsList || []).map((item) => {
              const act = item.id === activeEvent.id;
              return `<button onclick="${evSelectFn}('${escapeJs(item.id)}')"
                              class="px-2.5 py-0.5 text-[10.5px] font-bold rounded-full shrink-0 transition ${act ? 'bg-emerald-400 text-slate-950 shadow-sm' : 'bg-white/10 text-slate-200 hover:bg-white/20'}">
                        ${escapeHtml(item.name)}
                      </button>`;
            }).join('')}
          </div>
          <div class="shrink-0 flex items-center gap-1.5 text-[10px] font-bold text-slate-300">
            <span>進行 ${finishedCnt}/${totalMatches}</span>
            <span class="text-emerald-300">${progressPct}%</span>
          </div>
        </div>
        <div class="w-full h-1 rounded-full bg-white/10 overflow-hidden">
          <div class="h-full rounded-full bg-gradient-to-r from-emerald-400 to-teal-300 transition-all duration-300" style="width:${progressPct}%"></div>
        </div>
      </div>

      <!-- 運営者アナウンスバナー（高視認性） -->
      ${renderAnnouncementBannersHtml(tourAnnouncement, oop.announcement, oop.venue || activeEvent.venue)}

      <!-- マイチェック（省スペース1行ライブバー） -->
      ${myCheckBarHtml}

      <!-- タブ本体 -->
      <div class="flex-1 min-h-0 overflow-y-auto overscroll-contain p-3 pb-6 space-y-3 bg-slate-100/90">
        ${bodyHtml}
      </div>

      <!-- スマホ親指最適化 ボトムドック3タブナビ（iOSセーフエリア対応） -->
      <nav class="spectator-bottom-nav bg-white border-t border-slate-200 grid grid-cols-3 shrink-0 shadow-[0_-4px_12px_rgba(15,23,42,0.08)] z-20">
        <button onclick="${tabSelectFn}('courts')"
                class="pt-2 pb-1.5 flex flex-col items-center justify-center gap-0.5 transition ${tab === 'courts' ? 'text-teal-700 bg-teal-50/70' : 'text-slate-500 hover:text-slate-800'}">
          <div class="relative flex items-center justify-center">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="16" rx="2"/><line x1="12" y1="4" x2="12" y2="20"/><line x1="3" y1="12" x2="21" y2="12"/></svg>
            ${liveCourtCnt > 0 ? `<span class="ml-1 px-1 rounded-full bg-rose-600 text-white text-[9px] font-extrabold leading-tight">${liveCourtCnt}</span>` : ''}
          </div>
          <span class="text-[10px] font-extrabold">コート状況</span>
        </button>
        <button onclick="${tabSelectFn}('bracket')"
                class="pt-2 pb-1.5 flex flex-col items-center justify-center gap-0.5 transition ${tab === 'bracket' ? 'text-teal-700 bg-teal-50/70' : 'text-slate-500 hover:text-slate-800'}">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M4 6h5v5H4M4 18h5v-5H4M9 8.5h5v7H9M14 12h6"/></svg>
          <span class="text-[10px] font-extrabold">トーナメント</span>
        </button>
        <button onclick="${tabSelectFn}('search')"
                class="pt-2 pb-1.5 flex flex-col items-center justify-center gap-0.5 transition ${tab === 'search' ? 'text-teal-700 bg-teal-50/70' : 'text-slate-500 hover:text-slate-800'}">
          <div class="relative flex items-center justify-center">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path stroke-linecap="round" d="M20 20l-3.5-3.5"/></svg>
            ${checkedEntries.length > 0 ? `<span class="ml-1 px-1 rounded-full bg-amber-400 text-slate-950 text-[9px] font-extrabold leading-tight">★${checkedEntries.length}</span>` : ''}
          </div>
          <span class="text-[10px] font-extrabold">選手・学校</span>
        </button>
      </nav>
    `;
  }

  function renderSpectatorView() {
    const frame = document.getElementById('spectator-portal-frame');
    if (!frame || !opsUi.bundle) return;

    const pub = opsUi.bundle.publish || { enabled: true, passcode: '', nameDisplayMode: 'full' };
    const summaryEl = document.getElementById('spectator-privacy-summary');
    if (summaryEl) {
      summaryEl.textContent = `公開: ${pub.enabled ? 'ON' : 'OFF'} / パスコード: ${pub.passcode ? 'あり' : 'なし'} / 氏名: ${pub.nameDisplayMode === 'surname' ? '名字のみ' : 'フルネーム'}`;
    }

    // 非公開ゲート
    if (!pub.enabled && !opsUi.spectator.organizerBypassClosed) {
      frame.innerHTML = `
        <div class="flex-1 flex flex-col items-center justify-center p-8 text-center space-y-4 bg-white">
          <div class="w-12 h-12 rounded-2xl bg-slate-100 text-slate-500 flex items-center justify-center text-xl">🔒</div>
          <div class="space-y-1">
            <h3 class="text-base font-extrabold text-slate-900">${escapeHtml(opsUi.bundle.name)}</h3>
            <p class="text-xs text-slate-500">現在、この大会の公開ページは非公開に設定されています。</p>
          </div>
          <button onclick="bypassSpectatorGate('closed')"
                  class="px-4 py-2 text-xs font-bold bg-slate-900 text-white rounded-lg">管理者プレビューで中を見る</button>
        </div>
      `;
      return;
    }

    // パスコードゲート
    if (pub.passcode && opsUi.spectator.unlockedPasscode !== pub.passcode) {
      frame.innerHTML = `
        <div class="flex-1 flex flex-col items-center justify-center p-8 text-center space-y-4 bg-white">
          <div class="w-12 h-12 rounded-2xl bg-emerald-50 text-emerald-700 border border-emerald-200 flex items-center justify-center text-xl">🔑</div>
          <div class="space-y-1">
            <div class="text-[10px] font-bold text-emerald-700">生徒・保護者向け 大会ポータル</div>
            <h3 class="text-base font-extrabold text-slate-900">${escapeHtml(opsUi.bundle.name)}</h3>
            <p class="text-xs text-slate-500">閲覧パスコード（合言葉）を入力してください</p>
          </div>
          <div class="w-full max-w-xs space-y-2">
            <input type="password" id="spec-passcode-try" placeholder="パスコードを入力"
                   onkeydown="if(event.key==='Enter') submitSpectatorPasscode()"
                   class="w-full px-3 py-2 text-sm font-bold text-center border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-500" />
            <button onclick="submitSpectatorPasscode()" class="w-full py-2 text-xs font-bold bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg">閲覧する</button>
            <button onclick="bypassSpectatorGate('passcode')" class="w-full py-1.5 text-[11px] font-bold text-slate-500 hover:text-slate-800">管理者プレビューで解除</button>
          </div>
        </div>
      `;
      return;
    }

    const ev = getSpectatorEvent();
    if (!ev || !ev.data) return;
    renderUnifiedSpectatorPortal(frame, {
      tournamentName: opsUi.bundle.name,
      eventsList: opsUi.bundle.events,
      activeEvent: ev,
      evData: ev.data,
      nameMode: pub.nameDisplayMode || 'full',
      tourAnnouncement: opsUi.bundle.announcement,
      isStandalone: false
    });
  }

  function renderSpectatorCourtsTab(catalog, oop, pairFmt, checkedEntryIdSet) {
    const nCourts = oop.courtCount || 4;
    const checkedIds = checkedEntryIdSet || new Set();
    let playing = 0;
    let calling = 0;
    let finished = 0;
    catalog.matches.forEach((m) => {
      const st = oop.statuses[m.key] || 'waiting';
      if (st === 'playing') playing++;
      else if (st === 'calling') calling++;
      else if (st === 'finished' || st === 'wo') finished++;
    });

    const isMatchChecked = (m) => Boolean(
      m && ((m.p1Entry && checkedIds.has(m.p1Entry.id)) || (m.p2Entry && checkedIds.has(m.p2Entry.id)))
    );

    const filter = opsUi.spectator.courtFilter || 'ALL';
    let cardsHtml = '';
    let visibleCourtCount = 0;

    for (let c = 1; c <= nCourts; c++) {
      const cKey = String(c);
      const courtKeys = oop.courts[cKey] || [];
      const realCourtTotal = courtKeys.filter((k) => !OopEngine.isBlankKey(k)).length;
      const live = OopEngine.getCourtLiveSlots(courtKeys, catalog.matchMap, oop.statuses);
      const hasCheckedInCourt = courtKeys.some((k) => !OopEngine.isBlankKey(k) && isMatchChecked(catalog.matchMap[k]) && !((catalog.matchMap[k] && catalog.matchMap[k].isFinished)));
      const isActiveCourt = Boolean(live.current);

      if (filter === 'ACTIVE' && !isActiveCourt) continue;
      if (filter === 'CHECKED' && !hasCheckedInCourt) continue;
      visibleCourtCount++;

      let statusBadge = `<span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 text-slate-600">待機中</span>`;
      let cardBorder = hasCheckedInCourt ? 'border-amber-300 ring-1 ring-amber-200' : 'border-slate-200';
      if (live.current && live.current.status === 'playing') {
        statusBadge = `<span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-rose-600 text-white flex items-center gap-1"><span class="w-1.5 h-1.5 rounded-full bg-white live-dot"></span>試合中</span>`;
        if (!hasCheckedInCourt) cardBorder = 'border-rose-200';
      } else if (live.current && live.current.status === 'calling') {
        statusBadge = `<span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-500 text-white">コール中</span>`;
      } else if (live.leadingBlankKey) {
        const bLbl = (oop.blankLabels && oop.blankLabels[live.leadingBlankKey]) || '進行調整中';
        statusBadge = `<span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-100 text-amber-900">${escapeHtml(bLbl)}</span>`;
      } else if (live.finished.length === realCourtTotal && realCourtTotal > 0) {
        statusBadge = `<span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-100 text-emerald-800">全試合終了</span>`;
      }

      const renderLiveBox = (slotItem, title, tone) => {
        if (!slotItem) return '';
        const m = slotItem.match;
        const ump = OopEngine.getUmpireLabel(courtKeys, slotItem.indexInCourt, catalog.matchMap, oop, pairFmt);
        const star1 = m.p1Entry && checkedIds.has(m.p1Entry.id) ? '<span class="text-amber-500 mr-0.5">★</span>' : '';
        const star2 = m.p2Entry && checkedIds.has(m.p2Entry.id) ? '<span class="text-amber-500 mr-0.5">★</span>' : '';
        return `
          <div class="p-2.5 rounded-xl border ${tone} space-y-1">
            <div class="flex items-center justify-between text-[10px] font-bold">
              <span>${title} ・ ${escapeHtml(m.matchNo)}（${escapeHtml(m.roundLabel)}）</span>
            </div>
            <div class="text-xs font-extrabold text-slate-900 leading-snug">${star1}${escapeHtml(m.p1Label)}</div>
            <div class="text-[10px] font-bold text-slate-400">vs</div>
            <div class="text-xs font-extrabold text-slate-900 leading-snug">${star2}${escapeHtml(m.p2Label)}</div>
            ${ump.text ? `<div class="text-[10px] text-slate-600 pt-1 border-t border-slate-200/70">審判: ${escapeHtml(ump.text)}</div>` : ''}
          </div>
        `;
      };

      const renderCompactRow = (slotItem, tag) => {
        if (!slotItem) return '';
        const m = slotItem.match;
        const checked = isMatchChecked(m);
        return `
          <div class="flex items-start gap-2 text-[11px] py-1.5 border-t border-slate-100 ${checked ? 'bg-amber-50/70 -mx-1 px-1 rounded' : ''}">
            <span class="px-1.5 py-0.5 rounded ${tag === '次試合' ? 'bg-blue-100 text-blue-800' : 'bg-slate-100 text-slate-600'} font-bold text-[9px] shrink-0 mt-0.5">${tag}</span>
            <div class="min-w-0 flex-1 leading-snug">
              ${checked ? '<span class="text-amber-500 font-bold mr-0.5">★</span>' : ''}
              <span class="font-bold text-slate-500 mr-1">${escapeHtml(m.matchNo)}</span>
              <span class="font-bold text-slate-800">${escapeHtml(m.p1Label)}</span>
              <span class="text-slate-400 mx-1">vs</span>
              <span class="font-bold text-slate-800">${escapeHtml(m.p2Label)}</span>
            </div>
          </div>
        `;
      };

      const isExpanded = Boolean(opsUi.spectator.expandedCourts[cKey]);
      const fullScheduleHtml = isExpanded ? `
        <div class="pt-2 border-t border-slate-200 space-y-1">
          <div class="text-[10px] font-extrabold text-slate-500">コート ${c} 全予定（第1〜第${courtKeys.length}試合）</div>
          ${courtKeys.map((k, idx) => {
            if (OopEngine.isBlankKey(k)) {
              const bLbl = (oop.blankLabels && oop.blankLabels[k]) || '空き（進行調整）';
              return `<div class="flex items-center justify-between gap-2 px-2 py-1 rounded bg-amber-50/70 border border-dashed border-amber-300 text-[10px] text-amber-900 font-bold">
                <span>第${idx + 1}試合</span>
                <span>${escapeHtml(bLbl)}</span>
              </div>`;
            }
            const m = catalog.matchMap[k];
            if (!m) return '';
            const st = oop.statuses[k] || 'waiting';
            const checked = isMatchChecked(m);
            const stBadge = m.result
              ? `<span class="px-1.5 py-0.2 rounded bg-emerald-100 text-emerald-800 font-extrabold">${escapeHtml(m.result.g1)}-${escapeHtml(m.result.g2)}</span>`
              : (st === 'playing'
                ? `<span class="px-1.5 py-0.2 rounded bg-rose-600 text-white font-bold">試合中</span>`
                : (st === 'calling'
                  ? `<span class="px-1.5 py-0.2 rounded bg-amber-500 text-white font-bold">コール</span>`
                  : `<span class="px-1.5 py-0.2 rounded bg-slate-100 text-slate-500 font-bold">予定</span>`));
            return `<div class="flex items-center justify-between gap-2 px-2 py-1 rounded border text-[10px] ${checked ? 'bg-amber-50 border-amber-300' : (m.isFinished ? 'bg-slate-50 border-slate-200 text-slate-500' : 'bg-white border-slate-200')}">
              <div class="min-w-0 flex-1 truncate">
                <span class="font-bold text-slate-400 mr-1">第${idx + 1}</span>
                ${checked ? '<span class="text-amber-500 font-bold">★</span>' : ''}
                <span class="font-bold text-slate-700">${escapeHtml(m.matchNo)}</span>
                <span class="mx-1 text-slate-800 font-semibold">${escapeHtml(m.p1Label)} vs ${escapeHtml(m.p2Label)}</span>
              </div>
              <div class="shrink-0">${stBadge}</div>
            </div>`;
          }).join('')}
        </div>
      ` : '';

      cardsHtml += `
        <div data-spec-court="${cKey}" class="bg-white rounded-2xl border ${cardBorder} shadow-sm p-3 space-y-2">
          <div class="flex items-center justify-between">
            <div class="flex items-center gap-1.5">
              <span class="text-sm font-extrabold text-slate-900">コート ${c}</span>
              ${hasCheckedInCourt ? `<span class="px-1.5 py-0.2 rounded bg-amber-100 text-amber-900 text-[9.5px] font-extrabold">★チェック関連</span>` : ''}
            </div>
            <div class="flex items-center gap-2">
              <span class="text-[10px] text-slate-400 font-bold">${live.finished.length}/${realCourtTotal}完了</span>
              ${statusBadge}
            </div>
          </div>
          ${live.current
            ? renderLiveBox(live.current, live.current.status === 'playing' ? '現在 試合中' : 'コートへ集合（コール中）', live.current.status === 'playing' ? 'bg-rose-50/80 border-rose-300 text-rose-800' : 'bg-amber-50/80 border-amber-300 text-amber-800')
            : (live.next ? renderLiveBox(live.next, '次の試合予定', 'bg-sky-50/70 border-sky-200 text-sky-900') : `<div class="text-xs text-slate-400 py-2 text-center">予定されている試合はありません</div>`)}
          ${live.current ? renderCompactRow(live.next, '次試合') : ''}
          ${renderCompactRow(live.current ? live.onDeck : live.upcoming[1], '次々試合')}
          ${courtKeys.length > 0 ? `
            <button onclick="toggleSpectatorCourtExpand('${cKey}')"
                    class="w-full pt-1.5 border-t border-slate-100 text-[10.5px] font-bold text-teal-700 hover:text-teal-900 flex items-center justify-center gap-1">
              <span>${isExpanded ? '全予定を閉じる ▴' : `このコートの全予定を見る（${courtKeys.length}枠） ▾`}</span>
            </button>
          ` : ''}
          ${fullScheduleHtml}
        </div>
      `;
    }

    const filterBtnCls = (k) => filter === k
      ? 'px-2.5 py-1 rounded-full text-[11px] font-bold bg-slate-900 text-white shadow-sm transition'
      : 'px-2.5 py-1 rounded-full text-[11px] font-bold bg-white border border-slate-200 text-slate-600 hover:bg-slate-50 transition';

    return `
      <div class="grid grid-cols-3 gap-2 text-center">
        <div class="bg-white rounded-xl border border-slate-200 p-2 shadow-sm">
          <div class="text-[10px] font-bold text-slate-500">試合中</div>
          <div class="text-base font-extrabold text-rose-600">${playing}<span class="text-[10px] text-slate-400 ml-0.5">面</span></div>
        </div>
        <div class="bg-white rounded-xl border border-slate-200 p-2 shadow-sm">
          <div class="text-[10px] font-bold text-slate-500">コール中</div>
          <div class="text-base font-extrabold text-amber-600">${calling}<span class="text-[10px] text-slate-400 ml-0.5">面</span></div>
        </div>
        <div class="bg-white rounded-xl border border-slate-200 p-2 shadow-sm">
          <div class="text-[10px] font-bold text-slate-500">終了</div>
          <div class="text-base font-extrabold text-teal-700">${finished}<span class="text-[10px] text-slate-400 ml-0.5">/${catalog.matches.length}</span></div>
        </div>
      </div>

      <!-- ③ コート絞り込みフィルタ -->
      <div class="flex items-center gap-1.5 overflow-x-auto pb-0.5">
        <button onclick="setSpectatorCourtFilter('ALL')" class="${filterBtnCls('ALL')}">すべて (${nCourts}面)</button>
        <button onclick="setSpectatorCourtFilter('ACTIVE')" class="${filterBtnCls('ACTIVE')}">進行中のみ (${playing + calling})</button>
        <button onclick="setSpectatorCourtFilter('CHECKED')" class="${filterBtnCls('CHECKED')}">★チェック関連</button>
      </div>

      ${visibleCourtCount > 0 ? `
        <div class="grid grid-cols-1 ${opsUi.spectator.frameMode === 'full' ? 'md:grid-cols-2 lg:grid-cols-3' : ''} gap-2.5">
          ${cardsHtml}
        </div>
      ` : `
        <div class="bg-white rounded-2xl border border-slate-200 p-6 text-center text-xs text-slate-500 space-y-2">
          <div>${filter === 'CHECKED' ? '★チェック中のペアが予定されているコートがありません' : '現在、試合中・コール中のコートはありません'}</div>
          <button onclick="setSpectatorCourtFilter('ALL')" class="px-3 py-1 rounded-lg bg-slate-900 text-white text-[11px] font-bold">全コートを表示</button>
        </div>
      `}
    `;
  }

  /**
   * ② トーナメント表タブ：ブロック分割（A/B/C/D）＋ ベスト32/16/8 絞り込み表示対応
   */
  function renderSpectatorBracketTab(evId, evData, catalog, nameMode, shortSchool, checkedEntryIdSet) {
    const size = evData.bracketSize || 16;
    const totalRounds = Math.round(Math.log2(size));
    const tree = catalog.tree;
    const checkedIds = checkedEntryIdSet || new Set();

    let stage = opsUi.spectator.bracketStage || 'ALL';
    let blk = opsUi.spectator.bracketBlock || 'ALL';
    if (stage !== 'ALL' && parseInt(stage, 10) >= size) stage = 'ALL';
    if (blk !== 'ALL' && size < 8) blk = 'ALL';

    const labelW = 212;
    let blackLines = '';
    let redLines = '';
    let scoresSvg = '';
    let labelsHtml = '';
    const pushLine = (x1, y1, x2, y2, isRed, w) => {
      const line = `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${isRed ? '#dc2626' : '#334155'}" stroke-width="${w || (isRed ? 2.4 : 1.3)}" stroke-linecap="square" />`;
      if (isRed) redLines += line;
      else blackLines += line;
    };
    const reachedByWin = (node) => {
      if (!node || !node.active) return false;
      if (node.round === undefined) return false;
      if (node.isBye) return reachedByWin(node.c0.active ? node.c0 : node.c1);
      return Boolean(node.result);
    };

    let totalW = 480;
    let totalH = 300;
    let viewHeading = '全回戦';

    if (stage !== 'ALL') {
      // ベスト32 / ベスト16 / ベスト8 絞り込み表示
      const targetSize = parseInt(stage, 10);
      const subRounds = Math.round(Math.log2(targetSize));
      const baseRound = totalRounds - subRounds; // tree[baseRound] に targetSize 個のノードが並ぶ
      const rowH = targetSize >= 32 ? 28 : 34;
      const stepW = Math.max(32, Math.min(48, Math.floor(190 / Math.max(1, subRounds))));
      totalW = labelW + subRounds * stepW + 140;
      totalH = targetSize * rowH + 28;
      viewHeading = `ベスト${targetSize}〜決勝`;

      const subY = [];
      subY[0] = [];
      for (let i = 0; i < targetSize; i++) subY[0][i] = i * rowH + rowH / 2 + 12;
      for (let sr = 1; sr <= subRounds; sr++) {
        const cnt = targetSize / Math.pow(2, sr);
        subY[sr] = [];
        for (let m = 0; m < cnt; m++) {
          subY[sr][m] = (subY[sr - 1][m * 2] + subY[sr - 1][m * 2 + 1]) / 2;
        }
      }

      for (let i = 0; i < targetSize; i++) {
        const baseNode = tree[baseRound][i];
        const y = subY[0][i];
        const entry = baseNode ? baseNode.entry : null;
        const isFav = Boolean(entry && checkedIds.has(entry.id));
        let labelStr = '未定';
        if (entry) {
          labelStr = OopEngine.formatSpectatorPairName(entry, nameMode, shortSchool);
        } else if (baseNode && baseNode.sourceMatchKey && catalog.matchMap[baseNode.sourceMatchKey]) {
          labelStr = `${catalog.matchMap[baseNode.sourceMatchKey].matchNo}の勝者`;
        }
        labelsHtml += `<div ${isFav ? 'data-spec-checked="1"' : ''}
                            style="left:4px; top:${y - 11}px; width:${labelW - 10}px; height:22px;"
                            class="absolute flex items-center gap-1.5 text-[11px] leading-snug rounded px-1 ${isFav ? 'bg-amber-200/90 ring-1 ring-amber-400' : ''}">
          <span class="w-5 text-right font-bold ${isFav ? 'text-amber-900' : 'text-slate-400'} shrink-0">${i + 1}</span>
          <span class="${entry ? 'font-bold text-slate-900' : 'font-medium text-slate-400 italic'} truncate">${isFav ? '★' : ''}${escapeHtml(labelStr)}</span>
        </div>`;
      }

      for (let sr = 1; sr <= subRounds; sr++) {
        const r = baseRound + sr;
        const cnt = targetSize / Math.pow(2, sr);
        const xCurr = labelW + sr * stepW;
        const xPrev = labelW + (sr - 1) * stepW;
        for (let m = 0; m < cnt; m++) {
          const node = tree[r][m];
          if (!node || !node.active) continue;
          const y0 = subY[sr - 1][m * 2];
          const y1 = subY[sr - 1][m * 2 + 1];
          const yMid = subY[sr][m];
          const res = node.result;
          const w0 = Boolean(res && res.winner === 0);
          const w1 = Boolean(res && res.winner === 1);
          pushLine(xPrev, y0, xCurr, y0, w0 || (sr > 1 && reachedByWin(node.c0)));
          pushLine(xCurr, y0, xCurr, yMid, w0);
          pushLine(xPrev, y1, xCurr, y1, w1 || (sr > 1 && reachedByWin(node.c1)));
          pushLine(xCurr, yMid, xCurr, y1, w1);
          if (res) {
            scoresSvg += `<text x="${xCurr - 7}" y="${y0 - 3}" text-anchor="middle" fill="#dc2626" font-size="10" font-weight="bold">${escapeHtml(res.g1)}</text>`;
            scoresSvg += `<text x="${xCurr - 7}" y="${y1 + 10}" text-anchor="middle" fill="#dc2626" font-size="10" font-weight="bold">${escapeHtml(res.g2)}</text>`;
          }
        }
      }

      const finalNode = tree[totalRounds][0];
      const finalX = labelW + subRounds * stepW;
      const finalY = subY[subRounds][0];
      pushLine(finalX, finalY, finalX + 20, finalY, Boolean(finalNode && finalNode.result));
      if (finalNode && finalNode.entry) {
        labelsHtml += `<div style="left:${finalX + 24}px; top:${finalY - 10}px;" class="absolute text-[11px] font-extrabold text-rose-700 bg-rose-50 border border-rose-200 px-2 py-0.5 rounded whitespace-nowrap">${escapeHtml(OopEngine.formatSpectatorPairName(finalNode.entry, nameMode, shortSchool))}</div>`;
      }
    } else {
      // 全体 または A/B/C/D ブロック表示
      let startSlot = 0;
      let endSlot = size;
      let maxRound = totalRounds;
      if (blk !== 'ALL' && size >= 8) {
        const qMap = { A: 0, D: 1, C: 2, B: 3 };
        const qIdx = qMap[blk] ?? 0;
        const qSize = size / 4;
        startSlot = qIdx * qSize;
        endSlot = startSlot + qSize;
        maxRound = Math.max(1, totalRounds - 2);
        viewHeading = `${blk}ブロック`;
      }

      const activeIndices = [];
      for (let i = startSlot; i < endSlot; i++) {
        if (tree[0][i] && tree[0][i].active) activeIndices.push(i);
      }
      const slotToDisplayNo = {};
      let dispCnt = 1;
      for (let i = 0; i < size; i++) {
        if (tree[0][i] && tree[0][i].active) slotToDisplayNo[i] = dispCnt++;
      }

      const rowH = activeIndices.length >= 48 ? 26 : 30;
      const stepW = Math.max(24, Math.min(40, Math.floor(180 / Math.max(1, maxRound))));
      totalW = labelW + maxRound * stepW + 135;
      totalH = Math.max(180, activeIndices.length * rowH + 28);

      const nodeY = { 0: {} };
      activeIndices.forEach((slotIdx, vRow) => {
        nodeY[0][slotIdx] = vRow * rowH + rowH / 2 + 12;
      });
      for (let r = 1; r <= maxRound; r++) {
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
      const effX = (node, round) => {
        if (round === 0) return labelW;
        if (node.isBye) return effX(node.c0.active ? node.c0 : node.c1, round - 1);
        return labelW + round * stepW;
      };

      activeIndices.forEach((slotIdx) => {
        const leaf = tree[0][slotIdx];
        const y = nodeY[0][slotIdx];
        const nameStr = leaf.entry ? OopEngine.formatSpectatorPairName(leaf.entry, nameMode, shortSchool) : '空き';
        const isFav = Boolean(leaf.entry && checkedIds.has(leaf.entry.id));
        labelsHtml += `<div ${isFav ? 'data-spec-checked="1"' : ''}
                            style="left:4px; top:${y - 11}px; width:${labelW - 10}px; height:22px;"
                            class="absolute flex items-center gap-1.5 text-[11px] leading-snug rounded px-1 ${isFav ? 'bg-amber-200/90 ring-1 ring-amber-400' : ''}">
          <span class="w-5 text-right font-bold ${isFav ? 'text-amber-900' : 'text-slate-400'} shrink-0">${slotToDisplayNo[slotIdx]}</span>
          <span class="font-bold text-slate-900 truncate">${isFav ? '★' : ''}${escapeHtml(nameStr)}</span>
        </div>`;
      });

      for (let r = 1; r <= maxRound; r++) {
        const xCurr = labelW + r * stepW;
        const mStart = startSlot / Math.pow(2, r);
        const mEnd = endSlot / Math.pow(2, r);
        for (let m = mStart; m < mEnd; m++) {
          const node = tree[r][m];
          if (!node.active || node.isBye) continue;
          const y0 = nodeY[r - 1][m * 2];
          const y1 = nodeY[r - 1][m * 2 + 1];
          const yMid = nodeY[r][m];
          const x0 = effX(node.c0, r - 1);
          const x1 = effX(node.c1, r - 1);
          const res = node.result;
          const w0 = Boolean(res && res.winner === 0);
          const w1 = Boolean(res && res.winner === 1);
          pushLine(x0, y0, xCurr, y0, w0 || reachedByWin(node.c0));
          pushLine(xCurr, y0, xCurr, yMid, w0);
          pushLine(x1, y1, xCurr, y1, w1 || reachedByWin(node.c1));
          pushLine(xCurr, yMid, xCurr, y1, w1);
          if (res) {
            scoresSvg += `<text x="${xCurr - 7}" y="${y0 - 3}" text-anchor="middle" fill="#dc2626" font-size="10" font-weight="bold">${escapeHtml(res.g1)}</text>`;
            scoresSvg += `<text x="${xCurr - 7}" y="${y1 + 10}" text-anchor="middle" fill="#dc2626" font-size="10" font-weight="bold">${escapeHtml(res.g2)}</text>`;
          }
        }
      }

      const endM = startSlot / Math.pow(2, maxRound);
      const endNode = tree[maxRound][endM];
      const finalX = labelW + maxRound * stepW;
      const finalY = nodeY[maxRound][endM];
      if (endNode && endNode.active) {
        pushLine(finalX, finalY, finalX + 20, finalY, Boolean(endNode.result));
        if (endNode.entry && (endNode.result || endNode.isBye)) {
          labelsHtml += `<div style="left:${finalX + 24}px; top:${finalY - 10}px;" class="absolute text-[11px] font-extrabold text-rose-700 bg-rose-50 border border-rose-200 px-2 py-0.5 rounded whitespace-nowrap">${escapeHtml(OopEngine.formatSpectatorPairName(endNode.entry, nameMode, shortSchool))}</div>`;
        }
      }
    }

    const z = opsUi.spectator.bracketZoom;
    const stageOptions = [];
    if (size > 32) stageOptions.push('32');
    if (size > 16) stageOptions.push('16');
    if (size > 8) stageOptions.push('8');

    const pillCls = (act) => act
      ? 'px-2 py-0.5 rounded-md bg-teal-700 text-white text-[10.5px] font-bold shadow-sm'
      : 'px-2 py-0.5 rounded-md bg-slate-100 hover:bg-slate-200 text-slate-700 text-[10.5px] font-bold';

    return `
      <div class="bg-white rounded-2xl border border-slate-200 p-2.5 space-y-2 shadow-sm">
        <!-- ② ブロック分割 ＆ ベスト32/16/8 絞り込み -->
        <div class="flex flex-wrap items-center justify-between gap-1.5">
          <div class="flex flex-wrap items-center gap-1">
            <button onclick="setSpectatorBracketBlock('ALL'); setSpectatorBracketStage('ALL');" class="${pillCls(blk === 'ALL' && stage === 'ALL')}">全回戦</button>
            ${stageOptions.map((st) => `
              <button onclick="setSpectatorBracketStage('${st}')" class="${pillCls(stage === st)}">ベスト${st}</button>
            `).join('')}
          </div>
          ${size >= 8 ? `
            <div class="flex items-center gap-1">
              <span class="text-[10px] font-bold text-slate-400">山:</span>
              ${['A', 'B', 'C', 'D'].map((b) => `
                <button onclick="setSpectatorBracketBlock('${b}')" class="${pillCls(blk === b && stage === 'ALL')}">${b}</button>
              `).join('')}
            </div>` : ''}
        </div>

        <!-- ズーム・フィット・★ジャンプ -->
        <div class="flex flex-wrap items-center justify-between gap-2 pt-1.5 border-t border-slate-100">
          <span class="text-[10.5px] font-extrabold text-teal-800">${escapeHtml(viewHeading)}</span>
          <div class="flex items-center gap-1">
            ${checkedIds.size > 0 ? `<button onclick="jumpSpectatorToFirstChecked()" class="px-2 py-0.5 rounded-md bg-amber-100 hover:bg-amber-200 text-amber-900 text-[10px] font-bold">★位置へ</button>` : ''}
            <button onclick="fitSpectatorBracketZoom()" class="px-2 py-0.5 rounded-md bg-slate-100 hover:bg-slate-200 text-slate-700 text-[10px] font-bold">フィット</button>
            <button onclick="adjustSpectatorBracketZoom(-0.12)" class="w-6 h-6 rounded bg-slate-100 font-bold text-slate-700">-</button>
            <span class="text-[10.5px] font-bold text-slate-700 w-9 text-center">${Math.round(z * 100)}%</span>
            <button onclick="adjustSpectatorBracketZoom(0.12)" class="w-6 h-6 rounded bg-slate-100 font-bold text-slate-700">+</button>
          </div>
        </div>
      </div>

      <div id="spec-bracket-scroll" class="bg-white rounded-2xl border border-slate-200 p-3 overflow-auto shadow-sm" style="max-height:560px;">
        <div style="width:${Math.ceil(totalW * z)}px; height:${Math.ceil(totalH * z)}px; position:relative;">
          <div id="spec-bracket-inner" data-raw-width="${totalW}" style="transform:scale(${z}); transform-origin:top left; width:${totalW}px; height:${totalH}px; position:relative;">
            <svg width="${totalW}" height="${totalH}" class="absolute inset-0 pointer-events-none">${blackLines}${redLines}${scoresSvg}</svg>
            ${labelsHtml}
          </div>
        </div>
      </div>
    `;
  }

  function renderSpectatorSearchTab(ev, evData, catalog, oop, nameMode, shortSchool, pairFmt, metrics) {
    const entries = Array.isArray(evData.entries) ? evData.entries : [];
    const schools = Array.from(new Set(entries.map((e) => e.school).filter(Boolean)));
    const q = opsUi.spectator.searchQuery.trim().toLowerCase();
    const { matchCourtMap, getEntryStatus } = metrics || buildSpectatorEntryMetrics(ev.id, evData, catalog, oop);
    const makeFavKey = (entry) => `${ev.id}::${entry.id}`;

    const renderEntryCard = (entry) => {
      const favKey = makeFavKey(entry);
      const isFav = opsUi.spectator.favorites.has(favKey);
      const st = getEntryStatus(entry);
      const dispName = pairFmt(entry);
      return `
        <div class="bg-white rounded-xl border ${isFav ? 'border-amber-300 ring-1 ring-amber-200' : 'border-slate-200'} p-2.5 space-y-1.5 shadow-sm">
          <div class="flex items-start justify-between gap-2">
            <div class="min-w-0">
              <div class="text-xs font-extrabold text-slate-900 leading-snug">${escapeHtml(dispName)}</div>
              <div class="mt-1 inline-flex items-center gap-1.5 flex-wrap">
                <span class="px-2 py-0.5 rounded-full text-[10px] font-bold ${st.badge.cls}">${escapeHtml(st.badge.text)}</span>
                ${st.targetCourt ? `<button onclick="focusSpectatorCourt('${escapeJs(st.targetCourt)}')" class="text-[10px] font-bold text-teal-700 hover:underline">コート${st.targetCourt}を見る ▸</button>` : ''}
              </div>
            </div>
            <button onclick="toggleSpectatorFavorite('${escapeJs(favKey)}')"
                    class="px-2 py-1 rounded-lg text-sm leading-none transition ${isFav ? 'bg-amber-100 text-amber-500' : 'bg-slate-100 text-slate-400 hover:text-amber-500'}"
                    title="マイチェックに登録（画面上部に固定表示）">
              ${isFav ? '★' : '☆'}
            </button>
          </div>
          ${st.related.length > 0 ? `
            <div class="pt-1.5 border-t border-slate-100 space-y-1 text-[10px]">
              ${st.related.map((m) => {
                const loc = matchCourtMap[m.key];
                const isP1 = m.p1Entry && m.p1Entry.id === entry.id;
                const opp = isP1 ? m.p2Label : m.p1Label;
                const resTxt = m.result ? `${m.result.g1}-${m.result.g2}` : (oop.statuses[m.key] === 'playing' ? '試合中' : '予定');
                return `<div class="flex items-center justify-between gap-2 text-slate-600">
                  <span class="truncate">${escapeHtml(m.roundLabel)} (コート${loc ? loc.court : '-'}) vs ${escapeHtml(opp)}</span>
                  <span class="font-bold shrink-0 ${m.result ? 'text-rose-600' : 'text-slate-500'}">${escapeHtml(resTxt)}</span>
                </div>`;
              }).join('')}
            </div>` : ''}
        </div>
      `;
    };

    const favEntries = entries.filter((e) => opsUi.spectator.favorites.has(makeFavKey(e)));
    const filtered = entries.filter((e) => {
      if (opsUi.spectator.schoolFilter !== 'ALL' && e.school !== opsUi.spectator.schoolFilter) return false;
      if (!q) return true;
      const hay = `${pairFmt(e)} ${e.school || ''}`.toLowerCase();
      return hay.includes(q);
    });

    // オリジナル要素：学校選択時のリアルタイム戦績パルスバー
    let schoolPulseHtml = '';
    if (opsUi.spectator.schoolFilter !== 'ALL') {
      const scEntries = entries.filter((e) => e.school === opsUi.spectator.schoolFilter);
      let alive = 0;
      let scPlaying = 0;
      let scNext = 0;
      scEntries.forEach((e) => {
        const s = getEntryStatus(e);
        if (s.stateKind !== 'eliminated') alive++;
        if (s.stateKind === 'playing' || s.stateKind === 'calling') scPlaying++;
        if (s.stateKind === 'next') scNext++;
      });
      schoolPulseHtml = `
        <div class="rounded-xl bg-gradient-to-r from-teal-900 to-slate-900 text-white px-3 py-2 flex items-center justify-between gap-2 text-[11px]">
          <div class="font-extrabold truncate">${escapeHtml(shortSchool(opsUi.spectator.schoolFilter))} ライブ状況</div>
          <div class="flex items-center gap-2 shrink-0 text-[10px] font-bold">
            <span class="px-1.5 py-0.5 rounded bg-white/15">勝ち残り ${alive}/${scEntries.length}</span>
            ${scPlaying > 0 ? `<span class="px-1.5 py-0.5 rounded bg-rose-500">試合中 ${scPlaying}</span>` : ''}
            ${scNext > 0 ? `<span class="px-1.5 py-0.5 rounded bg-blue-500">次試合 ${scNext}</span>` : ''}
          </div>
        </div>
      `;
    }

    return `
      <div class="space-y-2">
        <input type="text" id="spec-search-input" value="${escapeHtml(opsUi.spectator.searchQuery)}"
               oninput="setSpectatorSearchQuery(this.value)"
               placeholder="選手名・学校名で検索..."
               class="w-full px-3 py-2 text-xs font-bold bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-teal-600 shadow-sm" />
        <div class="flex items-center gap-1.5 overflow-x-auto pb-1">
          <button onclick="setSpectatorSchoolFilter('ALL')"
                  class="px-2.5 py-1 rounded-full text-[11px] font-bold shrink-0 ${opsUi.spectator.schoolFilter === 'ALL' ? 'bg-slate-900 text-white' : 'bg-white border border-slate-200 text-slate-600'}">
            すべて (${entries.length})
          </button>
          ${schools.map((sc) => `
            <button onclick="setSpectatorSchoolFilter('${escapeJs(sc)}')"
                    class="px-2.5 py-1 rounded-full text-[11px] font-bold shrink-0 ${opsUi.spectator.schoolFilter === sc ? 'bg-teal-700 text-white' : 'bg-white border border-slate-200 text-slate-600'}">
              ${escapeHtml(shortSchool(sc))}
            </button>
          `).join('')}
        </div>
        ${schoolPulseHtml}
      </div>

      ${favEntries.length > 0 ? `
        <div class="space-y-1.5">
          <div class="text-[11px] font-extrabold text-amber-800 flex items-center gap-1">
            <span>★ マイチェック登録（${favEntries.length}組）</span>
          </div>
          ${favEntries.map(renderEntryCard).join('')}
        </div>` : ''}

      <div class="space-y-1.5">
        <div class="text-[11px] font-bold text-slate-500">出場ペア一覧（${filtered.length}組）</div>
        ${filtered.length > 0 ? filtered.map(renderEntryCard).join('') : `<div class="text-center py-6 text-xs text-slate-400">一致するペアが見つかりません</div>`}
      </div>
    `;
  }

  function jumpSpectatorBracketMountain(quarterIdx) {
    const box = document.getElementById('spec-bracket-scroll');
    if (!box) return;
    const maxY = Math.max(0, box.scrollHeight - box.clientHeight);
    box.scrollTo({ top: Math.round((box.scrollHeight * quarterIdx) / 4), behavior: 'smooth' });
    if (quarterIdx === 3 && maxY > 0) {
      box.scrollTo({ top: Math.min(maxY, Math.round((box.scrollHeight * 3) / 4)), behavior: 'smooth' });
    }
  }

  // =========================================================================
  // 5b. Firebase フェーズ2連携ブリッジ（マイ大会・別端末公開ポータル・会場スタッフモード）
  // =========================================================================
  function getCurrentTournamentBundle() {
    if (!opsUi.bundle) return null;
    const curEv = getActiveEvent();
    if (curEv) {
      curEv.data = captureActiveStateSnapshot();
      curEv.venue = (curEv.data.oop && curEv.data.oop.venue) || curEv.venue;
    }
    return cloneDeep(opsUi.bundle);
  }

  function replaceTournamentBundle(newBundle) {
    if (!newBundle || !Array.isArray(newBundle.events) || !newBundle.events.length) return;
    opsUi.bundle = cloneDeep(newBundle);
    if (!opsUi.bundle.publish) opsUi.bundle.publish = { enabled: true, passcode: '', nameDisplayMode: 'full' };
    opsUi.oopUndoStack = [];
    const activeEv = getActiveEvent();
    opsUi.suppressBundleSync = true;
    hydrateStateFromSnapshot(activeEv.data);
    opsUi.suppressBundleSync = false;
    try { localStorage.setItem(BUNDLE_STORAGE_KEY, JSON.stringify(opsUi.bundle)); } catch (_) {}
    renderEventSwitcherBar();
    updateOopHeaderBadge();
    updatePublishIndicator();
    if (state.slots.length > 0) switchMainView('tournament');
    else switchMainView('registration');
  }

  function createFreshTournamentBundle(newName) {
    const evId = randomId('ev');
    const fresh = {
      version: '6.0-bundle',
      id: randomId('tour'),
      name: newName || '新しい大会',
      publish: { enabled: true, passcode: '', nameDisplayMode: 'full' },
      activeEventId: evId,
      events: [
        {
          id: evId,
          name: '男子個人',
          venue: '第1会場',
          staffPin: generateStaffPin(),
          pinVersion: 1,
          data: createEmptyEventSnapshot('男子個人', '第1会場', 4)
        }
      ]
    };
    replaceTournamentBundle(fresh);
    switchMainView('registration');
    showToast(`新しい大会「${fresh.name}」を作成しました`, 'success');
  }

  function applyRemoteLivePatchToActiveEvent(patch) {
    if (!patch) return;
    opsUi.suppressBundleSync = true;
    try {
      if (patch.results) state.results = cloneDeep(patch.results);
      if (patch.consolation) state.consolation = cloneDeep(patch.consolation);
      if (patch.oop) state.oop = OopEngine.createDefaultOopState(patch.oop);
      const activeEv = getActiveEvent();
      if (activeEv) activeEv.data = captureActiveStateSnapshot();
      try { localStorage.setItem(BUNDLE_STORAGE_KEY, JSON.stringify(opsUi.bundle)); } catch (_) {}
      updateOopHeaderBadge();
      if (state.mainView === 'tournament') renderSheet();
      else if (state.mainView === 'oop') renderOopView();
      else if (state.mainView === 'spectator') renderSpectatorView();
    } finally {
      opsUi.suppressBundleSync = false;
    }
  }

  function hydrateStaffSingleEvent(payload) {
    if (!payload) return;
    opsUi.suppressBundleSync = true;
    try {
      const snap = Object.assign(captureActiveStateSnapshot(), {
        title: `${payload.tournamentName} ${payload.eventName}`,
        matchFormat: payload.matchFormat || 7,
        bracketSize: payload.bracketSize || 16,
        slots: payload.slots || [],
        entries: payload.entries || [],
        activeSlotMask: payload.activeSlotMask || [],
        schoolAliases: payload.schoolAliases || {},
        results: payload.results || {},
        consolation: payload.consolation || state.consolation,
        oop: OopEngine.createDefaultOopState(payload.oop)
      });
      hydrateStateFromSnapshot(snap);
      if (state.mainView !== 'oop' && state.mainView !== 'tournament') {
        switchMainView('oop');
      } else if (state.mainView === 'oop') {
        renderOopView();
      } else {
        renderSheet();
      }
    } finally {
      opsUi.suppressBundleSync = false;
    }
  }

  function renderStandalonePublicPortal(meta, liveEventObj, verifyPasscodeFn) {
    const prevUnlocked = opsUi.standalonePublic && opsUi.standalonePublic.meta && opsUi.standalonePublic.meta.id === meta.id
      ? opsUi.standalonePublic.unlocked
      : false;
    opsUi.standalonePublic = { meta, liveEvent: liveEventObj, verifyPasscodeFn, unlocked: prevUnlocked };
    opsUi.spectator.frameMode = 'full';
    renderStandalonePublicFrame();
  }

  function renderStandalonePublicFrame() {
    const ctx = opsUi.standalonePublic;
    const frame = document.getElementById('spectator-portal-frame');
    if (!ctx || !frame) return;
    const { meta, liveEvent } = ctx;

    if (meta.hasPasscode && !ctx.unlocked) {
      frame.innerHTML = `
        <div class="flex-1 flex flex-col items-center justify-center p-8 text-center space-y-4 bg-white">
          <div class="w-12 h-12 rounded-2xl bg-emerald-50 text-emerald-700 border border-emerald-200 flex items-center justify-center text-xl">🔑</div>
          <div class="space-y-1">
            <div class="text-[10px] font-bold text-emerald-700">生徒・保護者向け LIVE 大会ポータル</div>
            <h3 class="text-base font-extrabold text-slate-900">${escapeHtml(meta.name)}</h3>
            <p class="text-xs text-slate-500">大会本部から案内されたパスコード（合言葉）を入力してください</p>
          </div>
          <div class="w-full max-w-xs space-y-2">
            <input type="password" id="standalone-pass-input" placeholder="パスコードを入力"
                   onkeydown="if(event.key==='Enter') submitStandalonePasscode()"
                   class="w-full px-3 py-2 text-sm font-bold text-center border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-500" />
            <div id="standalone-pass-err" class="text-[11px] font-bold text-rose-600 min-h-[16px]"></div>
            <button onclick="submitStandalonePasscode()" class="w-full py-2.5 text-xs font-bold bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg shadow">閲覧する</button>
          </div>
        </div>
      `;
      return;
    }

    renderUnifiedSpectatorPortal(frame, {
      tournamentName: meta.name,
      eventsList: meta.events || [],
      activeEvent: liveEvent,
      evData: liveEvent.data,
      nameMode: meta.nameDisplayMode || 'full',
      tourAnnouncement: meta.announcement,
      isStandalone: true
    });
  }

  async function submitStandalonePasscode() {
    const ctx = opsUi.standalonePublic;
    if (!ctx || typeof ctx.verifyPasscodeFn !== 'function') return;
    const inp = document.getElementById('standalone-pass-input');
    const errEl = document.getElementById('standalone-pass-err');
    const ok = await ctx.verifyPasscodeFn(inp ? String(inp.value || '').trim() : '');
    if (ok) {
      ctx.unlocked = true;
      renderStandalonePublicFrame();
    } else if (errEl) {
      errEl.textContent = 'パスコードが一致しません。再度ご確認ください。';
    }
  }

  function selectStandalonePublicEvent(evId) {
    if (!opsUi.standalonePublic || !window.BracketCloud) return;
    opsUi.spectator.schoolFilter = 'ALL';
    window.BracketCloud.subscribePublicEvent(opsUi.standalonePublic.meta.id, evId);
  }

  function setStandaloneSpectatorTab(tab) {
    opsUi.spectator.tab = tab;
    if (opsUi.standalonePublic) renderStandalonePublicFrame();
    else renderSpectatorView();
  }

  // =========================================================================
  // 6. app.js からのライフサイクルフック接続 ＆ グローバル公開
  // =========================================================================
  window.onAppStateSaved = function () {
    if (opsUi.suppressBundleSync || !opsUi.bundle) return;
    const ev = getActiveEvent();
    if (ev) {
      ev.data = captureActiveStateSnapshot();
      ev.venue = (ev.data.oop && ev.data.oop.venue) || ev.venue;
      saveBundleToLocalStorage();
      renderEventSwitcherBar();
      updateOopHeaderBadge();
    }
  };

  window.onBracketResultsChanged = function () {
    syncOopWithBracket();
    saveToLocalStorage();
    updateOopHeaderBadge();
    if (state.mainView === 'oop') renderOopView();
    if (state.mainView === 'spectator') renderSpectatorView();
  };

  window.onEnterOopView = function () {
    renderOopView();
  };

  window.onEnterSpectatorView = function () {
    opsUi.spectator.organizerBypassClosed = false;
    renderSpectatorView();
  };

  // HTML onclick ハンドラ向けに公開
  window.updateParentTournamentName = updateParentTournamentName;
  window.switchActiveEvent = switchActiveEvent;
  window.openEventModal = openEventModal;
  window.closeEventModal = closeEventModal;
  window.submitEventModal = submitEventModal;
  window.duplicateActiveEvent = duplicateActiveEvent;
  window.deleteActiveEvent = deleteActiveEvent;

  window.exportBundleJsonState = exportBundleJsonState;
  window.importBundleJsonState = importBundleJsonState;

  window.updateOopVenue = updateOopVenue;
  window.changeOopCourtCount = changeOopCourtCount;
  window.setOopAssignMode = setOopAssignMode;
  window.regenerateOopSchedule = regenerateOopSchedule;
  window.toggleOopUmpireRule = toggleOopUmpireRule;
  window.setOopViewStyle = setOopViewStyle;
  window.toggleOopCourtFinished = toggleOopCourtFinished;
  window.setOopMatchStatus = setOopMatchStatus;
  window.openOopScoreInput = openOopScoreInput;
  window.toggleOopWoChooser = toggleOopWoChooser;
  window.recordOopWalkover = recordOopWalkover;
  window.selectOopMatchForMove = selectOopMatchForMove;
  window.cancelOopMoveSelection = cancelOopMoveSelection;
  window.moveOopMatchToCourt = moveOopMatchToCourt;
  window.undoOopAction = undoOopAction;
  window.toggleOopHideFinished = toggleOopHideFinished;
  window.toggleOopQueueDrawer = toggleOopQueueDrawer;
  window.reorderOopMatch = reorderOopMatch;
  window.handleOopDragStart = handleOopDragStart;
  window.handleOopDragEnd = handleOopDragEnd;
  window.handleOopDragOver = handleOopDragOver;
  window.handleOopDragLeave = handleOopDragLeave;
  window.handleOopDrop = handleOopDrop;
  window.handleOopInsertZoneClick = handleOopInsertZoneClick;
  window.handleOopInsertDragOver = handleOopInsertDragOver;
  window.handleOopInsertDragLeave = handleOopInsertDragLeave;
  window.handleOopInsertDrop = handleOopInsertDrop;
  window.handleOopCardDragOver = handleOopCardDragOver;
  window.handleOopCardDragLeave = handleOopCardDragLeave;
  window.handleOopCardDrop = handleOopCardDrop;
  window.autoAssignUnassignedMatches = autoAssignUnassignedMatches;

  // =========================================================================
  // 5c. 運営者リアルタイム・アナウンス機能
  // =========================================================================
  function renderAnnouncementBannersHtml(tourAnn, evAnn, venueName) {
    const items = [];
    if (tourAnn && String(tourAnn.text || '').trim()) {
      items.push({ scope: '大会本部より', ann: tourAnn });
    }
    if (evAnn && String(evAnn.text || '').trim()) {
      items.push({ scope: `${venueName || '会場'}より`, ann: evAnn });
    }
    if (!items.length) return '';

    const themeMap = {
      urgent: {
        wrap: 'bg-gradient-to-r from-rose-600 via-red-600 to-rose-700 text-white border-y-2 border-rose-800',
        pill: 'bg-white text-rose-700',
        pillText: '緊急・重要',
        iconBox: 'bg-white/20 text-white border border-white/30',
        dot: '<span class="w-2 h-2 rounded-full bg-white live-dot shrink-0"></span>',
        timeCls: 'bg-black/20 text-rose-100'
      },
      warn: {
        wrap: 'bg-gradient-to-r from-amber-400 via-yellow-300 to-amber-300 text-slate-950 border-y-2 border-amber-600',
        pill: 'bg-slate-950 text-amber-300',
        pillText: '重要なお知らせ',
        iconBox: 'bg-amber-950/15 text-slate-950 border border-amber-800/30',
        dot: '<span class="w-2 h-2 rounded-full bg-rose-600 live-dot shrink-0"></span>',
        timeCls: 'bg-amber-950/15 text-slate-900'
      },
      info: {
        wrap: 'bg-gradient-to-r from-blue-600 via-blue-600 to-indigo-600 text-white border-y-2 border-blue-800',
        pill: 'bg-white text-blue-700',
        pillText: 'お知らせ',
        iconBox: 'bg-white/20 text-white border border-white/30',
        dot: '',
        timeCls: 'bg-black/20 text-blue-100'
      }
    };

    return `<div class="shrink-0 shadow-md relative z-10 divide-y divide-black/15">${items.map(({ scope, ann }) => {
      const th = themeMap[ann.level] || themeMap.warn;
      const tm = ann.updatedAt ? formatLocalClock(ann.updatedAt) : '';
      return `<div class="px-3 py-2.5 ${th.wrap}">
        <div class="flex items-start gap-2.5">
          <div class="w-7 h-7 rounded-lg ${th.iconBox} flex items-center justify-center shrink-0 mt-0.5 text-sm font-bold shadow-sm" aria-hidden="true">📢</div>
          <div class="min-w-0 flex-1 space-y-1">
            <div class="flex items-center justify-between gap-1.5 flex-wrap">
              <div class="flex items-center gap-1.5">
                ${th.dot}
                <span class="px-1.5 py-0.2 rounded text-[10px] font-black tracking-wide ${th.pill}">${th.pillText}</span>
                <span class="text-[11px] font-extrabold opacity-95">${escapeHtml(scope)}</span>
              </div>
              ${tm ? `<span class="px-1.5 py-0.2 rounded text-[10px] font-bold ${th.timeCls}">${escapeHtml(tm)} 更新</span>` : ''}
            </div>
            <div class="text-[13px] font-extrabold leading-snug whitespace-pre-wrap break-words">${escapeHtml(ann.text)}</div>
          </div>
        </div>
      </div>`;
    }).join('')}</div>`;
  }

  function openAnnouncementModal() {
    const isStaff = window.BracketCloud && window.BracketCloud.isStaffMode();
    const scopeRow = document.getElementById('announce-scope-row');
    if (scopeRow) scopeRow.classList.toggle('hidden', Boolean(isStaff));
    document.querySelectorAll('input[name="announce-scope"]').forEach((r) => {
      r.checked = isStaff ? (r.value === 'event') : (r.value === 'tournament');
    });
    loadAnnouncementScopeDraft();
    document.getElementById('announcement-modal')?.classList.remove('hidden');
  }

  function closeAnnouncementModal() {
    document.getElementById('announcement-modal')?.classList.add('hidden');
  }

  function getSelectedAnnouncementScope() {
    if (window.BracketCloud && window.BracketCloud.isStaffMode()) return 'event';
    const r = document.querySelector('input[name="announce-scope"]:checked');
    return r ? r.value : 'tournament';
  }

  function loadAnnouncementScopeDraft() {
    const scope = getSelectedAnnouncementScope();
    const current = scope === 'event'
      ? (state.oop && state.oop.announcement)
      : (opsUi.bundle && opsUi.bundle.announcement);
    const ta = document.getElementById('announce-text-input');
    if (ta) ta.value = (current && current.text) || '';
    const lv = (current && current.level) || 'warn';
    document.querySelectorAll('input[name="announce-level"]').forEach((r) => {
      r.checked = (r.value === lv);
    });
  }

  function applyAnnouncementPreset(level, text) {
    document.querySelectorAll('input[name="announce-level"]').forEach((r) => {
      r.checked = (r.value === level);
    });
    const ta = document.getElementById('announce-text-input');
    if (ta) ta.value = text;
  }

  function saveAnnouncementFromModal() {
    const scope = getSelectedAnnouncementScope();
    const text = String(document.getElementById('announce-text-input')?.value || '').trim();
    const checkedLv = document.querySelector('input[name="announce-level"]:checked');
    const level = checkedLv ? checkedLv.value : 'warn';
    const payload = text ? { text, level, updatedAt: new Date().toISOString() } : null;

    if (scope === 'event') {
      if (!state.oop) state.oop = OopEngine.createDefaultOopState();
      state.oop.announcement = payload;
      saveToLocalStorage();
    } else if (opsUi.bundle) {
      opsUi.bundle.announcement = payload;
      saveBundleToLocalStorage();
    }
    if (window.BracketCloud && typeof window.BracketCloud.scheduleCloudSync === 'function') {
      window.BracketCloud.scheduleCloudSync(true);
    }
    closeAnnouncementModal();
    if (state.mainView === 'oop') renderOopView();
    if (state.mainView === 'spectator') renderSpectatorView();
    showToast(payload ? 'アナウンスを配信しました（公開画面へ即時反映）' : 'アナウンスを解除しました', 'success');
  }

  function clearAnnouncementFromModal() {
    const ta = document.getElementById('announce-text-input');
    if (ta) ta.value = '';
    saveAnnouncementFromModal();
  }

  // =========================================================================
  // 5d. 軽量「大会ホーム」ダッシュボード（A案）
  // =========================================================================
  function onEnterHomeView() {
    renderHomeView();
    if (window.BracketCloud && window.BracketCloud.isLoggedIn() && typeof window.BracketCloud.fetchMyTournaments === 'function') {
      window.BracketCloud.fetchMyTournaments(false);
    }
  }

  function setHomeSearchQuery(val) {
    opsUi.homeSearchQuery = String(val || '');
    renderHomeView();
    const inp = document.getElementById('home-tourney-search');
    if (inp) {
      inp.focus();
      inp.setSelectionRange(inp.value.length, inp.value.length);
    }
  }

  function startSampleFromHome() {
    if (typeof loadRealisticSampleData === 'function') {
      loadRealisticSampleData(false);
      generateTournamentFromEntries(false, false);
    }
    switchMainView('tournament');
    showToast('サンプルデータでトーナメントを作成しました', 'success');
  }

  function openJsonImportFromHome() {
    const inp = document.getElementById('json-import-input');
    if (inp) inp.click();
  }

  function renderHomeView() {
    const root = document.getElementById('home-dashboard-root');
    if (!root || !opsUi.bundle) return;

    const curEv = getActiveEvent();
    if (curEv) curEv.data = captureActiveStateSnapshot();
    const pub = opsUi.bundle.publish || { enabled: true };
    const isLogged = Boolean(window.BracketCloud && window.BracketCloud.isLoggedIn());
    const user = isLogged ? window.BracketCloud.getUser() : null;
    const cloudList = isLogged ? window.BracketCloud.getMyTournaments() : [];
    const q = opsUi.homeSearchQuery.trim().toLowerCase();
    const filteredCloud = cloudList.filter((t) => {
      if (!q) return true;
      return `${t.name} ${(t.eventNames || []).join(' ')}`.toLowerCase().includes(q);
    });

    const eventsSummary = (opsUi.bundle.events || []).map((ev) => {
      const cnt = (ev.data && Array.isArray(ev.data.entries)) ? ev.data.entries.length : 0;
      return `${ev.name} (${cnt}組)`;
    }).join(' ／ ');

    const bottomHtml = isLogged ? `
      <div class="bg-white rounded-2xl border border-slate-200 shadow-sm p-4 sm:p-5 space-y-3">
        <div class="flex flex-wrap items-center justify-between gap-2">
          <div class="flex items-center gap-2">
            <h3 class="text-sm font-extrabold text-slate-900">マイ大会一覧（クラウド保存）</h3>
            <span class="px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-700 border border-indigo-200 text-[11px] font-bold">${filteredCloud.length}件</span>
            ${user ? `<span class="text-[11px] text-slate-400 hidden sm:inline">${escapeHtml(user.email || user.displayName || '')}</span>` : ''}
          </div>
          <div class="flex items-center gap-2">
            <input id="home-tourney-search" type="text" value="${escapeHtml(opsUi.homeSearchQuery)}"
                   oninput="setHomeSearchQuery(this.value)"
                   placeholder="大会名・種目で検索..."
                   class="px-3 py-1 text-xs font-bold border border-slate-300 rounded-lg bg-slate-50 focus:bg-white focus:outline-none focus:border-blue-500 w-44 sm:w-56" />
            <button onclick="syncNowToCloud()" class="px-2.5 py-1 text-xs font-bold bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg shadow-sm transition">今すぐ同期</button>
          </div>
        </div>

        ${filteredCloud.length === 0 ? `
          <div class="py-8 text-center text-xs text-slate-400 border border-dashed border-slate-200 rounded-xl">
            ${cloudList.length === 0 ? 'クラウド保存された大会はまだありません。右上の「今すぐ同期」で現在の大会を保存できます。' : '一致する大会がありません'}
          </div>
        ` : `
          <div class="grid grid-cols-1 md:grid-cols-2 gap-2.5">
            ${filteredCloud.map((t) => {
              const isCur = t.id === opsUi.bundle.id;
              const dt = formatLocalDateTime(t.updatedAt);
              return `
                <div class="p-3 rounded-xl border ${isCur ? 'border-emerald-400 bg-emerald-50/40' : 'border-slate-200 bg-slate-50/70'} flex items-center justify-between gap-3">
                  <div class="min-w-0 flex-1">
                    <div class="flex items-center gap-1.5 flex-wrap">
                      <span class="text-xs font-extrabold text-slate-900 truncate">${escapeHtml(t.name)}</span>
                      ${isCur ? `<span class="px-1.5 py-0.2 rounded bg-emerald-600 text-white text-[9px] font-bold">編集中</span>` : ''}
                      <span class="px-1.5 py-0.2 rounded ${t.publishEnabled ? 'bg-blue-100 text-blue-800' : 'bg-slate-200 text-slate-600'} text-[9px] font-bold">${t.publishEnabled ? '公開中' : '非公開'}</span>
                    </div>
                    <div class="text-[11px] text-slate-600 truncate mt-0.5">${escapeHtml((t.eventNames || []).join(' / ') || '1種目')}</div>
                    <div class="text-[10px] text-slate-400 mt-0.5">更新: ${escapeHtml(dt)}</div>
                  </div>
                  <div class="flex flex-col gap-1 shrink-0">
                    ${!isCur ? `<button onclick="loadCloudTournament('${escapeJs(t.id)}')" class="px-3 py-1 text-[11px] font-bold bg-blue-600 hover:bg-blue-500 text-white rounded-lg transition">開く</button>` : `<button onclick="switchMainView('oop')" class="px-3 py-1 text-[11px] font-bold bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg transition">運営へ</button>`}
                    <div class="flex gap-1">
                      <button onclick="duplicateCloudTournament('${escapeJs(t.id)}')" class="flex-1 px-2 py-0.5 text-[10px] font-bold bg-white hover:bg-slate-100 border border-slate-300 text-slate-700 rounded">複製</button>
                      <button onclick="deleteCloudTournament('${escapeJs(t.id)}', '${escapeJs(t.name)}')" class="px-2 py-0.5 text-[10px] font-bold bg-rose-50 hover:bg-rose-100 text-rose-700 rounded">削除</button>
                    </div>
                  </div>
                </div>
              `;
            }).join('')}
          </div>
        `}
      </div>
    ` : `
      <div class="grid grid-cols-1 md:grid-cols-12 gap-4">
        <div class="md:col-span-5 bg-white rounded-2xl border border-blue-200 shadow-sm p-5 flex flex-col justify-between gap-3">
          <div class="space-y-1.5">
            <div class="text-xs font-extrabold text-blue-700">Google ログインで全機能が解放</div>
            <p class="text-xs text-slate-600 leading-relaxed">
              過去の大会をクラウド保存・呼び出しでき、生徒・保護者スマホへのQRリアルタイム速報や、別会場スタッフPIN同期も無料で使えます。
            </p>
          </div>
          <button onclick="signInWithGoogle()"
                  class="w-full py-2.5 px-4 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold shadow transition">
            G Googleアカウントでログイン
          </button>
        </div>
        <div class="md:col-span-7 bg-white rounded-2xl border border-slate-200 shadow-sm p-5 space-y-3">
          <div class="text-xs font-extrabold text-slate-500">BracketFlow の使い方（3ステップ）</div>
          <div class="grid grid-cols-1 sm:grid-cols-3 gap-2.5 text-xs">
            <div class="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-1">
              <div class="text-[10px] font-extrabold text-blue-600">STEP 1</div>
              <div class="font-extrabold text-slate-900">選手登録・自動抽選</div>
              <div class="text-[11px] text-slate-500 leading-snug">Excel読込から4山・シード規則で自動配置＆微調整</div>
            </div>
            <div class="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-1">
              <div class="text-[10px] font-extrabold text-emerald-600">STEP 2</div>
              <div class="font-extrabold text-slate-900">進行表・コート運営</div>
              <div class="text-[11px] text-slate-500 leading-snug">T字ドラッグ＆空き枠で進行表を作成しリアルタイム管理</div>
            </div>
            <div class="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-1">
              <div class="text-[10px] font-extrabold text-rose-600">STEP 3</div>
              <div class="font-extrabold text-slate-900">QR速報・アナウンス</div>
              <div class="text-[11px] text-slate-500 leading-snug">生徒・保護者スマホへ自動更新配信＆雨天アナウンス</div>
            </div>
          </div>
        </div>
      </div>
    `;

    root.innerHTML = `
      <div class="grid grid-cols-1 md:grid-cols-12 gap-4">
        <!-- 続きから再開カード -->
        <div class="md:col-span-8 bg-white rounded-2xl border border-slate-200 shadow-sm p-5 flex flex-col justify-between gap-4">
          <div class="flex items-start justify-between gap-3">
            <div class="min-w-0">
              <div class="flex items-center gap-2">
                <span class="text-[11px] font-bold text-slate-500">最後に開いていた大会</span>
                <span class="px-2 py-0.5 rounded-full text-[10px] font-bold ${pub.enabled ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-200 text-slate-600'}">${pub.enabled ? '公開設定 ON' : '非公開'}</span>
              </div>
              <h2 class="text-lg font-extrabold text-slate-900 truncate mt-0.5">${escapeHtml(opsUi.bundle.name || '大会')}</h2>
              <div class="text-xs text-slate-600 mt-1">${escapeHtml(eventsSummary)}</div>
            </div>
            <button onclick="openPublishModal()"
                    class="px-3 py-1.5 text-xs font-bold bg-slate-100 hover:bg-slate-200 border border-slate-300 text-slate-800 rounded-lg shrink-0 transition">
              公開設定・QR
            </button>
          </div>

          <div class="grid grid-cols-3 gap-2">
            <button onclick="switchMainView('registration')"
                    class="py-2.5 px-3 rounded-xl border border-slate-200 bg-slate-50 hover:bg-blue-50 hover:border-blue-300 text-xs font-extrabold text-slate-800 transition">
              1. 選手登録・設定
            </button>
            <button onclick="switchMainView('tournament')"
                    class="py-2.5 px-3 rounded-xl border border-slate-200 bg-slate-50 hover:bg-blue-50 hover:border-blue-300 text-xs font-extrabold text-slate-800 transition">
              2. トーナメント表
            </button>
            <button onclick="switchMainView('oop')"
                    class="py-2.5 px-3 rounded-xl bg-blue-600 hover:bg-blue-500 text-white text-xs font-extrabold shadow-sm transition">
              3. 進行表・運営 →
            </button>
          </div>
        </div>

        <!-- 新しく始めるカード -->
        <div class="md:col-span-4 bg-white rounded-2xl border border-slate-200 shadow-sm p-5 flex flex-col justify-between gap-2.5">
          <div class="text-xs font-extrabold text-slate-500">新しく始める</div>
          <button onclick="createNewCloudTournament()"
                  class="w-full py-2.5 px-4 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-extrabold shadow-sm transition">
            ＋ 新しい大会を作成
          </button>
          <div class="grid grid-cols-2 gap-2">
            <button onclick="openJsonImportFromHome()"
                    class="py-2 px-2.5 rounded-xl border border-slate-300 bg-slate-50 hover:bg-slate-100 text-xs font-bold text-slate-700 transition">
              JSON読込
            </button>
            <button onclick="startSampleFromHome()"
                    class="py-2 px-2.5 rounded-xl border border-slate-300 bg-slate-50 hover:bg-slate-100 text-xs font-bold text-slate-700 transition">
              サンプル開始
            </button>
          </div>
        </div>
      </div>

      ${bottomHtml}
    `;
  }

  window.onEnterHomeView = onEnterHomeView;
  window.renderHomeView = renderHomeView;
  window.setHomeSearchQuery = setHomeSearchQuery;
  window.startSampleFromHome = startSampleFromHome;
  window.openJsonImportFromHome = openJsonImportFromHome;

  function toggleOopSettingsMenu() {
    const pop = document.getElementById('oop-settings-popover');
    const exp = document.getElementById('oop-export-popover');
    if (exp) exp.classList.add('hidden');
    if (pop) pop.classList.toggle('hidden');
  }

  function toggleOopExportMenu() {
    const pop = document.getElementById('oop-export-popover');
    const set = document.getElementById('oop-settings-popover');
    if (set) set.classList.add('hidden');
    if (pop) pop.classList.toggle('hidden');
  }

  function closeOopMenus() {
    document.getElementById('oop-settings-popover')?.classList.add('hidden');
    document.getElementById('oop-export-popover')?.classList.add('hidden');
  }

  document.addEventListener('click', (e) => {
    const setPop = document.getElementById('oop-settings-popover');
    const setBtn = document.getElementById('btn-oop-settings-menu');
    if (setPop && !setPop.classList.contains('hidden')) {
      if (!setPop.contains(e.target) && (!setBtn || !setBtn.contains(e.target))) {
        setPop.classList.add('hidden');
      }
    }
    const expPop = document.getElementById('oop-export-popover');
    const expBtn = document.getElementById('btn-oop-export-menu');
    if (expPop && !expPop.classList.contains('hidden')) {
      if (!expPop.contains(e.target) && (!expBtn || !expBtn.contains(e.target))) {
        expPop.classList.add('hidden');
      }
    }
  });

  window.toggleOopSettingsMenu = toggleOopSettingsMenu;
  window.toggleOopExportMenu = toggleOopExportMenu;
  window.closeOopMenus = closeOopMenus;

  window.setOopSplitBlock = setOopSplitBlock;
  window.toggleOopConflictsDrawer = toggleOopConflictsDrawer;
  window.highlightOopCard = highlightOopCard;
  window.jumpToOopCard = jumpToOopCard;
  window.selectNewOopBlankForInsert = selectNewOopBlankForInsert;
  window.duplicateOopBlankRight = duplicateOopBlankRight;
  window.editOopBlankLabel = editOopBlankLabel;
  window.removeOopItem = removeOopItem;
  window.alignOopQuarterfinals = alignOopQuarterfinals;
  window.clearAllOopAssignments = clearAllOopAssignments;

  window.previewUmpireByNumber = previewUmpireByNumber;
  window.applyUmpireByNumber = applyUmpireByNumber;
  window.applyUmpireFromSelect = applyUmpireFromSelect;

  window.openAnnouncementModal = openAnnouncementModal;
  window.closeAnnouncementModal = closeAnnouncementModal;
  window.loadAnnouncementScopeDraft = loadAnnouncementScopeDraft;
  window.applyAnnouncementPreset = applyAnnouncementPreset;
  window.saveAnnouncementFromModal = saveAnnouncementFromModal;
  window.clearAnnouncementFromModal = clearAnnouncementFromModal;

  window.openUmpireModal = openUmpireModal;
  window.closeUmpireModal = closeUmpireModal;
  window.saveUmpireOverride = saveUmpireOverride;
  window.clearUmpireOverride = clearUmpireOverride;

  window.printOopSheet = printOopSheet;
  window.exportOopPdf = exportOopPdf;
  window.exportOopExcel = exportOopExcel;

  window.openPublishModal = openPublishModal;
  window.closePublishModal = closePublishModal;
  window.updatePublishSettingsFromModal = updatePublishSettingsFromModal;
  window.copyPortalUrl = copyPortalUrl;
  window.copyStaffUrl = copyStaffUrl;
  window.copyStaffPin = copyStaffPin;
  window.regenerateStaffPin = regenerateStaffPin;
  window.refreshPublishCloudBanner = refreshPublishCloudBanner;

  window.getCurrentTournamentBundle = getCurrentTournamentBundle;
  window.replaceTournamentBundle = replaceTournamentBundle;
  window.createFreshTournamentBundle = createFreshTournamentBundle;
  window.applyRemoteLivePatchToActiveEvent = applyRemoteLivePatchToActiveEvent;
  window.hydrateStaffSingleEvent = hydrateStaffSingleEvent;
  window.renderStandalonePublicPortal = renderStandalonePublicPortal;
  window.submitStandalonePasscode = submitStandalonePasscode;
  window.selectStandalonePublicEvent = selectStandalonePublicEvent;
  window.setStandaloneSpectatorTab = setStandaloneSpectatorTab;
  window.jumpSpectatorBracketMountain = jumpSpectatorBracketMountain;

  window.setSpectatorFrameMode = setSpectatorFrameMode;
  window.setSpectatorTab = setSpectatorTab;
  window.selectSpectatorEvent = selectSpectatorEvent;
  window.submitSpectatorPasscode = submitSpectatorPasscode;
  window.bypassSpectatorGate = bypassSpectatorGate;
  window.toggleSpectatorFavorite = toggleSpectatorFavorite;
  window.setSpectatorSearchQuery = setSpectatorSearchQuery;
  window.setSpectatorSchoolFilter = setSpectatorSchoolFilter;
  window.adjustSpectatorBracketZoom = adjustSpectatorBracketZoom;
  window.setSpectatorCourtFilter = setSpectatorCourtFilter;
  window.toggleSpectatorCourtExpand = toggleSpectatorCourtExpand;
  window.setSpectatorBracketBlock = setSpectatorBracketBlock;
  window.setSpectatorBracketStage = setSpectatorBracketStage;
  window.fitSpectatorBracketZoom = fitSpectatorBracketZoom;
  window.jumpSpectatorToFirstChecked = jumpSpectatorToFirstChecked;
  window.toggleSpectatorMyCheckCollapsed = toggleSpectatorMyCheckCollapsed;
  window.focusSpectatorCourt = focusSpectatorCourt;

  // 初期化（app.js の init 完了後に実行）
  initBundle();

  // 通常起動時は軽量「大会ホーム」を表示（?public= / ?staff= / #public / #staff は直接目的画面へ）
  const launchParams = new URLSearchParams(window.location.search || '');
  const isSpecialBoot = Boolean(
    launchParams.get('public') ||
    (launchParams.get('staff') && launchParams.get('e')) ||
    (window.location.hash && (window.location.hash.startsWith('#public') || window.location.hash.startsWith('#staff')))
  );
  if (window.location.hash === '#public') {
    switchMainView('spectator');
  } else if (!isSpecialBoot) {
    switchMainView('home');
  }
})();

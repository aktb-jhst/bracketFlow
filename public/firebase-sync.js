/* global firebase, OopEngine, state, showToast, escapeHtml, escapeJs, formatSchoolShortName */
'use strict';

/**
 * BracketFlow — Phase 2 Firebase Auth / Cloud Save / Realtime Sync / Staff PIN Module
 * Spark 無料プラン完全対応（1大会1メタ + 1種目1ライブドキュメント設計）
 */
(function () {
  const cloud = {
    initialized: false,
    auth: null,
    db: null,
    user: null,             // Googleログイン済み管理者（匿名スタッフは除外）
    isStaffMode: false,
    staffSession: null,     // { tid, evId, docId, pinVersion, venue, eventName, tournamentName }
    isPublicStandalone: false,
    publicTid: null,
    publicMeta: null,
    publicEventUnsub: null,
    adminLiveUnsub: null,
    adminLiveDocId: null,
    suppressRemoteEcho: false,
    syncTimer: null,
    lastSyncedAt: null,
    syncStatus: 'local',    // 'local' | 'syncing' | 'synced' | 'error'
    myTournaments: []
  };

  function isFirebaseReady() {
    return typeof firebase !== 'undefined' && Boolean(window.BRACKETFLOW_FIREBASE_CONFIG && window.BRACKETFLOW_FIREBASE_CONFIG.apiKey);
  }

  function initFirebase() {
    if (!isFirebaseReady()) {
      updateCloudStatusUi('local', 'オフライン（ローカル保存）');
      return false;
    }
    try {
      if (!firebase.apps || firebase.apps.length === 0) {
        firebase.initializeApp(window.BRACKETFLOW_FIREBASE_CONFIG);
      }
      cloud.auth = firebase.auth();
      cloud.db = firebase.firestore();
      cloud.initialized = true;

      // 認証状態監視
      cloud.auth.onAuthStateChanged((u) => {
        if (u && !u.isAnonymous) {
          cloud.user = u;
          updateAuthHeaderUi();
          updateCloudStatusUi('synced', 'クラウド連携中');
          if (!cloud.isPublicStandalone && !cloud.isStaffMode) {
            fetchMyTournaments(false);
            attachAdminLiveListener();
          }
        } else {
          cloud.user = null;
          updateAuthHeaderUi();
          if (!cloud.isStaffMode && !cloud.isPublicStandalone) {
            detachAdminLiveListener();
            updateCloudStatusUi('local', '未ログイン（ローカル保存）');
          }
        }
        if (typeof window.refreshPublishCloudBanner === 'function') {
          window.refreshPublishCloudBanner();
        }
      });
      return true;
    } catch (err) {
      console.warn('Firebase init failed:', err);
      updateCloudStatusUi('error', 'Firebase初期化エラー');
      return false;
    }
  }

  async function sha256Hex(str) {
    const s = String(str || '');
    if (!s) return '';
    if (window.crypto && window.crypto.subtle) {
      const buf = await window.crypto.subtle.digest('SHA-256', new TextEncoder().encode(s));
      return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, '0')).join('');
    }
    // フォールバック（非HTTPSローカル環境）
    let h = 2166136261;
    for (let i = 0; i < s.length; i++) {
      h ^= s.charCodeAt(i);
      h = Math.imul(h, 16777619);
    }
    return `fnv_${(h >>> 0).toString(16)}`;
  }

  function updateAuthHeaderUi() {
    const btnLogin = document.getElementById('btn-google-login');
    const userBox = document.getElementById('header-user-box');
    const userName = document.getElementById('header-user-name');
    const userAvatar = document.getElementById('header-user-avatar');

    if (cloud.user) {
      if (btnLogin) btnLogin.classList.add('hidden');
      if (userBox) userBox.classList.remove('hidden');
      if (userName) userName.textContent = cloud.user.displayName || cloud.user.email || 'ログイン中';
      if (userAvatar) {
        if (cloud.user.photoURL) {
          userAvatar.src = cloud.user.photoURL;
          userAvatar.classList.remove('hidden');
        } else {
          userAvatar.classList.add('hidden');
        }
      }
    } else {
      if (btnLogin) btnLogin.classList.remove('hidden');
      if (userBox) userBox.classList.add('hidden');
    }
  }

  function updateCloudStatusUi(status, label) {
    cloud.syncStatus = status;
    const dot = document.getElementById('cloud-sync-dot');
    const text = document.getElementById('cloud-sync-text');
    const map = {
      local: 'bg-slate-400',
      syncing: 'bg-amber-400 live-dot',
      synced: 'bg-emerald-400',
      error: 'bg-rose-500'
    };
    if (dot) dot.className = `w-1.5 h-1.5 rounded-full ${map[status] || map.local}`;
    if (text && label) text.textContent = label;
  }

  // =========================================================================
  // 1. Google ログイン / ログアウト
  // =========================================================================
  async function signInWithGoogle() {
    if (!cloud.initialized && !initFirebase()) {
      alert('Firebase SDK が読み込まれていません。通信環境をご確認ください。');
      return;
    }
    const provider = new firebase.auth.GoogleAuthProvider();
    provider.setCustomParameters({ prompt: 'select_account' });
    try {
      await cloud.auth.signInWithPopup(provider);
      showToast('Googleアカウントでログインしました', 'success');
      // 初回ログイン時に現在の大会をクラウド同期するか確認せずすぐ同期できるよう案内
      scheduleCloudSync(true);
    } catch (err) {
      if (err && err.code === 'auth/popup-blocked') {
        await cloud.auth.signInWithRedirect(provider);
        return;
      }
      if (err && err.code === 'auth/unauthorized-domain') {
        alert(
          `このドメイン（${window.location.hostname}）は Firebase Authentication の「承認済みドメイン」にまだ追加されていません。\n\n` +
          `Firebase コンソール → Authentication → Settings → 承認済みドメイン に「${window.location.hostname}」を追加してから再度お試しください。`
        );
        return;
      }
      if (err && err.code === 'auth/operation-not-allowed') {
        alert(
          'Firebase コンソールで Google ログインがまだ有効化されていません。\n' +
          'Authentication → Sign-in method → Google を有効にしてから再度お試しください。'
        );
        return;
      }
      if (err && err.code === 'auth/popup-closed-by-user') return;
      alert('ログインに失敗しました: ' + (err.message || err));
    }
  }

  async function signOutGoogle() {
    if (!cloud.initialized || !cloud.auth) return;
    if (!confirm('ログアウトしますか？（現在の画面データはローカルにも残ります）')) return;
    try {
      detachAdminLiveListener();
      await cloud.auth.signOut();
      cloud.myTournaments = [];
      showToast('ログアウトしました（ローカル保存モードに切替）', 'info');
    } catch (err) {
      alert('ログアウトに失敗しました: ' + err.message);
    }
  }

  // =========================================================================
  // 2. クラウド保存（tournaments + live_tournaments + live_events + staff_keys）
  // =========================================================================
  function maskSlotsForPublic(slots, nameMode) {
    if (!Array.isArray(slots)) return [];
    if (nameMode !== 'surname') return slots;
    return slots.map((entry) => {
      if (!entry) return null;
      return Object.assign({}, entry, {
        p1: OopEngine.extractSurname(entry.p1),
        p2: OopEngine.extractSurname(entry.p2)
      });
    });
  }

  function scheduleCloudSync(immediate) {
    if (cloud.suppressRemoteEcho) return;
    // 会場スタッフモード中はスタッフ権限の live_events のみ更新
    if (cloud.isStaffMode && cloud.staffSession) {
      if (cloud.syncTimer) clearTimeout(cloud.syncTimer);
      cloud.syncTimer = setTimeout(() => pushStaffLiveUpdate(), immediate ? 150 : 600);
      return;
    }
    if (!cloud.initialized || !cloud.user || cloud.isPublicStandalone) return;
    if (cloud.syncTimer) clearTimeout(cloud.syncTimer);
    updateCloudStatusUi('syncing', '同期待機中...');
    cloud.syncTimer = setTimeout(() => syncActiveBundleToCloud(false), immediate ? 250 : 1800);
  }

  async function syncActiveBundleToCloud(notifyUser) {
    if (!cloud.initialized || !cloud.user) {
      if (notifyUser) showToast('クラウド保存するにはGoogleログインしてください', 'warn');
      return false;
    }
    const bundle = typeof window.getCurrentTournamentBundle === 'function'
      ? window.getCurrentTournamentBundle()
      : null;
    if (!bundle || !bundle.id) return false;

    updateCloudStatusUi('syncing', 'クラウド保存中...');
    try {
      const uid = cloud.user.uid;
      const tid = bundle.id;
      const nowIso = new Date().toISOString();
      const pub = bundle.publish || { enabled: true, passcode: '', nameDisplayMode: 'full' };
      const passcodeHash = pub.passcode ? await sha256Hex(`${tid}:${pub.passcode}`) : '';

      const batch = cloud.db.batch();

      // 1. 管理者専用フルバックアップ（1ドキュメント・JSON文字列化でネスト配列制限を回避）
      const tourRef = cloud.db.collection('tournaments').doc(tid);
      batch.set(tourRef, {
        id: tid,
        ownerUid: uid,
        name: bundle.name || '無題の大会',
        eventCount: (bundle.events || []).length,
        eventNames: (bundle.events || []).map((e) => e.name),
        publishEnabled: Boolean(pub.enabled),
        updatedAt: nowIso,
        bundleJson: JSON.stringify(bundle)
      });

      // 2. 公開ポータル用メタ情報（1大会=1ドキュメント）
      const liveTourRef = cloud.db.collection('live_tournaments').doc(tid);
      batch.set(liveTourRef, {
        id: tid,
        ownerUid: uid,
        name: bundle.name || '大会',
        publishEnabled: Boolean(pub.enabled),
        passcodeHash,
        hasPasscode: Boolean(pub.passcode),
        nameDisplayMode: pub.nameDisplayMode || 'full',
        events: (bundle.events || []).map((ev) => ({
          id: ev.id,
          name: ev.name,
          venue: ev.venue || (ev.data && ev.data.oop && ev.data.oop.venue) || '第1会場',
          pinVersion: ev.pinVersion || 1
        })),
        updatedAt: nowIso
      });

      // 3. 種目別ライブドキュメント ＆ スタッフPIN検証キー
      (bundle.events || []).forEach((ev) => {
        const docId = `${tid}__${ev.id}`;
        const evData = ev.data || {};
        const pinVersion = ev.pinVersion || 1;
        const maskedSlots = maskSlotsForPublic(evData.slots || [], pub.nameDisplayMode || 'full');
        const maskedEntries = maskSlotsForPublic(evData.entries || [], pub.nameDisplayMode || 'full');

        const liveEvRef = cloud.db.collection('live_events').doc(docId);
        batch.set(liveEvRef, {
          ownerUid: uid,
          tid,
          evId: ev.id,
          name: ev.name,
          venue: ev.venue || (evData.oop && evData.oop.venue) || '第1会場',
          matchFormat: evData.matchFormat || 7,
          bracketSize: evData.bracketSize || 16,
          slotsJson: JSON.stringify(maskedSlots),
          entriesJson: JSON.stringify(maskedEntries),
          activeSlotMask: evData.activeSlotMask || [],
          schoolAliases: evData.schoolAliases || {},
          results: evData.results || {},
          consolation: evData.consolation || { preset: 'none', title: '順位決定戦', matchCount: 0, slots: [], results: {} },
          oop: evData.oop || OopEngine.createDefaultOopState(),
          publishEnabled: Boolean(pub.enabled),
          pinVersion,
          updatedAt: nowIso
        });

        // スタッフPIN検証用キー（外部読取不可）
        if (ev.staffPin) {
          const keyRef = cloud.db.collection('staff_keys').doc(`${docId}__${ev.staffPin}`);
          batch.set(keyRef, {
            ownerUid: uid,
            docId,
            pinVersion,
            updatedAt: nowIso
          });
        }
      });

      await batch.commit();
      cloud.lastSyncedAt = nowIso;
      updateCloudStatusUi('synced', `クラウド保存済 (${nowIso.slice(11, 16)})`);
      attachAdminLiveListener();
      if (typeof window.refreshPublishCloudBanner === 'function') window.refreshPublishCloudBanner();
      if (notifyUser) showToast('クラウドへ保存し、公開URL・スタッフ画面へ即時反映しました', 'success');
      return true;
    } catch (err) {
      console.error('Cloud sync error:', err);
      updateCloudStatusUi('error', '同期エラー');
      if (notifyUser) {
        if (err && err.code === 'permission-denied') {
          alert('Firestore のセキュリティルールまたは認証権限で保存が拒否されました。Firebase設定をご確認ください。');
        } else {
          alert('クラウド保存に失敗しました: ' + (err.message || err));
        }
      }
      return false;
    }
  }

  async function deleteStaffKeyFromCloud(tid, evId, oldPin) {
    if (!cloud.initialized || !cloud.user || !oldPin) return;
    try {
      await cloud.db.collection('staff_keys').doc(`${tid}__${evId}__${oldPin}`).delete();
    } catch (_) {
      // 無視
    }
  }

  // 管理者が開いているアクティブ種目に対し、別会場スタッフからの更新をリアルタイム受信
  function detachAdminLiveListener() {
    if (cloud.adminLiveUnsub) {
      cloud.adminLiveUnsub();
      cloud.adminLiveUnsub = null;
      cloud.adminLiveDocId = null;
    }
  }

  function attachAdminLiveListener() {
    if (!cloud.initialized || !cloud.user || cloud.isPublicStandalone || cloud.isStaffMode) return;
    const bundle = typeof window.getCurrentTournamentBundle === 'function' ? window.getCurrentTournamentBundle() : null;
    if (!bundle || !bundle.id || !bundle.activeEventId) return;
    const docId = `${bundle.id}__${bundle.activeEventId}`;
    if (cloud.adminLiveDocId === docId && cloud.adminLiveUnsub) return;

    detachAdminLiveListener();
    cloud.adminLiveDocId = docId;
    cloud.adminLiveUnsub = cloud.db.collection('live_events').doc(docId).onSnapshot((snap) => {
      if (!snap.exists) return;
      const remote = snap.data();
      if (!remote || !remote.updatedAt) return;
      if (cloud.lastSyncedAt && remote.updatedAt <= cloud.lastSyncedAt) return;
      // 別端末（会場スタッフ等）からの進行・スコア更新をローカルへ即時マージ
      cloud.suppressRemoteEcho = true;
      try {
        if (typeof window.applyRemoteLivePatchToActiveEvent === 'function') {
          window.applyRemoteLivePatchToActiveEvent({
            results: remote.results || {},
            consolation: remote.consolation || null,
            oop: remote.oop || null,
            updatedAt: remote.updatedAt
          });
        }
        cloud.lastSyncedAt = remote.updatedAt;
        updateCloudStatusUi('synced', `他端末より同期 (${remote.updatedAt.slice(11, 16)})`);
      } finally {
        cloud.suppressRemoteEcho = false;
      }
    }, () => {
      // 権限エラー等は静かに無視
    });
  }

  // =========================================================================
  // 3. マイ大会一覧モーダル（過去の大会を呼び出し・新規作成・削除）
  // =========================================================================
  async function fetchMyTournaments(renderModal) {
    if (!cloud.initialized || !cloud.user) return [];
    const listEl = document.getElementById('my-tournaments-list');
    if (renderModal && listEl) {
      listEl.innerHTML = `<div class="text-center py-8 text-xs text-slate-400">クラウドから大会一覧を読み込んでいます...</div>`;
    }
    try {
      const snap = await cloud.db.collection('tournaments')
        .where('ownerUid', '==', cloud.user.uid)
        .get();
      const rows = [];
      snap.forEach((doc) => {
        const d = doc.data();
        rows.push({
          id: doc.id,
          name: d.name || '無題の大会',
          eventCount: d.eventCount || 1,
          eventNames: Array.isArray(d.eventNames) ? d.eventNames : [],
          publishEnabled: Boolean(d.publishEnabled),
          updatedAt: d.updatedAt || ''
        });
      });
      rows.sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
      cloud.myTournaments = rows;
      if (renderModal) renderMyTournamentsList();
      return rows;
    } catch (err) {
      if (renderModal && listEl) {
        listEl.innerHTML = `<div class="p-4 text-xs text-rose-700 bg-rose-50 border border-rose-200 rounded-lg">読み込みに失敗しました: ${escapeHtml(err.message || err)}</div>`;
      }
      return [];
    }
  }

  function renderMyTournamentsList() {
    const listEl = document.getElementById('my-tournaments-list');
    if (!listEl) return;
    const curBundle = typeof window.getCurrentTournamentBundle === 'function' ? window.getCurrentTournamentBundle() : null;
    const curId = curBundle ? curBundle.id : null;

    if (!cloud.myTournaments.length) {
      listEl.innerHTML = `
        <div class="text-center py-8 space-y-2 text-slate-500">
          <div class="text-xs font-bold">クラウドに保存された大会はまだありません</div>
          <div class="text-[11px]">下の「現在の大会をクラウド保存」を押すと、ここに記録されていつでも呼び出せます。</div>
        </div>`;
      return;
    }

    listEl.innerHTML = cloud.myTournaments.map((t) => {
      const isCur = t.id === curId;
      const dt = t.updatedAt ? t.updatedAt.replace('T', ' ').slice(0, 16) : '-';
      return `
        <div class="p-3 rounded-xl border ${isCur ? 'border-emerald-400 bg-emerald-50/40' : 'border-slate-200 bg-white'} flex items-center justify-between gap-3">
          <div class="min-w-0 flex-1">
            <div class="flex items-center gap-1.5 flex-wrap">
              <span class="text-xs font-extrabold text-slate-900 truncate">${escapeHtml(t.name)}</span>
              ${isCur ? `<span class="px-1.5 py-0.2 rounded bg-emerald-600 text-white text-[9px] font-bold">編集中</span>` : ''}
              <span class="px-1.5 py-0.2 rounded ${t.publishEnabled ? 'bg-blue-100 text-blue-800' : 'bg-slate-100 text-slate-500'} text-[9px] font-bold">${t.publishEnabled ? '公開中' : '非公開'}</span>
            </div>
            <div class="text-[10px] text-slate-500 mt-0.5 truncate">
              種目(${t.eventCount}): ${escapeHtml(t.eventNames.join(' / ') || '1種目')} ・ 更新: ${escapeHtml(dt)}
            </div>
          </div>
          <div class="flex items-center gap-1.5 shrink-0">
            ${!isCur ? `<button onclick="loadCloudTournament('${escapeJs(t.id)}')" class="px-2.5 py-1 text-[11px] font-bold bg-blue-600 hover:bg-blue-500 text-white rounded-lg transition">開く</button>` : ''}
            <button onclick="deleteCloudTournament('${escapeJs(t.id)}', '${escapeJs(t.name)}')" class="px-2 py-1 text-[11px] font-bold bg-rose-50 hover:bg-rose-100 text-rose-700 rounded-lg transition">削除</button>
          </div>
        </div>
      `;
    }).join('');
  }

  function openMyTournamentsModal() {
    if (!cloud.user) {
      if (confirm('過去の大会をクラウドから呼び出すには Google ログインが必要です。ログイン画面を開きますか？')) {
        signInWithGoogle();
      }
      return;
    }
    const modal = document.getElementById('my-tournaments-modal');
    if (modal) modal.classList.remove('hidden');
    fetchMyTournaments(true);
  }

  function closeMyTournamentsModal() {
    const modal = document.getElementById('my-tournaments-modal');
    if (modal) modal.classList.add('hidden');
  }

  async function loadCloudTournament(tid) {
    if (!cloud.initialized || !cloud.user) return;
    try {
      const doc = await cloud.db.collection('tournaments').doc(tid).get();
      if (!doc.exists) {
        alert('指定された大会データが見つかりません。');
        return;
      }
      const d = doc.data();
      const parsed = JSON.parse(d.bundleJson);
      cloud.suppressRemoteEcho = true;
      if (typeof window.replaceTournamentBundle === 'function') {
        window.replaceTournamentBundle(parsed);
      }
      cloud.lastSyncedAt = d.updatedAt || new Date().toISOString();
      cloud.suppressRemoteEcho = false;
      attachAdminLiveListener();
      closeMyTournamentsModal();
      showToast(`大会「${parsed.name}」をクラウドから呼び出しました`, 'success');
    } catch (err) {
      cloud.suppressRemoteEcho = false;
      alert('大会の読み込みに失敗しました: ' + err.message);
    }
  }

  async function createNewCloudTournament() {
    const name = prompt('新しく作成する大会名を入力してください', '令和8年度 秋季地区大会');
    if (!name || !String(name).trim()) return;
    if (typeof window.createFreshTournamentBundle === 'function') {
      window.createFreshTournamentBundle(String(name).trim());
    }
    closeMyTournamentsModal();
    if (cloud.user) {
      await syncActiveBundleToCloud(true);
      fetchMyTournaments(false);
    }
  }

  async function deleteCloudTournament(tid, name) {
    if (!cloud.initialized || !cloud.user) return;
    if (!confirm(`クラウド上の大会「${name}」と公開ページを完全に削除しますか？`)) return;
    try {
      const doc = await cloud.db.collection('tournaments').doc(tid).get();
      const batch = cloud.db.batch();
      batch.delete(cloud.db.collection('tournaments').doc(tid));
      batch.delete(cloud.db.collection('live_tournaments').doc(tid));
      if (doc.exists) {
        const parsed = JSON.parse(doc.data().bundleJson || '{}');
        (parsed.events || []).forEach((ev) => {
          const docId = `${tid}__${ev.id}`;
          batch.delete(cloud.db.collection('live_events').doc(docId));
          if (ev.staffPin) batch.delete(cloud.db.collection('staff_keys').doc(`${docId}__${ev.staffPin}`));
        });
      }
      await batch.commit();
      showToast(`大会「${name}」をクラウドから削除しました`, 'warn');
      fetchMyTournaments(true);
    } catch (err) {
      alert('削除に失敗しました: ' + err.message);
    }
  }

  // =========================================================================
  // 4. 独立 公開ポータルモード（#public?t=<tid>：生徒・保護者の別端末リアルタイム閲覧）
  // =========================================================================
  async function startPublicStandaloneMode(tid) {
    cloud.isPublicStandalone = true;
    cloud.publicTid = tid;
    // 管理用ヘッダー・バーを完全に隠す
    const hdr = document.querySelector('header');
    const evBar = document.getElementById('event-switcher-bar');
    const previewBar = document.querySelector('#view-spectator > div:first-child');
    if (hdr) hdr.classList.add('hidden');
    if (evBar) evBar.classList.add('hidden');
    if (previewBar) previewBar.classList.add('hidden');

    document.getElementById('view-registration')?.classList.add('hidden');
    document.getElementById('view-tournament')?.classList.add('hidden');
    document.getElementById('view-oop')?.classList.add('hidden');
    const specView = document.getElementById('view-spectator');
    if (specView) specView.classList.remove('hidden');

    const frame = document.getElementById('spectator-portal-frame');
    if (frame) {
      frame.className = 'w-full max-w-4xl min-h-screen bg-slate-50 text-slate-900 shadow-none border-0 flex flex-col';
      frame.parentElement.className = 'flex-1 overflow-auto bg-slate-100 flex justify-center items-start';
      frame.innerHTML = `<div class="p-10 text-center text-xs font-bold text-slate-500">大会データをリアルタイム読込中...</div>`;
    }

    if (!cloud.initialized && !initFirebase()) {
      if (frame) frame.innerHTML = `<div class="p-10 text-center text-xs text-rose-700 font-bold">通信エラー：Firebase へ接続できませんでした。</div>`;
      return;
    }

    try {
      const metaSnap = await cloud.db.collection('live_tournaments').doc(tid).get();
      if (!metaSnap.exists) {
        if (frame) {
          frame.innerHTML = `<div class="p-10 text-center space-y-2"><div class="text-sm font-extrabold text-slate-800">大会ページが見つかりません</div><div class="text-xs text-slate-500">URLをご確認いただくか、大会本部がクラウド公開を完了するまでお待ちください。</div></div>`;
        }
        return;
      }
      cloud.publicMeta = metaSnap.data();
      if (!cloud.publicMeta.publishEnabled) {
        if (frame) {
          frame.innerHTML = `<div class="p-10 text-center space-y-2"><div class="text-2xl">🔒</div><div class="text-sm font-extrabold text-slate-800">${escapeHtml(cloud.publicMeta.name)}</div><div class="text-xs text-slate-500">現在この大会は非公開です。</div></div>`;
        }
        return;
      }
      const firstEv = (cloud.publicMeta.events && cloud.publicMeta.events[0]) || null;
      if (firstEv) {
        subscribePublicEvent(tid, firstEv.id);
      }
    } catch (err) {
      if (frame) {
        frame.innerHTML = `<div class="p-10 text-center space-y-2"><div class="text-sm font-extrabold text-rose-700">読み込みエラー</div><div class="text-xs text-slate-500">${escapeHtml(err.message || '現在この大会は非公開または閲覧できません')}</div></div>`;
      }
    }
  }

  function subscribePublicEvent(tid, evId) {
    if (cloud.publicEventUnsub) {
      cloud.publicEventUnsub();
      cloud.publicEventUnsub = null;
    }
    const docId = `${tid}__${evId}`;
    cloud.publicEventUnsub = cloud.db.collection('live_events').doc(docId).onSnapshot((snap) => {
      if (!snap.exists) return;
      const d = snap.data();
      const parsedSlots = JSON.parse(d.slotsJson || '[]');
      const parsedEntries = JSON.parse(d.entriesJson || '[]');
      const liveEventObj = {
        id: d.evId,
        name: d.name,
        venue: d.venue,
        data: {
          bracketSize: d.bracketSize || 16,
          matchFormat: d.matchFormat || 7,
          slots: parsedSlots,
          entries: parsedEntries,
          activeSlotMask: d.activeSlotMask || [],
          schoolAliases: d.schoolAliases || {},
          results: d.results || {},
          consolation: d.consolation || null,
          oop: d.oop || OopEngine.createDefaultOopState()
        }
      };
      if (typeof window.renderStandalonePublicPortal === 'function') {
        window.renderStandalonePublicPortal(cloud.publicMeta, liveEventObj, async (tryPass) => {
          const h = await sha256Hex(`${tid}:${tryPass}`);
          return h === (cloud.publicMeta.passcodeHash || '');
        });
      }
    }, (err) => {
      const frame = document.getElementById('spectator-portal-frame');
      if (frame) {
        frame.innerHTML = `<div class="p-10 text-center text-xs text-slate-500">リアルタイム接続が終了しました（${escapeHtml(err.code || '非公開')}）</div>`;
      }
    });
  }

  // =========================================================================
  // 5. 会場運営スタッフPINモード（#staff?t=<tid>&e=<evId>：ログイン不要で別会場から進行更新）
  // =========================================================================
  async function startStaffPinMode(tid, evId) {
    cloud.isStaffMode = true;
    const modal = document.getElementById('staff-pin-gate-modal');
    const infoEl = document.getElementById('staff-gate-info');
    if (modal) modal.classList.remove('hidden');

    if (!cloud.initialized && !initFirebase()) {
      if (infoEl) infoEl.textContent = 'Firebase に接続できませんでした。';
      return;
    }

    try {
      const docId = `${tid}__${evId}`;
      const [metaSnap, evSnap] = await Promise.all([
        cloud.db.collection('live_tournaments').doc(tid).get(),
        cloud.db.collection('live_events').doc(docId).get()
      ]);
      if (!metaSnap.exists || !evSnap.exists) {
        if (infoEl) infoEl.textContent = '指定された大会・種目が見つからないか、まだクラウド公開されていません。';
        return;
      }
      const meta = metaSnap.data();
      const evDoc = evSnap.data();
      cloud.staffSession = {
        tid,
        evId,
        docId,
        pinVersion: evDoc.pinVersion || 1,
        venue: evDoc.venue || '第1会場',
        eventName: evDoc.name || '種目',
        tournamentName: meta.name || '大会',
        authenticated: false
      };
      if (infoEl) {
        infoEl.innerHTML = `<b>${escapeHtml(meta.name)}</b> ／ ${escapeHtml(evDoc.name)}（${escapeHtml(evDoc.venue || '第1会場')}）`;
      }
    } catch (err) {
      if (infoEl) infoEl.textContent = '大会情報の取得に失敗しました: ' + (err.message || err);
    }
  }

  async function submitStaffPinLogin() {
    if (!cloud.staffSession || !cloud.initialized) return;
    const pinInput = document.getElementById('staff-gate-pin-input');
    const errEl = document.getElementById('staff-gate-error');
    const pin = String(pinInput ? pinInput.value : '').trim();
    if (!pin) return;
    if (errEl) errEl.textContent = '認証中...';

    try {
      // 匿名認証で安全にセッションUIDを取得（Googleログイン不要・無料枠）
      if (!cloud.auth.currentUser) {
        await cloud.auth.signInAnonymously();
      }
      const uid = cloud.auth.currentUser.uid;
      const { docId, pinVersion } = cloud.staffSession;

      // Firestore ルールが staff_keys/{docId}__{pin} の存在をサーバー側で検証
      await cloud.db.collection('live_events').doc(docId).collection('staff_sessions').doc(uid).set({
        pin,
        pinVersion,
        verifiedAt: new Date().toISOString()
      });

      cloud.staffSession.authenticated = true;
      document.getElementById('staff-pin-gate-modal')?.classList.add('hidden');
      enterStaffOopWorkspace();
      showToast(`${cloud.staffSession.eventName}（${cloud.staffSession.venue}）の運営モードに入りました`, 'success');
    } catch (err) {
      if (err && err.code === 'auth/operation-not-allowed') {
        if (errEl) errEl.textContent = 'Firebase Console で「匿名（Anonymous）」認証を有効にしてください。';
      } else {
        if (errEl) errEl.textContent = 'PINが正しくないか、権限がありません。本部に最新PINをご確認ください。';
      }
    }
  }

  function enterStaffOopWorkspace() {
    const { docId, tournamentName, eventName, venue } = cloud.staffSession;
    // 名簿・抽選タブは隠し、進行表・トーナメント確認に集中するスタッフUIへ
    document.getElementById('main-tab-reg')?.classList.add('hidden');
    document.getElementById('event-switcher-bar')?.classList.add('hidden');
    document.getElementById('btn-open-publish-modal')?.classList.add('hidden');
    document.getElementById('btn-google-login')?.classList.add('hidden');
    updateCloudStatusUi('synced', `会場スタッフ: ${eventName} (${venue})`);

    // リアルタイム購読で初回ロード＆他スタッフ変更を即時同期
    if (cloud.adminLiveUnsub) cloud.adminLiveUnsub();
    cloud.adminLiveUnsub = cloud.db.collection('live_events').doc(docId).onSnapshot((snap) => {
      if (!snap.exists) return;
      const d = snap.data();
      cloud.suppressRemoteEcho = true;
      try {
        if (typeof window.hydrateStaffSingleEvent === 'function') {
          window.hydrateStaffSingleEvent({
            tournamentName,
            eventName: d.name || eventName,
            venue: d.venue || venue,
            matchFormat: d.matchFormat || 7,
            bracketSize: d.bracketSize || 16,
            slots: JSON.parse(d.slotsJson || '[]'),
            entries: JSON.parse(d.entriesJson || '[]'),
            activeSlotMask: d.activeSlotMask || [],
            schoolAliases: d.schoolAliases || {},
            results: d.results || {},
            consolation: d.consolation || null,
            oop: d.oop || OopEngine.createDefaultOopState()
          });
        }
      } finally {
        cloud.suppressRemoteEcho = false;
      }
    });
  }

  async function pushStaffLiveUpdate() {
    if (!cloud.isStaffMode || !cloud.staffSession || !cloud.staffSession.authenticated) return;
    updateCloudStatusUi('syncing', '本部へ送信中...');
    try {
      const nowIso = new Date().toISOString();
      await cloud.db.collection('live_events').doc(cloud.staffSession.docId).update({
        results: state.results || {},
        consolation: state.consolation || { preset: 'none', title: '順位決定戦', matchCount: 0, slots: [], results: {} },
        oop: state.oop || OopEngine.createDefaultOopState(),
        updatedAt: nowIso
      });
      updateCloudStatusUi('synced', `同期済 (${nowIso.slice(11, 16)})`);
    } catch (err) {
      updateCloudStatusUi('error', '送信エラー');
      showToast('進行状況の送信に失敗しました: ' + (err.message || err), 'error');
    }
  }

  // URLハッシュ解析（#public?t=... / #staff?t=...&e=...）
  function checkSpecialLaunchHash() {
    const hash = window.location.hash || '';
    if (hash.startsWith('#public')) {
      const qIdx = hash.indexOf('?');
      const params = new URLSearchParams(qIdx >= 0 ? hash.slice(qIdx + 1) : '');
      const tid = params.get('t');
      if (tid) {
        startPublicStandaloneMode(tid);
        return true;
      }
    }
    if (hash.startsWith('#staff')) {
      const qIdx = hash.indexOf('?');
      const params = new URLSearchParams(qIdx >= 0 ? hash.slice(qIdx + 1) : '');
      const tid = params.get('t');
      const evId = params.get('e');
      if (tid && evId) {
        startStaffPinMode(tid, evId);
        return true;
      }
    }
    return false;
  }

  // グローバルAPI公開
  window.BracketCloud = {
    isLoggedIn: () => Boolean(cloud.user),
    isSynced: () => Boolean(cloud.user && cloud.lastSyncedAt),
    isStaffMode: () => Boolean(cloud.isStaffMode),
    isPublicStandalone: () => Boolean(cloud.isPublicStandalone),
    getUser: () => cloud.user,
    scheduleCloudSync,
    syncActiveBundleToCloud,
    deleteStaffKeyFromCloud,
    attachAdminLiveListener,
    subscribePublicEvent
  };

  window.signInWithGoogle = signInWithGoogle;
  window.signOutGoogle = signOutGoogle;
  window.openMyTournamentsModal = openMyTournamentsModal;
  window.closeMyTournamentsModal = closeMyTournamentsModal;
  window.loadCloudTournament = loadCloudTournament;
  window.createNewCloudTournament = createNewCloudTournament;
  window.deleteCloudTournament = deleteCloudTournament;
  window.syncNowToCloud = () => syncActiveBundleToCloud(true);
  window.submitStaffPinLogin = submitStaffPinLogin;

  initFirebase();
  checkSpecialLaunchHash();
})();

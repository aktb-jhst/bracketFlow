/**
 * BracketFlow — Order of Play (進行表 / OOP) & Operations Engine
 * Browser / Node.js UMD dual-compatible pure logic module.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.OopEngine = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const DEFAULT_OOP_STATE = {
    venue: '第1会場',
    courtCount: 4,
    courtNames: [],
    assignMode: 'block',   // 'block' | 'round'
    umpireRule: 'loser',   // 'loser' | 'none'
    viewStyle: 'board',    // 'board' | 'table'
    courts: {},            // { "1": [matchKey, ...], ... }
    statuses: {},          // { [matchKey]: 'waiting' | 'calling' | 'playing' | 'finished' | 'wo' }
    umpires: {},           // { [matchKey]: 'カスタム審判メモ' }
    updatedAt: null
  };

  function createDefaultOopState(overrides) {
    return Object.assign(
      {},
      DEFAULT_OOP_STATE,
      {
        courtNames: [],
        courts: {},
        statuses: {},
        umpires: {},
        updatedAt: new Date().toISOString()
      },
      overrides || {}
    );
  }

  function getRoundLabel(roundIdx, totalRounds) {
    const remaining = Math.pow(2, totalRounds - roundIdx + 1);
    if (remaining === 2) return '決勝';
    if (remaining === 4) return '準決勝';
    if (remaining === 8) return '準々決勝';
    return `${roundIdx}回戦`;
  }

  /**
   * 枠インデックスとドローサイズから4つの山（A〜D）を判定
   * 左上=A, 左下=D, 右上=C, 右下=B（BracketEngine の配置規則と一致）
   */
  function getMountainCode(slotIdx, bracketSize) {
    if (!bracketSize || bracketSize < 4) return 'A';
    const q = Math.floor((slotIdx * 4) / bracketSize);
    if (q === 0) return 'A';
    if (q === 1) return 'D';
    if (q === 2) return 'C';
    return 'B';
  }

  /**
   * 選手名文字列から名字（姓）のみを抽出
   * 例: "佐藤 翔太" -> "佐藤", "佐藤翔太" -> 先頭2文字(空白無し時), "佐藤・鈴木" -> "佐藤・鈴木"
   */
  function extractSurname(rawName) {
    const s = String(rawName || '').trim();
    if (!s) return '';
    // 半角・全角スペースで区切られている場合は先頭を姓とみなす
    const parts = s.split(/[\s\u3000]+/).filter(Boolean);
    if (parts.length >= 2) return parts[0];
    // スペースなし＆4文字以上の日本語氏名なら先頭2文字を姓の近似とする（3文字以下はそのまま）
    if (s.length >= 4 && /^[\u4E00-\u9FFF\u3040-\u309F\u30A0-\u30FF]+$/.test(s)) {
      return s.slice(0, 2);
    }
    return s;
  }

  /**
   * 生徒・保護者向け画面用のペア名フォーマット
   * mode: 'full' | 'surname'
   */
  function formatSpectatorPairName(entry, mode, schoolShortFormatter) {
    if (!entry) return '未定';
    const p1 = mode === 'surname' ? extractSurname(entry.p1) : String(entry.p1 || '').trim();
    const p2 = mode === 'surname' ? extractSurname(entry.p2) : String(entry.p2 || '').trim();
    const pairText = p2 ? `${p1}・${p2}` : p1;
    if (!schoolShortFormatter) return pairText;
    const school = schoolShortFormatter(entry.school);
    return school ? `${pairText}（${school}）` : pairText;
  }

  /**
   * トーナメントツリーと順位決定戦から全試合カタログを生成
   * BYE（片側欠け）は試合として生成せず、自動的に勝ち上がりを解決する。
   */
  function buildMatchCatalog(params) {
    const bracketSize = Math.max(2, Number(params.bracketSize) || 16);
    const totalRounds = Math.round(Math.log2(bracketSize));
    const slots = Array.isArray(params.slots) ? params.slots : [];
    const activeSlotMask = Array.isArray(params.activeSlotMask) ? params.activeSlotMask : [];
    const results = params.results || {};
    const consolation = params.consolation || null;
    const labelFormatter = typeof params.labelFormatter === 'function'
      ? params.labelFormatter
      : (entry) => (entry ? (entry.p2 ? `${entry.p1}・${entry.p2}` : entry.p1) : '未定');

    const tree = [];
    tree[0] = [];
    for (let i = 0; i < bracketSize; i++) {
      tree[0][i] = {
        slotIdx: i,
        minSlot: i,
        maxSlot: i,
        active: Boolean(activeSlotMask[i]),
        entry: slots[i] || null,
        sourceMatchKey: null
      };
    }

    const matches = [];
    const matchMap = {};
    let seqNo = 1;

    for (let r = 1; r <= totalRounds; r++) {
      const count = bracketSize / Math.pow(2, r);
      tree[r] = [];
      for (let m = 0; m < count; m++) {
        const c0 = tree[r - 1][m * 2];
        const c1 = tree[r - 1][m * 2 + 1];
        const active = c0.active || c1.active;
        const isBye = (c0.active && !c1.active) || (!c0.active && c1.active);
        const matchKey = `${r}_${m}`;
        const res = results[matchKey] || null;

        let advEntry = null;
        let loserEntry = null;
        if (isBye) {
          advEntry = c0.active ? c0.entry : c1.entry;
        } else if (res) {
          advEntry = res.winner === 0 ? c0.entry : c1.entry;
          loserEntry = res.winner === 0 ? c1.entry : c0.entry;
        }

        const node = {
          round: r,
          match: m,
          minSlot: c0.minSlot,
          maxSlot: c1.maxSlot,
          active,
          isBye,
          c0,
          c1,
          result: res,
          entry: advEntry,
          loserEntry,
          sourceMatchKey: isBye ? (c0.active ? c0.sourceMatchKey : c1.sourceMatchKey) : (active ? matchKey : null)
        };
        tree[r][m] = node;

        if (!active || isBye) continue;

        const m0 = getMountainCode(c0.minSlot, bracketSize);
        const m1 = getMountainCode(c1.maxSlot, bracketSize);
        let mountain = m0 === m1 ? m0 : (r === totalRounds ? '決勝' : (m0 === 'A' || m0 === 'D' ? '左山' : '右山'));

        const p1Entry = c0.entry || null;
        const p2Entry = c1.entry || null;
        const p1Label = p1Entry
          ? labelFormatter(p1Entry)
          : (c0.sourceMatchKey && matchMap[c0.sourceMatchKey] ? `${matchMap[c0.sourceMatchKey].matchNo}の勝者` : '未定');
        const p2Label = p2Entry
          ? labelFormatter(p2Entry)
          : (c1.sourceMatchKey && matchMap[c1.sourceMatchKey] ? `${matchMap[c1.sourceMatchKey].matchNo}の勝者` : '未定');

        const isReady = Boolean(p1Entry && p2Entry);
        const matchObj = {
          key: matchKey,
          matchNo: `M${seqNo++}`,
          seqOrder: seqNo - 1,
          isConsolation: false,
          round: r,
          totalRounds,
          roundLabel: getRoundLabel(r, totalRounds),
          matchIndex: m,
          mountain,
          minSlot: c0.minSlot,
          p1Entry,
          p2Entry,
          p1Label,
          p2Label,
          sourceKey0: c0.sourceMatchKey,
          sourceKey1: c1.sourceMatchKey,
          result: res,
          winnerEntry: advEntry,
          loserEntry,
          isReady,
          isFinished: Boolean(res),
          isWo: Boolean(res && res.wo)
        };
        matches.push(matchObj);
        matchMap[matchKey] = matchObj;
      }
    }

    // 順位決定戦を進行表キューに統合
    if (consolation && consolation.preset && consolation.preset !== 'none') {
      const count = Math.max(0, parseInt(consolation.matchCount, 10) || 0);
      const conSlots = Array.isArray(consolation.slots) ? consolation.slots : [];
      const conResults = consolation.results || {};
      const conTitle = consolation.title || '順位決定戦';
      for (let i = 0; i < count; i++) {
        const key = `con_${i}`;
        const raw1 = String(conSlots[i * 2] || '').trim();
        const raw2 = String(conSlots[i * 2 + 1] || '').trim();
        const res = conResults[String(i)] || conResults[i] || null;
        const p1Label = raw1 || '未定';
        const p2Label = raw2 || '未定';
        const isReady = Boolean(raw1 && raw2);
        const matchObj = {
          key,
          consolationIndex: i,
          matchNo: `順位${i + 1}`,
          seqOrder: 10000 + i,
          isConsolation: true,
          round: totalRounds + 1,
          totalRounds,
          roundLabel: count > 1 ? `${conTitle} #${i + 1}` : conTitle,
          matchIndex: i,
          mountain: '順位戦',
          minSlot: bracketSize + i,
          p1Entry: raw1 ? { id: `con_p1_${i}`, p1: raw1, p2: '', school: '' } : null,
          p2Entry: raw2 ? { id: `con_p2_${i}`, p1: raw2, p2: '', school: '' } : null,
          p1Label,
          p2Label,
          sourceKey0: null,
          sourceKey1: null,
          result: res,
          winnerEntry: res ? { p1: res.winner === 0 ? p1Label : p2Label, p2: '', school: '' } : null,
          loserEntry: res ? { p1: res.winner === 0 ? p2Label : p1Label, p2: '', school: '' } : null,
          isReady,
          isFinished: Boolean(res),
          isWo: Boolean(res && res.wo)
        };
        matches.push(matchObj);
        matchMap[key] = matchObj;
      }
    }

    return { matches, matchMap, tree };
  }

  /**
   * 試合カタログに対してハイブリッド進行表（コート別割当）を自動生成
   * mode: 'block' (山・ブロック優先) | 'round' (ラウンド順均等)
   * preserveActive: true の場合、calling/playing/finished/wo の試合は現在のコートに固定して再配分
   */
  function generateCourtAssignments(matches, courtCount, mode, existingOop, preserveActive) {
    const nCourts = Math.max(1, Math.min(24, parseInt(courtCount, 10) || 4));
    const courts = {};
    for (let c = 1; c <= nCourts; c++) courts[String(c)] = [];

    const statuses = (existingOop && existingOop.statuses) ? Object.assign({}, existingOop.statuses) : {};
    const pinnedSet = new Set();

    if (preserveActive && existingOop && existingOop.courts) {
      const validKeys = new Set(matches.map((m) => m.key));
      Object.keys(existingOop.courts).forEach((cKey) => {
        const cNum = Math.min(nCourts, Math.max(1, parseInt(cKey, 10) || 1));
        const list = Array.isArray(existingOop.courts[cKey]) ? existingOop.courts[cKey] : [];
        list.forEach((mKey) => {
          if (!validKeys.has(mKey) || pinnedSet.has(mKey)) return;
          const st = statuses[mKey] || 'waiting';
          if (st === 'calling' || st === 'playing' || st === 'finished' || st === 'wo') {
            courts[String(cNum)].push(mKey);
            pinnedSet.add(mKey);
          }
        });
      });
    }

    const remaining = matches.filter((m) => !pinnedSet.has(m.key));

    if (mode === 'block') {
      // 山の並び: A -> D -> C -> B -> 左山 -> 右山 -> 決勝 -> 順位戦
      const mountainRank = { A: 0, D: 1, C: 2, B: 3, 左山: 4, 右山: 5, 決勝: 6, 順位戦: 7 };
      const sorted = remaining.slice().sort((a, b) => {
        if (a.round !== b.round) return a.round - b.round;
        const ma = mountainRank[a.mountain] ?? 9;
        const mb = mountainRank[b.mountain] ?? 9;
        if (ma !== mb) return ma - mb;
        return a.minSlot - b.minSlot;
      });

      // 山ごとに担当コート帯を決定（4コート以上なら A/D/C/B を均等分割）
      const mountains = ['A', 'D', 'C', 'B'];
      const mountainCourts = {};
      if (nCourts >= 4) {
        for (let i = 0; i < 4; i++) {
          const start = Math.floor((i * nCourts) / 4) + 1;
          const end = Math.floor(((i + 1) * nCourts) / 4);
          mountainCourts[mountains[i]] = [];
          for (let c = start; c <= Math.max(start, end); c++) mountainCourts[mountains[i]].push(String(c));
        }
      } else {
        mountains.forEach((mt, idx) => {
          mountainCourts[mt] = [String((idx % nCourts) + 1)];
        });
      }

      const pickLeastLoaded = (candidateCourts) => {
        let best = candidateCourts[0];
        let bestLen = courts[best].length;
        for (let i = 1; i < candidateCourts.length; i++) {
          const c = candidateCourts[i];
          if (courts[c].length < bestLen) {
            best = c;
            bestLen = courts[c].length;
          }
        }
        return best;
      };

      const allCourtKeys = Object.keys(courts);
      sorted.forEach((m) => {
        let targetCandidates = allCourtKeys;
        if (mountainCourts[m.mountain]) {
          targetCandidates = mountainCourts[m.mountain];
        } else if (m.mountain === '左山') {
          targetCandidates = [...mountainCourts.A, ...mountainCourts.D];
        } else if (m.mountain === '右山') {
          targetCandidates = [...mountainCourts.C, ...mountainCourts.B];
        } else if (m.mountain === '決勝') {
          targetCandidates = ['1'];
        }
        const chosen = pickLeastLoaded(targetCandidates);
        courts[chosen].push(m.key);
      });
    } else {
      // ラウンド順均等割当
      const sorted = remaining.slice().sort((a, b) => {
        if (a.round !== b.round) return a.round - b.round;
        return a.minSlot - b.minSlot;
      });
      const allCourtKeys = Object.keys(courts);
      sorted.forEach((m) => {
        let best = allCourtKeys[0];
        let bestLen = courts[best].length;
        for (let i = 1; i < allCourtKeys.length; i++) {
          const c = allCourtKeys[i];
          if (courts[c].length < bestLen) {
            best = c;
            bestLen = courts[c].length;
          }
        }
        courts[best].push(m.key);
      });
    }

    return courts;
  }

  /**
   * 既存のOOP状態を最新の試合カタログと整合させる
   * - 存在しなくなった試合キーを削除
   * - トーナメント結果に応じて finished / wo ステータスを同期
   * - 未割当リスト（unassigned）を返す
   */
  function reconcileOopState(oop, matchCatalog) {
    const next = createDefaultOopState(oop);
    const nCourts = Math.max(1, Math.min(24, parseInt(next.courtCount, 10) || 4));
    next.courtCount = nCourts;

    const validMap = matchCatalog.matchMap || {};
    const validKeys = new Set(Object.keys(validMap));

    // コート配列の正規化
    const cleanCourts = {};
    const assigned = new Set();
    for (let c = 1; c <= nCourts; c++) {
      const cKey = String(c);
      const rawList = Array.isArray(next.courts[cKey]) ? next.courts[cKey] : [];
      cleanCourts[cKey] = [];
      rawList.forEach((mKey) => {
        if (validKeys.has(mKey) && !assigned.has(mKey)) {
          cleanCourts[cKey].push(mKey);
          assigned.add(mKey);
        }
      });
    }
    // もしコート数を減らした直後で範囲外コートに試合が残っていたら最小コートへ退避
    Object.keys(next.courts || {}).forEach((cKey) => {
      const cNum = parseInt(cKey, 10);
      if (cNum > nCourts && Array.isArray(next.courts[cKey])) {
        next.courts[cKey].forEach((mKey) => {
          if (validKeys.has(mKey) && !assigned.has(mKey)) {
            let best = '1';
            for (let c = 2; c <= nCourts; c++) {
              if (cleanCourts[String(c)].length < cleanCourts[best].length) best = String(c);
            }
            cleanCourts[best].push(mKey);
            assigned.add(mKey);
          }
        });
      }
    });
    next.courts = cleanCourts;

    // ステータス同期
    const cleanStatuses = {};
    matchCatalog.matches.forEach((m) => {
      const prev = next.statuses[m.key];
      if (m.isFinished) {
        cleanStatuses[m.key] = m.isWo ? 'wo' : 'finished';
      } else if (prev === 'calling' || prev === 'playing') {
        cleanStatuses[m.key] = prev;
      } else {
        cleanStatuses[m.key] = 'waiting';
      }
    });
    next.statuses = cleanStatuses;

    // 手動審判メモの整理
    const cleanUmpires = {};
    Object.keys(next.umpires || {}).forEach((mKey) => {
      if (validKeys.has(mKey) && String(next.umpires[mKey] || '').trim() !== '') {
        cleanUmpires[mKey] = String(next.umpires[mKey]).trim();
      }
    });
    next.umpires = cleanUmpires;

    // 初回（どのコートにも1試合も割り当てられていない場合）は自動で初期進行表を生成
    if (assigned.size === 0 && matchCatalog.matches.length > 0) {
      next.courts = generateCourtAssignments(matchCatalog.matches, nCourts, next.assignMode, next, false);
      assigned.clear();
      Object.values(next.courts).forEach((list) => list.forEach((k) => assigned.add(k)));
    }

    const unassigned = matchCatalog.matches
      .filter((m) => !assigned.has(m.key))
      .map((m) => m.key);

    return { oop: next, unassigned };
  }

  /**
   * 審判割当メモを計算（手動上書き優先、未設定時は敗者審判ルールで自動生成）
   */
  function getUmpireLabel(courtMatchKeys, idxInCourt, matchMap, oop, labelFormatter) {
    const mKey = courtMatchKeys[idxInCourt];
    if (oop && oop.umpires && oop.umpires[mKey]) {
      return { text: oop.umpires[mKey], isCustom: true };
    }
    if (!oop || oop.umpireRule === 'none') {
      return { text: '', isCustom: false };
    }
    // 敗者審判ルール
    if (idxInCourt === 0) {
      const nextKey = courtMatchKeys[1];
      const nextM = nextKey ? matchMap[nextKey] : null;
      if (nextM) {
        return { text: `初期審判（${nextM.matchNo}待機ペア）`, isCustom: false };
      }
      return { text: '初期審判（本部指定）', isCustom: false };
    }
    const prevKey = courtMatchKeys[idxInCourt - 1];
    const prevM = prevKey ? matchMap[prevKey] : null;
    if (!prevM) return { text: '前試合敗者', isCustom: false };
    if (prevM.isFinished && prevM.loserEntry) {
      const loserText = typeof labelFormatter === 'function'
        ? labelFormatter(prevM.loserEntry)
        : (prevM.loserEntry.p2 ? `${prevM.loserEntry.p1}・${prevM.loserEntry.p2}` : prevM.loserEntry.p1);
      return { text: `敗者審判: ${loserText}`, isCustom: false, resolved: true };
    }
    return { text: `${prevM.matchNo}の敗者`, isCustom: false };
  }

  /**
   * コート内の進行状況サマリー（現在試合中/コール中・ネクスト・オンデック）を取得
   */
  function getCourtLiveSlots(courtMatchKeys, matchMap, statuses) {
    const keys = Array.isArray(courtMatchKeys) ? courtMatchKeys : [];
    let current = null;
    const upcoming = [];
    const finished = [];

    keys.forEach((k, idx) => {
      const m = matchMap[k];
      if (!m) return;
      const st = statuses[k] || 'waiting';
      const item = { key: k, match: m, status: st, indexInCourt: idx };
      if (st === 'finished' || st === 'wo') {
        finished.push(item);
      } else if ((st === 'playing' || st === 'calling') && !current) {
        current = item;
      } else {
        upcoming.push(item);
      }
    });

    return {
      current,
      next: upcoming[0] || null,
      onDeck: upcoming[1] || null,
      upcoming,
      finished
    };
  }

  /**
   * 試合を指定コートの任意の位置（試合と試合の間 insertIndex）へ正確に挿入移動
   * 同一コート内での後方移動時のインデックスずれも正しく補正する
   */
  function moveMatchInCourts(courts, matchKey, targetCourtKey, insertIndex) {
    const next = {};
    let sourceCourtKey = null;
    let sourceIdx = -1;

    Object.keys(courts || {}).forEach((cKey) => {
      const list = Array.isArray(courts[cKey]) ? courts[cKey].slice() : [];
      const idx = list.indexOf(matchKey);
      if (idx >= 0) {
        sourceCourtKey = cKey;
        sourceIdx = idx;
        list.splice(idx, 1);
      }
      next[cKey] = list;
    });

    const tKey = String(targetCourtKey);
    if (!Array.isArray(next[tKey])) next[tKey] = [];
    const targetList = next[tKey];

    let finalIdx = (typeof insertIndex === 'number' && Number.isFinite(insertIndex))
      ? Math.floor(insertIndex)
      : targetList.length;

    if (sourceCourtKey === tKey && sourceIdx >= 0 && sourceIdx < finalIdx) {
      finalIdx -= 1;
    }
    finalIdx = Math.max(0, Math.min(targetList.length, finalIdx));
    targetList.splice(finalIdx, 0, matchKey);

    return { courts: next, insertedIndex: finalIdx };
  }

  return {
    DEFAULT_OOP_STATE,
    createDefaultOopState,
    getRoundLabel,
    getMountainCode,
    extractSurname,
    formatSpectatorPairName,
    buildMatchCatalog,
    generateCourtAssignments,
    reconcileOopState,
    getUmpireLabel,
    getCourtLiveSlots,
    moveMatchInCourts
  };
});

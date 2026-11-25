/**
 * ソフトテニス個人戦 公式足長トーナメント 配置・検証エンジン
 * 優先順位ベース（枠優先順位 1..256 × 学校優先順位 × 上位4番手規則性 × 5番手以下下位枠配置）
 * ブラウザ（window.BracketEngine）および Node.js（module.exports）両対応
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.BracketEngine = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const DEFAULT_RULES = {
    enforceTop4Pattern: true,       // 各校1〜4番手の4山＆in/out規則性（A-B-C-D等 & in-out-out-in等）
    avoidTop4UnderSeed: true,       // 各校1〜4番手を1〜8シードの下（1〜2回戦で当たる枠）に入れない
    lowRankFromBottomSlots: true,   // 5番手以下は最下位番手から1〜8シードの下→低優先枠の順に配置
    balanceLowerMountainAndSide: true, // 5番手以降も学校全体でA〜Dの4山・in/outのバランスを保つ
    spreadUnder8To8Blocks: true,    // 8ペア以下の学校は8ブロック（A1〜B1）に重複なく分散
    avoidRound1SameSchool: true,    // 1回戦の同校対決を最優先回避
    avoidTripleSameCard: true,      // 1回戦の同一カード3対戦以上を回避
    avoidRound2SameSchool: true     // 2回戦（4枠内）の同校接近を回避
  };

  // 1番手の山に応じた上位4番手の山順パターン
  const MOUNTAIN_PATTERNS = {
    A: ['A', 'B', 'C', 'D'],
    B: ['B', 'A', 'D', 'C'],
    C: ['C', 'D', 'A', 'B'],
    D: ['D', 'C', 'B', 'A']
  };

  // 1番手の内外に応じた上位4番手の内外順パターン
  const SIDE_PATTERNS = {
    in:  ['in', 'out', 'out', 'in'],
    out: ['out', 'in', 'in', 'out']
  };

  const BLOCK_CODE_MAP = {
    'A:out': 'A1',
    'A:in':  'A2',
    'D:in':  'D2',
    'D:out': 'D1',
    'C:out': 'C1',
    'C:in':  'C2',
    'B:in':  'B2',
    'B:out': 'B1'
  };

  function getTop4ExpectedBlocks(anchorMountain, anchorSide) {
    const mSeq = MOUNTAIN_PATTERNS[anchorMountain];
    const sSeq = SIDE_PATTERNS[anchorSide];
    if (!mSeq || !sSeq) return null;
    return [0, 1, 2, 3].map((idx) => {
      const mountain = mSeq[idx];
      const side = sSeq[idx];
      const code = BLOCK_CODE_MAP[`${mountain}:${side}`];
      return { rank: idx + 1, mountain, side, code };
    });
  }

  /**
   * 同一学校内の複数シード特例を含む、上位4番手（rank 1〜4）の目標ブロック計画を算出
   * ・シードが0〜1組：1番手基準の標準規則（A-B-C-D等 × in-out-out-in / out-in-in-out）
   * ・シードが2組以上：シードが占有している山以外の山へ残りの上位番手を振り分け、
   *   かつ in / out のバランスが均等（2:2）になるように割り当てる
   */
  function computeSchoolTop4Plan(school, slots, size, schoolTopRanksSet) {
    if (size < 8) return null;

    // 盤面上の当該校の上位4番手シードおよび1番手アンカーを収集
    const seededTop4 = [];
    let anchorInfo = null;
    let anchorEntry = null;
    let anchorSlotIdx = -1;

    for (let i = 0; i < size; i++) {
      const s = slots[i];
      if (!s || s.school !== school) continue;
      if (s.schoolRank === 1 && !anchorInfo) {
        anchorInfo = getBlockInfo(i, size);
        anchorEntry = s;
        anchorSlotIdx = i;
      }
      if (s.seed && s.schoolRank >= 1 && s.schoolRank <= 4) {
        seededTop4.push({ slotIdx: i, entry: s, info: getBlockInfo(i, size) });
      }
    }

    if (!anchorInfo && seededTop4.length < 2) return null;

    const baseMountain = anchorInfo ? anchorInfo.mountain : seededTop4[0].info.mountain;
    const baseSide = anchorInfo ? anchorInfo.side : seededTop4[0].info.side;
    const standardExpected = getTop4ExpectedBlocks(baseMountain, baseSide);
    const byRank = {};

    // 標準モード（同一校内の上位4番手シードが0〜1組）
    if (seededTop4.length < 2) {
      standardExpected.forEach((b) => { byRank[b.rank] = b; });
      return {
        isMultiSeedException: false,
        anchorInfo,
        anchorEntry,
        anchorSlotIdx,
        byRank
      };
    }

    // 特例モード（同一校内に上位4番手シードが2組以上）
    const usedMountains = new Set();
    let curIn = 0;
    let curOut = 0;
    const seededRanks = new Set();

    seededTop4.forEach((item) => {
      seededRanks.add(item.entry.schoolRank);
      usedMountains.add(item.info.mountain);
      if (item.info.side === 'in') curIn++;
      else curOut++;
      byRank[item.entry.schoolRank] = {
        rank: item.entry.schoolRank,
        mountain: item.info.mountain,
        side: item.info.side,
        code: item.info.code,
        isSeedFixed: true
      };
    });

    // シードが占有していない山を、1番手基準の順序で抽出
    const mOrder = MOUNTAIN_PATTERNS[baseMountain] || ['A', 'B', 'C', 'D'];
    const remainingMountains = mOrder.filter((m) => !usedMountains.has(m));
    // もし複数シードが同じ山に重複していて残りの山が足りない場合は、使用回数が少ない山を後ろに補完
    if (remainingMountains.length < 4 - seededRanks.size) {
      const mUseCnt = { A: 0, B: 0, C: 0, D: 0 };
      seededTop4.forEach((item) => { mUseCnt[item.info.mountain]++; });
      const extra = mOrder
        .filter((m) => usedMountains.has(m))
        .sort((a, b) => mUseCnt[a] - mUseCnt[b]);
      remainingMountains.push(...extra);
    }

    // 残りの非シード上位番手（1〜4のうちシード以外）に、空いている山＋in/outバランスで割り当て
    const unseededRanks = [1, 2, 3, 4].filter(
      (r) => !seededRanks.has(r) && (!schoolTopRanksSet || schoolTopRanksSet.has(r))
    );
    let mCursor = 0;
    unseededRanks.forEach((r) => {
      const mountain = remainingMountains[mCursor % remainingMountains.length];
      mCursor++;
      let side;
      if (curIn < curOut) {
        side = 'in';
      } else if (curOut < curIn) {
        side = 'out';
      } else {
        side = SIDE_PATTERNS[baseSide][r - 1];
      }
      if (side === 'in') curIn++;
      else curOut++;
      byRank[r] = {
        rank: r,
        mountain,
        side,
        code: BLOCK_CODE_MAP[`${mountain}:${side}`],
        isSeedFixed: false
      };
    });

    return {
      isMultiSeedException: true,
      anchorInfo,
      anchorEntry,
      anchorSlotIdx,
      seededMountains: Array.from(usedMountains),
      byRank
    };
  }

  function nextPowerOfTwo(n) {
    let p = 4;
    while (p < n) p *= 2;
    return Math.min(256, p);
  }

  function shuffleArray(arr, rng) {
    const rand = rng || Math.random;
    const copy = arr.slice();
    for (let i = copy.length - 1; i > 0; i--) {
      const j = Math.floor(rand() * (i + 1));
      const tmp = copy[i];
      copy[i] = copy[j];
      copy[j] = tmp;
    }
    return copy;
  }

  /**
   * 標準1位〜256位のシード優先順位マップを生成
   * 返り値: seedOrder[priorityRank - 1] = slotIndex
   * 例 (size=16): [0, 15, 8, 7, 4, 11, 12, 3, 2, 13, 10, 5, 6, 9, 14, 1]
   */
  function buildBracketOrder(size) {
    let seeds = [0, 1];
    let currentSize = 2;
    while (currentSize < size) {
      currentSize *= 2;
      const nextSeeds = [];
      for (let i = 0; i < seeds.length; i++) {
        const s = seeds[i];
        const comp = currentSize - 1 - s;
        if (i % 2 === 0) {
          nextSeeds.push(s);
          nextSeeds.push(comp);
        } else {
          nextSeeds.push(comp);
          nextSeeds.push(s);
        }
      }
      seeds = nextSeeds;
    }
    return seeds;
  }

  /**
   * 各スロットの優先順位（1 = 最上位 1位枠 〜 size = 最下位枠）を返す配列
   */
  function buildSlotPriorityMap(size) {
    const seedOrder = buildBracketOrder(size);
    const slotPriority = new Array(size);
    seedOrder.forEach((slotIdx, idx) => {
      slotPriority[slotIdx] = idx + 1; // 1-based
    });
    return slotPriority;
  }

  /**
   * 「1〜8シードの下」マップを構築する
   * 返り値: underMap[slotIdx] = null または { seedRank: 1..8, type: 'R1' | 'R2' }
   * ・type 'R1' : 1回戦で直接1〜8シードと当たる枠（1〜8シードが足長でない場合）
   * ・type 'R2' : 1回戦を勝ち上がったら2回戦で1〜8シードと当たる枠（シードの下）
   * ※ 1〜8シード自身の枠（slotPriority <= topSeedLimit）は対象外
   */
  function buildUnderTop8SeedMap(size, mask, slotPriority) {
    const underMap = new Array(size).fill(null);
    if (size < 8) return underMap;
    const seedOrder = buildBracketOrder(size);
    const pMap = slotPriority || buildSlotPriorityMap(size);
    const topSeedLimit = Math.min(8, size / 2);

    // 1. まず1回戦直下（R1）を登録（S1→S8の順）
    for (let k = 1; k <= topSeedLimit; k++) {
      const seedSlot = seedOrder[k - 1];
      const oppR1 = (seedSlot % 2 === 0) ? seedSlot + 1 : seedSlot - 1;
      if (mask[oppR1] && pMap[oppR1] > topSeedLimit && !underMap[oppR1]) {
        underMap[oppR1] = { seedRank: k, type: 'R1' };
      }
    }

    // 2. 続いて「1回戦勝ち上がりで2回戦に1〜8シードと当たる枠（R2）」を登録（S1→S8の順）
    if (size >= 16) {
      for (let k = 1; k <= topSeedLimit; k++) {
        const seedSlot = seedOrder[k - 1];
        const qBase = Math.floor(seedSlot / 4) * 4;
        const otherPairBase = (seedSlot - qBase < 2) ? (qBase + 2) : qBase;
        [otherPairBase, otherPairBase + 1].forEach((idx) => {
          if (mask[idx] && pMap[idx] > topSeedLimit && !underMap[idx]) {
            underMap[idx] = { seedRank: k, type: 'R2' };
          }
        });
      }
    }

    return underMap;
  }

  /**
   * 調整済み枠優先順位マップ（1 〜 有効枠数N）を構築する
   * ・1〜8シードの下が最下位側になるよう再配列：
   *   - 1シードの下（2枠）＝ 最下位(N), 最下位の次(N-1)
   *   - 2シードの下（2枠）＝ 最下位から3番目(N-2), 4番目(N-3)
   *   - … 8シードの下（2枠）＝ N-14, N-15
   *   ※ 満枠等で1〜8シードの1回戦直下(R1)も存在する場合は、R1直下(S1→S8)を最下位、その上にR2(S1→S8)を配置
   * ・それ以外のアクティブ枠は標準シード順のまま 1 〜 K を採番
   */
  function buildAdjustedSlotPriorityMap(size, mask) {
    const rawPriority = buildSlotPriorityMap(size);
    const activeMask = Array.isArray(mask) && mask.length === size ? mask : new Array(size).fill(true);
    const underMap = buildUnderTop8SeedMap(size, activeMask, rawPriority);

    const normalSlots = [];
    const r2UnderSlots = [];
    const r1UnderSlots = [];

    for (let i = 0; i < size; i++) {
      if (!activeMask[i]) continue;
      const u = underMap[i];
      if (!u) {
        normalSlots.push(i);
      } else if (u.type === 'R2') {
        r2UnderSlots.push(i);
      } else {
        r1UnderSlots.push(i);
      }
    }

    // 通常枠：標準シード優先順位の昇順（1位, 2位, ...）
    normalSlots.sort((a, b) => rawPriority[a] - rawPriority[b]);

    // 2回戦で1〜8シードに当たる枠：S8下 → S7下 → … → S1下 の順（結果として S1下 が最下位・その次になる）
    r2UnderSlots.sort((a, b) => {
      const ua = underMap[a];
      const ub = underMap[b];
      if (ub.seedRank !== ua.seedRank) return ub.seedRank - ua.seedRank;
      return rawPriority[a] - rawPriority[b];
    });

    // 1回戦で直接1〜8シードに当たる枠：S8直下 → … → S1直下 の順（S1直下が絶対最下位）
    r1UnderSlots.sort((a, b) => {
      const ua = underMap[a];
      const ub = underMap[b];
      if (ub.seedRank !== ua.seedRank) return ub.seedRank - ua.seedRank;
      return rawPriority[a] - rawPriority[b];
    });

    const orderedActive = [...normalSlots, ...r2UnderSlots, ...r1UnderSlots];
    const adjustedMap = new Array(size).fill(null);
    orderedActive.forEach((slotIdx, idx) => {
      adjustedMap[slotIdx] = idx + 1; // 1 〜 N（連番）
    });
    return adjustedMap;
  }

  /**
   * 左右完全バランス＆公式シード優先順の有効枠（BYE足長）マスク生成
   */
  function createBalancedActiveMask(totalPairs, size) {
    if (totalPairs >= size) {
      return new Array(size).fill(true);
    }

    const mask = new Array(size).fill(false);
    const seedOrder = buildBracketOrder(size);
    const slotPriority = buildSlotPriorityMap(size);

    const numBoxes = size / 2;
    const boxPriorityOrder = [];
    const seenBox = new Set();
    for (let i = 0; i < seedOrder.length; i++) {
      const boxIdx = Math.floor(seedOrder[i] / 2);
      if (!seenBox.has(boxIdx)) {
        seenBox.add(boxIdx);
        boxPriorityOrder.push(boxIdx);
      }
    }

    const numByes = Math.min(numBoxes, Math.max(0, size - totalPairs));
    const byeBoxes = new Set(boxPriorityOrder.slice(0, numByes));

    for (let b = 0; b < numBoxes; b++) {
      const s0 = b * 2;
      const s1 = b * 2 + 1;
      if (byeBoxes.has(b)) {
        if (slotPriority[s0] <= slotPriority[s1]) {
          mask[s0] = true;
        } else {
          mask[s1] = true;
        }
      } else {
        mask[s0] = true;
        mask[s1] = true;
      }
    }

    return mask;
  }

  function getBlockInfo(slotIdx, size) {
    const octantSize = Math.max(1, size / 8);
    const oct = Math.min(7, Math.floor(slotIdx / octantSize));
    const map = [
      { code: 'A1', mountain: 'A', side: 'out', half: 'L' },
      { code: 'A2', mountain: 'A', side: 'in',  half: 'L' },
      { code: 'D2', mountain: 'D', side: 'in',  half: 'L' },
      { code: 'D1', mountain: 'D', side: 'out', half: 'L' },
      { code: 'C1', mountain: 'C', side: 'out', half: 'R' },
      { code: 'C2', mountain: 'C', side: 'in',  half: 'R' },
      { code: 'B2', mountain: 'B', side: 'in',  half: 'R' },
      { code: 'B1', mountain: 'B', side: 'out', half: 'R' }
    ];
    return map[oct];
  }

  function isByeSlot(slotIdx, mask) {
    const oppIdx = (slotIdx % 2 === 0) ? slotIdx + 1 : slotIdx - 1;
    return Boolean(mask[slotIdx]) && !Boolean(mask[oppIdx]);
  }

  function makeCardKey(schoolA, schoolB) {
    return schoolA < schoolB ? `${schoolA}||${schoolB}` : `${schoolB}||${schoolA}`;
  }

  /**
   * 学校の優先順位マップを構築（未設定の学校はシード順→組数多い順→登場順で自動補完）
   */
  function resolveSchoolPriorityMap(entries, schoolOrder) {
    const schools = [];
    const seen = new Set();
    const schoolBestSeed = {};
    const schoolCount = {};

    entries.forEach((e) => {
      const sc = e.school || '無所属';
      if (!seen.has(sc)) {
        seen.add(sc);
        schools.push(sc);
      }
      schoolCount[sc] = (schoolCount[sc] || 0) + 1;
      if (e.seed && (!schoolBestSeed[sc] || e.seed < schoolBestSeed[sc])) {
        schoolBestSeed[sc] = e.seed;
      }
    });

    const ordered = [];
    const used = new Set();

    // 1. ユーザー指定の学校優先順位配列（schoolOrder）を最優先で反映
    if (Array.isArray(schoolOrder)) {
      schoolOrder.forEach((sc) => {
        if (seen.has(sc) && !used.has(sc)) {
          ordered.push(sc);
          used.add(sc);
        }
      });
    }

    // 2. 指定に含まれていない学校はシード上位→組数多い順→登場順で補完
    const rest = schools.filter((sc) => !used.has(sc));
    rest.sort((a, b) => {
      const sa = schoolBestSeed[a] || 9999;
      const sb = schoolBestSeed[b] || 9999;
      if (sa !== sb) return sa - sb;
      if (schoolCount[b] !== schoolCount[a]) return schoolCount[b] - schoolCount[a];
      return schools.indexOf(a) - schools.indexOf(b);
    });
    ordered.push(...rest);

    const priorityMap = {};
    ordered.forEach((sc, idx) => {
      priorityMap[sc] = idx + 1; // 1 = 最優先学校
    });
    return { ordered, priorityMap };
  }

  function buildAnchorMap(slots, size) {
    const anchors = {};
    for (let i = 0; i < size; i++) {
      const s = slots[i];
      if (!s || !s.school) continue;
      if (s.schoolRank === 1 && !anchors[s.school]) {
        const info = getBlockInfo(i, size);
        anchors[s.school] = {
          slotIdx: i,
          entry: s,
          mountain: info.mountain,
          side: info.side,
          code: info.code,
          expected: getTop4ExpectedBlocks(info.mountain, info.side)
        };
      }
    }
    return anchors;
  }

  function buildRound1CardCounts(slots, size) {
    const cardMap = {};
    for (let i = 0; i < size; i += 2) {
      const p1 = slots[i];
      const p2 = slots[i + 1];
      if (p1 && p2 && p1.school && p2.school && p1.school !== p2.school) {
        const key = makeCardKey(p1.school, p2.school);
        if (!cardMap[key]) {
          cardMap[key] = {
            key,
            schoolA: p1.school < p2.school ? p1.school : p2.school,
            schoolB: p1.school < p2.school ? p2.school : p1.school,
            matches: []
          };
        }
        cardMap[key].matches.push({ slotA: i, slotB: i + 1, p1, p2 });
      }
    }
    return cardMap;
  }

  /**
   * 優先順位ベース・トーナメント自動配置アルゴリズム
   * @param {Array} entries - 出場ペア配列
   * @param {Object} userRules - 配置ルールオプション
   * @param {Object} options - { schoolOrder?: string[], rng?: Function }
   */
  function generateBracket(entries, userRules, options) {
    const opts = options || {};
    const rules = Object.assign({}, DEFAULT_RULES, userRules || {});
    const n = entries.length;
    if (n < 2) {
      return {
        bracketSize: 4,
        activeSlotMask: [false, false, false, false],
        slots: [null, null, null, null],
        slotPriorityMap: [1, 4, 3, 2],
        schoolOrder: []
      };
    }

    const size = nextPowerOfTwo(n);
    const mask = createBalancedActiveMask(n, size);
    const slots = new Array(size).fill(null);
    const seedOrder = buildBracketOrder(size);
    const rawSlotPriority = buildSlotPriorityMap(size);
    const underSeedMap = buildUnderTop8SeedMap(size, mask, rawSlotPriority);
    const slotPriority = buildAdjustedSlotPriorityMap(size, mask);
    const { ordered: resolvedSchoolOrder, priorityMap: schoolPriority } = resolveSchoolPriorityMap(entries, opts.schoolOrder);

    const placedIds = new Set();
    const schoolTotalCounts = {};
    const schoolTopRanksPresent = {};
    entries.forEach((e) => {
      const sc = e.school || '無所属';
      schoolTotalCounts[sc] = (schoolTotalCounts[sc] || 0) + 1;
      if (!schoolTopRanksPresent[sc]) schoolTopRanksPresent[sc] = new Set();
      if (e.schoolRank >= 1 && e.schoolRank <= 4) {
        schoolTopRanksPresent[sc].add(e.schoolRank);
      }
    });

    const schoolStats = {};
    const schoolAnchors = {};
    const r1CardCounts = {};

    function ensureSchoolStat(school) {
      if (!schoolStats[school]) {
        schoolStats[school] = {
          half: { L: 0, R: 0 },
          mountain: { A: 0, B: 0, C: 0, D: 0 },
          block: { A1: 0, A2: 0, D2: 0, D1: 0, C1: 0, C2: 0, B2: 0, B1: 0 },
          side: { out: 0, in: 0 }
        };
      }
      return schoolStats[school];
    }

    function placeEntry(idx, entry) {
      slots[idx] = Object.assign({}, entry);
      placedIds.add(entry.id);
      const sc = entry.school || '無所属';
      const st = ensureSchoolStat(sc);
      const info = getBlockInfo(idx, size);
      st.half[info.half]++;
      st.mountain[info.mountain]++;
      st.block[info.code]++;
      st.side[info.side]++;

      if (entry.schoolRank === 1 && !schoolAnchors[sc]) {
        schoolAnchors[sc] = {
          slotIdx: idx,
          entry: entry,
          mountain: info.mountain,
          side: info.side,
          code: info.code,
          expected: getTop4ExpectedBlocks(info.mountain, info.side)
        };
      }

      const oppIdx = (idx % 2 === 0) ? idx + 1 : idx - 1;
      const opp = slots[oppIdx];
      if (opp && opp.school && sc && opp.school !== sc) {
        const key = makeCardKey(sc, opp.school);
        r1CardCounts[key] = (r1CardCounts[key] || 0) + 1;
      }
    }

    function countOpenInBlock(blockCode) {
      let c = 0;
      for (let i = 0; i < size; i++) {
        if (mask[i] && !slots[i] && getBlockInfo(i, size).code === blockCode) c++;
      }
      return c;
    }

    function getReservedCountsByPlacedAnchors(excludeSchool) {
      const reserved = {};
      const allSchools = new Set([
        ...Object.keys(schoolAnchors),
        ...Object.keys(schoolTopRanksPresent)
      ]);
      allSchools.forEach((sc) => {
        if (sc === excludeSchool) return;
        const plan = computeSchoolTop4Plan(sc, slots, size, schoolTopRanksPresent[sc]);
        if (!plan) return;
        [1, 2, 3, 4].forEach((r) => {
          if (!schoolTopRanksPresent[sc]?.has(r)) return;
          const alreadyPlaced = slots.some((s) => s && s.school === sc && s.schoolRank === r);
          if (!alreadyPlaced && plan.byRank[r]) {
            const code = plan.byRank[r].code;
            reserved[code] = (reserved[code] || 0) + 1;
          }
        });
      });
      return reserved;
    }

    function getOpenActiveIndices() {
      const list = [];
      for (let i = 0; i < size; i++) {
        if (mask[i] && !slots[i]) list.push(i);
      }
      return list;
    }

    // ---------------------------------------------------------
    // ステップ0：シード指定ペアの固定配置
    // ---------------------------------------------------------
    const seededEntries = entries
      .filter((e) => e.seed && e.seed >= 1)
      .sort((a, b) => a.seed - b.seed);

    seededEntries.forEach((entry) => {
      const targetIdx = seedOrder[entry.seed - 1];
      if (targetIdx === undefined) return;
      const pairBase = Math.floor(targetIdx / 2) * 2;
      if (mask[targetIdx] && !slots[targetIdx]) {
        placeEntry(targetIdx, entry);
      } else if (mask[pairBase] && !slots[pairBase]) {
        placeEntry(pairBase, entry);
      } else if (mask[pairBase + 1] && !slots[pairBase + 1]) {
        placeEntry(pairBase + 1, entry);
      }
    });

    // ---------------------------------------------------------
    // ステップ1〜4：各校の1番手 → 2番手 → 3番手 → 4番手を
    // 「学校優先順位が高い順」に「各山の優先順位が高い空き枠」へ配置
    // ---------------------------------------------------------
    const useTop4 = rules.enforceTop4Pattern && size >= 8;

    for (let targetRank = 1; targetRank <= 4; targetRank++) {
      // 対象番手の未配置ペアを「学校優先順位が高い順」にソート
      const tierEntries = entries
        .filter((e) => !placedIds.has(e.id) && e.schoolRank === targetRank)
        .sort((a, b) => {
          const pa = schoolPriority[a.school || '無所属'] || 9999;
          const pb = schoolPriority[b.school || '無所属'] || 9999;
          if (pa !== pb) return pa - pb;
          return (a.serial || 0) - (b.serial || 0);
        });

      tierEntries.forEach((cand) => {
        const openIndices = getOpenActiveIndices();
        if (openIndices.length === 0) return;

        const sc = cand.school || '無所属';
        const st = ensureSchoolStat(sc);
        const reservedMap = useTop4 ? getReservedCountsByPlacedAnchors(sc) : {};
        const schoolPlan = useTop4 ? computeSchoolTop4Plan(sc, slots, size, schoolTopRanksPresent[sc]) : null;

        // 複数シード特例 or 1番手基準の目標ブロック
        let targetBlock = null;
        if (useTop4 && schoolPlan && schoolPlan.byRank[targetRank]) {
          // 1番手配置時でも、すでに2〜4番手に複数シードがある場合は特例計画を採用
          if (targetRank >= 2 || schoolPlan.isMultiSeedException) {
            targetBlock = schoolPlan.byRank[targetRank];
          }
        }

        // 1番手配置時（特例でない場合）：2〜4番手にシード固定が1つある場合は整合する1番手ブロックを逆算
        let reverseRank1Code = null;
        if (useTop4 && targetRank === 1 && !targetBlock) {
          for (let i = 0; i < size; i++) {
            const s = slots[i];
            if (s && s.school === sc && s.schoolRank >= 2 && s.schoolRank <= 4) {
              const sInfo = getBlockInfo(i, size);
              const matchedKey = Object.keys(BLOCK_CODE_MAP).find((k) => {
                const [m, sd] = k.split(':');
                return getTop4ExpectedBlocks(m, sd)[s.schoolRank - 1].code === sInfo.code;
              });
              if (matchedKey) {
                reverseRank1Code = BLOCK_CODE_MAP[matchedKey];
                break;
              }
            }
          }
        }

        let bestIdx = openIndices[0];
        let bestScore = -Infinity;

        for (let oi = 0; oi < openIndices.length; oi++) {
          const idx = openIndices[oi];
          const info = getBlockInfo(idx, size);
          const oppIdx = (idx % 2 === 0) ? idx + 1 : idx - 1;
          const opp = slots[oppIdx];
          const pRank = slotPriority[idx]; // 1(最上位) 〜 size(最下位)

          // 基本スコア：枠の優先順位が高い（pRankが小さい）ほど高得点（厳密・決定論的）
          let score = (size + 1 - pRank) * 20;

          // (1) 1回戦同校対決の絶対回避
          if (rules.avoidRound1SameSchool && opp && opp.school === sc) {
            score -= 500000;
          }

          // (2) 1回戦同一カード3対戦以上の回避
          if (rules.avoidTripleSameCard && opp && opp.school && opp.school !== sc) {
            const cCnt = r1CardCounts[makeCardKey(sc, opp.school)] || 0;
            if (cCnt >= 2) score -= 180000 * (cCnt - 1);
            else if (cCnt === 1) score -= 8; // 枠優先順を壊さない程度の微調整
          }

          // (3) 上位4番手の4山＆in/out規則性
          if (useTop4) {
            if (targetBlock) {
              if (info.code === targetBlock.code) {
                score += 200000; // 指定ブロック（山＋内外）に完全一致
              } else if (info.mountain === targetBlock.mountain) {
                score += 40000;  // 満席時のフォールバック：同じ山
              } else if (st.mountain[info.mountain] === 0) {
                score += 10000;  // まだ自校のいない山
              } else {
                score -= 50000;
              }
            } else if (targetRank === 1) {
              if (reverseRank1Code && info.code === reverseRank1Code) {
                score += 180000;
              }
              // 1番手枠自体が他校予約で埋まり切っていないか確認
              const selfFree = countOpenInBlock(info.code) - (reservedMap[info.code] || 0);
              if (selfFree <= 0) {
                score -= 90000;
              }
              // 自校の2〜4番手の予定ブロックに空きがあるか確認
              const expList = getTop4ExpectedBlocks(info.mountain, info.side);
              [2, 3, 4].forEach((r) => {
                if (schoolTopRanksPresent[sc]?.has(r)) {
                  const needCode = expList[r - 1].code;
                  const freeAfterReserve = countOpenInBlock(needCode) - (reservedMap[needCode] || 0);
                  if (freeAfterReserve <= 0) {
                    score -= 120000;
                  }
                }
              });
            }
          } else {
            // size < 8 等で8ブロック規則性を使わない場合も、4山・左右へ分散
            score -= st.mountain[info.mountain] * 30000;
            score -= st.half[info.half] * 15000;
          }

          // (4) 2回戦（4枠ブロック内）での同校接近回避
          if (rules.avoidRound2SameSchool && size >= 4) {
            const r2Base = Math.floor(idx / 4) * 4;
            let r2Same = 0;
            for (let k = r2Base; k < r2Base + 4; k++) {
              if (k === idx || k === oppIdx) continue;
              if (slots[k] && slots[k].school === sc) r2Same++;
            }
            if (r2Same > 0) score -= 35000 * r2Same;
          }

          // (5) 各校の1〜4番手を「1〜8シードの下」に入れないルール
          //     ※指定ブロック内にシード下ではない空き枠があれば必ずそちらを優先
          //     ※やむを得ず入る場合でも、1回戦直下(R1)より2回戦側(R2)、上位シード下より下位シード下を優先
          if (rules.avoidTop4UnderSeed && underSeedMap[idx]) {
            const u = underSeedMap[idx];
            if (u.type === 'R1') {
              score -= 28000 + (9 - u.seedRank) * 200;
            } else {
              score -= 18000 + (9 - u.seedRank) * 200;
            }
          }

          if (score > bestScore) {
            bestScore = score;
            bestIdx = idx;
          }
        }

        placeEntry(bestIdx, cand);
      });
    }

    // ---------------------------------------------------------
    // ステップ5：5番手以下（および残余ペア）を
    // 「最下位番手から」「優先順位が低い枠（1〜8番シードの下）から」配置
    // ・ペア数が8ペア以下（<8および8）の学校は8ブロックに重複なく分散
    // ・8ペア以上の学校も同校対戦を避けて均等分散
    // ---------------------------------------------------------
    const lowerEntries = entries
      .filter((e) => !placedIds.has(e.id))
      .sort((a, b) => {
        // 最下位番手（schoolRankが大きい順）から配置
        const ra = Number.isFinite(a.schoolRank) ? a.schoolRank : 999;
        const rb = Number.isFinite(b.schoolRank) ? b.schoolRank : 999;
        if (rules.lowRankFromBottomSlots) {
          if (rb !== ra) return rb - ra;
          // 同じ番手同士は学校優先順位が低い学校から先に低い枠（シード下）へ入れる
          const pa = schoolPriority[a.school || '無所属'] || 9999;
          const pb = schoolPriority[b.school || '無所属'] || 9999;
          if (pb !== pa) return pb - pa;
        } else {
          if (ra !== rb) return ra - rb;
          const pa = schoolPriority[a.school || '無所属'] || 9999;
          const pb = schoolPriority[b.school || '無所属'] || 9999;
          if (pa !== pb) return pa - pb;
        }
        return (a.serial || 0) - (b.serial || 0);
      });

    lowerEntries.forEach((cand) => {
      const openIndices = getOpenActiveIndices();
      if (openIndices.length === 0) return;

      const sc = cand.school || '無所属';
      const st = ensureSchoolStat(sc);
      const totalPairsOfSchool = schoolTotalCounts[sc] || 1;
      const isUnderOrEq8 = totalPairsOfSchool <= 8;

      let bestIdx = openIndices[0];
      let bestScore = -Infinity;

      for (let oi = 0; oi < openIndices.length; oi++) {
        const idx = openIndices[oi];
        const info = getBlockInfo(idx, size);
        const oppIdx = (idx % 2 === 0) ? idx + 1 : idx - 1;
        const opp = slots[oppIdx];
        const pRank = slotPriority[idx]; // 大きいほど優先順位が低い枠（1〜8番シード直下側）

        // 基本スコア：
        // ・lowRankFromBottomSlots ONの場合：
        //   ①「1〜8シードの下」（R1直下→1回戦勝上がりで2回戦に1〜8シードと当たる枠）を
        //     S1の下 → S2の下 → … → S8の下 の順で最優先
        //   ② それ以外の枠は優先順位が低い枠（pRankが大きい枠）から順に埋める
        let score;
        if (rules.lowRankFromBottomSlots) {
          score = pRank * 20;
          const u = underSeedMap[idx];
          if (u) {
            if (u.type === 'R1') {
              score += 24000 + (9 - u.seedRank) * 500;
            } else {
              score += 15000 + (9 - u.seedRank) * 500;
            }
          }
        } else {
          score = (size + 1 - pRank) * 20;
        }

        // (1) 1回戦同校対決の絶対回避（8ペア以上でも必須）
        if (rules.avoidRound1SameSchool && opp && opp.school === sc) {
          score -= 500000;
        }

        // (2) 1回戦同一カード3対戦以上の回避
        if (rules.avoidTripleSameCard && opp && opp.school && opp.school !== sc) {
          const cCnt = r1CardCounts[makeCardKey(sc, opp.school)] || 0;
          if (cCnt >= 2) score -= 180000 * (cCnt - 1);
          else if (cCnt === 1) score -= 8;
        }

        // (3) 2回戦（4枠ブロック内）での同校接近回避
        if (rules.avoidRound2SameSchool && size >= 4) {
          const r2Base = Math.floor(idx / 4) * 4;
          let r2Same = 0;
          for (let k = r2Base; k < r2Base + 4; k++) {
            if (k === idx || k === oppIdx) continue;
            if (slots[k] && slots[k].school === sc) r2Same++;
          }
          if (r2Same > 0) score -= 35000 * r2Same;
        }

        // (4) 8つの山（8ブロック：A1〜B1）への分散 ＋ 学校全体でのA〜D・in/outバランス
        //     ・8ペア以下の学校は「まだ自校が入っていないブロック」を強く優先し、必ず8ブロックに散らす
        //     ・5番手以降も学校全体で A〜D（差1以内）・in/out（差1以内）・左右のバランスを維持
        if (size >= 8) {
          const blockCnt = st.block[info.code];
          if (rules.spreadUnder8To8Blocks && isUnderOrEq8) {
            if (blockCnt > 0) {
              score -= 120000 * blockCnt;
            }
          } else {
            score -= blockCnt * 14000;
          }

          if (rules.balanceLowerMountainAndSide) {
            const minM = Math.min(st.mountain.A, st.mountain.B, st.mountain.C, st.mountain.D);
            const mExcess = st.mountain[info.mountain] - minM;
            if (mExcess > 0) {
              score -= mExcess * 26000;
            }

            const otherSide = info.side === 'in' ? 'out' : 'in';
            const sideDiff = st.side[info.side] - st.side[otherSide];
            if (sideDiff > 0) {
              score -= sideDiff * 24000;
            } else if (sideDiff < 0) {
              score += 6000;
            }

            const otherHalf = info.half === 'L' ? 'R' : 'L';
            const halfDiff = st.half[info.half] - st.half[otherHalf];
            if (halfDiff > 0) {
              score -= halfDiff * 14000;
            }
          } else {
            score -= st.mountain[info.mountain] * 3000;
            score -= st.half[info.half] * 1500;
          }
        } else {
          score -= st.mountain[info.mountain] * 15000;
          score -= st.half[info.half] * 8000;
        }

        if (score > bestScore) {
          bestScore = score;
          bestIdx = idx;
        }
      }

      placeEntry(bestIdx, cand);
    });

    // ---------------------------------------------------------
    // ステップ6：最終整合スワップ最適化
    // ---------------------------------------------------------
    optimizeBracketBySwap(slots, mask, size, rules, schoolTotalCounts, slotPriority, underSeedMap);

    return {
      bracketSize: size,
      activeSlotMask: mask,
      slots: slots,
      slotPriorityMap: slotPriority,
      underSeedMap: underSeedMap,
      schoolOrder: resolvedSchoolOrder
    };
  }

  function evaluateGlobalPenalty(slots, mask, size, rules, schoolTotalCounts, slotPriority, underSeedMap) {
    let hardPenalty = 0;

    if (rules.avoidRound1SameSchool) {
      for (let i = 0; i < size; i += 2) {
        if (slots[i] && slots[i + 1] && slots[i].school === slots[i + 1].school) {
          hardPenalty += 500000;
        }
      }
    }

    if (rules.avoidTripleSameCard) {
      const cards = buildRound1CardCounts(slots, size);
      Object.keys(cards).forEach((k) => {
        const cnt = cards[k].matches.length;
        if (cnt >= 3) hardPenalty += (cnt - 2) * 180000;
      });
    }

    if (rules.enforceTop4Pattern && size >= 8) {
      // シード複数配置の特例適用通知（isMultiSeedException && entry.seed）は固定シード起因のためスワップ対象外
      const top4V = detectTop4PatternViolations(slots, size).filter((v) => !(v.isMultiSeedException && v.entry.seed));
      hardPenalty += top4V.length * 140000;
    }

    if (rules.spreadUnder8To8Blocks && size >= 8 && schoolTotalCounts) {
      const schoolBlockCounts = {};
      for (let i = 0; i < size; i++) {
        const s = slots[i];
        if (!s || !s.school) continue;
        if ((schoolTotalCounts[s.school] || 0) <= 8) {
          const code = getBlockInfo(i, size).code;
          if (!schoolBlockCounts[s.school]) schoolBlockCounts[s.school] = {};
          schoolBlockCounts[s.school][code] = (schoolBlockCounts[s.school][code] || 0) + 1;
          if (schoolBlockCounts[s.school][code] > 1) {
            hardPenalty += 60000;
          }
        }
      }
    }

    if (rules.avoidRound2SameSchool && size >= 4) {
      for (let base = 0; base < size; base += 4) {
        const left = [slots[base], slots[base + 1]].filter(Boolean);
        const right = [slots[base + 2], slots[base + 3]].filter(Boolean);
        for (let a = 0; a < left.length; a++) {
          for (let b = 0; b < right.length; b++) {
            if (left[a].school === right[b].school) hardPenalty += 20000;
          }
        }
      }
    }

    // 5番手以降を含む学校全体の A〜D（4山）・in/out バランス評価
    if (rules.balanceLowerMountainAndSide && size >= 8) {
      const scBalance = {};
      for (let i = 0; i < size; i++) {
        const s = slots[i];
        if (!s || !s.school) continue;
        if (!scBalance[s.school]) {
          scBalance[s.school] = { m: { A: 0, B: 0, C: 0, D: 0 }, in: 0, out: 0 };
        }
        const info = getBlockInfo(i, size);
        scBalance[s.school].m[info.mountain]++;
        scBalance[s.school][info.side]++;
      }
      Object.values(scBalance).forEach((b) => {
        const mVals = [b.m.A, b.m.B, b.m.C, b.m.D];
        const mSpread = Math.max(...mVals) - Math.min(...mVals);
        if (mSpread > 1) hardPenalty += (mSpread - 1) * 28000;
        const sideSpread = Math.abs(b.in - b.out);
        if (sideSpread > 1) hardPenalty += (sideSpread - 1) * 24000;
      });
    }

    // 枠優先順位の忠実度 ＆ 1〜4番手のシード下回避（スワップ候補の選択基準）
    let priorityCost = 0;
    if (slotPriority) {
      for (let i = 0; i < size; i++) {
        const s = slots[i];
        if (!s) continue;
        const pRank = slotPriority[i];
        const u = underSeedMap ? underSeedMap[i] : null;
        if (s.schoolRank <= 4) {
          priorityCost += pRank * 20;
          if (rules.avoidTop4UnderSeed && u && !s.seed) {
            priorityCost += (u.type === 'R1' ? 12000 : 8000);
          }
        } else if (rules.lowRankFromBottomSlots) {
          priorityCost += (size + 1 - pRank) * 20;
          if (!u) {
            priorityCost += 5000;
          }
        }
      }
    }

    return { hardPenalty, totalPenalty: hardPenalty + priorityCost };
  }

  function optimizeBracketBySwap(slots, mask, size, rules, schoolTotalCounts, slotPriority, underSeedMap) {
    let current = evaluateGlobalPenalty(slots, mask, size, rules, schoolTotalCounts, slotPriority, underSeedMap);
    if (current.hardPenalty === 0) return;

    const movableIndices = [];
    for (let i = 0; i < size; i++) {
      if (mask[i] && slots[i] && !slots[i].seed) movableIndices.push(i);
    }
    if (movableIndices.length < 2) return;

    const maxPasses = 25;
    for (let pass = 0; pass < maxPasses; pass++) {
      if (current.hardPenalty === 0) break;
      const diag = analyzeBracketAlerts(slots, mask, size);
      const hotSlots = new Set();
      if (rules.avoidRound1SameSchool) {
        diag.round1SameSchool.forEach((c) => { hotSlots.add(c.slotA); hotSlots.add(c.slotB); });
      }
      if (rules.avoidTripleSameCard) {
        diag.tripleSameCards.forEach((card) => {
          card.matches.forEach((m) => { hotSlots.add(m.slotA); hotSlots.add(m.slotB); });
        });
      }
      if (rules.enforceTop4Pattern && size >= 8) {
        diag.top4Violations.forEach((v) => { hotSlots.add(v.slotIdx); hotSlots.add(v.anchorSlotIdx); });
      }
      if (rules.avoidRound2SameSchool) {
        diag.round2SameSchool.forEach((c) => { hotSlots.add(c.slotA); hotSlots.add(c.slotB); });
      }
      if (rules.spreadUnder8To8Blocks && size >= 8 && schoolTotalCounts) {
        const scBlocks = {};
        for (let idx = 0; idx < size; idx++) {
          const s = slots[idx];
          if (!s || !s.school || (schoolTotalCounts[s.school] || 0) > 8) continue;
          const code = getBlockInfo(idx, size).code;
          const k = `${s.school}||${code}`;
          if (!scBlocks[k]) scBlocks[k] = [];
          scBlocks[k].push(idx);
        }
        Object.values(scBlocks).forEach((arr) => {
          if (arr.length > 1) arr.forEach((idx) => hotSlots.add(idx));
        });
      }
      if (rules.balanceLowerMountainAndSide && size >= 8) {
        const scBal = {};
        for (let idx = 0; idx < size; idx++) {
          const s = slots[idx];
          if (!s || !s.school) continue;
          if (!scBal[s.school]) scBal[s.school] = { m: { A: 0, B: 0, C: 0, D: 0 }, in: 0, out: 0, lowerSlots: [] };
          const info = getBlockInfo(idx, size);
          scBal[s.school].m[info.mountain]++;
          scBal[s.school][info.side]++;
          if (s.schoolRank >= 5 && !s.seed) scBal[s.school].lowerSlots.push(idx);
        }
        Object.values(scBal).forEach((b) => {
          const mSpread = Math.max(b.m.A, b.m.B, b.m.C, b.m.D) - Math.min(b.m.A, b.m.B, b.m.C, b.m.D);
          if (mSpread > 1 || Math.abs(b.in - b.out) > 1) {
            b.lowerSlots.forEach((idx) => hotSlots.add(idx));
          }
        });
      }

      const candidateStarts = hotSlots.size > 0 ? Array.from(hotSlots) : movableIndices;
      let bestSwap = null;
      let bestEval = current;

      for (let si = 0; si < candidateStarts.length; si++) {
        const i = candidateStarts[si];
        if (!slots[i] || slots[i].seed) continue;
        for (let m = 0; m < movableIndices.length; m++) {
          const j = movableIndices[m];
          if (i === j) continue;
          const tmp = slots[i];
          slots[i] = slots[j];
          slots[j] = tmp;

          const ev = evaluateGlobalPenalty(slots, mask, size, rules, schoolTotalCounts, slotPriority, underSeedMap);
          // 違反（hardPenalty）を確実に減らすスワップのみ採用し、その中で枠優先忠実度が最も高いものを選ぶ
          if (
            ev.hardPenalty < bestEval.hardPenalty ||
            (ev.hardPenalty < current.hardPenalty && ev.hardPenalty === bestEval.hardPenalty && ev.totalPenalty < bestEval.totalPenalty)
          ) {
            bestEval = ev;
            bestSwap = [i, j];
          }

          slots[j] = slots[i];
          slots[i] = tmp;
        }
      }

      if (!bestSwap) break;
      const [a, b] = bestSwap;
      const tmp = slots[a];
      slots[a] = slots[b];
      slots[b] = tmp;
      current = bestEval;
    }
  }

  /**
   * 各校上位4番手の規則性違反を検出（同一校複数シード特例にも対応）
   */
  function detectTop4PatternViolations(slots, size) {
    if (size < 8) return [];
    const violations = [];
    const schools = new Set();
    const schoolTopRanks = {};
    for (let i = 0; i < size; i++) {
      const s = slots[i];
      if (!s || !s.school) continue;
      schools.add(s.school);
      if (s.schoolRank >= 1 && s.schoolRank <= 4) {
        if (!schoolTopRanks[s.school]) schoolTopRanks[s.school] = new Set();
        schoolTopRanks[s.school].add(s.schoolRank);
      }
    }

    schools.forEach((sc) => {
      const plan = computeSchoolTop4Plan(sc, slots, size, schoolTopRanks[sc]);
      if (!plan) return;

      for (let i = 0; i < size; i++) {
        const s = slots[i];
        if (!s || s.school !== sc) continue;
        if (s.schoolRank < 1 || s.schoolRank > 4) continue;
        // 標準モードでは1番手は基準点なのでスキップ、複数シード特例時は非シード1番手も検証
        if (!plan.isMultiSeedException && s.schoolRank === 1) continue;

        const actual = getBlockInfo(i, size);
        const exp = plan.byRank[s.schoolRank];
        if (!exp) continue;

        // 複数シード特例において、シード固定されている2〜4番手自体は
        // 標準規則から外れた場合に「シード特例適用中」として1回だけ通知する
        if (plan.isMultiSeedException && s.seed && s.schoolRank >= 2 && plan.anchorInfo) {
          const stdExp = getTop4ExpectedBlocks(plan.anchorInfo.mountain, plan.anchorInfo.side)[s.schoolRank - 1];
          if (stdExp && (actual.mountain !== stdExp.mountain || actual.side !== stdExp.side)) {
            violations.push({
              school: sc,
              entry: s,
              slotIdx: i,
              rank: s.schoolRank,
              anchorEntry: plan.anchorEntry,
              anchorSlotIdx: plan.anchorSlotIdx,
              anchorCode: plan.anchorInfo.code,
              anchorMountain: plan.anchorInfo.mountain,
              anchorSide: plan.anchorInfo.side,
              expectedCode: stdExp.code,
              expectedMountain: stdExp.mountain,
              expectedSide: stdExp.side,
              actualCode: actual.code,
              actualMountain: actual.mountain,
              actualSide: actual.side,
              causedBySeed: true,
              isMultiSeedException: true
            });
          }
          continue;
        }

        if (actual.mountain !== exp.mountain || actual.side !== exp.side) {
          violations.push({
            school: sc,
            entry: s,
            slotIdx: i,
            rank: s.schoolRank,
            anchorEntry: plan.anchorEntry,
            anchorSlotIdx: plan.anchorSlotIdx >= 0 ? plan.anchorSlotIdx : i,
            anchorCode: plan.anchorInfo ? plan.anchorInfo.code : '-',
            anchorMountain: plan.anchorInfo ? plan.anchorInfo.mountain : '-',
            anchorSide: plan.anchorInfo ? plan.anchorInfo.side : '-',
            expectedCode: exp.code,
            expectedMountain: exp.mountain,
            expectedSide: exp.side,
            actualCode: actual.code,
            actualMountain: actual.mountain,
            actualSide: actual.side,
            causedBySeed: Boolean(s.seed || (plan.anchorEntry && plan.anchorEntry.seed)),
            isMultiSeedException: plan.isMultiSeedException
          });
        }
      }
    });

    return violations;
  }

  /**
   * 全通知チェック
   */
  function analyzeBracketAlerts(slots, activeSlotMask, size) {
    const round1SameSchool = [];
    const round2SameSchool = [];

    for (let i = 0; i < size; i += 2) {
      const p1 = slots[i];
      const p2 = slots[i + 1];
      if (p1 && p2 && p1.school && p2.school && p1.school === p2.school) {
        round1SameSchool.push({ slotA: i, slotB: i + 1, school: p1.school, p1, p2 });
      }
    }

    const cardMap = buildRound1CardCounts(slots, size);
    const tripleSameCards = Object.values(cardMap)
      .filter((card) => card.matches.length >= 3)
      .sort((a, b) => b.matches.length - a.matches.length);

    const top4Violations = detectTop4PatternViolations(slots, size);

    if (size >= 4) {
      for (let base = 0; base < size; base += 4) {
        const leftSide = [base, base + 1].filter((idx) => slots[idx]);
        const rightSide = [base + 2, base + 3].filter((idx) => slots[idx]);
        const reported = new Set();
        leftSide.forEach((li) => {
          rightSide.forEach((ri) => {
            const pL = slots[li];
            const pR = slots[ri];
            if (pL && pR && pL.school && pL.school === pR.school) {
              const key = `${li}_${ri}`;
              if (!reported.has(key)) {
                reported.add(key);
                const bothBye = isByeSlot(li, activeSlotMask) && isByeSlot(ri, activeSlotMask);
                round2SameSchool.push({
                  slotA: li,
                  slotB: ri,
                  school: pL.school,
                  p1: pL,
                  p2: pR,
                  directMatch: bothBye
                });
              }
            }
          });
        });
      }
    }

    return { round1SameSchool, tripleSameCards, top4Violations, round2SameSchool };
  }

  return {
    DEFAULT_RULES,
    MOUNTAIN_PATTERNS,
    SIDE_PATTERNS,
    getTop4ExpectedBlocks,
    computeSchoolTop4Plan,
    nextPowerOfTwo,
    shuffleArray,
    buildBracketOrder,
    buildSlotPriorityMap,
    buildUnderTop8SeedMap,
    buildAdjustedSlotPriorityMap,
    createBalancedActiveMask,
    getBlockInfo,
    isByeSlot,
    resolveSchoolPriorityMap,
    generateBracket,
    detectTop4PatternViolations,
    analyzeBracketAlerts
  };
});

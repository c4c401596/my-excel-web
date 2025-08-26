// src/ExcelUploader.jsx
import React, { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
const APP_VERSION = "";
// ===== 공통 유틸 =====
const normalize = (s) =>
  String(s ?? "").toLowerCase().normalize("NFKC").replace(/\s|\/|\\|[-_,.·()\[\]{}:;|]/g, "");
const rowIsAllEmpty = (row) => Object.values(row).every((v) => String(v ?? "").trim() === "");
const getMappedValue = (row, col) => (col ? row?.[col] ?? "" : "");
const text = (v) => String(v ?? "").trim();
const colToLetter = (n) => { let s=""; while(n>0){n--; s=String.fromCharCode(65+(n%26))+s; n=Math.floor(n/26);} return s; };
const asNumber = (v) => {
  if (v == null) return null;
  if (typeof v === "number" && Number.isFinite(v)) return v;
  const n = Number(String(v).replace(/[^0-9.\-]/g, ""));
  return Number.isFinite(n) ? n : null;
};
const isDone = (v) => {
  if (typeof v === "number") return v > 1; // 엑셀 직렬일 1초과를 날짜로 간주
  const s = text(v).toLowerCase();
  if (!s) return false;
  return !/^(0|-|없음|null|n\/a|na|x)$/i.test(s);
};

function levenshteinRatio(a, b) {
  a = normalize(a); b = normalize(b);
  const m = a.length, n = b.length;
  if (!m && !n) return 1;
  if (!m || !n) return 0;
  const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++) {
    const cost = a[i - 1] === b[j - 1] ? 0 : 1;
    dp[i][j] = Math.min(dp[i-1][j]+1, dp[i][j-1]+1, dp[i-1][j-1]+cost);
  }
  const dist = dp[m][n];
  return 1 - dist / Math.max(m, n);
}
function smartPick(headers, hints, threshold = 0.62) {
  const hnorm = headers.map((h) => normalize(h));
  for (const hint of hints) {
    const hn = normalize(hint);
    const idx = hnorm.findIndex((h) => h.includes(hn) || hn.includes(h));
    if (idx >= 0) return headers[idx];
  }
  let bestIdx = -1, bestSim = 0;
  for (let i = 0; i < headers.length; i++) for (const hint of hints) {
    const sim = levenshteinRatio(headers[i], hint);
    if (sim > bestSim) { bestSim = sim; bestIdx = i; }
  }
  return bestSim >= threshold && bestIdx >= 0 ? headers[bestIdx] : "";
}

function aoaToObjectsFirstRow(aoa) {
  if (!aoa.length) return [];
  const rawHeader = (aoa[0] || []).map((h) => String(h || "").trim());
  const seen = {};
  const header = rawHeader.map((h) => {
    let name = h || "열";
    if (seen[name] == null) { seen[name] = 1; return name; }
    seen[name] += 1; return `${name}_${seen[name]}`;
  });
  const body = aoa.slice(1).filter((r) => r.some((c) => String(c ?? "").trim() !== ""));
  return body.map((row) => {
    const o = {}; for (let j = 0; j < header.length; j++) o[header[j]] = row[j] ?? ""; return o;
  });
}

// ===== 매핑 키 관련 =====
const STANDARD_KEYS = [
  { std: "공장", hints: ["공장", "공장번호", "factory", "site", "mill"] },
  { std: "본수", hints: ["본수", "본", "ends"] },
  { std: "사종", hints: ["사종", "사종명", "원사", "사종코드", "yarn"] },
  { std: "상하", hints: ["상/하", "상하", "상하구분", "빔구분", "구분", "upper", "lower"] },
  { std: "투입명", hints: ["투입설비명", "투입명", "설비명", "품명", "제품명", "item", "product"] },
  { std: "설비번호", hints: ["직기코드", "설비번호", "직기호수", "직기", "machine", "loom"] },
  { std: "제품코드", hints: ["제품코드", "코드", "product code", "item code"] },
];
const DEFAULT_ENABLED_KEYS = ["공장", "본수", "사종", "상하"];

function canonicalValue(stdKey, valRaw) {
  const val = text(valRaw);
  if (stdKey === "공장" || stdKey === "본수") {
    const n = Number(val.replace(/[^0-9.]/g, ""));
    return Number.isFinite(n) ? String(n % 1 === 0 ? Math.trunc(n) : n) : val;
  }
  if (stdKey === "사종") return val; // 원문 유지
  if (stdKey === "상하") {
    const t = val.replace(/\s|\/|\\|[-_,.·]/g, "");
    if (t.startsWith("상")) return "상";
    if (t.startsWith("하")) return "하";
    return t || val;
  }
  return val;
}
const buildKey = (row, map) =>
  DEFAULT_ENABLED_KEYS.map((k) => canonicalValue(k, getMappedValue(row, map[k] || ""))).join("||");

// ===== 기준파일(1) 전처리: 헤더 찾기 + A열 공장 삽입 =====
function findHeaderRow(aoa) {
  const top = Math.min(15, aoa.length);
  let best = 0, bestScore = -Infinity;
  for (let i = 0; i < top; i++) {
    const row = (aoa[i] || []).map((x) => String(x ?? ""));
    if (!row.length) continue;
    let score = 0;
    for (const { hints } of STANDARD_KEYS) {
      const hit = row.some((cell) => {
        const c = normalize(cell);
        return hints.some((h) => {
          const hn = normalize(h);
          return c.includes(hn) || hn.includes(c) || levenshteinRatio(c, hn) >= 0.82;
        });
      });
      if (hit) score += 2;
    }
    const nonEmpty = row.filter((c) => String(c).trim() !== "").length;
    const numericish = row.filter((c) => /^-?\d+(\.\d+)?$/.test(String(c).trim())).length;
    score += Math.max(0, nonEmpty - numericish * 0.7);
    if (score > bestScore) { bestScore = score; best = i; }
  }
  return best;
}
const TUIP_HINTS = ["투입설비명", "투입명", "설비명"];
function inferFactoryCodeFromText(s) {
  const m = String(s ?? "").match(/([123])\s*공장\s*(\d+)\s*호/i);
  if (!m) return "";
  const g = Number(m[1]), ho = Number(m[2]);
  if (g === 1 && ho >= 1 && ho <= 20) return "1";
  if (g === 2 && ho >= 1 && ho <= 32) return "2";
  if (g === 3 && ho >= 1 && ho <= 10) return "3";
  return "";
}
function inferFactoryCodeFromRow(row, header) {
  let idx = -1;
  for (let i = 0; i < header.length; i++) {
    const h = String(header[i] ?? "");
    const hit = TUIP_HINTS.some((hint) => normalize(h).includes(normalize(hint)));
    if (hit) { idx = i; break; }
  }
  if (idx >= 0) {
    const code = inferFactoryCodeFromText(row[idx]);
    if (code) return code;
  }
  for (const cell of row) {
    const code = inferFactoryCodeFromText(cell);
    if (code) return code;
  }
  return "";
}
function preprocessBaseAOA_HeaderDetect_InsertFactory(aoa) {
  const arr = aoa.map((r) => [...r]);
  if (arr.length === 0) return arr;
  const h = findHeaderRow(arr);
  if (h > 0) arr.splice(0, h);
  const header = (arr[0] || []).map((x) => String(x ?? ""));
  const out = [];
  out.push(["공장", ...header]); // A열 공장
  for (let r = 1; r < arr.length; r++) {
    const row = arr[r] || [];
    const code = inferFactoryCodeFromRow(row, header);
    out.push([code, ...row]);
  }
  return out;
}

// ===== 출고파일(3) 전처리: 병합 해제/빈 행 삭제/헤더 승격 =====
const SHIP_HINTS = ["창고코드","상/하구분","총본수","사종코드","수량","요청일","빔작업완료일","완료일","완료"];
function findHeaderRowGeneric(aoa, hints) {
  const top = Math.min(15, aoa.length);
  let best = 0, bestScore = -Infinity;
  for (let i = 0; i < top; i++) {
    const row = (aoa[i] || []).map((x) => String(x ?? ""));
    if (!row.length) continue;
    let score = 0;
    for (const hint of hints) {
      const hit = row.some((cell) => {
        const c = normalize(cell); const hn = normalize(hint);
        return c.includes(hn) || hn.includes(c) || levenshteinRatio(c, hn) >= 0.8;
      });
      if (hit) score += 2;
    }
    const nonEmpty = row.filter((c) => String(c).trim() !== "").length;
    const numericish = row.filter((c) => /^-?\d+(\.\d+)?$/.test(String(c).trim())).length;
    score += Math.max(0, nonEmpty - numericish * 0.6);
    if (score > bestScore) { bestScore = score; best = i; }
  }
  return best;
}
function preprocessShipAOA(aoa) {
  if (!aoa?.length) return aoa;
  let arr = aoa.filter(r => (r||[]).some(c => String(c ?? "").trim() !== ""));
  if (!arr.length) return arr;
  const h = findHeaderRowGeneric(arr, SHIP_HINTS);
  if (h > 0) arr.splice(0, h);
  const header = (arr[0] || []).map(x => String(x ?? "").trim());
  const cols = header.length;

  // 위쪽부터 병합해제 효과: 빈칸은 위 값으로 채우되, 완료/수량/삭제 계열은 제외
  const skipFill = header.map(hd => {
    const n = normalize(hd);
    return n.includes("완료") || n.includes("수량") || n.includes("삭제") || n.includes("done") || n.includes("finish");
  });
  for (let r = 1; r < arr.length; r++) {
    arr[r] = arr[r] || [];
    for (let c = 0; c < cols; c++) {
      if (skipFill[c]) continue;
      const cur = String(arr[r][c] ?? "").trim();
      if (!cur) arr[r][c] = arr[r-1]?.[c] ?? "";
    }
  }

  // 매우 비어있는 행 제거
  const idxQty  = header.findIndex(hd => normalize(hd).includes(normalize("수량")));
  const idxWH   = header.findIndex(hd => normalize(hd).includes(normalize("창고코드")));
  const idxYarn = header.findIndex(hd => normalize(hd).includes(normalize("사종코드")));
  arr = arr.filter((row, i) => {
    if (i === 0) return true;
    const qty = idxQty >= 0 ? asNumber(row[idxQty]) : null;
    const hasWH = idxWH >= 0 ? String(row[idxWH] ?? "").trim() !== "" : false;
    const hasYarn = idxYarn >= 0 ? String(row[idxYarn] ?? "").trim() !== "" : false;
    return hasWH || hasYarn || (qty !== null && qty > 0);
  });
  return arr;
}

// ===== 리더 =====
function sheetToObjectsSmart(ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!aoa.length) return [];
  let bestIdx = 0, bestScore = -Infinity;
  const limit = Math.min(10, aoa.length);
  for (let i = 0; i < limit; i++) {
    const row = aoa[i] || [];
    let nonEmpty = 0, numericish = 0;
    for (const cell of row) {
      const s = String(cell ?? "").trim();
      if (s) { nonEmpty++; if (/^-?\d+(\.\d+)?$/.test(s)) numericish++; }
    }
    const score = nonEmpty - numericish * 0.5;
    if (score > bestScore) { bestScore = score; bestIdx = i; }
  }
  const rawHeader = (aoa[bestIdx] || []).map((h) => String(h || "").trim());
  const seen = {};
  const header = rawHeader.map((h) => {
    let name = h || "열";
    if (seen[name] == null) { seen[name] = 1; return name; }
    seen[name] += 1; return `${name}_${seen[name]}`;
  });
  const body = aoa.slice(bestIdx + 1).filter((r) => r.some((c) => String(c ?? "").trim() !== ""));
  return body.map((row) => {
    const o = {}; for (let j = 0; j < header.length; j++) o[header[j]] = row[j] ?? ""; return o;
  });
}
function readXlsxSmart(file, { preprocessor = null, headerMode = "auto" } = {}) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = (e) => reject(e);
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const sheet = wb.SheetNames[0];
        const ws = wb.Sheets[sheet];

        let aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        if (typeof preprocessor === "function") aoa = preprocessor(aoa);

        let rows = [];
        if (headerMode === "firstRow") rows = aoaToObjectsFirstRow(aoa);
        else rows = sheetToObjectsSmart(XLSX.utils.aoa_to_sheet(aoa));

        resolve({ rows, meta: { file: file.name, sheet, count: rows.length } });
      } catch (err) { reject(err); }
    };
    reader.readAsArrayBuffer(file);
  });
}

// ===== 기준 헤더 보정 =====
function fixBaseHeaderColumns(rows) {
  if (!rows || !rows.length) return rows;
  return rows.map((row) => {
    const out = { ...row };
    ["열", "열1", "열_1"].forEach((k) => { if (k in out) delete out[k]; });
    if (!("재고일수" in out) || text(out["재고일수"]) === "") {
      const cand = ["열2", "열_2", "열 2"].find((k) => k in out);
      if (cand) { out["재고일수"] = out[cand]; delete out[cand]; }
    }
    return out;
  });
}

// ===== 컴포넌트 =====
export default function ExcelUploader() {
  const baseInputRef = useRef(null);
  const stockInputRef = useRef(null);
  const shipInputRef = useRef(null);

  const [baseRows, setBaseRows] = useState([]);
  const [stockRows, setStockRows] = useState([]);
  const [shipRows, setShipRows] = useState([]);

  const [baseMeta, setBaseMeta] = useState(null);
  const [stockMeta, setStockMeta] = useState(null);
  const [shipMeta, setShipMeta] = useState(null);

  const baseHeaders = useMemo(() => (baseRows.length ? Object.keys(baseRows[0]) : []), [baseRows]);
  const stockHeaders = useMemo(() => (stockRows.length ? Object.keys(stockRows[0]) : []), [stockRows]);
  const shipHeaders  = useMemo(() => (shipRows.length ? Object.keys(shipRows[0])  : []), [shipRows]);

  const [baseMap, setBaseMap] = useState({});
  const [stockMap, setStockMap] = useState({});
  const [shipMap, setShipMap] = useState({});

  const [basePreviewAOA, setBasePreviewAOA] = useState(null);
  const [factoryStats, setFactoryStats] = useState({ filled: 0, empty: 0 });

  const [showFilterRow, setShowFilterRow] = useState(true);
  const [globalQuery, setGlobalQuery] = useState("");
  const [columnFilters, setColumnFilters] = useState({});
  const [sortKey, setSortKey] = useState("");
  const [sortDir, setSortDir] = useState(null);
  const [errorMsg, setErrorMsg] = useState("");
  const [showOnlyNeedClaim, setShowOnlyNeedClaim] = useState(false);

  // 업로드
  async function onPickBase(e) {
    const file = e.target.files?.[0]; if (!file) return;
    setErrorMsg("");
    try {
      let { rows, meta } = await readXlsxSmart(file, {
        preprocessor: preprocessBaseAOA_HeaderDetect_InsertFactory,
        headerMode: "firstRow",
      });
      rows = fixBaseHeaderColumns(rows);
      setBaseRows(rows); setBaseMeta(meta);

      const headersLocal = rows.length ? Object.keys(rows[0]) : [];
      const preview = [headersLocal];
      let filled = 0, empty = 0;
      rows.forEach((r) => { if (text(r["공장"])) filled++; else empty++; preview.push(headersLocal.map((h) => r[h])); });
      setBasePreviewAOA(preview);
      setFactoryStats({ filled, empty });

      setGlobalQuery(""); setColumnFilters({}); setSortKey(""); setSortDir(null);
    } catch (err) { setErrorMsg("기준 파일 읽기 오류: " + (err?.message || String(err))); }
    finally { if (baseInputRef.current) baseInputRef.current.value = ""; }
  }
  async function onPickStock(e) {
    const file = e.target.files?.[0]; if (!file) return;
    setErrorMsg("");
    try {
      const { rows, meta } = await readXlsxSmart(file, { headerMode: "auto" });
      setStockRows(rows); setStockMeta(meta);
      setGlobalQuery(""); setColumnFilters({}); setSortKey(""); setSortDir(null);
    } catch (err) { setErrorMsg("재고 파일 읽기 오류: " + (err?.message || String(err))); }
    finally { if (stockInputRef.current) stockInputRef.current.value = ""; }
  }
  async function onPickShip(e) {
    const file = e.target.files?.[0]; if (!file) return;
    setErrorMsg("");
    try {
      const { rows, meta } = await readXlsxSmart(file, { preprocessor: preprocessShipAOA, headerMode: "firstRow" });
      setShipRows(rows); setShipMeta(meta);
      setGlobalQuery(""); setColumnFilters({}); setSortKey(""); setSortDir(null);
    } catch (err) { setErrorMsg("출고 파일 읽기 오류: " + (err?.message || String(err))); }
    finally { if (shipInputRef.current) shipInputRef.current.value = ""; }
  }

  // 자동 매핑
  useEffect(() => {
    if (!baseRows.length) return;
    setBaseMap((prev) => {
      const m = { ...prev };
      STANDARD_KEYS.forEach(({ std, hints }) => { if (!m[std]) m[std] = smartPick(baseHeaders, hints); });
      return m;
    });
  }, [baseRows, baseHeaders.join("|")]);
  useEffect(() => {
    if (!stockRows.length) return;
    setStockMap((prev) => {
      const m = { ...prev };
      STANDARD_KEYS.forEach(({ std, hints }) => { if (!m[std]) m[std] = smartPick(stockHeaders, hints); });
      return m;
    });
  }, [stockRows, stockHeaders.join("|")]);
  useEffect(() => {
    if (!shipRows.length) return;
    setShipMap((prev) => {
      const m = { ...prev };
      m["공장"] = smartPick(shipHeaders, ["창고코드","공장","factory","site"]);
      m["상하"] = smartPick(shipHeaders, ["상/하구분","상하구분","상/하","상하","구분"]);
      m["본수"] = smartPick(shipHeaders, ["총본수","본수","본","ends"]);
      m["사종"] = smartPick(shipHeaders, ["사종코드","사종","사종명","원사"]);
      return m;
    });
  }, [shipRows, shipHeaders.join("|")]);

  const effectiveKeys = useMemo(
    () => DEFAULT_ENABLED_KEYS.filter((k) => !!baseMap[k] && !!stockMap[k]),
    [baseMap, stockMap]
  );

  // 출고 보조 헤더
  const shipDoneCol = useMemo(
    () => smartPick(shipHeaders, ["빔작업완료일","작업완료일","완료일","완료일자","완료","finish","done"], 0.5),
    [shipHeaders.join("|")]
  );
  const shipQtyCol = useMemo(
    () => smartPick(shipHeaders, ["수량","총본수","본수","qty","quantity"], 0.5),
    [shipHeaders.join("|")]
  );
  const shipDeleteCol = useMemo(
    () => smartPick(shipHeaders, ["삭제","delete","del"], 0.5),
    [shipHeaders.join("|")]
  );

  // 기준 "재고일수"
  const DAYS_HINTS = ["재고일수", "재고 일수", "일수", "days", "재고days", "재고(d)"];
  const baseDaysCol = useMemo(() => smartPick(baseHeaders, DAYS_HINTS, 0.55), [baseHeaders.join("|")]);
  const parseDays = (val) => {
    const n = asNumber(val);
    return Number.isFinite(n) ? n : Number.POSITIVE_INFINITY;
  };

  // 클린업
  const baseClean = useMemo(() => {
    let arr = baseRows.slice();
    arr = arr.filter((r) => !rowIsAllEmpty(r));
    if (effectiveKeys.length)
      arr = arr.filter((r) => effectiveKeys.every((std) => text(getMappedValue(r, baseMap[std] || "")) !== ""));
    return arr;
  }, [baseRows, effectiveKeys, baseMap]);
  const stockClean = useMemo(() => {
    let arr = stockRows.slice();
    arr = arr.filter((r) => !rowIsAllEmpty(r));
    if (effectiveKeys.length)
      arr = arr.filter((r) => effectiveKeys.every((std) => text(getMappedValue(r, stockMap[std] || "")) !== ""));
    return arr;
  }, [stockRows, effectiveKeys, stockMap]);

  const shipClean = useMemo(() => {
    let arr = shipRows.slice();
    arr = arr.filter((r) => !rowIsAllEmpty(r));
    if (shipDeleteCol) {
      arr = arr.filter(r => {
        const v = text(getMappedValue(r, shipDeleteCol)).toLowerCase();
        return !(v === "1" || v === "y" || v === "yes" || v === "true" || v === "삭제");
      });
    }
    return arr;
  }, [shipRows, shipDeleteCol]);

  // 출고 대기건(완료일 미기입 + 수량>0)
  const shipPendingRows = useMemo(() => {
    if (!shipClean.length) return [];
    return shipClean
      .filter((r) => {
        const doneVal = shipDoneCol ? getMappedValue(r, shipDoneCol) : "";
        const qtyVal  = shipQtyCol  ? getMappedValue(r, shipQtyCol)  : "";
        const qty = asNumber(qtyVal);
        const hasDone = isDone(doneVal);
        return !hasDone && qty !== null && qty > 0;
      })
      .map((r) => ({
        공장: canonicalValue("공장", getMappedValue(r, shipMap["공장"])),
        상하: canonicalValue("상하", getMappedValue(r, shipMap["상하"])),
        본수: canonicalValue("본수", getMappedValue(r, shipMap["본수"])),
        사종: getMappedValue(r, shipMap["사종"]),
        수량: asNumber(getMappedValue(r, shipQtyCol)) || 0,
        매칭키: buildKey(r, shipMap),
      }));
  }, [shipClean, shipMap, shipQtyCol, shipDoneCol]);

  const shipKeySet = useMemo(() => {
    const set = new Set();
    for (const r of shipPendingRows) set.add(r.매칭키);
    return set;
  }, [shipPendingRows]);

  // 재고 수량(키별)
  const stockKeyCount = useMemo(() => {
    const cnt = new Map();
    if (!effectiveKeys.length) return cnt;
    for (const r of stockClean) {
      const k = buildKey(r, stockMap);
      if (!k) continue;
      cnt.set(k, (cnt.get(k) || 0) + 1);
    }
    return cnt;
  }, [stockClean, effectiveKeys, stockMap]);

  // 1:1 매칭 (재고일수 오름차순)
  const oneToOneMatchFlags = useMemo(() => {
    const flag = new Array(baseClean.length).fill(0);
    if (!effectiveKeys.length) return flag;

    const groups = new Map();
    baseClean.forEach((row, idx) => {
      const key = buildKey(row, baseMap);
      const days = parseDays(baseDaysCol ? row[baseDaysCol] : undefined);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push({ idx, days });
    });

    for (const [key, arr] of groups.entries()) {
      const available = stockKeyCount.get(key) || 0;
      arr.sort((a, b) => a.days - b.days);
      for (let i = 0; i < arr.length; i++) flag[arr[i].idx] = i < available ? 1 : 0;
    }
    return flag;
  }, [baseClean, baseMap, stockKeyCount, baseDaysCol, effectiveKeys.length]);

  // 결과 + 구분(있음/없음/청구) + 매칭키
  const mergedRows = useMemo(() => {
    if (!effectiveKeys.length) return [];
    return baseClean.map((r, i) => {
      const matched = oneToOneMatchFlags[i] === 1;
      const baseKey = buildKey(r, baseMap);

      let status;
      if (matched) {
        status = "있음";                // 재고 있음
      } else if (shipKeySet.has(baseKey)) {
        status = "없음";                // 재고 없지만 출고 예정 있음 → 청구 불필요
      } else {
        status = "청구";                // 재고 없고 출고 예정도 없음 → 청구 필요
      }

      return { ...r, 매칭키: baseKey, 재고유무: status, 재고매칭수: matched ? 1 : 0 };
    });
  }, [baseClean, oneToOneMatchFlags, shipKeySet, baseMap, effectiveKeys.length]);

  // 보기용 별칭 보강
  const alias = {
    "투입설비명": ["투입설비명", "투입명", "설비명", "설비번호", "직기호수"],
    "상/하": ["상/하", "상하", "상/하구분", "상하구분", "구분"],
    "직기코드": ["직기코드", "설비번호", "직기호수", "직기"],
    "사용량(cm)": ["사용량(cm)", "사용량", "사용", "일사용량"],
    "잔량(cm)": ["잔량(cm)", "잔량", "잔량cm"],
    "잔량(%)": ["잔량(%)", "잔량%", "잔량비율"],
    "재고일수": ["재고일수", "재고 일수", "일수", "days"],
  };
  const pickFromRow = (row, list) => list.find((c) => row[c] != null && text(row[c]) !== "");
  const augmentedRows = useMemo(() => {
    return mergedRows.map((row) => {
      const out = { ...row };
      Object.keys(alias).forEach((k) => {
        if (out[k] == null || out[k] === "") {
          const src = pickFromRow(row, alias[k] || []);
          if (src) out[k] = row[src];
        }
      });
      return out;
    });
  }, [mergedRows]);

  // 화면 컬럼
  const displayPreferred = ["공장","투입설비명","상/하","본수","사종","사용량(cm)","잔량(cm)","잔량(%)","재고일수","재고유무","매칭키","직기코드"];
  const headers = useMemo(() => {
    const cols = augmentedRows[0] ? Object.keys(augmentedRows[0]) : [];
    const pref = displayPreferred.filter((h) => cols.includes(h));
    const rest = cols.filter((c) => !pref.includes(c));
    return [...pref, ...rest];
  }, [augmentedRows]);

  // 화면 데이터
  const isNumberLike = (v) => v !== "" && v !== null && !Number.isNaN(Number(v));
  const visibleData = useMemo(() => {
    let rows = augmentedRows;
    if (showOnlyNeedClaim) rows = rows.filter(r => r["재고유무"] === "청구");
    const q = globalQuery.trim().toLowerCase();
    if (q) rows = rows.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q)));
    const active = Object.entries(columnFilters).filter(([, v]) => (v ?? "").trim() !== "");
    if (active.length) rows = rows.filter((r) => active.every(([c, v]) => String(r[c] ?? "").toLowerCase().includes(String(v).toLowerCase())));
    if (sortKey && sortDir) {
      rows = rows.slice().sort((a, b) => {
        const av = a[sortKey], bv = b[sortKey];
        if (isNumberLike(av) && isNumberLike(bv)) {
          const diff = Number(av) - Number(bv);
          return sortDir === "asc" ? diff : -diff;
        }
        const diff = String(av ?? "").localeCompare(String(bv ?? ""), undefined, { numeric: true, sensitivity: "base" });
        return sortDir === "asc" ? diff : -diff;
      });
    }
    return rows;
  }, [augmentedRows, showOnlyNeedClaim, globalQuery, columnFilters, sortKey, sortDir]);

  // 재고일수 정렬본
  const rowsSortedByDays = useMemo(() => {
    const getDays = (row) => {
      const raw = row["재고일수"] ?? (baseDaysCol ? row[baseDaysCol] : undefined);
      const n = asNumber(raw);
      return Number.isFinite(n) ? n : Number.POSITIVE_INFINITY;
    };
    const tmp = visibleData.map((r) => ({ ...r, __d: getDays(r) }));
    tmp.sort((a, b) => a.__d - b.__d);
    return tmp.map(({ __d, ...rest }) => rest);
  }, [visibleData, baseDaysCol]);

  // ===== 엑셀 내보내기 =====
  async function exportXLSX() {
    try {
      if (!visibleData.length) return alert("내보낼 데이터가 없습니다.");
      const wb = new ExcelJS.Workbook();

      // 시트1: 결과(현재보기)
      const ws1 = wb.addWorksheet("결과(현재보기)");
      ws1.addRow(headers);
      ws1.getRow(1).font = { bold: true };
      ws1.getRow(1).alignment = { vertical: "middle", horizontal: "center" };
      visibleData.forEach((r) => ws1.addRow(headers.map((h) => r[h])));
      ws1.columns = headers.map(() => ({ width: 16 }));
      ws1.autoFilter = { from: "A1", to: `${colToLetter(headers.length)}1` };

      // 시트2: 재고현황(보고서)
      const ws2 = wb.addWorksheet("재고현황(보고서)");
      const now = new Date(); const mm = now.getMonth() + 1; const dd = now.getDate();
      ws2.mergeCells(1, 1, 1, 7);
      const titleCell = ws2.getCell(1, 1);
      titleCell.value = `${mm}월  ${dd}일  오전 (09시 40분 )  기준 빔 현황`;
      titleCell.alignment = { horizontal: "center", vertical: "middle" };
      ws2.getCell("H1").value = "재고있음 :";
      ws2.getCell("I1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
      ws2.getCell("K1").value = "재고없음 :";
      ws2.getCell("L1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF92D050" } };
      ws2.getCell("M1").value = "청구 :";
      ws2.getCell("N1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } };

      const header2 = [
        "공장","투입설비명","상/하","본수","사종","사용량(cm)","잔량(cm)","잔량(%)","재고일수",
        "구분","빔하대\n예상시간\n(9시 40분 기준)","사종","다음계획"
      ];
      ws2.addRow(header2);
      const hr = ws2.getRow(2);
      hr.height = 48; hr.font = { bold: true };
      hr.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      ws2.autoFilter = { from: "A2", to: "M2" };

      const pxToChar = (px) => px / 7.1;
      ws2.getColumn('A').width = 4.30;
      ws2.getColumn('B').width = pxToChar(123);
      ws2.getColumn('C').width = pxToChar(66);
      ws2.getColumn('D').width = pxToChar(66);
      ws2.getColumn('E').width = pxToChar(170);
      ['F','G','H','I','J'].forEach(c => ws2.getColumn(c).width = pxToChar(60));
      ws2.getColumn('K').width = pxToChar(204);
      ['L','M'].forEach(c => ws2.getColumn(c).width = pxToChar(160));
      ;['A','B','C','D','E','F','G','H','I','J','K','L','M'].forEach(c => {
        ws2.getColumn(c).alignment = { vertical: "middle", horizontal: "center", shrinkToFit: true };
      });

      const numInt = "#,##0", num1 = "0.0";
      const dateFmt = 'm"월" d"일" hh"시" mm"분"';

      rowsSortedByDays.forEach((r, i) => {
        const excelRow = 3 + i;
        const baseDaysNum = asNumber(r["재고일수"]);

        const row = ws2.addRow([
          r["공장"] ?? "",
          r["투입설비명"] ?? "",
          r["상/하"] ?? r["상하"] ?? "",
          asNumber(r["본수"]),
          r["사종"] ?? "",
          asNumber(r["사용량(cm)"]),
          asNumber(r["잔량(cm)"]),
          asNumber(r["잔량(%)"]),
          null,
          r["재고유무"] ?? "",
          null,
          "",
          ""
        ]);

        if (typeof row.getCell(4).value === "number") row.getCell(4).numFmt = numInt;
        if (typeof row.getCell(6).value === "number") row.getCell(6).numFmt = numInt;
        if (typeof row.getCell(7).value === "number") row.getCell(7).numFmt = numInt;
        if (typeof row.getCell(8).value === "number") row.getCell(8).numFmt = num1;

        if (Number.isFinite(baseDaysNum)) { row.getCell(9).value = baseDaysNum; row.getCell(9).numFmt = num1; }
        else { row.getCell(9).value = "O"; }

        row.getCell(11).value = { formula: `TODAY()+TIME(9,40,0)+IF(C${excelRow}="상", G${excelRow}/11.11, IF(C${excelRow}="하", G${excelRow}/3.75, ""))` };
        row.getCell(11).numFmt = dateFmt;

        const j = row.getCell(10);
        const v = String(j.value ?? "");
        if (v === "있음") j.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
        else if (v === "없음") j.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF92D050" } };
        else if (v === "청구") j.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } };

        for (let c = 1; c <= header2.length; c++) {
          row.getCell(c).alignment = { vertical: "middle", horizontal: "center", shrinkToFit: true };
          row.getCell(c).border = { top:{style:"thin"}, left:{style:"thin"}, bottom:{style:"thin"}, right:{style:"thin"} };
        }
      });

      // 시트3: 출고예정(요약)
      const ws3 = wb.addWorksheet("출고예정(요약)");
      ws3.addRow(["공장","상/하","본수","사종","수량"]);
      ws3.getRow(1).font = { bold: true };
      ws3.getRow(1).alignment = { vertical: "middle", horizontal: "center" };

      const summaryMap = new Map();
      for (const r of shipPendingRows) {
        const key = `${r.공장}||${r.상하}||${r.본수}||${r.사종}`;
        const cur = summaryMap.get(key) || { 공장: r.공장, 상하: r.상하, 본수: r.본수, 사종: r.사종, 수량: 0 };
        cur.수량 += r.수량;
        summaryMap.set(key, cur);
      }
      const summaryRows = Array.from(summaryMap.values())
        .sort((a,b) =>
          String(a.공장).localeCompare(String(b.공장), undefined, { numeric:true }) ||
          String(a.사종).localeCompare(String(b.사종), undefined, { numeric:true }) ||
          String(a.본수).localeCompare(String(b.본수), undefined, { numeric:true }) ||
          String(a.상하).localeCompare(String(b.상하), undefined, { numeric:true })
        );
      for (const r of summaryRows) {
        const row = ws3.addRow([r.공장, r.상하, Number(r.본수), r.사종, Number(r.수량)]);
        row.getCell(3).numFmt = "#,##0";
        row.getCell(5).numFmt = "#,##0";
      }
      ws3.columns = [
        { width: 8 }, { width: 8 }, { width: 10 }, { width: 30 }, { width: 10 }
      ];
      ws3.autoFilter = { from: "A1", to: "E1" };
      for (let r = 1; r <= ws3.rowCount; r++) {
        for (let c = 1; c <= 5; c++) {
          ws3.getCell(r, c).alignment = { vertical: "middle", horizontal: "center" };
          ws3.getCell(r, c).border = { top:{style:"thin"}, left:{style:"thin"}, bottom:{style:"thin"}, right:{style:"thin"} };
        }
      }

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = "재고매칭_결과.xlsx";
      document.body.appendChild(a); a.click(); a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      setErrorMsg("엑셀 내보내기 오류: " + (err?.message || String(err)));
    }
  }

  const mappingStatus =
    (baseRows.length || stockRows.length || shipRows.length)
      ? `자동매핑 — 기준:${DEFAULT_ENABLED_KEYS.map(k=>`${k}:${baseMap[k]||"?"}`).join(" / ")} · 재고:${DEFAULT_ENABLED_KEYS.map(k=>`${k}:${stockMap[k]||"?"}`).join(" / ")} · 출고:${DEFAULT_ENABLED_KEYS.map(k=>`${k}:${shipMap[k]||"?"}`).join(" / ")}`
      : "파일 업로드 대기 중";

  return (
    <div>
      <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
        <div style={{ fontSize: 12, opacity: 0.7, padding: "2px 6px", border: "1px dashed #bbb", borderRadius: 6 }}>{APP_VERSION}</div>

        <div><div style={{ fontWeight: 600, marginBottom: 6 }}>① 기준 파일</div>
          <input ref={baseInputRef} type="file" accept=".xlsx,.xls" onChange={onPickBase} />
          {baseMeta && <div style={{ fontSize: 12, opacity: .85, marginTop: 4 }}>📄 {baseMeta.file} / {baseMeta.sheet} ({baseMeta.count.toLocaleString()}행)</div>}
        </div>

        <div><div style={{ fontWeight: 600, marginBottom: 6 }}>② 재고 파일</div>
          <input ref={stockInputRef} type="file" accept=".xlsx,.xls" onChange={onPickStock} />
          {stockMeta && <div style={{ fontSize: 12, opacity: .85, marginTop: 4 }}>📦 {stockMeta.file} / {stockMeta.sheet} ({stockMeta.count.toLocaleString()}행)</div>}
        </div>

        <div><div style={{ fontWeight: 600, marginBottom: 6 }}>③ 출고 파일(작업요청서)</div>
          <input ref={shipInputRef} type="file" accept=".xlsx,.xls" onChange={onPickShip} />
          {shipMeta && <div style={{ fontSize: 12, opacity: .85, marginTop: 4 }}>🚚 {shipMeta.file} / {shipMeta.sheet} ({shipMeta.count.toLocaleString()}행)</div>}
        </div>

        <button onClick={() => setShowFilterRow((v) => !v)} style={{ padding: "8px 10px" }}>
          열 필터 {showFilterRow ? "숨기기" : "보이기"}
        </button>

        <input
          placeholder="전역 검색"
          value={globalQuery}
          onChange={(e) => setGlobalQuery(e.target.value)}
          style={{ padding: "8px 10px", border: "1px solid #ddd", borderRadius: 6, minWidth: 200 }}
        />

        <label style={{ display:"inline-flex", alignItems:"center", gap:8 }}>
          <input type="checkbox" checked={showOnlyNeedClaim} onChange={(e)=>setShowOnlyNeedClaim(e.target.checked)} />
          청구만 보기
        </label>

        <button onClick={exportXLSX} style={{ padding: "8px 12px", border: "1px solid #ddd", borderRadius: 8, background: "#fafafa" }}>
          엑셀(xlsx) 내보내기
        </button>
      </div>

      {(baseRows.length || stockRows.length || shipRows.length) ? (
        <div style={{ marginTop: 8, fontSize: 13, opacity: 0.9 }}>{mappingStatus}</div>
      ) : null}

      {errorMsg && (
        <div style={{ marginTop: 8, padding: "8px 12px", border: "1px solid #ffa39e", background: "#fff1f0", color: "#cf1322", borderRadius: 6 }}>
          ⚠️ {errorMsg}
        </div>
      )}

      {/* 간단 테이블 미리보기 (옵션) */}
      {visibleData.length > 0 && (
        <div style={{ overflowX: "auto", marginTop: 16 }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 14, border: "1px solid #e5e5e5" }}>
            <thead>
              <tr style={{ background: "#f6f6f6" }}>
                {headers.map((h) => (
                  <th
                    key={h}
                    onClick={() => {
                      if (sortKey !== h) { setSortKey(h); setSortDir("asc"); }
                      else setSortDir((d) => (d === "asc" ? "desc" : d === "desc" ? null : "asc"));
                    }}
                    style={{
                      padding: "10px 8px", borderBottom: "1px solid #e5e5e5",
                      textAlign: "center", whiteSpace: "nowrap", cursor: "pointer",
                      userSelect: "none", position: "sticky", top: 0, background: "#f6f6f6",
                    }}
                    title="헤더 클릭: 정렬(오름/내림/해제)"
                  >
                    {h}
                  </th>
                ))}
              </tr>
              {showFilterRow && (
                <tr>
                  {headers.map((h) => (
                    <th key={h} style={{ padding: 4, borderBottom: "1px solid #e5e5e5", textAlign: "center" }}>
                      <input
                        value={columnFilters[h] ?? ""}
                        onChange={(e) => setColumnFilters((p)=>({ ...p, [h]: e.target.value }))}
                        placeholder="이 열만 검색"
                        style={{ width: "100%", boxSizing: "border-box", padding: "6px 8px", border: "1px solid #ddd", borderRadius: 6, textAlign: "center" }}
                      />
                    </th>
                  ))}
                </tr>
              )}
            </thead>
            <tbody>
              {visibleData.map((row, rIdx) => (
                <tr
                  key={rIdx}
                  style={{
                    borderTop: "1px solid #f0f0f0",
                    background:
                      row["재고유무"] === "있음" ? "#f6ffed" :
                      row["재고유무"] === "없음" ? "#fff7f6" :
                      row["재고유무"] === "청구" ? "#fff1f0" : "transparent",
                  }}
                >
                  {headers.map((h) => (
                    <td key={h} style={{ padding: "8px 8px", verticalAlign: "top", textAlign: "center" }}>
                      {String(row[h] ?? "")}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

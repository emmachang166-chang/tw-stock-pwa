import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ── Supabase Config ──
const SUPA_URL = "https://vbibpfapxvyfpvxoqwah.supabase.co";
const SUPA_KEY = "sb_publishable_dg96WUQavjlew-BxLhFeCg_l853-Y5D";

const supa = async (path, opts = {}) => {
  const { prefer, headers: extraHeaders, ...fetchOpts } = opts;
  // For write ops, use return=representation to get JSON back; for reads no Prefer needed
  const isWrite = ["POST","PATCH","PUT","DELETE"].includes((fetchOpts.method||"GET").toUpperCase());
  const preferHeader = prefer || (isWrite ? "return=representation" : "");
  const res = await fetch(`${SUPA_URL}/rest/v1/${path}`, {
    headers: {
      "apikey": SUPA_KEY,
      "Authorization": `Bearer ${SUPA_KEY}`,
      "Content-Type": "application/json",
      ...(preferHeader ? { "Prefer": preferHeader } : {}),
      ...(extraHeaders || {}),
    },
    ...fetchOpts,
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(err);
  }
  const text = await res.text();
  if (!text || text === "null") return null;
  try { return JSON.parse(text); } catch { return null; }
};

// ── Storage Keys (for session only) ──
const SESSION_KEY = "tw_stock_session";

const BASE_FEE_RATE = 0.001425;
const calcFee = (qty, price, disc) =>
  Math.max(1, Math.round(qty * price * BASE_FEE_RATE * disc));

const AUTO_REFRESH_MS = 5 * 60 * 1000;
const PRICE_TS_KEY = (uid) => `tw_price_ts_${uid}`;

const BROKERS = [
  { name: "元大證券",         defaultDiscount: 0.6  },
  { name: "富邦證券",         defaultDiscount: 0.6  },
  { name: "國泰證券",         defaultDiscount: 0.6  },
  { name: "凱基證券",         defaultDiscount: 0.6  },
  { name: "永豐金證券",       defaultDiscount: 0.6  },
  { name: "玉山證券",         defaultDiscount: 0.6  },
  { name: "中信證券",         defaultDiscount: 0.6  },
  { name: "台新證券",         defaultDiscount: 0.6  },
  { name: "第一金證券",       defaultDiscount: 0.6  },
  { name: "華南永昌",         defaultDiscount: 0.6  },
  { name: "兆豐證券",         defaultDiscount: 0.6  },
  { name: "群益證券",         defaultDiscount: 0.65 },
  { name: "統一證券",         defaultDiscount: 0.65 },
  { name: "新光證券",         defaultDiscount: 0.6  },
  { name: "網路券商(無折扣)", defaultDiscount: 1.0  },
  { name: "其他",             defaultDiscount: 0.85 },
];

const fmtNum   = (n, d = 0) =>
  Number(n).toLocaleString("zh-TW", { minimumFractionDigits: d, maximumFractionDigits: d });
const fmtMoney = (n) => `${n >= 0 ? "+" : ""}${fmtNum(n)}`;

const simpleHash = (s) => {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (Math.imul(31, h) + s.charCodeAt(i)) | 0;
  return h.toString(36);
};

const STOCK_DB = {
  "2330":"台積電","2303":"聯電","2454":"聯發科","3711":"日月光投控","2379":"瑞昱",
  "3034":"聯詠","2408":"南亞科","2344":"華邦電","2449":"京元電子","6415":"矽力-KY",
  "5269":"祥碩","3443":"創意電子","3529":"力旺","4966":"譜瑞-KY","3661":"世芯-KY",
  "2388":"威盛","2368":"金像電","6245":"立端","3037":"欣興","2367":"燿華",
  "2317":"鴻海","2382":"廣達","3231":"緯創","2324":"仁寶","4938":"和碩",
  "2357":"華碩","2353":"宏碁","2308":"台達電","2395":"研華","2327":"國巨",
  "2301":"光寶科","2376":"技嘉","2377":"微星","2356":"英業達","2352":"佳世達",
  "2409":"友達","3481":"群創","2498":"宏達電",
  "2412":"中華電","3045":"台灣大","4904":"遠傳",
  "2882":"國泰金","2881":"富邦金","2891":"中信金","2886":"兆豐金","2884":"玉山金",
  "2892":"第一金","2880":"華南金","2887":"台新金","2890":"永豐金","2885":"元大金",
  "2883":"開發金","2888":"新光金","5880":"合庫金",
  "1301":"台塑","1303":"南亞","1326":"台化","6505":"台塑化",
  "1402":"遠東新","1476":"儒鴻","1477":"聚陽",
  "2002":"中鋼","1101":"台泥","1102":"亞泥",
  "1216":"統一","2912":"統一超","2903":"遠百","5903":"全家",
  "2603":"長榮","2609":"陽明","2615":"萬海","2610":"華航","2618":"長榮航",
  "3008":"大立光","2474":"可成","6669":"緯穎","2354":"鴻準",
  "6176":"瑞儀","6202":"盛群","6271":"同欣電","6278":"台表科",
  "0050":"元大台灣50","0056":"元大高股息","00878":"國泰永續高股息","00881":"國泰台灣5G+",
  "00900":"富邦特選高股息30","00919":"群益台灣精選高息","00929":"復華台灣科技優息",
  "006208":"富邦台灣采吉50","00713":"元大台灣高息低波","00830":"國泰費城半導體",
  "00850":"元大臺灣ESG永續","00861":"元大全球AI","00891":"中信關鍵半導體",
  "00894":"中信小資高息ETF","00927":"群益半導體收益","00930":"永豐智能車供應鏈",
  "AAPL":"Apple","MSFT":"Microsoft","GOOGL":"Alphabet","AMZN":"Amazon",
  "NVDA":"NVIDIA","META":"Meta","TSLA":"Tesla","AVGO":"Broadcom",
  "TSM":"台積電ADR","AMD":"AMD","INTC":"Intel","QCOM":"Qualcomm",
  "SPY":"S&P 500 ETF","QQQ":"Nasdaq 100 ETF","VTI":"全市場ETF",
};

const NAME_TO_CODE = Object.fromEntries(Object.entries(STOCK_DB).map(([c, n]) => [n, c]));
const lookupNameByCode = (code) => STOCK_DB[String(code || "").trim()] || null;
const lookupCodeByName = (name) => {
  if (!name) return null;
  const n = name.trim();
  if (NAME_TO_CODE[n]) return NAME_TO_CODE[n];
  for (const [nm, cd] of Object.entries(NAME_TO_CODE)) {
    if (n.includes(nm) || nm.includes(n)) return cd;
  }
  return null;
};

const toYahooSymbol = (code) => {
  const c = String(code).trim().toUpperCase();
  if (/^[A-Z]{1,5}$/.test(c)) return c;
  if (/^\d{4,6}[A-Z]?$/.test(c)) return `${c}.TW`;
  return c;
};

const matchHeader = (h) => {
  const s = (h || "").replace(/[\s"'\u3000\r\n]/g, "").toLowerCase();
  if (/^(日期|成交日|交割日|date|tradedate|交易日)/.test(s)) return "date";
  if (/^(股票代號|代號|証券代號|代碼|stockcode|code|symbol|ticker|證券代號)/.test(s)) return "stock";
  if (/^(股票名稱|名稱|証券名稱|stockname|name|商品名稱|證券名稱)/.test(s)) return "name";
  if (/^(類型|買賣別|交易別|買賣|type|action|side|buysell)/.test(s)) return "type";
  if (/^(數量|成交股數|股數|qty|quantity|shares|成交量)/.test(s)) return "qty";
  if (/^(成交價|成交價格|每股價格|price|unitprice|tradeprice|單價)/.test(s)) return "price";
  if (/^(手續費|佣金|fee|commission|brokerage)/.test(s)) return "fee";
  if (/^(交易稅|證交稅|tax|transactiontax|稅額)/.test(s)) return "tax";
  return null;
};

const parseType = (v) => {
  if (!v) return "buy";
  const s = String(v).trim().toLowerCase();
  if (["賣", "賣出", "sell", "s", "sale"].some((k) => s.includes(k))) return "sell";
  return "buy";
};

const parseDate = (v) => {
  if (v === null || v === undefined) return "";
  if (typeof v === "number" && v > 1000) {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
  }
  const s = String(v).trim();
  if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const roc = s.match(/^(\d{2,3})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (roc) { const y = parseInt(roc[1]) + 1911; return `${y}-${roc[2].padStart(2, "0")}-${roc[3].padStart(2, "0")}`; }
  if (/^\d{8}$/.test(s)) return `${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6, 8)}`;
  if (/^\d{7}$/.test(s)) { const y = parseInt(s.slice(0, 3)) + 1911; return `${y}-${s.slice(3, 5)}-${s.slice(5, 7)}`; }
  return s;
};

const parseNum = (v) => {
  if (typeof v === "number") return v;
  return parseFloat(String(v || "0").replace(/,/g, "")) || 0;
};

const parseCsvLine = (line) => {
  const cols = []; let cur = ""; let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') { if (inQ && line[i + 1] === '"') { cur += '"'; i++; } else inQ = !inQ; }
    else if (c === ',' && !inQ) { cols.push(cur.trim()); cur = ""; }
    else cur += c;
  }
  cols.push(cur.trim()); return cols;
};

const csvTextToRows = (text) => {
  let t = text;
  if (t.charCodeAt(0) === 0xFEFF) t = t.slice(1);
  return t.trim().split(/\r?\n/).map(parseCsvLine);
};

const xlsxBufToRows = (buf) => {
  const wb = XLSX.read(buf, { type: "array", cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
};

const rowsToTrades = (rows, filename = "") => {
  if (rows.length < 2) throw new Error(`「${filename}」沒有足夠的資料列`);
  let headerIdx = 0;
  for (let i = 0; i < Math.min(8, rows.length); i++) {
    const matched = rows[i].filter((c) => matchHeader(String(c ?? "")) !== null).length;
    if (matched >= 2) { headerIdx = i; break; }
  }
  const header = rows[headerIdx].map((c) => String(c ?? ""));
  const idxMap = {};
  header.forEach((h, i) => { const k = matchHeader(h); if (k && idxMap[k] === undefined) idxMap[k] = i; });
  if (idxMap.qty === undefined || idxMap.price === undefined)
    throw new Error(`「${filename}」缺少必要欄位（數量/成交價）`);
  if (idxMap.stock === undefined && idxMap.name === undefined)
    throw new Error(`「${filename}」找不到股票代號或名稱欄位`);
  const results = [];
  rows.slice(headerIdx + 1).forEach((row, idx) => {
    if (!row || row.every((c) => c === null || c === undefined || c === "")) return;
    const get = (key) => row[idxMap[key]] ?? null;
    const rawStock = idxMap.stock !== undefined ? String(get("stock") ?? "").replace(/\s/g, "") : "";
    const rawName  = idxMap.name  !== undefined ? String(get("name")  ?? "").trim() : "";
    let stock = rawStock;
    if (!stock && rawName) stock = lookupCodeByName(rawName) || rawName;
    const name  = rawName || stock;
    const qty   = parseNum(get("qty"));
    const price = parseNum(get("price"));
    if (qty === 0 || price === 0) return;
    const date = parseDate(idxMap.date !== undefined ? get("date") : null);
    const type = parseType(idxMap.type !== undefined ? get("type") : null);
    const fee  = idxMap.fee !== undefined ? parseNum(get("fee")) : 0;
    const tax  = idxMap.tax !== undefined ? parseNum(get("tax")) : 0;
    results.push({ id: Date.now() + idx * 7, date, stock, name, type, qty, price, fee, tax });
  });
  if (results.length === 0) throw new Error(`「${filename}」沒有解析到有效的交易資料`);
  return results;
};

const csvEscape = (v) => {
  const s = String(v ?? "");
  return s.includes(",") || s.includes('"') ? `"${s.replace(/"/g, '""')}"` : s;
};
const tradesToCsv = (trades) => {
  const hdr = ["日期", "股票代號", "股票名稱", "類型", "數量", "成交價", "手續費", "交易稅"];
  const rows = trades.map((t) => [t.date, t.stock, t.name, t.type === "buy" ? "買進" : "賣出", t.qty, t.price, t.fee, t.tax || 0]);
  return [hdr, ...rows].map((r) => r.map(csvEscape).join(",")).join("\n");
};

const fetchQuotes = async (codes) => {
  if (!codes || codes.length === 0) return {};
  const symbols = codes.map(toYahooSymbol).join(",");
  try {
    const res = await fetch(`/api/quote?symbols=${encodeURIComponent(symbols)}`);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (!data.quotes) return {};
    const result = {};
    codes.forEach((code) => {
      const ySym = toYahooSymbol(code);
      if (data.quotes[ySym]) result[code] = data.quotes[ySym];
    });
    return result;
  } catch (e) {
    console.warn("fetchQuotes error:", e);
    return {};
  }
};

// ═══════════════════════
//  AuthScreen
// ═══════════════════════
function AuthScreen({ onLogin }) {
  const [tab, setTab]   = useState("login");
  const [f, setF]       = useState({ username: "", password: "", display: "", confirm: "", broker: BROKERS[0].name, feeDiscount: BROKERS[0].defaultDiscount });
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [shake, setShake] = useState(false);
  const doShake = () => { setShake(true); setTimeout(() => setShake(false), 500); };
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const handleBroker = (name) => {
    const b = BROKERS.find((x) => x.name === name);
    setF((p) => ({ ...p, broker: name, feeDiscount: b?.defaultDiscount ?? p.feeDiscount }));
  };

  const doLogin = async () => {
    setError(""); setLoading(true);
    try {
      const rows = await supa(`users?username=eq.${encodeURIComponent(f.username.trim())}&select=*`);
      const user = rows?.[0];
      if (!user || user.password_hash !== simpleHash(f.password)) {
        setError("帳號或密碼錯誤"); doShake(); setLoading(false); return;
      }
      const sess = { id: user.id, username: user.username, displayName: user.display_name, broker: user.broker, feeDiscount: Number(user.fee_discount) };
      localStorage.setItem(SESSION_KEY, JSON.stringify(sess));
      onLogin(sess);
    } catch (e) {
      setError("連線失敗：" + e.message); doShake();
    }
    setLoading(false);
  };

  const doRegister = async () => {
    setError(""); 
    if (!f.username.trim() || !f.password || !f.display.trim()) { setError("請填入所有必要欄位"); doShake(); return; }
    if (f.password.length < 6) { setError("密碼至少需要 6 個字元"); doShake(); return; }
    if (f.password !== f.confirm) { setError("兩次輸入密碼不一致"); doShake(); return; }
    const d = Number(f.feeDiscount);
    if (isNaN(d) || d <= 0 || d > 1) { setError("折數請輸入 0.01 ~ 1.00"); doShake(); return; }
    setLoading(true);
    try {
      const existing = await supa(`users?username=eq.${encodeURIComponent(f.username.trim())}&select=id`);
      if (existing?.length > 0) { setError("此帳號名稱已被使用"); doShake(); setLoading(false); return; }
      const uid = `u${Date.now()}`;
      await supa("users", {
        method: "POST",
        prefer: "return=minimal",
        body: JSON.stringify({ id: uid, username: f.username.trim(), password_hash: simpleHash(f.password), display_name: f.display.trim(), broker: f.broker, fee_discount: d }),
      });
      const sess = { id: uid, username: f.username.trim(), displayName: f.display.trim(), broker: f.broker, feeDiscount: d };
      localStorage.setItem(SESSION_KEY, JSON.stringify(sess));
      onLogin(sess);
    } catch (e) {
      setError("註冊失敗：" + e.message); doShake();
    }
    setLoading(false);
  };

  const prevFee = !isNaN(Number(f.feeDiscount))
    ? Math.max(1, Math.round(100000 * BASE_FEE_RATE * Number(f.feeDiscount))) : "—";

  return (
    <div style={{ fontFamily: "'Noto Sans TC','Segoe UI',sans-serif", background: "#0d1117", minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&family=JetBrains+Mono:wght@400;600&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        .ai{background:#0d1117;border:1px solid #30363d;color:#e6edf3;border-radius:8px;padding:11px 14px;font-size:14px;font-family:inherit;outline:none;width:100%;transition:border-color .15s}
        .ai:focus{border-color:#58a6ff;box-shadow:0 0 0 3px #58a6ff18}
        .ab{width:100%;padding:12px;border:none;border-radius:8px;font-family:inherit;font-size:14px;font-weight:600;cursor:pointer;background:#238636;color:#fff;transition:background .15s}
        .ab:hover{background:#2ea043}.ab:disabled{opacity:.5;cursor:not-allowed}
        @keyframes shake{0%,100%{transform:translateX(0)}20%,60%{transform:translateX(-8px)}40%,80%{transform:translateX(8px)}}
        .shake{animation:shake .4s ease}
        @keyframes fadeUp{from{opacity:0;transform:translateY(16px)}to{opacity:1;transform:translateY(0)}}
        .fadeUp{animation:fadeUp .35s ease}
      `}</style>
      <div className={`fadeUp${shake ? " shake" : ""}`} style={{ background: "#161b22", border: "1px solid #21262d", borderRadius: 14, width: "100%", maxWidth: 420, padding: "36px 32px", boxShadow: "0 16px 48px rgba(0,0,0,.5)" }}>
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <div style={{ fontSize: 36, marginBottom: 8 }}>📈</div>
          <div style={{ fontSize: 20, fontWeight: 700 }}>台股追蹤</div>
          <div style={{ fontSize: 12, color: "#484f58", marginTop: 3 }}>Portfolio Tracker</div>
        </div>
        <div style={{ display: "flex", background: "#0d1117", borderRadius: 8, padding: 4, marginBottom: 24 }}>
          {["login", "register"].map((t) => (
            <button key={t} onClick={() => { setTab(t); setError(""); }}
              style={{ flex: 1, padding: "8px 0", border: "none", borderRadius: 6, cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 600, transition: "all .15s", background: tab === t ? "#21262d" : "transparent", color: tab === t ? "#e6edf3" : "#484f58" }}>
              {t === "login" ? "登入" : "註冊帳號"}
            </button>
          ))}
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          {tab === "register" && (
            <div>
              <label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 5 }}>顯示名稱 *</label>
              <input className="ai" placeholder="例：小明" value={f.display} onChange={(e) => set("display", e.target.value)} />
            </div>
          )}
          <div>
            <label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 5 }}>帳號 *</label>
            <input className="ai" placeholder="輸入帳號" value={f.username} onChange={(e) => set("username", e.target.value)} onKeyDown={(e) => e.key === "Enter" && tab === "login" && doLogin()} />
          </div>
          <div>
            <label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 5 }}>密碼 *</label>
            <input className="ai" type="password" placeholder={tab === "register" ? "至少 6 個字元" : "輸入密碼"} value={f.password} onChange={(e) => set("password", e.target.value)} onKeyDown={(e) => e.key === "Enter" && tab === "login" && doLogin()} />
          </div>
          {tab === "register" && (
            <>
              <div>
                <label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 5 }}>確認密碼 *</label>
                <input className="ai" type="password" placeholder="再輸入一次" value={f.confirm} onChange={(e) => set("confirm", e.target.value)} />
              </div>
              <div style={{ background: "#0d1117", borderRadius: 10, padding: 16, border: "1px solid #21262d" }}>
                <div style={{ fontSize: 12, color: "#58a6ff", fontWeight: 600, marginBottom: 12 }}>🏦 券商手續費設定</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  <div>
                    <label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 5 }}>使用券商</label>
                    <select className="ai" style={{ padding: "11px 14px" }} value={f.broker} onChange={(e) => handleBroker(e.target.value)}>
                      {BROKERS.map((b) => <option key={b.name} value={b.name}>{b.name}</option>)}
                    </select>
                  </div>
                  <div>
                    <label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 5 }}>手續費折數（0.01 ~ 1.00）</label>
                    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                      <input className="ai" type="number" step="0.01" min="0.01" max="1" value={f.feeDiscount} onChange={(e) => set("feeDiscount", e.target.value)} style={{ flex: 1 }} />
                      <span style={{ fontSize: 12, color: "#8b949e", whiteSpace: "nowrap" }}>{f.feeDiscount ? `${Math.round(Number(f.feeDiscount) * 10)}折` : "—"}</span>
                    </div>
                    <div style={{ fontSize: 11, color: "#484f58", marginTop: 5 }}>10 萬元手續費預估：<span style={{ color: "#e6edf3", fontFamily: "monospace" }}>NT$ {prevFee}</span></div>
                  </div>
                </div>
              </div>
            </>
          )}
          {error && <div style={{ background: "#3a1a1a", border: "1px solid #f8514940", borderRadius: 6, padding: "9px 13px", fontSize: 13, color: "#f85149" }}>⚠️ {error}</div>}
          <button className="ab" disabled={loading} onClick={tab === "login" ? doLogin : doRegister}>
            {loading ? "處理中…" : tab === "login" ? "登入" : "建立帳號"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════
//  Main App
// ═══════════════════════
export default function App() {
  const [currentUser, setCurrentUser] = useState(() => {
    try { const s = localStorage.getItem(SESSION_KEY); return s ? JSON.parse(s) : null; } catch { return null; }
  });
  const uid = currentUser?.id;
  const feeDiscount = currentUser?.feeDiscount ?? 0.85;

  const [trades, setTrades] = useState([]);
  const [tradesLoading, setTradesLoading] = useState(false);

  const [livePrices, setLivePrices]   = useState({});
  const [priceLoading, setPriceLoading] = useState(false);
  const [priceError, setPriceError]   = useState("");
  const [lastFetched, setLastFetched] = useState(() => {
    try { return localStorage.getItem(PRICE_TS_KEY(uid)) || ""; } catch { return ""; }
  });

  const [view, setView]             = useState("portfolio");
  const [showForm, setShowForm]     = useState(false);
  const [editTrade, setEditTrade]   = useState(null);
  const [filterStock, setFilter]    = useState("");
  const [sortBy, setSortBy]         = useState("date");
  const [notification, setNotif]   = useState(null);
  const [showUserMenu, setUserMenu] = useState(false);
  const [showSettings, setSettings] = useState(false);
  const [settingsTab, setStab]      = useState("broker");
  const [sf, setSf]                 = useState({ broker: "", feeDiscount: "", oldPw: "", newPw: "", confirmPw: "" });
  const [sfErr, setSfErr]           = useState("");
  const [selectedYear, setYear]     = useState("all");
  const [importError, setImportError] = useState("");
  const [selectedIds, setSelectedIds] = useState(new Set());

  const today = new Date().toISOString().slice(0, 10);
  const [rangeStart, setRangeStart] = useState(today.slice(0, 4) + "-01-01");
  const [rangeEnd,   setRangeEnd]   = useState(today);

  const emptyForm = () => ({ date: today, stock: "", name: "", type: "buy", qty: "", price: "", fee: "", tax: "" });
  const [form, setForm]             = useState(emptyForm());
  const [stockLookup, setStockLookup] = useState({ loading: false, error: "" });

  const fileInputRef = useRef();
  const autoRefreshTimer = useRef(null);

  // ── Load trades from Supabase ──
  const loadTrades = useCallback(async () => {
    if (!uid) return;
    setTradesLoading(true);
    try {
      const rows = await supa(`trades?user_id=eq.${uid}&order=date.desc&select=*`);
      setTrades((rows || []).map((r) => ({
        id: r.id, date: r.date, stock: r.stock, name: r.name,
        type: r.type, qty: Number(r.qty), price: Number(r.price),
        fee: Number(r.fee), tax: Number(r.tax),
      })));
    } catch (e) {
      notify("載入資料失敗：" + e.message, "error");
    }
    setTradesLoading(false);
  }, [uid]);

  useEffect(() => { if (uid) loadTrades(); }, [uid]);

  const heldCodes = useMemo(() => [...new Set(trades.map((t) => t.stock))], [trades]);

  const refreshPrices = useCallback(async (silent = false) => {
    if (heldCodes.length === 0) return;
    if (!silent) setPriceLoading(true);
    setPriceError("");
    try {
      const quotes = await fetchQuotes(heldCodes);
      if (Object.keys(quotes).length > 0) {
        setLivePrices(quotes);
        const ts = new Date().toLocaleString("zh-TW", { hour12: false });
        setLastFetched(ts);
        try { localStorage.setItem(PRICE_TS_KEY(uid), ts); } catch (_) {}
      } else {
        setPriceError("無法取得報價，可能是盤後時間或網路問題");
      }
    } catch (e) {
      setPriceError("報價取得失敗：" + e.message);
    }
    if (!silent) setPriceLoading(false);
  }, [heldCodes, uid]);

  useEffect(() => {
    if (!uid || heldCodes.length === 0) return;
    refreshPrices(false);
    autoRefreshTimer.current = setInterval(() => refreshPrices(true), AUTO_REFRESH_MS);
    return () => clearInterval(autoRefreshTimer.current);
  }, [uid, heldCodes.join(",")]); // eslint-disable-line

  const notify = (msg, type = "success") => { setNotif({ msg, type }); setTimeout(() => setNotif(null), 4000); };

  const lookupStock = useCallback((code) => {
    if (!code) return;
    const name = lookupNameByCode(code);
    if (name) { setForm((f) => ({ ...f, name })); setStockLookup({ loading: false, error: "" }); }
    else if (code.length >= 2) { setStockLookup({ loading: false, error: "代號不在資料庫，請手動輸入名稱" }); }
  }, []);

  // ── Portfolio calc ──
  const portfolio = useMemo(() => {
    const map = {};
    trades.forEach((t) => {
      if (!map[t.stock]) map[t.stock] = { stock: t.stock, name: t.name, qty: 0, cost: 0, realizedPnl: 0, realizedByYear: {}, buyCount: 0 };
      const p    = map[t.stock];
      const year = (t.date || "").slice(0, 4);
      if (t.type === "buy") {
        p.cost += t.qty * t.price + (t.fee || 0); p.qty += t.qty; p.buyCount++;
        if (t.name && t.name !== t.stock) p.name = t.name;
      } else {
        const avg = p.qty > 0 ? p.cost / p.qty : 0;
        const pnl = t.qty * t.price - (t.fee || 0) - (t.tax || 0) - avg * t.qty;
        p.realizedPnl += pnl; p.realizedByYear[year] = (p.realizedByYear[year] || 0) + pnl;
        p.cost -= avg * t.qty; p.qty -= t.qty;
      }
    });
    return Object.values(map).filter((p) => p.qty > 0 || p.realizedPnl !== 0).map((p) => {
      const avgCost  = p.qty > 0 ? p.cost / p.qty : 0;
      const liveData = livePrices[p.stock];
      const cur      = liveData ? liveData.price : avgCost;
      const mktVal   = p.qty * cur;
      const unreal   = mktVal - p.cost;
      const dbName   = lookupNameByCode(p.stock);
      const dispName = (p.name && p.name !== p.stock) ? p.name : (dbName || p.stock);
      return { ...p, name: dispName, avgCost, curPrice: cur, mktVal, unrealizedPnl: unreal, unrealizedPct: p.cost > 0 ? (unreal / p.cost) * 100 : 0, hasLivePrice: !!liveData, changePct: liveData?.changePct ?? null, currency: liveData?.currency ?? "TWD", marketState: liveData?.marketState ?? "CLOSED" };
    });
  }, [trades, livePrices]);

  const years = useMemo(() => {
    const ys = new Set(trades.map((t) => (t.date || "").slice(0, 4)).filter(Boolean));
    return ["all", ...Array.from(ys).sort().reverse()];
  }, [trades]);

  const totalMktVal     = portfolio.reduce((s, p) => s + p.mktVal, 0);
  const totalCost       = portfolio.reduce((s, p) => s + p.cost, 0);
  const totalUnrealized = totalMktVal - totalCost;
  const totalRealized   = portfolio.reduce((s, p) => s + p.realizedPnl, 0);

  const yearRealized = useMemo(() =>
    selectedYear === "all" ? totalRealized : portfolio.reduce((s, p) => s + (p.realizedByYear[selectedYear] || 0), 0),
    [portfolio, selectedYear, totalRealized]);

  const yearUnrealized = useMemo(() => {
    if (selectedYear === "all") return totalUnrealized;
    const bY = trades.filter((t) => t.type === "buy" && t.date.startsWith(selectedYear));
    const cY = bY.reduce((s, t) => s + t.qty * t.price + (t.fee || 0), 0);
    const tC = trades.filter((t) => t.type === "buy").reduce((s, t) => s + t.qty * t.price + (t.fee || 0), 0);
    return tC === 0 ? 0 : totalUnrealized * (cY / tC);
  }, [trades, selectedYear, totalUnrealized]);

  // ── 區間損益 ──
  const rangeStats = useMemo(() => {
    if (!rangeStart || !rangeEnd) return null;
    const inRange = (d) => d >= rangeStart && d <= rangeEnd;
    let realized = 0;
    const stockMap = {};
    trades.forEach((t) => {
      if (!stockMap[t.stock]) stockMap[t.stock] = { qty: 0, cost: 0 };
      const sm = stockMap[t.stock];
      if (t.type === "buy") { sm.cost += t.qty * t.price + (t.fee || 0); sm.qty += t.qty; }
      else {
        const avg = sm.qty > 0 ? sm.cost / sm.qty : 0;
        const pnl = t.qty * t.price - (t.fee || 0) - (t.tax || 0) - avg * t.qty;
        if (inRange(t.date)) realized += pnl;
        sm.cost -= avg * t.qty; sm.qty -= t.qty;
      }
    });
    const buyInRange = {};
    trades.filter((t) => t.type === "buy" && inRange(t.date)).forEach((t) => {
      if (!buyInRange[t.stock]) buyInRange[t.stock] = { qty: 0, cost: 0, name: t.name };
      buyInRange[t.stock].qty  += t.qty;
      buyInRange[t.stock].cost += t.qty * t.price + (t.fee || 0);
    });
    trades.filter((t) => t.type === "sell").forEach((t) => {
      if (buyInRange[t.stock]) {
        const sm = buyInRange[t.stock];
        const avgCost = sm.qty > 0 ? sm.cost / sm.qty : 0;
        sm.cost -= avgCost * Math.min(t.qty, sm.qty);
        sm.qty  -= Math.min(t.qty, sm.qty);
        if (sm.qty < 0) sm.qty = 0;
      }
    });
    let unrealCost = 0, unrealMkt = 0;
    const unrealRows = Object.entries(buyInRange).filter(([, v]) => v.qty > 0).map(([code, v]) => {
      const lp = livePrices[code];
      const cur = lp ? lp.price : (v.cost / v.qty);
      const mkt = v.qty * cur;
      const pnl = mkt - v.cost;
      const pct = v.cost > 0 ? (pnl / v.cost) * 100 : 0;
      unrealCost += v.cost; unrealMkt += mkt;
      const dbName = lookupNameByCode(code);
      return { code, name: (v.name && v.name !== code) ? v.name : (dbName || code), qty: v.qty, cost: v.cost, mkt, pnl, pct, hasLive: !!lp };
    });
    const unrealized = unrealMkt - unrealCost;
    const unrealPct  = unrealCost > 0 ? (unrealized / unrealCost) * 100 : 0;
    return { realized, unrealized, unrealPct, unrealRows, unrealCost };
  }, [trades, rangeStart, rangeEnd, livePrices]);

  // ── 賣出損益 ──
  const tradesWithPnl = useMemo(() => {
    const stockMap = {};
    return trades.slice().sort((a, b) => a.date.localeCompare(b.date)).map((t) => {
      if (!stockMap[t.stock]) stockMap[t.stock] = { qty: 0, cost: 0 };
      const sm = stockMap[t.stock];
      if (t.type === "buy") { sm.cost += t.qty * t.price + (t.fee || 0); sm.qty += t.qty; return { ...t, sellPnl: null, sellPct: null }; }
      else {
        const avg = sm.qty > 0 ? sm.cost / sm.qty : 0;
        const pnl = t.qty * t.price - (t.fee || 0) - (t.tax || 0) - avg * t.qty;
        const pct = avg > 0 ? ((t.price - avg) / avg) * 100 : 0;
        sm.cost -= avg * Math.min(t.qty, sm.qty); sm.qty -= Math.min(t.qty, sm.qty);
        if (sm.qty < 0) sm.qty = 0;
        return { ...t, sellPnl: pnl, sellPct: pct };
      }
    });
  }, [trades]);

  // ── Trade CRUD ──
  const submitTrade = async () => {
    if (!form.stock || !form.qty || !form.price) { notify("請填入必要欄位", "error"); return; }
    const fee = form.fee !== "" ? Number(form.fee) : calcFee(Number(form.qty), Number(form.price), feeDiscount);
    const t   = { ...form, qty: Number(form.qty), price: Number(form.price), fee, tax: Number(form.tax || 0) };
    try {
      if (editTrade) {
        await supa(`trades?id=eq.${editTrade.id}`, { method: "PATCH", prefer: "return=minimal", body: JSON.stringify({ ...t, user_id: uid }) });
        notify("已更新");
      } else {
        const id = Date.now();
        await supa("trades", { method: "POST", prefer: "return=minimal", body: JSON.stringify({ ...t, id, user_id: uid }) });
        notify("已新增 ✓");
      }
      await loadTrades();
    } catch (e) { notify("儲存失敗：" + e.message, "error"); return; }
    setShowForm(false); setEditTrade(null); setForm(emptyForm()); setStockLookup({ loading: false, error: "" });
  };

  const deleteTrade = async (id) => {
    try {
      await supa(`trades?id=eq.${id}`, { method: "DELETE" });
      setTrades((p) => p.filter((t) => t.id !== id));
      notify("已刪除");
    } catch (e) { notify("刪除失敗：" + e.message, "error"); }
  };

  const openEdit = (t) => { setEditTrade(t); setForm({ ...t, fee: t.fee ?? "" }); setShowForm(true); };

  const exportCsv = () => {
    const blob = new Blob(["\uFEFF" + tradesToCsv(trades)], { type: "text/csv;charset=utf-8;" });
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob);
    a.download = `台股交易記錄_${today}.csv`; a.click();
  };

  const importFile = (e) => {
    setImportError("");
    const files = Array.from(e.target.files || []);
    if (!files.length) return;
    let total = 0; const errors = [];
    const processRows = async (rows, filename) => {
      try {
        const imported = rowsToTrades(rows, filename);
        const existing = new Set(trades.map((t) => `${t.date}|${t.stock}|${t.qty}|${t.price}|${t.type}`));
        const fresh = imported.filter((t) => !existing.has(`${t.date}|${t.stock}|${t.qty}|${t.price}|${t.type}`));
        for (const t of fresh) {
          await supa("trades", { method: "POST", prefer: "return=minimal", body: JSON.stringify({ ...t, id: Date.now() + Math.random() * 1000 | 0, user_id: uid }) });
          total++;
        }
      } catch (err) { errors.push(err.message); }
    };
    const readFile = (file) => new Promise((resolve) => {
      const isExcel = /\.(xlsx|xls|xlsm|ods)$/i.test(file.name);
      const reader  = new FileReader();
      if (isExcel) {
        reader.onload  = async (ev) => { try { await processRows(xlsxBufToRows(new Uint8Array(ev.target.result)), file.name); } catch (err) { errors.push(`「${file.name}」Excel 讀取失敗`); } resolve(); };
        reader.readAsArrayBuffer(file);
      } else {
        reader.onload  = async (ev) => { try { let t = ev.target.result; if (t.charCodeAt(0) === 0xFEFF) t = t.slice(1); await processRows(csvTextToRows(t), file.name); } catch (err) { errors.push(err.message); } resolve(); };
        reader.readAsText(file, "UTF-8");
      }
    });
    Promise.all(files.map(readFile)).then(async () => {
      await loadTrades();
      if (total > 0) notify(`✓ 成功匯入 ${total} 筆交易`);
      if (errors.length) setImportError(errors.join(" ／ "));
      else if (total === 0) setImportError("沒有新資料被匯入");
    });
    e.target.value = "";
  };

  const openSettings = () => {
    setSf({ broker: currentUser.broker || "其他", feeDiscount: currentUser.feeDiscount ?? 0.85, oldPw: "", newPw: "", confirmPw: "" });
    setSfErr(""); setStab("broker"); setSettings(true); setUserMenu(false);
  };
  const saveBroker = async () => {
    setSfErr(""); const d = Number(sf.feeDiscount);
    if (isNaN(d) || d <= 0 || d > 1) { setSfErr("折數請輸入 0.01 ~ 1.00"); return; }
    try {
      await supa(`users?id=eq.${uid}`, { method: "PATCH", prefer: "return=minimal", body: JSON.stringify({ broker: sf.broker, fee_discount: d }) });
      const upd = { ...currentUser, broker: sf.broker, feeDiscount: d };
      setCurrentUser(upd); localStorage.setItem(SESSION_KEY, JSON.stringify(upd));
      notify("券商設定已更新 ✓"); setSettings(false);
    } catch (e) { setSfErr("儲存失敗：" + e.message); }
  };
  const changePw = async () => {
    setSfErr("");
    if (!sf.oldPw || !sf.newPw || !sf.confirmPw) { setSfErr("請填入所有欄位"); return; }
    if (sf.newPw.length < 6) { setSfErr("新密碼至少需要 6 個字元"); return; }
    if (sf.newPw !== sf.confirmPw) { setSfErr("兩次密碼不一致"); return; }
    try {
      const rows = await supa(`users?id=eq.${uid}&select=password_hash`);
      if (!rows?.[0] || rows[0].password_hash !== simpleHash(sf.oldPw)) { setSfErr("舊密碼不正確"); return; }
      await supa(`users?id=eq.${uid}`, { method: "PATCH", prefer: "return=minimal", body: JSON.stringify({ password_hash: simpleHash(sf.newPw) }) });
      notify("密碼已更新 ✓"); setSettings(false);
    } catch (e) { setSfErr("更新失敗：" + e.message); }
  };
  const handleLogout = () => { localStorage.removeItem(SESSION_KEY); setCurrentUser(null); setUserMenu(false); };

  if (!currentUser) {
    return <AuthScreen onLogin={(u) => { setCurrentUser(u); setView("portfolio"); }} />;
  }

  const filteredTrades = [...tradesWithPnl]
    .filter((t) => !filterStock || t.stock.includes(filterStock) || t.name.includes(filterStock))
    .sort((a, b) => sortBy === "date" ? b.date.localeCompare(a.date) : a.stock.localeCompare(b.stock));

  const prevFeeAmt = form.qty && form.price
    ? (form.fee !== "" ? Number(form.fee) : calcFee(Number(form.qty), Number(form.price), feeDiscount)) : null;
  const stPreviewFee = sf.feeDiscount && !isNaN(Number(sf.feeDiscount))
    ? Math.max(1, Math.round(100000 * BASE_FEE_RATE * Number(sf.feeDiscount))) : "—";

  return (
    <div style={{ fontFamily: "'Noto Sans TC','Segoe UI',sans-serif", background: "#0d1117", minHeight: "100vh", color: "#e6edf3" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&family=JetBrains+Mono:wght@400;600&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:6px}::-webkit-scrollbar-track{background:#161b22}::-webkit-scrollbar-thumb{background:#30363d;border-radius:3px}
        .btn{border:none;cursor:pointer;border-radius:6px;font-family:inherit;font-size:13px;font-weight:500;transition:all .15s}
        .btn-p{background:#238636;color:#fff;padding:7px 16px}.btn-p:hover{background:#2ea043}
        .btn-d{background:transparent;color:#f85149;border:1px solid #f85149;padding:4px 10px}.btn-d:hover{background:#f8514922}
        .btn-g{background:transparent;color:#8b949e;border:1px solid #30363d;padding:4px 10px}.btn-g:hover{background:#21262d;color:#e6edf3}
        .btn-s{background:#da3633;color:#fff;padding:7px 16px}.btn-s:hover{background:#f85149}
        .btn-b{background:#1f6feb;color:#fff;padding:6px 13px;font-size:12px}.btn-b:hover{background:#388bfd}
        input,select{background:#161b22;border:1px solid #30363d;color:#e6edf3;border-radius:6px;padding:8px 12px;font-size:13px;font-family:inherit;outline:none;width:100%;transition:border-color .15s}
        input:focus,select:focus{border-color:#58a6ff} select option{background:#161b22}
        .card{background:#161b22;border:1px solid #21262d;border-radius:10px;padding:20px}
        .tag-b{background:#1a3a2a;color:#3fb950;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600}
        .tag-s{background:#3a1a1a;color:#f85149;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600}
        .pp{color:#3fb950}.pn{color:#f85149}.pz{color:#8b949e}
        .nb{background:transparent;border:none;cursor:pointer;padding:8px 18px;font-family:inherit;font-size:14px;font-weight:500;border-radius:6px;transition:all .15s;color:#8b949e}
        .nb.act{background:#21262d;color:#e6edf3}.nb:hover:not(.act){color:#c9d1d9}
        .ov{position:fixed;inset:0;background:rgba(0,0,0,.75);display:flex;align-items:center;justify-content:center;z-index:100;backdrop-filter:blur(4px);padding:16px}
        .md{background:#161b22;border:1px solid #30363d;border-radius:12px;width:100%;max-width:500px;padding:26px;max-height:92vh;overflow-y:auto}
        .g2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
        .tbl{width:100%;border-collapse:collapse;font-size:13px}
        .tbl th{background:#21262d;color:#8b949e;font-weight:500;padding:10px 14px;text-align:left;white-space:nowrap}
        .tbl td{padding:10px 14px;border-bottom:1px solid #21262d}
        .tbl tr:hover td{background:#1c2128}
        .mono{font-family:'JetBrains Mono',monospace}
        .bdg{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600}
        .notif{position:fixed;top:20px;right:20px;padding:12px 20px;border-radius:8px;font-size:13px;font-weight:500;z-index:200;animation:sIn .2s ease;max-width:360px}
        .ns{background:#1a3a2a;color:#3fb950;border:1px solid #3fb95040}
        .ne{background:#3a1a1a;color:#f85149;border:1px solid #f8514940}
        .ni{background:#1a2a3a;color:#58a6ff;border:1px solid #58a6ff40}
        @keyframes sIn{from{transform:translateX(20px);opacity:0}to{transform:translateX(0);opacity:1}}
        .sc{background:#161b22;border:1px solid #21262d;border-radius:10px;padding:18px 22px}
        .emp{text-align:center;padding:48px;color:#484f58}
        .av{width:32px;height:32px;border-radius:50%;background:linear-gradient(135deg,#238636,#58a6ff);display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;cursor:pointer;user-select:none}
        .dd{position:absolute;top:48px;right:0;background:#161b22;border:1px solid #30363d;border-radius:10px;min-width:210px;overflow:hidden;box-shadow:0 8px 32px rgba(0,0,0,.5);z-index:50}
        .di{padding:10px 16px;font-size:13px;cursor:pointer;transition:background .1s;display:block;width:100%;text-align:left;border:none;font-family:inherit;color:#e6edf3;background:transparent}
        .di:hover{background:#21262d}
        .stb{background:transparent;border:none;cursor:pointer;padding:8px 16px;font-family:inherit;font-size:13px;font-weight:500;border-radius:6px;transition:all .15s;color:#8b949e}
        .stb.act{background:#21262d;color:#e6edf3}
        .fb{display:inline-block;background:#1a2a3a;color:#58a6ff;border:1px solid #58a6ff30;border-radius:20px;padding:2px 10px;font-size:11px;font-weight:600}
        .spin{display:inline-block;width:14px;height:14px;border:2px solid #30363d;border-top-color:#58a6ff;border-radius:50%;animation:spin .7s linear infinite;vertical-align:middle}
        @keyframes spin{to{transform:rotate(360deg)}}
        .yr-sel{background:#21262d;border:1px solid #30363d;color:#e6edf3;border-radius:6px;padding:4px 10px;font-size:12px;font-family:inherit;outline:none;cursor:pointer}
        .pulse{animation:pulse 2s ease-in-out infinite}
        @keyframes pulse{0%,100%{opacity:1}50%{opacity:.5}}
        @media(max-width:640px){.g2{grid-template-columns:1fr}.hm{display:none}}
      `}</style>

      {/* Header */}
      <div style={{ background: "#161b22", borderBottom: "1px solid #21262d", padding: "0 24px", paddingTop: "env(safe-area-inset-top, 0px)" }}>
        <div style={{ maxWidth: 1200, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", height: 56 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: 22 }}>📈</span>
            <span style={{ fontSize: 16, fontWeight: 700 }}>台股追蹤</span>
            <span className="hm" style={{ fontSize: 11, color: "#484f58", marginLeft: 4 }}>Portfolio Tracker</span>
          </div>
          <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
            {["portfolio", "history", "trades", "range"].map((v) => (
              <button key={v} className={`nb ${view === v ? "act" : ""}`} onClick={() => setView(v)}
                style={{ padding: "8px 10px", fontSize: 13 }}>
                {{ portfolio: "📊 持股", history: "🏆 歷史", trades: "📋 交易", range: "📅 區間" }[v]}
                <span className="hm" style={{ marginLeft: 2 }}>{{ portfolio: "明細", history: "損益", trades: "明細", range: "損益" }[v]}</span>
              </button>
            ))}
            <div style={{ position: "relative", marginLeft: 8 }} onClick={() => setUserMenu((x) => !x)}>
              <div className="av">{currentUser.displayName.slice(0, 1).toUpperCase()}</div>
              {showUserMenu && (
                <div className="dd" onClick={(e) => e.stopPropagation()}>
                  <div style={{ padding: "12px 16px", borderBottom: "1px solid #21262d" }}>
                    <div style={{ fontSize: 13, fontWeight: 600 }}>{currentUser.displayName}</div>
                    <div style={{ fontSize: 11, color: "#484f58" }}>@{currentUser.username}</div>
                    <div style={{ marginTop: 6, display: "flex", gap: 5, flexWrap: "wrap" }}>
                      <span className="fb">🏦 {currentUser.broker || "未設定"}</span>
                      <span className="fb">✂️ {Math.round((currentUser.feeDiscount ?? 0.85) * 10)}折</span>
                    </div>
                  </div>
                  <button className="di" onClick={openSettings}>⚙️ 帳戶設定</button>
                  <button className="di" style={{ color: "#f85149" }} onClick={handleLogout}>🚪 登出</button>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 1200, margin: "0 auto", padding: "24px 16px" }}>
        {/* Action Bar */}
        <div style={{ background: "#161b22", border: "1px solid #21262d", borderRadius: 8, padding: "10px 16px", marginBottom: 16, display: "flex", alignItems: "center", gap: 10, fontSize: 12, flexWrap: "wrap" }}>
          <span style={{ color: "#484f58" }}>手續費：</span>
          <span className="fb">🏦 {currentUser.broker || "未設定"}</span>
          <span className="fb">✂️ {Math.round(feeDiscount * 10)}折</span>
          {priceLoading
            ? <span className="fb pulse" style={{ background: "#1a2a3a", color: "#58a6ff", border: "1px solid #58a6ff40" }}><span className="spin" style={{ width: 10, height: 10, margin: "0 4px 0 0" }} />報價更新中…</span>
            : lastFetched
              ? <span className="fb" style={{ background: "#1a3a2a", color: "#3fb950", border: "1px solid #3fb95030" }}>🟢 報價 {lastFetched}</span>
              : <span className="fb" style={{ background: "#2a1a1a", color: "#f85149", border: "1px solid #f8514930" }}>⚠ 尚未取得報價</span>
          }
          {priceError && <span style={{ fontSize: 11, color: "#f85149" }}>⚠ {priceError}</span>}
          <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
            <button className="btn" style={{ background: "#388bfd22", color: "#58a6ff", border: "1px solid #58a6ff40", padding: "6px 12px", fontSize: 12, fontWeight: 600 }} onClick={() => refreshPrices(false)} disabled={priceLoading}>
              {priceLoading ? "⏳ 更新中" : "🔄 更新報價"}
            </button>
            <button className="btn btn-b" onClick={exportCsv}>⬇ 匯出紀錄</button>
            <button className="btn btn-b" onClick={() => fileInputRef.current?.click()}>⬆ 匯入交易</button>
            <input ref={fileInputRef} type="file" multiple accept=".csv,.xlsx,.xls,.xlsm,.ods,text/csv" style={{ display: "none" }} onChange={importFile} />
            <button className="btn btn-g" style={{ padding: "4px 11px", fontSize: 11 }} onClick={openSettings}>⚙ 設定</button>
          </div>
        </div>

        {importError && (
          <div style={{ background: "#3a1a1a", border: "1px solid #f8514940", borderRadius: 6, padding: "10px 14px", fontSize: 13, color: "#f85149", marginBottom: 12, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
            <span>⚠️ 匯入問題：{importError}</span>
            <button className="btn" style={{ color: "#8b949e", fontSize: 11, marginLeft: 12 }} onClick={() => setImportError("")}>✕</button>
          </div>
        )}

        {/* Summary Cards */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(190px,1fr))", gap: 12, marginBottom: 24 }}>
          {[
            { label: "市值合計", value: `$${fmtNum(totalMktVal)}`, sub: "新台幣" },
            { label: "持股成本", value: `$${fmtNum(totalCost)}`, sub: "含手續費" },
            { label: "未實現損益 (全部)", value: fmtMoney(Math.round(totalUnrealized)), sub: `${totalCost > 0 ? fmtMoney((totalUnrealized / totalCost * 100).toFixed(2)) : "0"}%`, pnl: totalUnrealized },
            { label: "已實現損益 (全部)", value: fmtMoney(Math.round(totalRealized)), sub: "歷史累計", pnl: totalRealized },
          ].map((c, i) => (
            <div key={i} className="sc">
              <div style={{ fontSize: 10, color: "#484f58", marginBottom: 6, fontWeight: 500, textTransform: "uppercase", letterSpacing: "1px" }}>{c.label}</div>
              <div style={{ fontSize: 20, fontWeight: 700, fontFamily: "'JetBrains Mono',monospace", color: c.pnl !== undefined ? (c.pnl >= 0 ? "#3fb950" : "#f85149") : "#e6edf3" }}>{c.value}</div>
              <div style={{ fontSize: 11, color: c.pnl !== undefined ? (c.pnl >= 0 ? "#3fb95099" : "#f8514999") : "#484f58", marginTop: 2 }}>{c.sub}</div>
            </div>
          ))}
          {[
            { label: "未實現損益", value: fmtMoney(Math.round(yearUnrealized)), pnl: yearUnrealized },
            { label: "已實現損益", value: fmtMoney(Math.round(yearRealized)), pnl: yearRealized },
          ].map((c, i) => (
            <div key={i} className="sc" style={{ border: "1px solid #2d4a6b" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
                <div style={{ fontSize: 10, color: "#58a6ff", fontWeight: 600, textTransform: "uppercase", letterSpacing: "1px" }}>{c.label}</div>
                {i === 0 && <select className="yr-sel" value={selectedYear} onChange={(e) => setYear(e.target.value)}>{years.map((y) => <option key={y} value={y}>{y === "all" ? "全部年度" : `${y} 年`}</option>)}</select>}
                {i === 1 && <span style={{ fontSize: 11, color: "#58a6ff" }}>{selectedYear === "all" ? "全部" : `${selectedYear} 年`}</span>}
              </div>
              <div style={{ fontSize: 20, fontWeight: 700, fontFamily: "'JetBrains Mono',monospace", color: c.pnl >= 0 ? "#3fb950" : "#f85149" }}>{c.value}</div>
              <div style={{ fontSize: 11, color: "#2d4a6b", marginTop: 2 }}>依年度篩選</div>
            </div>
          ))}
        </div>

        {tradesLoading && <div style={{ textAlign: "center", padding: 40, color: "#484f58" }}><span className="spin" style={{ width: 24, height: 24, borderWidth: 3 }} /></div>}

        {/* Portfolio View */}
        {!tradesLoading && view === "portfolio" && (
          <div className="card">
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, gap: 10 }}>
              <h2 style={{ fontSize: 15, fontWeight: 600 }}>持倉明細</h2>
              <button className="btn btn-p" onClick={() => { setEditTrade(null); setForm(emptyForm()); setStockLookup({ loading: false, error: "" }); setShowForm(true); }}>+ 新增交易</button>
            </div>
            {portfolio.filter((p) => p.qty > 0).length === 0
              ? <div className="emp">📭 尚無持股，點擊「新增交易」開始記錄</div>
              : (
                <div style={{ overflowX: "auto" }}>
                  <table className="tbl">
                    <thead><tr><th>股票</th><th>持股數</th><th>均成本</th><th>現價</th><th>今日漲跌</th><th>市值</th><th>未實現損益</th><th>損益%</th><th>已實現</th></tr></thead>
                    <tbody>
                      {portfolio.filter((p) => p.qty > 0).map((p) => {
                        const hasPx = p.hasLivePrice; const diff = hasPx ? p.curPrice - p.avgCost : 0; const isTW = !p.currency || p.currency === "TWD";
                        return (
                          <tr key={p.stock}>
                            <td><div style={{ fontWeight: 600 }}>{p.stock}</div><div style={{ fontSize: 11, color: "#8b949e" }}>{p.name}</div>{hasPx && <div style={{ fontSize: 10, color: "#484f58" }}>{p.marketState === "REGULAR" ? "🟢 交易中" : p.marketState === "PRE" ? "🌅 盤前" : p.marketState === "POST" ? "🌙 盤後" : "⚫ 收盤"}</div>}</td>
                            <td className="mono">{fmtNum(p.qty)}</td>
                            <td className="mono">{fmtNum(p.avgCost, 2)}</td>
                            <td>{hasPx ? <><div className="mono" style={{ fontWeight: 700, color: diff >= 0 ? "#3fb950" : "#f85149" }}>{isTW ? "" : "$"}{fmtNum(p.curPrice, 2)}{isTW ? "" : " USD"}</div><div style={{ fontSize: 10, color: diff >= 0 ? "#3fb95099" : "#f8514999", fontFamily: "monospace" }}>{diff >= 0 ? "+" : ""}{fmtNum(diff, 2)} vs 均成本</div></> : <span style={{ color: "#484f58", fontSize: 12 }}>{priceLoading ? <span className="spin" /> : "—"}</span>}</td>
                            <td>{hasPx && p.changePct != null ? <span className="bdg" style={{ background: p.changePct >= 0 ? "#1a3a2a" : "#3a1a1a", color: p.changePct >= 0 ? "#3fb950" : "#f85149" }}>{p.changePct >= 0 ? "▲" : "▼"} {Math.abs(p.changePct).toFixed(2)}%</span> : <span style={{ color: "#484f58" }}>—</span>}</td>
                            <td className="mono">{hasPx ? fmtNum(p.mktVal) : <span style={{ color: "#484f58" }}>—</span>}</td>
                            <td className={`mono ${p.unrealizedPnl >= 0 ? "pp" : "pn"}`}>{hasPx ? fmtMoney(Math.round(p.unrealizedPnl)) : <span style={{ color: "#484f58" }}>—</span>}</td>
                            <td>{hasPx ? <span className="bdg" style={{ background: p.unrealizedPct >= 0 ? "#1a3a2a" : "#3a1a1a", color: p.unrealizedPct >= 0 ? "#3fb950" : "#f85149" }}>{fmtMoney(p.unrealizedPct.toFixed(2))}%</span> : <span style={{ color: "#484f58" }}>—</span>}</td>
                            <td className={`mono ${p.realizedPnl > 0 ? "pp" : p.realizedPnl < 0 ? "pn" : "pz"}`}>{p.realizedPnl !== 0 ? fmtMoney(Math.round(p.realizedPnl)) : "—"}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
          </div>
        )}

        {/* History View */}
        {!tradesLoading && view === "history" && (() => {
          const allStock = [...new Set(trades.map((t) => t.stock))];
          const stockSummary = allStock.map((code) => {
            const related = trades.filter((t) => t.stock === code);
            const name = portfolio.find((p) => p.stock === code)?.name || lookupNameByCode(code) || code;
            let qty = 0, cost = 0, realized = 0;
            related.forEach((t) => {
              if (t.type === "buy") { cost += t.qty * t.price + (t.fee || 0); qty += t.qty; }
              else { const avg = qty > 0 ? cost / qty : 0; const pnl = t.qty * t.price - (t.fee || 0) - (t.tax || 0) - avg * t.qty; realized += pnl; cost -= avg * t.qty; qty -= t.qty; }
            });
            const totalBuy  = related.filter((t) => t.type === "buy").reduce((s, t) => s + t.qty * t.price + (t.fee || 0), 0);
            const totalSell = related.filter((t) => t.type === "sell").reduce((s, t) => s + t.qty * t.price - (t.fee || 0) - (t.tax || 0), 0);
            return { code, name, isHeld: qty > 0, qty, realized, totalBuy, totalSell };
          }).filter((s) => s.realized !== 0 || s.totalSell > 0);

          const yearMap = {};
          trades.forEach((t) => {
            if (t.type !== "sell") return;
            const yr = (t.date || "").slice(0, 4); if (!yr) return;
            if (!yearMap[yr]) yearMap[yr] = { year: yr, realized: 0, sells: 0 };
            const relBuys = trades.filter((x) => x.stock === t.stock && x.type === "buy" && x.date <= t.date);
            let q = 0, c = 0; relBuys.forEach((x) => { c += x.qty * x.price + (x.fee || 0); q += x.qty; });
            const avg = q > 0 ? c / q : 0;
            yearMap[yr].realized += t.qty * t.price - (t.fee || 0) - (t.tax || 0) - avg * t.qty;
            yearMap[yr].sells++;
          });
          const yearRows = Object.values(yearMap).sort((a, b) => b.year - a.year);

          return (
            <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
              <div className="card">
                <h2 style={{ fontSize: 15, fontWeight: 600, marginBottom: 14 }}>📅 歷年已實現損益</h2>
                {yearRows.length === 0 ? <div className="emp">尚無已實現損益資料</div> : (
                  <div style={{ overflowX: "auto" }}>
                    <table className="tbl">
                      <thead><tr><th>年度</th><th>賣出筆數</th><th>已實現損益</th><th>獲利狀態</th></tr></thead>
                      <tbody>
                        {yearRows.map((r) => (
                          <tr key={r.year}>
                            <td style={{ fontWeight: 600 }}>{r.year} 年</td>
                            <td className="mono pz">{r.sells} 筆</td>
                            <td className={`mono ${r.realized >= 0 ? "pp" : "pn"}`} style={{ fontSize: 15, fontWeight: 700 }}>{fmtMoney(Math.round(r.realized))}</td>
                            <td><span className="bdg" style={{ background: r.realized >= 0 ? "#1a3a2a" : "#3a1a1a", color: r.realized >= 0 ? "#3fb950" : "#f85149" }}>{r.realized >= 0 ? "獲利" : "虧損"}</span></td>
                          </tr>
                        ))}
                        <tr style={{ borderTop: "2px solid #30363d" }}>
                          <td style={{ fontWeight: 700 }}>合計</td>
                          <td className="mono pz">{yearRows.reduce((s, r) => s + r.sells, 0)} 筆</td>
                          <td className={`mono ${yearRows.reduce((s, r) => s + r.realized, 0) >= 0 ? "pp" : "pn"}`} style={{ fontSize: 16, fontWeight: 700 }}>{fmtMoney(Math.round(yearRows.reduce((s, r) => s + r.realized, 0)))}</td>
                          <td></td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
              <div className="card">
                <h2 style={{ fontSize: 15, fontWeight: 600, marginBottom: 14 }}>📊 各股歷史績效</h2>
                {stockSummary.length === 0 ? <div className="emp">尚無交易歷史資料</div> : (
                  <div style={{ overflowX: "auto" }}>
                    <table className="tbl">
                      <thead><tr><th>股票</th><th>狀態</th><th>買入總額</th><th>賣出收回</th><th>已實現損益</th><th>報酬率</th></tr></thead>
                      <tbody>
                        {stockSummary.sort((a, b) => Math.abs(b.realized) - Math.abs(a.realized)).map((s) => (
                          <tr key={s.code}>
                            <td><div style={{ fontWeight: 600 }}>{s.code}</div><div style={{ fontSize: 11, color: "#8b949e" }}>{s.name}</div></td>
                            <td><span className="bdg" style={{ background: s.isHeld ? "#1a2a3a" : "#21262d", color: s.isHeld ? "#58a6ff" : "#8b949e" }}>{s.isHeld ? "持有中" : "已出清"}</span></td>
                            <td className="mono pz">{fmtNum(Math.round(s.totalBuy))}</td>
                            <td className="mono pz">{s.totalSell > 0 ? fmtNum(Math.round(s.totalSell)) : "—"}</td>
                            <td className={`mono ${s.realized >= 0 ? "pp" : "pn"}`} style={{ fontWeight: 700 }}>{fmtMoney(Math.round(s.realized))}</td>
                            <td>{s.totalBuy > 0 ? <span className="bdg" style={{ background: s.realized >= 0 ? "#1a3a2a" : "#3a1a1a", color: s.realized >= 0 ? "#3fb950" : "#f85149" }}>{fmtMoney((s.realized / s.totalBuy * 100).toFixed(2))}%</span> : "—"}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </div>
          );
        })()}

        {/* Trades View */}
        {!tradesLoading && view === "trades" && (() => {
          const allIds = filteredTrades.map((t) => t.id);
          const allSelected = allIds.length > 0 && allIds.every((id) => selectedIds.has(id));
          const someSelected = selectedIds.size > 0;
          const toggleAll = () => { if (allSelected) setSelectedIds(new Set()); else setSelectedIds(new Set(allIds)); };
          const toggleOne = (id) => setSelectedIds((prev) => { const s = new Set(prev); s.has(id) ? s.delete(id) : s.add(id); return s; });
          const deleteSelected = async () => {
            if (!window.confirm(`確定要刪除選取的 ${selectedIds.size} 筆交易？`)) return;
            try {
              for (const id of selectedIds) { await supa(`trades?id=eq.${id}`, { method: "DELETE" }); }
              setTrades((p) => p.filter((t) => !selectedIds.has(t.id)));
              setSelectedIds(new Set());
              notify(`已刪除 ${selectedIds.size} 筆`);
            } catch (e) { notify("刪除失敗：" + e.message, "error"); }
          };
          return (
            <div className="card">
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
                <h2 style={{ fontSize: 15, fontWeight: 600, flex: 1 }}>交易明細</h2>
                <input placeholder="搜尋代號/名稱" value={filterStock} onChange={(e) => setFilter(e.target.value)} style={{ width: 140 }} />
                <select value={sortBy} onChange={(e) => setSortBy(e.target.value)} style={{ width: 110 }}>
                  <option value="date">依日期</option><option value="stock">依代號</option>
                </select>
                {someSelected && (
                  <button className="btn btn-d" onClick={deleteSelected}>🗑 刪除選取（{selectedIds.size}）</button>
                )}
                <button className="btn btn-p" onClick={() => { setEditTrade(null); setForm(emptyForm()); setStockLookup({ loading: false, error: "" }); setShowForm(true); }}>+ 新增</button>
              </div>
              {filteredTrades.length === 0 ? <div className="emp">📭 無交易記錄</div> : (
                <div style={{ overflowX: "auto" }}>
                  <table className="tbl">
                    <thead>
                      <tr>
                        <th style={{ width: 36 }}>
                          <input type="checkbox" checked={allSelected} onChange={toggleAll} style={{ width: 15, height: 15, cursor: "pointer", accentColor: "#58a6ff" }} />
                        </th>
                        <th>日期</th><th>股票</th><th>類型</th><th>數量</th><th>價格</th><th>手續費＋稅</th><th>金額</th><th>損益（賣出）</th><th>操作</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredTrades.map((t) => (
                        <tr key={t.id} style={{ background: selectedIds.has(t.id) ? "#1a2a3a" : undefined }}>
                          <td><input type="checkbox" checked={selectedIds.has(t.id)} onChange={() => toggleOne(t.id)} style={{ width: 15, height: 15, cursor: "pointer", accentColor: "#58a6ff" }} /></td>
                          <td style={{ color: "#8b949e" }}>{t.date}</td>
                          <td><div style={{ fontWeight: 600 }}>{t.stock}</div><div style={{ fontSize: 11, color: "#8b949e" }}>{t.name}</div></td>
                          <td><span className={t.type === "buy" ? "tag-b" : "tag-s"}>{t.type === "buy" ? "買進" : "賣出"}</span></td>
                          <td className="mono">{fmtNum(t.qty)}</td>
                          <td className="mono">{fmtNum(t.price, 2)}</td>
                          <td className="mono pz">{fmtNum((t.fee || 0) + (t.tax || 0))}</td>
                          <td className="mono">{fmtNum(t.qty * t.price)}</td>
                          <td>{t.type === "sell" && t.sellPnl !== null ? <div><div className={`mono ${t.sellPnl >= 0 ? "pp" : "pn"}`} style={{ fontWeight: 700 }}>{fmtMoney(Math.round(t.sellPnl))}</div><span className="bdg" style={{ background: t.sellPct >= 0 ? "#1a3a2a" : "#3a1a1a", color: t.sellPct >= 0 ? "#3fb950" : "#f85149" }}>{t.sellPct >= 0 ? "+" : ""}{t.sellPct.toFixed(2)}%</span></div> : <span style={{ color: "#484f58" }}>—</span>}</td>
                          <td><div style={{ display: "flex", gap: 6 }}><button className="btn btn-g" onClick={() => openEdit(t)}>✏️</button><button className="btn btn-d" onClick={() => deleteTrade(t.id)}>✕</button></div></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          );
        })()}

        {/* Range View */}
        {!tradesLoading && view === "range" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            <div className="card">
              <h2 style={{ fontSize: 15, fontWeight: 600, marginBottom: 16 }}>📅 自訂區間損益</h2>
              <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ fontSize: 13, color: "#8b949e", whiteSpace: "nowrap" }}>開始日期</span>
                  <input type="date" value={rangeStart} onChange={(e) => setRangeStart(e.target.value)} style={{ width: 160 }} />
                </div>
                <span style={{ color: "#484f58" }}>～</span>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ fontSize: 13, color: "#8b949e", whiteSpace: "nowrap" }}>結束日期</span>
                  <input type="date" value={rangeEnd} onChange={(e) => setRangeEnd(e.target.value)} style={{ width: 160 }} />
                </div>
                <button className="btn btn-g" style={{ fontSize: 12 }} onClick={() => { setRangeStart(today.slice(0, 4) + "-01-01"); setRangeEnd(today); }}>今年</button>
                <button className="btn btn-g" style={{ fontSize: 12 }} onClick={() => { const y = String(Number(today.slice(0, 4)) - 1); setRangeStart(y + "-01-01"); setRangeEnd(y + "-12-31"); }}>去年</button>
                <button className="btn btn-g" style={{ fontSize: 12 }} onClick={() => { const d = trades.length > 0 ? [...trades].sort((a,b)=>a.date.localeCompare(b.date))[0].date : today; setRangeStart(d); setRangeEnd(today); }}>全部</button>
              </div>
            </div>
            {rangeStats && (
              <>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 12 }}>
                  {[
                    { label: "區間已實現損益", value: fmtMoney(Math.round(rangeStats.realized)), sub: "區間內賣出已結算", pnl: rangeStats.realized },
                    { label: "區間未實現損益", value: fmtMoney(Math.round(rangeStats.unrealized)), sub: `${fmtMoney(rangeStats.unrealPct.toFixed(2))}%　以現價計算`, pnl: rangeStats.unrealized },
                    { label: "區間總損益", value: fmtMoney(Math.round(rangeStats.realized + rangeStats.unrealized)), sub: "已實現＋未實現合計", pnl: rangeStats.realized + rangeStats.unrealized },
                  ].map((c, i) => (
                    <div key={i} className="sc" style={{ border: "1px solid #2d4a6b" }}>
                      <div style={{ fontSize: 10, color: "#58a6ff", fontWeight: 600, marginBottom: 8, textTransform: "uppercase", letterSpacing: "1px" }}>{c.label}</div>
                      <div style={{ fontSize: 24, fontWeight: 700, fontFamily: "monospace", color: c.pnl >= 0 ? "#3fb950" : "#f85149" }}>{c.value}</div>
                      <div style={{ fontSize: 11, color: "#484f58", marginTop: 4 }}>{c.sub}</div>
                    </div>
                  ))}
                </div>
                {rangeStats.unrealRows.length > 0 && (
                  <div className="card">
                    <h3 style={{ fontSize: 14, fontWeight: 600, marginBottom: 14, color: "#8b949e" }}>區間買入、尚未賣出的持股（未實現損益明細）</h3>
                    <div style={{ overflowX: "auto" }}>
                      <table className="tbl">
                        <thead><tr><th>股票</th><th>持股數</th><th>買入成本</th><th>現值</th><th>未實現損益</th><th>損益%</th></tr></thead>
                        <tbody>
                          {rangeStats.unrealRows.map((r) => (
                            <tr key={r.code}>
                              <td><div style={{ fontWeight: 600 }}>{r.code}</div><div style={{ fontSize: 11, color: "#8b949e" }}>{r.name}</div></td>
                              <td className="mono">{fmtNum(r.qty)}</td>
                              <td className="mono pz">{fmtNum(Math.round(r.cost))}</td>
                              <td className="mono">{r.hasLive ? fmtNum(Math.round(r.mkt)) : <span style={{ color: "#484f58" }}>—（以成本估）</span>}</td>
                              <td className={`mono ${r.pnl >= 0 ? "pp" : "pn"}`} style={{ fontWeight: 700 }}>{fmtMoney(Math.round(r.pnl))}</td>
                              <td><span className="bdg" style={{ background: r.pct >= 0 ? "#1a3a2a" : "#3a1a1a", color: r.pct >= 0 ? "#3fb950" : "#f85149" }}>{fmtMoney(r.pct.toFixed(2))}%</span></td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
                {rangeStats.unrealRows.length === 0 && rangeStats.realized === 0 && <div className="card"><div className="emp">此區間內無交易記錄</div></div>}
              </>
            )}
          </div>
        )}
      </div>

      {/* Add/Edit Modal */}
      {showForm && (
        <div className="ov" onClick={(e) => e.target === e.currentTarget && setShowForm(false)}>
          <div className="md">
            <h3 style={{ marginBottom: 20, fontSize: 16 }}>{editTrade ? "✏️ 編輯交易" : "➕ 新增交易"}</h3>
            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              <div className="g2">
                <div><label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 5 }}>日期 *</label><input type="date" value={form.date} onChange={(e) => setForm((f) => ({ ...f, date: e.target.value }))} /></div>
                <div><label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 5 }}>類型 *</label><select value={form.type} onChange={(e) => setForm((f) => ({ ...f, type: e.target.value }))}><option value="buy">買進</option><option value="sell">賣出</option></select></div>
              </div>
              <div className="g2">
                <div>
                  <label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 5 }}>股票代號 *{stockLookup.error && <span style={{ color: "#f85149", fontSize: 11, marginLeft: 6 }}>⚠ {stockLookup.error}</span>}</label>
                  <input placeholder="例：2330 或 AAPL" value={form.stock} onChange={(e) => { const v = e.target.value; setForm((f) => ({ ...f, stock: v, name: "" })); setStockLookup({ loading: false, error: "" }); if (v.length >= 2) lookupStock(v); }} onBlur={(e) => lookupStock(e.target.value)} />
                </div>
                <div><label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 5 }}>股票名稱</label><input placeholder="代號輸入後自動帶入" value={form.name} onChange={(e) => setForm((f) => ({ ...f, name: e.target.value }))} /></div>
              </div>
              <div className="g2">
                <div><label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 5 }}>數量（股）*</label><input type="number" placeholder="1000" value={form.qty} onChange={(e) => setForm((f) => ({ ...f, qty: e.target.value }))} /></div>
                <div><label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 5 }}>成交價 *</label><input type="number" placeholder="0.00" step="0.01" value={form.price} onChange={(e) => setForm((f) => ({ ...f, price: e.target.value }))} /></div>
              </div>
              <div className="g2">
                <div><label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 5 }}>手續費 <span style={{ color: "#484f58" }}>（留空＝自動）</span></label><input type="number" placeholder={form.qty && form.price ? `自動：${calcFee(Number(form.qty), Number(form.price), feeDiscount)}` : "自動計算"} value={form.fee} onChange={(e) => setForm((f) => ({ ...f, fee: e.target.value }))} /></div>
                {form.type === "sell" && <div><label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 5 }}>交易稅</label><input type="number" placeholder="賣出才需填" value={form.tax} onChange={(e) => setForm((f) => ({ ...f, tax: e.target.value }))} /></div>}
              </div>
              {form.qty && form.price && (
                <div style={{ background: "#0d1117", borderRadius: 8, padding: "12px 14px", fontSize: 12, display: "flex", gap: 24, flexWrap: "wrap" }}>
                  <div><span style={{ color: "#484f58" }}>交易金額</span><br /><span style={{ color: "#e6edf3", fontWeight: 700, fontFamily: "monospace", fontSize: 15 }}>${fmtNum(Number(form.qty) * Number(form.price))}</span></div>
                  <div><span style={{ color: "#484f58" }}>手續費</span><br /><span style={{ color: "#58a6ff", fontWeight: 700, fontFamily: "monospace", fontSize: 15 }}>${fmtNum(prevFeeAmt)}</span></div>
                </div>
              )}
              <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 4 }}>
                <button className="btn btn-g" onClick={() => setShowForm(false)}>取消</button>
                <button className={`btn ${form.type === "sell" ? "btn-s" : "btn-p"}`} onClick={submitTrade}>{editTrade ? "儲存" : form.type === "buy" ? "確認買進" : "確認賣出"}</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Settings Modal */}
      {showSettings && (
        <div className="ov" onClick={(e) => e.target === e.currentTarget && setSettings(false)}>
          <div className="md" style={{ maxWidth: 420 }}>
            <h3 style={{ marginBottom: 16, fontSize: 16 }}>⚙️ 帳戶設定</h3>
            <div style={{ display: "flex", background: "#0d1117", borderRadius: 8, padding: 4, marginBottom: 20 }}>
              {[["broker", "🏦 券商設定"], ["password", "🔑 修改密碼"]].map(([t, lbl]) => (
                <button key={t} className={`stb ${settingsTab === t ? "act" : ""}`} style={{ flex: 1 }} onClick={() => { setStab(t); setSfErr(""); }}>{lbl}</button>
              ))}
            </div>
            {settingsTab === "broker" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                <div><label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 6 }}>使用券商</label><select value={sf.broker} onChange={(e) => { const b = BROKERS.find((x) => x.name === e.target.value); setSf((f) => ({ ...f, broker: e.target.value, feeDiscount: b ? b.defaultDiscount : f.feeDiscount })); }}>{BROKERS.map((b) => <option key={b.name} value={b.name}>{b.name}</option>)}</select></div>
                <div>
                  <label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 6 }}>手續費折數</label>
                  <div style={{ display: "flex", gap: 8, alignItems: "center" }}><input type="number" step="0.01" min="0.01" max="1" value={sf.feeDiscount} onChange={(e) => setSf((f) => ({ ...f, feeDiscount: e.target.value }))} style={{ flex: 1 }} /><span style={{ fontSize: 12, color: "#8b949e", whiteSpace: "nowrap" }}>{sf.feeDiscount ? `${Math.round(Number(sf.feeDiscount) * 10)}折` : "—"}</span></div>
                  <div style={{ fontSize: 11, color: "#484f58", marginTop: 5 }}>10 萬元手續費預估：<span style={{ color: "#e6edf3", fontFamily: "monospace" }}>NT$ {stPreviewFee}</span></div>
                </div>
                {sfErr && <div style={{ background: "#3a1a1a", border: "1px solid #f8514940", borderRadius: 6, padding: "9px 13px", fontSize: 13, color: "#f85149" }}>⚠️ {sfErr}</div>}
                <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}><button className="btn btn-g" onClick={() => setSettings(false)}>取消</button><button className="btn btn-p" onClick={saveBroker}>儲存設定</button></div>
              </div>
            )}
            {settingsTab === "password" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                {[["oldPw", "舊密碼", "輸入目前密碼"], ["newPw", "新密碼", "至少 6 個字元"], ["confirmPw", "確認新密碼", "再輸入一次"]].map(([k, lbl, ph]) => (
                  <div key={k}><label style={{ fontSize: 11, color: "#8b949e", display: "block", marginBottom: 6 }}>{lbl}</label><input type="password" placeholder={ph} value={sf[k]} onChange={(e) => setSf((f) => ({ ...f, [k]: e.target.value }))} /></div>
                ))}
                {sfErr && <div style={{ background: "#3a1a1a", border: "1px solid #f8514940", borderRadius: 6, padding: "9px 13px", fontSize: 13, color: "#f85149" }}>⚠️ {sfErr}</div>}
                <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}><button className="btn btn-g" onClick={() => setSettings(false)}>取消</button><button className="btn btn-p" onClick={changePw}>確認修改</button></div>
              </div>
            )}
          </div>
        </div>
      )}

      {notification && <div className={`notif ${notification.type === "error" ? "ne" : notification.type === "info" ? "ni" : "ns"}`}>{notification.msg}</div>}
    </div>
  );
}

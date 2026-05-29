import { useState, useEffect, useRef, useCallback } from "react";
import { api } from "./api.js";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, CartesianGrid, LineChart, Line, Legend,
} from "recharts";

const FONTS = `@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@300;400;500&display=swap');`;

const CAT_RULES = [
  // Groceries — Swiss & Israeli supermarkets
  { keywords: ["migros","coop","lidl","aldi","denner","spar","volg","leshop","pick pay","globus food",
    "supermarket","minimarkt","lebensmittel","superpharma",
    "rami levi","shufersal","victory","osher ad","yochananof","mega","tiv taam","makolet"], cat: "Groceries" },

  // Lunch / Dining — restaurants, takeaway, cafes
  { keywords: ["restaurant","cafe","café","tibits","mcdonalds","mcdonald","burger king","subway","starbucks",
    "pizza","sushi","kebab","falafel","shawarma","hummus","mittag","lunch","essen","takeaway","takeout",
    "lieferando","uber eats","wolt","just eat","foodora","ten bis","10bis",
    "bistro","brasserie","trattoria","osteria","canteen","kantine","mensa"], cat: "Lunch/Dining" },

  // Transport — Swiss & Israeli public transport + ride share
  { keywords: ["sbb","zvv","mobility","postbus","bls","postauto","tpg","bernmobil","rbs","sob","thurbo",
    "parking","parkhaus","parkplatz","vignette","autobahn",
    "uber","lyft","taxi","bolt","gett",
    "egged","dan bus","metropoline","kavim","rav kav"], cat: "Transport" },

  // Entertainment
  { keywords: ["spotify","netflix","disney","amazon prime","youtube premium","steam","playstation","xbox",
    "kino","cinema","theater","theatre","concert","museum","zoo","ticket","eventim","ticketmaster",
    "apple tv","hbo","paramount","peacock"], cat: "Entertainment" },

  // Health — Swiss & Israeli
  { keywords: ["apotheke","pharmacy","pharmacia","arzt","zahnarzt","physio","physiotherpeut",
    "hospital","spital","klinik","clinic","doktor","doctor","optiker","optician",
    "helsana","css","swica","sanitas","visana","concordia","atupri","groupe mutuel",
    "maccabi","clalit","leumit","meuhedet","kupat holim"], cat: "Health" },

  // Phone / Internet
  { keywords: ["swisscom","sunrise","salt mobile","salt.ch","upc","quickline","wingo","yallo",
    "partner communications","hot mobile","cellcom","bezeq","012","019"], cat: "Phone/Internet" },

  // Utilities
  { keywords: ["ewz","ewb","stadtwerke","strom","gas","elektrizitat","electricity","water","wasser",
    "heizung","heating","nebenkosten","utility","entsorgung"], cat: "Utilities" },

  // Shopping — clothes, electronics, home
  { keywords: ["zara","h&m","uniqlo","zalando","about you","asos","mango","pull&bear",
    "ikea","galaxus","digitec","manor","jelmoli","globus","fnac","mediamarkt","interdiscount",
    "amazon","aliexpress","ebay","wish",
    "castro","fox","renuar","golf","honigman"], cat: "Shopping" },

  // Finance — transfers, fees, ATM
  { keywords: ["twint","paypal","wise","transferwise","revolut","western union",
    "atm","bancomat","cash withdrawal","gebühr","fee","zinsen","interest",
    "kontoführung","annual fee","jahresgebühr"], cat: "Transfers" },

  // Israeli-specific
  { keywords: ["bit payment","pepper","max mastercard","leumi card","cal card","isracard","visa cal",
    "arnona","property tax","mas hakhnasa","bituah leumi"], cat: "Other" },

  // Child
  { keywords: ["schulgebühren","schulgeld","schulbedarf","kita","krippe","hort","nachhilfe","schulreise","klassenlager"], cat: "Child/School" },
  { keywords: ["kinderarzt","kinderkleidung","kinderschuhe","baby","spielzeug"], cat: "Child/Other" },
];

const CAT_COLORS = {
  "Groceries":"#F59E0B","Lunch/Dining":"#E4003A","Transport":"#8B5CF6",
  "Health":"#10B981","Phone/Internet":"#F97316","Utilities":"#06B6D4",
  "Entertainment":"#EC4899","Shopping":"#3B82F6","Other":"#4B5563",
};

const CAT_ICONS = {
  "Groceries":"🛒","Lunch/Dining":"🍽️","Transport":"🚆","Health":"🏥",
  "Phone/Internet":"📱","Utilities":"⚡","Entertainment":"🎭","Shopping":"🛍️",
  "Other":"💳","Insurance":"🛡️","Child/School":"📚","Child/Other":"🎒",
};

const DEFAULT_BUDGETS = {
  "Groceries":450,"Lunch/Dining":280,"Transport":200,"Health":100,
  "Phone/Internet":65,"Utilities":80,"Entertainment":100,"Shopping":150,"Other":80,
};


const FREQ_MULT = { weekly:52, monthly:12, quarterly:4, yearly:1 };

const DEFAULT_SUBS = [
  { id:"s1", name:"Helsana Health Insurance", amount:380,   freq:"monthly",   cat:"Health",        icon:"🏥", taxDeductible:true  },
  { id:"s2", name:"SBB GA Travelcard",         amount:3860,  freq:"yearly",    cat:"Transport",     icon:"🚆", taxDeductible:true  },
  { id:"s3", name:"Sunrise Mobile + Internet", amount:79,    freq:"monthly",   cat:"Phone/Internet",icon:"📱", taxDeductible:false },
  { id:"s4", name:"EWZ Electricity",           amount:62,    freq:"monthly",   cat:"Utilities",     icon:"⚡", taxDeductible:false },
  { id:"s5", name:"Spotify",                   amount:12.95, freq:"monthly",   cat:"Entertainment", icon:"🎵", taxDeductible:false },
  { id:"s6", name:"Netflix",                   amount:17.90, freq:"monthly",   cat:"Entertainment", icon:"🎬", taxDeductible:false },
];

const DEFAULT_INVESTMENTS = [
  { id:"i1", name:"Neon Invest — ETF (VWRL)", value:8420,  gain:+340,  pct:+4.2, currency:"CHF", icon:"📈" },
  { id:"i2", name:"Pillar 3a (Frankly)",       value:14800, gain:+680,  pct:+4.8, currency:"CHF", icon:"🏦" },
  { id:"i3", name:"Israeli Stocks (Max)",      value:9200,  gain:-180,  pct:-1.9, currency:"₪",  icon:"📉" },
];

const DEFAULT_ACCOUNTS = [
  { id:"max",   name:"Max Mastercard", bank:"Israeli Account",    balance:0, currency:"₪",  color:"#E4003A", icon:"🇮🇱" },
  { id:"neon",  name:"Neon",           bank:"Neon + Investments", balance:0, currency:"CHF", color:"#00C896", icon:"🟢" },
  { id:"twint", name:"Twint",          bank:"Mobile Payments",    balance:0, currency:"CHF", color:"#0066FF", icon:"💙" },
  { id:"sbb",   name:"SBB",            bank:"Travel Account",     balance:0, currency:"CHF", color:"#E84040", icon:"🚆" },
];

const TABS = [
  { id:"overview",     label:"Overview",     icon:"◎" },
  { id:"transactions", label:"Transactions", icon:"≡" },
  { id:"spending",     label:"Spending",     icon:"◈" },
  { id:"food",         label:"Food",         icon:"◆" },
  { id:"documents",    label:"Documents",    icon:"◉" },
  { id:"taxes",        label:"Tax Summary",  icon:"◐" },
  { id:"annual",       label:"Annual View",  icon:"◑" },
  { id:"family",       label:"Family",       icon:"◈" },
];

// Data persistence handled via api.js → backend → PostgreSQL

function categorize(desc = "") {
  const d = desc.toLowerCase();
  for (const rule of CAT_RULES) {
    if (rule.keywords.some(k => d.includes(k))) return rule.cat;
  }
  return "Other";
}

function parseCSV(text, accountId) {
  const result = Papa.parse(text.trim(), {
    header: true, skipEmptyLines: true,
    delimitersToGuess: [",", ";", "\t", "|"],
  });
  const rows = result.data;
  if (!rows.length) return [];
  const keys = Object.keys(rows[0]);
  const dateCol = keys.find(k => /date|datum|buchungsdatum/i.test(k)) || keys[0];
  const descCol = keys.find(k => /desc|buchung|besch|text|betreff|memo|verwendung/i.test(k)) || keys[1];
  const amtCol  = keys.find(k => /amount|betrag|credit|debit|summe|total/i.test(k)) || keys[2];
  return rows.map((row, i) => {
    const raw = String(row[amtCol] || "0").replace(/[^\d.\-,]/g, "").replace(",", ".");
    const amount = parseFloat(raw) || 0;
    if (amount === 0) return null;
    return {
      id: `${accountId}-${Date.now()}-${i}`,
      date: row[dateCol] || "", desc: row[descCol] || "",
      amount, account: accountId, category: categorize(row[descCol] || ""),
    };
  }).filter(Boolean);
}

function parseExcel(buffer, accountId) {
  const wb = XLSX.read(buffer, { type: "array", cellDates: true });
  const allRows = [];

  wb.SheetNames.forEach(sheetName => {
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
    if (!rows.length) return;

    const keys = Object.keys(rows[0]);
    const dateCol = keys.find(k => /date|datum|buchungsdatum|תאריך/i.test(k)) || keys[0];
    const descCol = keys.find(k => /desc|buchung|besch|text|betreff|memo|verwendung|פירוט|תיאור|סוג/i.test(k)) || keys[1];
    const amtCol  = keys.find(k => /amount|betrag|credit|debit|summe|total|סכום|זכות|חובה|קרדיט/i.test(k)) || keys[2];

    rows.forEach((row, i) => {
      let rawAmt = String(row[amtCol] || "0");
      // Handle negative amounts written as (1,234.56)
      const neg = rawAmt.startsWith("(") && rawAmt.endsWith(")");
      rawAmt = rawAmt.replace(/[^\d.\-,]/g, "").replace(",", ".");
      let amount = parseFloat(rawAmt) || 0;
      if (neg && amount > 0) amount = -amount;
      if (amount === 0) return;

      // Format date
      let date = "";
      const rawDate = row[dateCol];
      if (rawDate instanceof Date) {
        date = rawDate.toISOString().slice(0, 10);
      } else if (rawDate) {
        // Try to parse common date formats
        const d = new Date(rawDate);
        date = isNaN(d) ? String(rawDate) : d.toISOString().slice(0, 10);
      }

      const desc = String(row[descCol] || "").trim();
      allRows.push({
        id: `${accountId}-xl-${Date.now()}-${i}`,
        date, desc, amount,
        account: accountId,
        category: categorize(desc),
      });
    });
  });

  return allRows.filter(Boolean);
}

// OCR handled server-side via api.ocrBill / api.ocrGrocery

/* ── Reusable UI ── */
function SL({ children }) {
  return <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:10,color:"#A8A29E",letterSpacing:2,textTransform:"uppercase",marginBottom:12 }}>{children}</div>;
}
function StatCard({ label, value, sub, good }) {
  return (
    <div style={{ background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:12,padding:"16px 18px",boxShadow:"0 1px 3px rgba(28,25,23,0.06)" }}>
      <div style={{ fontFamily:"DM Mono",fontSize:9,color:"#A8A29E",textTransform:"uppercase",letterSpacing:1,marginBottom:8 }}>{label}</div>
      <div style={{ fontFamily:"Syne",fontWeight:800,fontSize:18,color:good?"#166534":"#1C1917",marginBottom:4 }}>{value}</div>
      <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#A8A29E" }}>{sub}</div>
    </div>
  );
}
function TipCard({ tip }) {
  const col = { warning:"#92400E",alert:"#991B1B",ok:"#166534",info:"#1E40AF" }[tip.level]||"#57534E";
  return (
    <div style={{ background:`${col}10`,border:`1px solid ${col}25`,borderLeft:`3px solid ${col}`,borderRadius:10,padding:"12px 16px",marginBottom:8 }}>
      <div style={{ display:"flex",alignItems:"center",gap:8,marginBottom:3 }}>
        <span>{tip.icon}</span>
        <span style={{ fontFamily:"Syne",fontWeight:600,fontSize:12,color:col }}>{tip.title}</span>
      </div>
      <div style={{ fontFamily:"DM Mono",fontSize:11,color:"#78716C",lineHeight:1.6 }}>{tip.desc}</div>
    </div>
  );
}
function AccountCard({ acc, active, onClick, onEditBalance }) {
  const [editing, setEditing] = useState(false);
  const [val, setVal] = useState(String(acc.balance));
  const commit = () => {
    const n = parseFloat(val.replace(/[^0-9.]/g,""));
    if (!isNaN(n)) onEditBalance(acc.id, n);
    setEditing(false);
  };
  return (
    <div onClick={()=>{ if(!editing) onClick(); }} style={{
      background:active?acc.color:"#FFFFFF",
      border:`1px solid ${active?acc.color:"#EAE7E1"}`,
      borderRadius:14,padding:"17px 19px",cursor:"pointer",
      transition:"all 0.22s",flex:"1 1 150px",minWidth:140,
      boxShadow:active?`0 8px 24px ${acc.color}35`:"0 1px 3px rgba(28,25,23,0.06)",
    }}>
      <div style={{ fontSize:18,marginBottom:8 }}>{acc.icon}</div>
      <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:12,color:active?"#fff":"#1C1917" }}>{acc.name}</div>
      <div style={{ fontFamily:"DM Mono",fontSize:10,color:active?"rgba(255,255,255,0.65)":"#A8A29E",marginBottom:10 }}>{acc.bank}</div>
      {editing ? (
        <div onClick={e=>e.stopPropagation()} style={{ display:"flex",alignItems:"center",gap:4 }}>
          <input autoFocus value={val} onChange={e=>setVal(e.target.value)}
            onBlur={commit} onKeyDown={e=>e.key==="Enter"&&commit()}
            style={{ width:"100%",fontFamily:"DM Mono",fontWeight:500,fontSize:15,
              background:"rgba(255,255,255,0.25)",border:"none",
              borderBottom:`2px solid ${active?"#fff":"#C8102E"}`,
              color:active?"#fff":"#1C1917",padding:"2px 0",outline:"none" }} />
        </div>
      ) : (
        <div style={{ display:"flex",alignItems:"baseline",gap:6 }}>
          <div style={{ fontFamily:"DM Mono",fontWeight:500,fontSize:17,color:active?"#fff":"#1C1917" }}>
            {acc.currency} {acc.balance>0 ? acc.balance.toLocaleString("de-CH",{minimumFractionDigits:2}) : "—"}
          </div>
          <div onClick={e=>{ e.stopPropagation(); setVal(String(acc.balance)); setEditing(true); }}
            style={{ fontSize:10,color:active?"rgba(255,255,255,0.6)":"#C4BDB7",cursor:"pointer" }}
            title="Click to update balance">✏️</div>
        </div>
      )}
    </div>
  );
}
function UploadZone({ onFile, label, accept=".pdf,.jpg,.jpeg,.png" }) {
  const ref = useRef();
  const [drag, setDrag] = useState(false);
  return (
    <div onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
      onDrop={e=>{e.preventDefault();setDrag(false);onFile(e.dataTransfer.files[0]);}}
      onClick={()=>ref.current.click()}
      style={{ border:`2px dashed ${drag?"#C8102E":"#D9D5CE"}`,borderRadius:12,
        padding:"22px 18px",textAlign:"center",cursor:"pointer",
        background:drag?"#FEF2F2":"#FAFAF8",transition:"all 0.18s" }}>
      <input ref={ref} type="file" accept={accept} style={{ display:"none" }} onChange={e=>onFile(e.target.files[0])} />
      <div style={{ fontSize:22,marginBottom:6 }}>📎</div>
      <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:12,color:"#44403C" }}>{label}</div>
      <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#A8A29E",marginTop:4 }}>Drag & drop or click to browse</div>
    </div>
  );
}
function Loader({ text }) {
  return <div style={{ textAlign:"center",padding:"22px",fontFamily:"DM Mono",fontSize:11,color:"#78716C" }}><div style={{ fontSize:18,marginBottom:6 }}>⏳</div>{text}</div>;
}
function ErrBox({ msg }) {
  return <div style={{ marginTop:12,background:"#FEF2F2",border:"1px solid #FECACA",borderRadius:8,padding:"10px 14px" }}><span style={{ fontFamily:"DM Mono",fontSize:11,color:"#991B1B" }}>⚠ {msg}</span></div>;
}
function EmptyState({ icon, title, desc }) {
  return (
    <div style={{ textAlign:"center",padding:"44px 20px" }}>
      <div style={{ fontSize:30,marginBottom:12 }}>{icon}</div>
      <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:14,color:"#44403C",marginBottom:8 }}>{title}</div>
      <div style={{ fontFamily:"DM Mono",fontSize:11,color:"#78716C",lineHeight:1.7 }}>{desc}</div>
    </div>
  );
}
function ChartBox({ title, children }) {
  return (
    <div style={{ background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:14,padding:"18px 16px",boxShadow:"0 1px 3px rgba(28,25,23,0.06)" }}>
      <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:12,color:"#78716C",marginBottom:14 }}>{title}</div>
      {children}
    </div>
  );
}
function FoodCard({ label, color, total, count, budget, avg }) {
  const over = count > 0 && total > budget;
  return (
    <div style={{ background:"#FFFFFF",border:`1px solid ${color}30`,borderRadius:14,padding:20,boxShadow:"0 1px 3px rgba(28,25,23,0.06)" }}>
      <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:10,color,letterSpacing:1,textTransform:"uppercase",marginBottom:10 }}>{label}</div>
      <div style={{ fontFamily:"Syne",fontWeight:800,fontSize:24,color:"#1C1917" }}>CHF {total.toFixed(2)}</div>
      <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#A8A29E",marginTop:3 }}>
        {count > 0 ? avg : "No data — import CSV or scan receipts"}
      </div>
      {count > 0 && <div style={{ marginTop:10,fontFamily:"DM Mono",fontSize:11,color:over?"#E4003A":"#10B981" }}>
        {over ? `⚠ CHF ${(total-budget).toFixed(2)} over budget` : "✓ Within budget"}
      </div>}
    </div>
  );
}

/* ── Main App ── */
export default function App() {
  const [tab, setTab]             = useState("overview");
  const [activeAcc, setActiveAcc] = useState("neon");
  const [transactions, setTx]     = useState([]);
  const [documents, setDocs]      = useState([
    { id:"demo1",vendor:"Helsana Health Insurance",date:"01 May 2025",amount:380,currency:"CHF",documentType:"insurance",taxDeductible:true,taxNote:"Declare under Versicherungsprämien — CHF 380/mo × 12 = CHF 4,560/yr",summary:"Monthly health insurance premium.",category:"Health" },
    { id:"demo2",vendor:"SBB GA Travelcard 2025",date:"01 Jan 2025",amount:3860,currency:"CHF",documentType:"contract",taxDeductible:true,taxNote:"Deductible as Berufskosten if used for commuting to work",summary:"Annual GA for all Swiss public transport.",category:"Transport" },
  ]);
  const [groceryReceipts, setGR]  = useState([]);
  const [budgets, setBudgets]     = useState(DEFAULT_BUDGETS);
  const [accounts, setAccounts]   = useState(DEFAULT_ACCOUNTS);
  const [editingAcc, setEditingAcc] = useState(null);
  const [loaded, setLoaded]       = useState(false);
  const [subs, setSubs]             = useState(DEFAULT_SUBS);
  const [investments, setInvestments] = useState(DEFAULT_INVESTMENTS);
  const [monthlyIncome, setMonthlyIncome] = useState(0);
  // Child / Family
  const [childName, setChildName]   = useState("My Child");
  const [childExpenses, setChildExp]= useState([]);
  const [showAddChild, setShowAddChild] = useState(false);
  const [newChild, setNewChild]     = useState({ desc:"", amount:"", cat:"Child/School", date: new Date().toISOString().slice(0,10), note:"" });
  // Quick Log
  const [showQL, setShowQL]         = useState(false);
  const [ql, setQl]                 = useState({ amount:"", cat:"Lunch/Dining", desc:"", acc:"neon" });
  const [showAddSub, setShowAddSub] = useState(false);
  const [newSub, setNewSub]         = useState({ name:"", amount:"", freq:"monthly", cat:"Other", icon:"💳", taxDeductible:false });

  const [showImport, setShowImport] = useState(false);
  const [csvAccId, setCsvAccId]     = useState("neon");
  const [csvPreview, setCsvPreview] = useState(null);

  const [ocrLoading, setOcrLoading] = useState(false);
  const [ocrResult, setOcrResult]   = useState(null);
  const [ocrErr, setOcrErr]         = useState(null);

  const [grLoading, setGrLoading]   = useState(false);
  const [grResult, setGrResult]     = useState(null);
  const [grErr, setGrErr]           = useState(null);

  const [viewDoc, setViewDoc]       = useState(null);
  const [search, setSearch]         = useState("");
  const [filterCat, setFilterCat]   = useState("All");
  const [filterAcc, setFilterAcc]   = useState("All");

  useEffect(() => {
    (async () => {
      const t = await store.get("ml_tx");
      const d = await store.get("ml_docs");
      const g = await store.get("ml_gr");
      const b = await store.get("ml_budgets");
      if (t?.length) setTx(t);
      if (d?.length) setDocs(d);
      if (g?.length) setGR(g);
      if (b) setBudgets(b);
      const s = await store.get("ml_subs");
      const inv = await store.get("ml_inv");
      const inc = await store.get("ml_income");
      if (s?.length) setSubs(s);
      if (inv?.length) setInvestments(inv);
      if (inc) setMonthlyIncome(inc);
      const ch = await store.get("ml_child");
      const cn = await store.get("ml_childname");
      if (ch?.length) setChildExp(ch);
      if (cn) setChildName(cn);
      setLoaded(true);
    })();
  }, []);

  useEffect(() => { if (loaded) api.setSetting("budgets", budgets).catch(()=>{}); }, [budgets, loaded]);
  useEffect(() => { if (loaded) api.setSetting("subs", subs).catch(()=>{}); }, [subs, loaded]);
  useEffect(() => { if (loaded) api.setSetting("investments", investments).catch(()=>{}); }, [investments, loaded]);
  useEffect(() => { if (loaded) api.setSetting("income", monthlyIncome).catch(()=>{}); }, [monthlyIncome, loaded]);
  useEffect(() => { if (loaded) api.setSetting("childname", childName).catch(()=>{}); }, [childName, loaded]);
  useEffect(() => { if (loaded) api.setSetting("accounts", accounts).catch(()=>{}); }, [accounts, loaded]);

  const now = new Date();
  const cm = now.getMonth(), cy = now.getFullYear();
  const monthlyTx = transactions.filter(t => { const d=new Date(t.date); return d.getMonth()===cm&&d.getFullYear()===cy&&t.amount<0; });
  const spendByCat = Object.entries(monthlyTx.reduce((acc,t)=>{ acc[t.category]=(acc[t.category]||0)+Math.abs(t.amount); return acc; },{}))
    .map(([cat,amount])=>({ cat,amount:Math.round(amount*100)/100,color:CAT_COLORS[cat]||CAT_COLORS.Other,budget:budgets[cat]||100 }))
    .sort((a,b)=>b.amount-a.amount);
  const totalSpend  = spendByCat.reduce((s,c)=>s+c.amount,0);
  const lunchTx     = monthlyTx.filter(t=>t.category==="Lunch/Dining");
  const groceryTx   = monthlyTx.filter(t=>t.category==="Groceries");
  const lunchTotal  = lunchTx.reduce((s,t)=>s+Math.abs(t.amount),0);
  const groceryTotal= groceryTx.reduce((s,t)=>s+Math.abs(t.amount),0);
  const taxDocs     = documents.filter(d=>d.taxDeductible);
  const taxTotal    = taxDocs.reduce((s,d)=>s+(d.amount||0),0);

  // ── Child calculations ──
  const CHILD_CATS = ["Child/School","Child/Other"];
  const childTxFromImport = transactions.filter(t => CHILD_CATS.includes(t.category));
  const allChildExpenses = [
    ...childExpenses.map(e => ({...e, source:"manual"})),
    ...childTxFromImport.map(t => ({...t, desc:t.desc, amount:Math.abs(t.amount), source:"import"}))
  ].sort((a,b) => new Date(b.date)-new Date(a.date));
  const childTotal = allChildExpenses.reduce((s,e)=>s+Number(e.amount),0);
  const childBySubcat = Object.entries(allChildExpenses.reduce((acc,e)=>{ const k=e.cat||"Other"; acc[k]=(acc[k]||0)+Number(e.amount); return acc; },{})).sort((a,b)=>b[1]-a[1]);
  const childThisMonth = allChildExpenses.filter(e=>{ const d=new Date(e.date); return d.getMonth()===cm&&d.getFullYear()===cy; }).reduce((s,e)=>s+Number(e.amount),0);
  // ── Annual calculations ──
  const annualIncome = monthlyIncome * 12;
  const subAnnual = subs.map(s => ({ ...s, annual: s.amount * FREQ_MULT[s.freq] }));
  const totalAnnualSubs = subAnnual.reduce((a,s) => a + s.annual, 0);
  const projectedAnnualSpend = totalSpend > 0 ? totalSpend * 12 : totalAnnualSubs;
  const annualBalance = annualIncome - projectedAnnualSpend;

  // ── Week calculations ──
  const startOfWeek = new Date(now); startOfWeek.setDate(now.getDate() - now.getDay() + 1); startOfWeek.setHours(0,0,0,0);
  const startOfLastWeek = new Date(startOfWeek); startOfLastWeek.setDate(startOfLastWeek.getDate() - 7);
  const endOfLastWeek = new Date(startOfWeek); endOfLastWeek.setMilliseconds(-1);
  const thisWeekTx = transactions.filter(t => { const d=new Date(t.date); return d>=startOfWeek && t.amount<0; });
  const lastWeekTx = transactions.filter(t => { const d=new Date(t.date); return d>=startOfLastWeek && d<=endOfLastWeek && t.amount<0; });
  const thisWeekTotal = thisWeekTx.reduce((s,t)=>s+Math.abs(t.amount),0);
  const lastWeekTotal = lastWeekTx.reduce((s,t)=>s+Math.abs(t.amount),0);
  const thisWeekByCat = Object.entries(thisWeekTx.reduce((acc,t)=>{ acc[t.category]=(acc[t.category]||0)+Math.abs(t.amount); return acc; },{}))
    .map(([cat,amount])=>({ cat, amount, color:CAT_COLORS[cat]||"#4B5563" })).sort((a,b)=>b.amount-a.amount);

  // ── Monthly projection chart data ──
  const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const monthlyChartData = monthNames.map((m,i) => {
    const mTx = transactions.filter(t => { const d=new Date(t.date); return d.getMonth()===i && d.getFullYear()===cy && t.amount<0; });
    const spent = mTx.reduce((s,t)=>s+Math.abs(t.amount),0);
    return { m, spent: Math.round(spent), income: monthlyIncome, subs: Math.round(totalAnnualSubs/12) };
  });

  const tips = [];
  spendByCat.forEach(c => {
    if (c.budget && c.amount > c.budget) {
      const pct = ((c.amount/c.budget-1)*100).toFixed(0);
      tips.push({ level:"alert",icon:CAT_ICONS[c.cat]||"⚠️",title:`${c.cat} is ${pct}% over budget`,desc:`CHF ${c.amount.toFixed(2)} spent vs CHF ${c.budget} budget.` });
    }
  });
  if (!tips.length && spendByCat.length) tips.push({ level:"ok",icon:"✅",title:"All categories within budget!",desc:"Great spending discipline this month." });
  if (!spendByCat.length) tips.push({ level:"info",icon:"📥",title:"No transactions imported yet",desc:"Click '+ Import CSV / Excel' to load your bank statement." });
  if (taxDocs.length) tips.push({ level:"info",icon:"🏛️",title:`CHF ${taxTotal.toLocaleString()} deductible tracked`,desc:"Don't forget to declare these in your Swiss annual tax return." });

  const allCats = ["All",...Object.keys(CAT_COLORS)];
  const filteredTx = transactions
    .filter(t=>(!search||t.desc.toLowerCase().includes(search.toLowerCase()))&&(filterCat==="All"||t.category===filterCat)&&(filterAcc==="All"||t.account===filterAcc))
    .sort((a,b)=>new Date(b.date)-new Date(a.date));

  const handleCSVFile = file => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (["xlsx","xls","ods"].includes(ext)) {
      // Excel / ODS path
      const reader = new FileReader();
      reader.onload = e => {
        const buffer = new Uint8Array(e.target.result);
        setCsvPreview(parseExcel(buffer, csvAccId));
      };
      reader.readAsArrayBuffer(file);
    } else {
      // CSV / TXT path
      const reader = new FileReader();
      reader.onload = e => setCsvPreview(parseCSV(e.target.result, csvAccId));
      reader.readAsText(file, "UTF-8");
    }
  };
  const confirmImport = async () => {
    try {
      const existing = await api.getTransactions();
      const seen = new Set(existing.map(t=>`${t.date}|${t.desc}|${t.amount}`));
      const news = (csvPreview||[]).filter(t=>!seen.has(`${t.date}|${t.desc}|${t.amount}`));
      if (news.length) {
        await api.addTransactions(news);
        const all = await api.getTransactions();
        setTx(all);
      }
    } catch(e) { console.error("Import error:", e); }
    setCsvPreview(null); setShowImport(false);
  };

  const submitQL = async () => {
    if (!ql.amount) return;
    const tx = {
      id:`ql-${Date.now()}`, date: new Date().toISOString().slice(0,10),
      desc: ql.desc || ql.cat, amount: -Math.abs(Number(ql.amount)),
      account: ql.acc, category: ql.cat,
    };
    try {
      await api.addTransactions([tx]);
      setTx(prev => [tx, ...prev]);
    } catch(e) { console.error("Quick log error:", e); }
    setQl({ amount:"", cat:"Lunch/Dining", desc:"", acc:"neon" });
    setShowQL(false);
  };

  const handleDocFile = async file => {
    if (!file) return;
    setOcrLoading(true); setOcrResult(null); setOcrErr(null);
    try { setOcrResult({...(await api.ocrBill(file)),_fn:file.name}); }
    catch(e) { setOcrErr(e.message); }
    finally { setOcrLoading(false); }
  };
  const saveDoc = async () => {
    if (!ocrResult) return;
    const doc = {...ocrResult, id:`d-${Date.now()}`, vendor:ocrResult.vendor||ocrResult._fn};
    try {
      await api.addDocument(doc);
      setDocs(prev=>[doc,...prev]);
    } catch(e) { console.error("Save doc error:", e); }
    setOcrResult(null);
  };

  const saveChildEntry = async () => {
    if (!newChild.amount || !newChild.desc) return;
    const entry = { ...newChild, id:`ce-${Date.now()}`, amount:Number(newChild.amount) };
    try {
      await api.addChildExpense(entry);
      setChildExp(prev => [entry, ...prev]);
    } catch(e) { console.error("Save child error:", e); }
    setNewChild({ desc:"", amount:"", cat:"Child/School", date: new Date().toISOString().slice(0,10), note:"" });
    setShowAddChild(false);
  };

  const handleGrFile = async file => {
    if (!file) return;
    setGrLoading(true); setGrResult(null); setGrErr(null);
    try { setGrResult(await api.ocrGrocery(file)); }
    catch(e) { setGrErr(e.message); }
    finally { setGrLoading(false); }
  };
  const saveGrocery = async () => {
    if (!grResult) return;
    const rec = {...grResult, id:`gr-${Date.now()}`};
    try {
      await api.addGroceryReceipt(rec);
      setGR(prev=>[rec,...prev]);
    } catch(e) { console.error("Save grocery error:", e); }
    setGrResult(null);
  };

  return (
    <>
      <style>{`
        ${FONTS}
        *{box-sizing:border-box;margin:0;padding:0;}
        body{background:#F2EFE9;}
        ::-webkit-scrollbar{width:4px;}
        ::-webkit-scrollbar-thumb{background:#C4BDB7;border-radius:2px;}
        ::-webkit-scrollbar-track{background:#F2EFE9;}
        .br:hover{background:#A30D25!important;}
        .bg:hover{background:#F0EDE7!important;}
        .tn:hover{color:#44403C!important;background:#EAE7E1!important;}
        .hr:hover{background:#F5F3EF!important;}
        input:focus,select:focus{outline:none;}
        select option{background:#fff;color:#1C1917;}
      `}</style>

      <div style={{ minHeight:"100vh",background:"#F2EFE9",color:"#1C1917",fontFamily:"DM Mono,monospace",maxWidth:940,margin:"0 auto",display:"flex",flexDirection:"column" }}>

        {/* Header */}
        <div style={{ padding:"24px 28px 0",borderBottom:"1px solid #EAE7E1" }}>
          <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18 }}>
            <div>
              <div style={{ fontFamily:"Syne",fontWeight:800,fontSize:20,color:"#1C1917",letterSpacing:"-0.5px" }}>💰 MoneyLens</div>
              <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#A8A29E",marginTop:2 }}>
                {loaded ? `${transactions.length} transactions · ${documents.length} documents · auto-saved` : "Loading…"}
              </div>
            </div>
            <div style={{ display:"flex",gap:8 }}>
              {transactions.length>0 && (
                <button className="bg" onClick={async()=>{ if(window.confirm("Clear all imported transactions?")){ try{ await api.clearTransactions(); setTx([]); }catch(e){} } }}
                  style={{ background:"#FFFFFF",border:"1px solid #D9D5CE",color:"#78716C",fontFamily:"Syne",fontWeight:600,fontSize:11,padding:"8px 13px",borderRadius:8,cursor:"pointer",transition:"background 0.15s" }}>
                  Clear TX
                </button>
              )}
              <button className="br" onClick={()=>setShowImport(true)}
                style={{ background:"#C8102E",border:"none",color:"#fff",fontFamily:"Syne",fontWeight:700,fontSize:12,padding:"9px 16px",borderRadius:8,cursor:"pointer",transition:"background 0.18s" }}>
                + Import CSV / Excel
              </button>
            </div>
          </div>
          <div style={{ display:"flex",gap:2,overflowX:"auto" }}>
            {TABS.map(t=>(
              <button key={t.id} className="tn" onClick={()=>setTab(t.id)} style={{
                background:tab===t.id?"#FFFFFF":"none",
                border:"none",borderBottom:`2px solid ${tab===t.id?"#C8102E":"transparent"}`,
                color:tab===t.id?"#C8102E":"#78716C",fontFamily:"Syne",fontWeight:600,fontSize:12,
                padding:"9px 14px",cursor:"pointer",whiteSpace:"nowrap",
                transition:"all 0.15s",borderRadius:"7px 7px 0 0",
              }}>{t.icon} {t.label}</button>
            ))}
          </div>
        </div>

        <div style={{ padding:"24px 28px",flex:1 }}>

          {/* OVERVIEW */}
          {tab==="overview" && <>
            <SL>Accounts</SL>
            <div style={{ display:"flex",gap:10,flexWrap:"wrap",marginBottom:24 }}>
              {accounts.map(a=><AccountCard key={a.id} acc={a} active={activeAcc===a.id} onClick={()=>setActiveAcc(a.id)}
                onEditBalance={(id,val)=>setAccounts(prev=>prev.map(x=>x.id===id?{...x,balance:val}:x))} />)}
            </div>
            <div style={{ display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:24 }}>
              <StatCard label="Spent This Month" value={transactions.length?`CHF ${totalSpend.toFixed(0)}`:"—"} sub={`${monthlyTx.length} transactions`} />
              <StatCard label="Tax Deductible" value={`CHF ${taxTotal.toLocaleString()}`} sub={`${taxDocs.length} documents`} good />
              <StatCard label="Lunch avg/meal" value={lunchTx.length?`CHF ${(lunchTotal/lunchTx.length).toFixed(2)}`:"—"} sub={`${lunchTx.length} meals logged`} />
            </div>
            <SL>Smart Tips</SL>
            {tips.map((t,i)=><TipCard key={i} tip={t} />)}
          </>}

          {/* TRANSACTIONS */}
          {tab==="transactions" && <>
            <div style={{ display:"flex",gap:10,marginBottom:18,flexWrap:"wrap" }}>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search transactions…"
                style={{ flex:1,minWidth:160,background:"#FFFFFF",border:"1px solid #D9D5CE",borderRadius:8,color:"#1C1917",fontFamily:"DM Mono",fontSize:12,padding:"8px 12px" }} />
              <select value={filterCat} onChange={e=>setFilterCat(e.target.value)}
                style={{ background:"#FFFFFF",border:"1px solid #D9D5CE",borderRadius:8,color:"#78716C",fontFamily:"DM Mono",fontSize:11,padding:"8px 10px",cursor:"pointer" }}>
                {allCats.map(c=><option key={c} value={c}>{c}</option>)}
              </select>
              <select value={filterAcc} onChange={e=>setFilterAcc(e.target.value)}
                style={{ background:"#FFFFFF",border:"1px solid #D9D5CE",borderRadius:8,color:"#78716C",fontFamily:"DM Mono",fontSize:11,padding:"8px 10px",cursor:"pointer" }}>
                <option value="All">All Accounts</option>
                {ACCOUNTS.map(a=><option key={a.id} value={a.id}>{a.name}</option>)}
              </select>
            </div>
            {filteredTx.length===0
              ? <EmptyState icon="📥" title="No transactions" desc={transactions.length?"Try adjusting your filters.":"Click '+ Import CSV / Excel' to load your bank statement."} />
              : filteredTx.slice(0,150).map(t=>(
                <div key={t.id} className="hr" style={{ display:"flex",alignItems:"center",gap:12,padding:"10px 13px",background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:10,marginBottom:4,transition:"background 0.12s",boxShadow:"0 1px 2px rgba(28,25,23,0.04)" }}>
                  <div style={{ width:34,height:34,borderRadius:8,flexShrink:0,background:`${CAT_COLORS[t.category]||"#4B5563"}18`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14 }}>
                    {CAT_ICONS[t.category]||"💳"}
                  </div>
                  <div style={{ flex:1,minWidth:0 }}>
                    <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:12,color:"#1C1917",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis" }}>{t.desc}</div>
                    <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#A8A29E",marginTop:2 }}>
                      {t.date} · <span style={{ color:CAT_COLORS[t.category]||"#555" }}>{t.category}</span>
                    </div>
                  </div>
                  <select value={t.category} onChange={async e=>{ const cat=e.target.value; try{ await api.updateTxCategory(t.id,cat); setTx(prev=>prev.map(x=>x.id===t.id?{...x,category:cat}:x)); }catch(e){} }}
                    style={{ background:"transparent",border:"none",color:CAT_COLORS[t.category]||"#555",fontFamily:"DM Mono",fontSize:10,cursor:"pointer",maxWidth:110 }}>
                    {Object.keys(CAT_COLORS).map(c=><option key={c} value={c}>{c}</option>)}
                  </select>
                  <div style={{ fontFamily:"DM Mono",fontSize:13,fontWeight:500,flexShrink:0,color:t.amount<0?"#991B1B":"#166534" }}>
                    {t.amount<0?"-":"+"}CHF {Math.abs(t.amount).toFixed(2)}
                  </div>
                  <button onClick={async()=>{ try{ await api.deleteTransaction(t.id); setTx(prev=>prev.filter(x=>x.id!==t.id)); }catch(e){} }}
                    style={{ background:"none",border:"none",color:"#D9D5CE",cursor:"pointer",fontSize:13,padding:"2px 4px" }}>✕</button>
                </div>
              ))
            }
            {filteredTx.length>150&&<div style={{ textAlign:"center",fontFamily:"DM Mono",fontSize:10,color:"#2a2a2a",padding:"10px 0" }}>Showing 150 of {filteredTx.length}</div>}
          </>}

          {/* SPENDING */}
          {tab==="spending" && (
            spendByCat.length===0
              ? <EmptyState icon="📊" title="No spending data" desc="Import your CSV transactions to see spending charts." />
              : <>
                <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:18,marginBottom:22 }}>
                  <ChartBox title="By Category">
                    <ResponsiveContainer width="100%" height={185}>
                      <PieChart>
                        <Pie data={spendByCat} cx="50%" cy="50%" innerRadius={50} outerRadius={78} dataKey="amount" strokeWidth={0}>
                          {spendByCat.map((e,i)=><Cell key={i} fill={e.color||"#4B5563"} />)}
                        </Pie>
                        <Tooltip formatter={v=>[`CHF ${Number(v).toFixed(2)}`,"Spent"]} contentStyle={{ background:"#1a1a1a",border:"1px solid #222",borderRadius:8,fontFamily:"DM Mono",fontSize:11 }} />
                      </PieChart>
                    </ResponsiveContainer>
                    <div style={{ textAlign:"center",fontFamily:"Syne",fontWeight:800,fontSize:20,color:"#1C1917",marginTop:-8 }}>CHF {totalSpend.toFixed(2)}</div>
                    <div style={{ textAlign:"center",fontFamily:"DM Mono",fontSize:10,color:"#A8A29E" }}>total this month</div>
                    <div style={{ display:"flex",flexWrap:"wrap",gap:"5px 10px",marginTop:12 }}>
                      {spendByCat.map((c,i)=>(
                        <div key={i} style={{ display:"flex",alignItems:"center",gap:5 }}>
                          <div style={{ width:7,height:7,borderRadius:"50%",background:c.color||"#4B5563",flexShrink:0 }} />
                          <span style={{ fontFamily:"DM Mono",fontSize:10,color:"#A8A29E" }}>{c.cat}</span>
                        </div>
                      ))}
                    </div>
                  </ChartBox>
                  <ChartBox title="Spent vs Budget (top 6)">
                    <ResponsiveContainer width="100%" height={215}>
                      <BarChart data={spendByCat.slice(0,6)} layout="vertical" margin={{ left:0,right:8,top:0,bottom:0 }}>
                        <CartesianGrid strokeDasharray="2 4" stroke="#F0EDE7" horizontal={false} />
                        <XAxis type="number" tick={{ fill:"#A8A29E",fontSize:9,fontFamily:"DM Mono" }} axisLine={false} tickLine={false} />
                        <YAxis type="category" dataKey="cat" tick={{ fill:"#A8A29E",fontSize:9,fontFamily:"DM Mono" }} axisLine={false} tickLine={false} width={72} />
                        <Tooltip formatter={v=>[`CHF ${Number(v).toFixed(2)}`]} contentStyle={{ background:"#1a1a1a",border:"1px solid #222",borderRadius:8,fontFamily:"DM Mono",fontSize:11 }} />
                        <Bar dataKey="budget" fill="#EAE7E1" radius={3} name="Budget" />
                        <Bar dataKey="amount" radius={3} name="Spent">
                          {spendByCat.slice(0,6).map((e,i)=><Cell key={i} fill={e.amount>e.budget?"#E4003A":e.color} />)}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </ChartBox>
                </div>
                <SL>All Categories · click budget to edit</SL>
                <div style={{ background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:12,padding:"6px 16px",boxShadow:"0 1px 3px rgba(28,25,23,0.06)" }}>
                  {spendByCat.map((c,i)=>{
                    const pct=Math.round((c.amount/(c.budget||1))*100);
                    const over=c.amount>(c.budget||0);
                    return (
                      <div key={i} className="hr" style={{ display:"grid",gridTemplateColumns:"118px 1fr 90px 76px",alignItems:"center",gap:12,padding:"9px 4px",borderBottom:i<spendByCat.length-1?"1px solid #F0EDE7":"none",transition:"background 0.12s" }}>
                        <div style={{ display:"flex",alignItems:"center",gap:7 }}>
                          <span style={{ fontSize:12 }}>{CAT_ICONS[c.cat]||"💳"}</span>
                          <span style={{ fontFamily:"Syne",fontSize:11,color:"#44403C" }}>{c.cat}</span>
                        </div>
                        <div style={{ height:5,background:"#EAE7E1",borderRadius:3,overflow:"hidden" }}>
                          <div style={{ height:"100%",width:`${Math.min(pct,100)}%`,background:over?"#E4003A":c.color||"#4B5563",borderRadius:3,transition:"width 0.3s" }} />
                        </div>
                        <div style={{ fontFamily:"DM Mono",fontSize:11,color:over?"#991B1B":"#44403C",textAlign:"right" }}>CHF {c.amount.toFixed(2)}</div>
                        <div style={{ display:"flex",alignItems:"center",gap:3 }}>
                          <span style={{ fontFamily:"DM Mono",fontSize:9,color:"#C4BDB7" }}>/ </span>
                          <input type="number" value={c.budget||100}
                            onChange={e=>setBudgets(prev=>({...prev,[c.cat]:Number(e.target.value)}))}
                            style={{ width:52,background:"transparent",border:"none",color:"#78716C",fontFamily:"DM Mono",fontSize:11,textAlign:"right" }} />
                        </div>
                      </div>
                    );
                  })}
                </div>
              </>
          )}

          {/* FOOD */}
          {tab==="food" && <>
            <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:22 }}>
              <FoodCard label="🍽️ Mittag / Lunch" color="#E4003A" total={lunchTotal} count={lunchTx.length}
                budget={budgets["Lunch/Dining"]||280}
                avg={lunchTx.length?`${lunchTx.length} meals · avg CHF ${(lunchTotal/lunchTx.length).toFixed(2)}`:`${lunchTx.length} meals`} />
              <FoodCard label="🛒 Groceries" color="#F59E0B" total={groceryTotal} count={groceryTx.length}
                budget={budgets["Groceries"]||450}
                avg={groceryTx.length?`${groceryTx.length} shops this month`:""} />
            </div>
            {(lunchTx.length>0||groceryTx.length>0) && <>
              <SL>Recent Food Transactions</SL>
              {[...lunchTx,...groceryTx].sort((a,b)=>new Date(b.date)-new Date(a.date)).slice(0,20).map((t,i)=>(
                <div key={i} className="hr" style={{ display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 14px",background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:10,marginBottom:5,transition:"background 0.12s",boxShadow:"0 1px 2px rgba(28,25,23,0.04)" }}>
                  <div style={{ display:"flex",alignItems:"center",gap:10 }}>
                    <span style={{ fontSize:16 }}>{t.category==="Lunch/Dining"?"🍽️":"🛒"}</span>
                    <div>
                      <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:12,color:"#1C1917" }}>{t.desc}</div>
                      <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#A8A29E" }}>{t.date}</div>
                    </div>
                  </div>
                  <span style={{ fontFamily:"DM Mono",fontSize:13,color:t.category==="Lunch/Dining"?"#E4003A":"#F59E0B" }}>CHF {Math.abs(t.amount).toFixed(2)}</span>
                </div>
              ))}
            </>}
            <div style={{ marginTop:22 }}>
              <SL>Scan Grocery Receipt (AI OCR)</SL>
              <UploadZone label="Upload grocery receipt photo or PDF" onFile={handleGrFile} />
              {grLoading && <Loader text="Claude is reading the receipt and extracting items…" />}
              {grErr && <ErrBox msg={grErr} />}
              {grResult&&!grErr && (
                <div style={{ marginTop:14,background:"#FFFBEB",border:"1px solid #FDE68A",borderRadius:12,padding:18 }}>
                  <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14 }}>
                    <div>
                      <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:14,color:"#F59E0B" }}>{grResult.store}</div>
                      <div style={{ fontFamily:"DM Mono",fontSize:11,color:"#3d3d3d" }}>{grResult.date}</div>
                    </div>
                    <div style={{ fontFamily:"DM Mono",fontSize:18,color:"#F59E0B" }}>CHF {grResult.total?.toFixed(2)}</div>
                  </div>
                  {(grResult.items||[]).map((item,i)=>(
                    <div key={i} style={{ display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:i<grResult.items.length-1?"1px solid #F0EDE7":"none" }}>
                      <div style={{ display:"flex",gap:8,alignItems:"center" }}>
                        <span style={{ fontFamily:"DM Mono",fontSize:11,color:"#1C1917" }}>{item.name}</span>
                        {item.qty>1&&<span style={{ fontFamily:"DM Mono",fontSize:9,color:"#C4BDB7" }}>×{item.qty}</span>}
                        <span style={{ fontFamily:"Syne",fontSize:9,color:"#78716C",padding:"1px 5px",background:"#F0EDE7",borderRadius:3 }}>{item.category}</span>
                      </div>
                      <span style={{ fontFamily:"DM Mono",fontSize:11,color:"#78716C" }}>CHF {(item.unitPrice*item.qty).toFixed(2)}</span>
                    </div>
                  ))}
                  <button className="br" onClick={saveGrocery}
                    style={{ marginTop:14,background:"#C8102E",border:"none",color:"#fff",fontFamily:"Syne",fontWeight:700,fontSize:12,padding:"8px 16px",borderRadius:8,cursor:"pointer",transition:"background 0.18s" }}>
                    Save Receipt
                  </button>
                </div>
              )}
              {groceryReceipts.length>0 && <>
                <div style={{ marginTop:20 }}><SL>Saved Receipts</SL></div>
                {groceryReceipts.map((r,i)=>(
                  <div key={i} className="hr" style={{ display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 14px",background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:10,marginBottom:6,transition:"background 0.12s",boxShadow:"0 1px 2px rgba(28,25,23,0.04)" }}>
                    <div>
                      <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:12,color:"#1C1917" }}>{r.store}</div>
                      <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#A8A29E" }}>{r.date} · {r.items?.length||0} items</div>
                    </div>
                    <div style={{ display:"flex",alignItems:"center",gap:10 }}>
                      <span style={{ fontFamily:"DM Mono",fontSize:13,color:"#92400E" }}>CHF {r.total?.toFixed(2)}</span>
                      <button onClick={async()=>{ try{ await api.deleteGroceryReceipt(r.id); setGR(prev=>prev.filter(x=>x.id!==r.id)); }catch(e){} }} style={{ background:"none",border:"none",color:"#D9D5CE",cursor:"pointer",fontSize:13 }}>✕</button>
                    </div>
                  </div>
                ))}
              </>}
            </div>
          </>}

          {/* DOCUMENTS */}
          {tab==="documents" && <>
            <SL>Saved Documents</SL>
            {documents.map(doc=>(
              <div key={doc.id} style={{ display:"flex",alignItems:"center",gap:13,padding:"13px 15px",background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:12,marginBottom:8,boxShadow:"0 1px 3px rgba(28,25,23,0.06)" }}>
                <div style={{ width:38,height:38,borderRadius:9,flexShrink:0,background:doc.taxDeductible?"#F0FDF4":"#F5F3EF",border:`1px solid ${doc.taxDeductible?"#BBF7D0":"#E2DED8"}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:15 }}>
                  {CAT_ICONS[doc.category]||"📄"}
                </div>
                <div style={{ flex:1,minWidth:0 }}>
                  <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:13,color:"#1C1917" }}>{doc.vendor}</div>
                  <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#A8A29E" }}>{doc.date} · {doc.currency} {doc.amount}</div>
                  {doc.taxDeductible&&(
                    <div style={{ marginTop:5,display:"inline-flex",alignItems:"center",background:"#DCFCE7",border:"1px solid #BBF7D0",borderRadius:4,padding:"2px 7px" }}>
                      <span style={{ color:"#166534",fontSize:9,fontFamily:"Syne",fontWeight:700 }}>✓ TAX DEDUCTIBLE</span>
                    </div>
                  )}
                </div>
                <button className="bg" onClick={()=>setViewDoc(doc)} style={{ background:"none",border:"1px solid #D9D5CE",color:"#78716C",borderRadius:7,padding:"5px 12px",fontFamily:"Syne",fontSize:11,cursor:"pointer",transition:"background 0.15s" }}>View</button>
                <button onClick={async()=>{ try{ await api.deleteDocument(doc.id); setDocs(prev=>prev.filter(d=>d.id!==doc.id)); }catch(e){} }} style={{ background:"none",border:"none",color:"#D9D5CE",cursor:"pointer",fontSize:13,padding:"2px 4px" }}>✕</button>
              </div>
            ))}
            <div style={{ marginTop:22 }}>
              <SL>Upload & Scan (AI OCR via Claude)</SL>
              <UploadZone label="Upload bill, receipt, insurance or contract" onFile={handleDocFile} />
              {ocrLoading && <Loader text="Claude is analyzing your document…" />}
              {ocrErr && <ErrBox msg={ocrErr} />}
              {ocrResult&&(
                <div style={{ marginTop:14,background:"#F0FDF4",border:"1px solid #BBF7D0",borderRadius:12,padding:20 }}>
                  <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:13,color:"#166534",marginBottom:16 }}>✓ Document analyzed by Claude</div>
                  {[["Vendor",ocrResult.vendor],["Amount",`${ocrResult.currency} ${ocrResult.amount}`],["Date",ocrResult.date],["Type",ocrResult.documentType],["Summary",ocrResult.summary],["Tax note",ocrResult.taxNote]].map(([k,v])=>(
                    <div key={k} style={{ display:"flex",gap:12,marginBottom:10 }}>
                      <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:10,color:"#A8A29E",width:68,flexShrink:0,paddingTop:1 }}>{k}</div>
                      <div style={{ fontFamily:"DM Mono",fontSize:11,color:"#44403C",lineHeight:1.6 }}>{String(v)}</div>
                    </div>
                  ))}
                  {ocrResult.taxDeductible&&(
                    <div style={{ background:"#DCFCE7",border:"1px solid #BBF7D0",borderRadius:8,padding:"8px 14px",marginTop:8,marginBottom:14 }}>
                      <span style={{ color:"#166534",fontFamily:"Syne",fontWeight:700,fontSize:11 }}>✓ Likely tax deductible — will appear in Tax Summary</span>
                    </div>
                  )}
                  <button className="br" onClick={saveDoc}
                    style={{ background:"#C8102E",border:"none",color:"#fff",fontFamily:"Syne",fontWeight:700,fontSize:12,padding:"8px 18px",borderRadius:8,cursor:"pointer",transition:"background 0.18s" }}>
                    Save Document
                  </button>
                </div>
              )}
            </div>
            {viewDoc&&(
              <div onClick={()=>setViewDoc(null)} style={{ position:"fixed",inset:0,background:"rgba(28,25,23,0.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:999,padding:24 }}>
                <div onClick={e=>e.stopPropagation()} style={{ background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:16,padding:28,maxWidth:440,width:"100%",boxShadow:"0 20px 60px rgba(28,25,23,0.15)" }}>
                  <div style={{ fontFamily:"Syne",fontWeight:800,fontSize:16,color:"#1C1917",marginBottom:20 }}>{viewDoc.vendor}</div>
                  {[["Date",viewDoc.date],["Amount",`${viewDoc.currency} ${viewDoc.amount}`],["Type",viewDoc.documentType],["Summary",viewDoc.summary],["Tax note",viewDoc.taxNote]].map(([k,v])=>(
                    <div key={k} style={{ marginBottom:14 }}>
                      <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:10,color:"#A8A29E",letterSpacing:1,textTransform:"uppercase",marginBottom:4 }}>{k}</div>
                      <div style={{ fontFamily:"DM Mono",fontSize:12,color:"#44403C",lineHeight:1.6 }}>{String(v)}</div>
                    </div>
                  ))}
                  {viewDoc.taxDeductible&&(
                    <div style={{ background:"#DCFCE7",border:"1px solid #BBF7D0",borderRadius:8,padding:"10px 14px",marginTop:8 }}>
                      <span style={{ color:"#166534",fontFamily:"Syne",fontWeight:700,fontSize:12 }}>✓ Include in Swiss annual tax return</span>
                    </div>
                  )}
                  <button onClick={()=>setViewDoc(null)} style={{ marginTop:20,width:"100%",background:"#F5F3EF",border:"1px solid #EAE7E1",color:"#78716C",borderRadius:8,padding:"10px",fontFamily:"Syne",fontSize:12,cursor:"pointer" }}>Close</button>
                </div>
              </div>
            )}
          </>}

          {/* TAXES */}
          {tab==="taxes" && <>
            <div style={{ background:"#F0FDF4",border:"1px solid #BBF7D0",borderRadius:14,padding:22,marginBottom:24 }}>
              <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#166534",marginBottom:6,letterSpacing:1 }}>ESTIMATED DEDUCTIONS · TAX YEAR 2025</div>
              <div style={{ fontFamily:"Syne",fontWeight:800,fontSize:28,color:"#166534" }}>CHF {taxTotal.toLocaleString()}</div>
              <div style={{ fontFamily:"DM Mono",fontSize:11,color:"#78716C",marginTop:4 }}>{taxDocs.length} deductible documents tracked</div>
            </div>
            <SL>Deductible Documents</SL>
            {taxDocs.length===0
              ? <EmptyState icon="🏛️" title="No tax documents yet" desc="Upload bills and receipts in the Documents tab — Claude will flag what is deductible." />
              : taxDocs.map((doc,i)=>(
                <div key={i} style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start",background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:10,padding:"14px 16px",marginBottom:8,boxShadow:"0 1px 2px rgba(28,25,23,0.04)" }}>
                  <div>
                    <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:13,color:"#1C1917",marginBottom:4 }}>{doc.vendor}</div>
                    <div style={{ fontFamily:"DM Mono",fontSize:11,color:"#78716C",marginBottom:8,lineHeight:1.6 }}>{doc.taxNote}</div>
                    <span style={{ padding:"2px 8px",background:"#DCFCE7",border:"1px solid #BBF7D0",borderRadius:4 }}>
                      <span style={{ color:"#166534",fontSize:9,fontFamily:"Syne",fontWeight:700 }}>{doc.category||doc.documentType}</span>
                    </span>
                  </div>
                  <div style={{ fontFamily:"DM Mono",fontSize:14,color:"#166534",flexShrink:0,marginLeft:16 }}>{doc.currency} {Number(doc.amount).toLocaleString()}</div>
                </div>
              ))
            }
            <div style={{ marginTop:24,background:"#EFF6FF",border:"1px solid #BFDBFE",borderRadius:12,padding:18 }}>
              <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:13,color:"#1E40AF",marginBottom:12 }}>💡 Swiss Tax Checklist 2025</div>
              {["Krankenkasse premiums → Versicherungsprämien (max CHF 2,700 single / CHF 5,400 couple)","SBB GA → Berufskosten if used for commuting to your workplace","Home office → CHF 2,400 flat rate if you work from home regularly","Professional training (Weiterbildung) → deductible if related to current job","Donations to Swiss NPOs → deductible up to 20% of net income","Pillar 3a → CHF 7,056 max (2025) for employed; CHF 35,280 for self-employed"].map((tip,i)=>(
                <div key={i} style={{ fontFamily:"DM Mono",fontSize:11,color:"#44403C",marginBottom:8,paddingLeft:12,borderLeft:"2px solid #BFDBFE",lineHeight:1.7 }}>{tip}</div>
              ))}
            </div>
          </>}


          {/* ANNUAL VIEW */}
          {tab==="annual" && <>

            {/* ── Income input ── */}
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:24 }}>
              <div style={{ background:"rgba(16,185,129,0.06)", border:"1px solid rgba(16,185,129,0.18)", borderRadius:14, padding:20 }}>
                <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:10,color:"#166534",letterSpacing:1,marginBottom:10,textTransform:"uppercase" }}>💼 Monthly Net Income</div>
                <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                  <span style={{ fontFamily:"DM Mono",fontSize:16,color:"#555" }}>CHF</span>
                  <input type="number" value={monthlyIncome||""} placeholder="e.g. 7500"
                    onChange={e=>setMonthlyIncome(Number(e.target.value))}
                    style={{ flex:1, background:"transparent", border:"none", borderBottom:"1px solid rgba(255,255,255,0.12)", color:"#fff", fontFamily:"Syne", fontWeight:700, fontSize:22, padding:"4px 0", width:"100%" }} />
                </div>
                <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#A8A29E",marginTop:8 }}>
                  {monthlyIncome > 0 ? `CHF ${(monthlyIncome*12).toLocaleString()} / year` : "Enter your net monthly salary"}
                </div>
              </div>
              <div style={{ background: annualBalance >= 0 ? "rgba(16,185,129,0.06)" : "rgba(228,0,58,0.06)",
                border:`1px solid ${annualBalance>=0?"rgba(16,185,129,0.18)":"rgba(228,0,58,0.18)"}`, borderRadius:14, padding:20 }}>
                <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:10,color:annualBalance>=0?"#10B981":"#E4003A",letterSpacing:1,marginBottom:10,textTransform:"uppercase" }}>
                  {annualBalance>=0?"✅ Annual Balance":"⚠ Annual Balance"}
                </div>
                <div style={{ fontFamily:"Syne",fontWeight:800,fontSize:22,color:annualBalance>=0?"#10B981":"#E4003A" }}>
                  {annualBalance>=0?"+":""}{annualIncome>0?`CHF ${annualBalance.toLocaleString()}`:"—"}
                </div>
                <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#A8A29E",marginTop:8 }}>
                  CHF {projectedAnnualSpend.toLocaleString()} projected spend / year
                </div>
              </div>
            </div>

            {/* ── Weekly Summary ── */}
            <SL>📅 This Week vs Last Week</SL>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:12, marginBottom:22 }}>
              {[
                { label:"This Week", value:`CHF ${thisWeekTotal.toFixed(2)}`, sub:`${thisWeekTx.length} transactions`, col:"#fff" },
                { label:"Last Week", value:`CHF ${lastWeekTotal.toFixed(2)}`, sub:`${lastWeekTx.length} transactions`, col:"#555" },
                { label:"Change", value: lastWeekTotal>0 ? `${thisWeekTotal>lastWeekTotal?"+":""}${(((thisWeekTotal-lastWeekTotal)/lastWeekTotal)*100).toFixed(1)}%` : "—",
                  sub: thisWeekTotal>lastWeekTotal ? "spending up ↑" : "spending down ↓",
                  col: thisWeekTotal<=lastWeekTotal ? "#10B981" : "#E4003A" },
              ].map((s,i)=>(
                <div key={i} style={{ background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 3px rgba(28,25,23,0.06)" }}>
                  <div style={{ fontFamily:"DM Mono",fontSize:9,color:"#A8A29E",textTransform:"uppercase",letterSpacing:1,marginBottom:8 }}>{s.label}</div>
                  <div style={{ fontFamily:"Syne",fontWeight:800,fontSize:17,color:s.col==="#fff"?"#1C1917":s.col,marginBottom:3 }}>{s.value}</div>
                  <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#A8A29E" }}>{s.sub}</div>
                </div>
              ))}
            </div>

            {/* Weekly breakdown by cat */}
            {thisWeekByCat.length > 0 && (
              <div style={{ background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:12,padding:"12px 16px",marginBottom:24,boxShadow:"0 1px 3px rgba(28,25,23,0.06)" }}>
                <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:11,color:"#78716C",marginBottom:12 }}>This week by category</div>
                {thisWeekByCat.map((c,i)=>(
                  <div key={i} style={{ display:"flex",alignItems:"center",gap:12,padding:"6px 0",borderBottom:i<thisWeekByCat.length-1?"1px solid #F0EDE7":"none" }}>
                    <span style={{ fontSize:13 }}>{CAT_ICONS[c.cat]||"💳"}</span>
                    <span style={{ fontFamily:"Syne",fontSize:12,color:"#44403C",flex:1 }}>{c.cat}</span>
                    <div style={{ width:90,height:4,background:"#EAE7E1",borderRadius:2,overflow:"hidden" }}>
                      <div style={{ height:"100%",background:c.color,borderRadius:2,width:`${Math.min((c.amount/thisWeekTotal)*100,100)}%` }} />
                    </div>
                    <span style={{ fontFamily:"DM Mono",fontSize:12,color:c.color,width:80,textAlign:"right" }}>CHF {c.amount.toFixed(2)}</span>
                  </div>
                ))}
                <div style={{ display:"flex",justifyContent:"space-between",marginTop:10,paddingTop:10,borderTop:"1px solid #EAE7E1" }}>
                  <span style={{ fontFamily:"Syne",fontWeight:700,fontSize:12,color:"#78716C" }}>Total this week</span>
                  <span style={{ fontFamily:"DM Mono",fontSize:13,color:"#1C1917",fontWeight:500 }}>CHF {thisWeekTotal.toFixed(2)}</span>
                </div>
              </div>
            )}
            {thisWeekTx.length===0 && <EmptyState icon="📅" title="No transactions this week" desc="Import your CSV to see weekly spending breakdowns." />}

            {/* ── Monthly chart ── */}
            {monthlyIncome > 0 && (
              <>
                <SL>📊 Income vs Spending — Full Year</SL>
                <div style={{ background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:14,padding:"18px 16px",marginBottom:24,boxShadow:"0 1px 3px rgba(28,25,23,0.06)" }}>
                  <ResponsiveContainer width="100%" height={220}>
                    <LineChart data={monthlyChartData} margin={{ left:0,right:8,top:4,bottom:0 }}>
                      <CartesianGrid strokeDasharray="2 4" stroke="#F0EDE7" />
                      <XAxis dataKey="m" tick={{ fill:"#A8A29E",fontSize:10,fontFamily:"DM Mono" }} axisLine={false} tickLine={false} />
                      <YAxis tick={{ fill:"#A8A29E",fontSize:9,fontFamily:"DM Mono" }} axisLine={false} tickLine={false} />
                      <Tooltip contentStyle={{ background:"#1a1a1a",border:"1px solid #222",borderRadius:8,fontFamily:"DM Mono",fontSize:11 }}
                        formatter={(v,n) => [`CHF ${v.toLocaleString()}`, n]} />
                      <Legend wrapperStyle={{ fontFamily:"DM Mono",fontSize:10,color:"#78716C" }} />
                      <Line type="monotone" dataKey="income" name="Net Income" stroke="#10B981" strokeWidth={2} dot={false} strokeDasharray="6 3" />
                      <Line type="monotone" dataKey="subs"   name="Fixed Costs" stroke="#F59E0B" strokeWidth={2} dot={false} />
                      <Line type="monotone" dataKey="spent"  name="Actual Spent" stroke="#E4003A" strokeWidth={2} dot={{ fill:"#E4003A",r:3 }} />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </>
            )}

            {/* ── Subscriptions & Recurring ── */}
            <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12 }}>
              <SL>🔁 Subscriptions & Recurring Costs</SL>
              <button className="br" onClick={()=>setShowAddSub(true)}
                style={{ background:"#C8102E",border:"none",color:"#fff",fontFamily:"Syne",fontWeight:700,fontSize:11,padding:"6px 13px",borderRadius:7,cursor:"pointer",transition:"background 0.18s",marginBottom:12 }}>
                + Add
              </button>
            </div>

            {/* Annual total banner */}
            <div style={{ background:"#FFFBEB",border:"1px solid #FDE68A",borderRadius:12,padding:"14px 18px",marginBottom:16,display:"flex",justifyContent:"space-between",alignItems:"center" }}>
              <div>
                <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:10,color:"#92400E",letterSpacing:1,textTransform:"uppercase",marginBottom:4 }}>Total Committed / Year</div>
                <div style={{ fontFamily:"Syne",fontWeight:800,fontSize:24,color:"#92400E" }}>CHF {totalAnnualSubs.toLocaleString()}</div>
              </div>
              <div style={{ textAlign:"right" }}>
                <div style={{ fontFamily:"DM Mono",fontSize:11,color:"#A8A29E" }}>per month avg</div>
                <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:16,color:"#92400E" }}>CHF {(totalAnnualSubs/12).toFixed(0)}</div>
              </div>
            </div>

            {/* Sub list */}
            {subAnnual.map((s,i)=>(
              <div key={s.id} className="hr" style={{ display:"flex",alignItems:"center",gap:12,padding:"11px 14px",background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:10,marginBottom:5,transition:"background 0.12s",boxShadow:"0 1px 2px rgba(28,25,23,0.04)" }}>
                <div style={{ width:34,height:34,borderRadius:8,flexShrink:0,background:`${CAT_COLORS[s.cat]||"#4B5563"}18`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:15 }}>{s.icon}</div>
                <div style={{ flex:1,minWidth:0 }}>
                  <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:12,color:"#1C1917" }}>{s.name}</div>
                  <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#A8A29E",marginTop:2 }}>
                    CHF {s.amount} / {s.freq}
                    {s.taxDeductible && <span style={{ color:"#166534",marginLeft:8 }}>✓ tax deductible</span>}
                  </div>
                </div>
                <div style={{ textAlign:"right",flexShrink:0 }}>
                  <div style={{ fontFamily:"DM Mono",fontSize:13,color:"#92400E" }}>CHF {s.annual.toLocaleString()}</div>
                  <div style={{ fontFamily:"DM Mono",fontSize:9,color:"#C4BDB7" }}>/ year</div>
                </div>
                <button onClick={()=>setSubs(prev=>prev.filter(x=>x.id!==s.id))}
                  style={{ background:"none",border:"none",color:"#D9D5CE",cursor:"pointer",fontSize:13,padding:"2px 4px" }}>✕</button>
              </div>
            ))}

            {/* ── Investments ── */}
            <div style={{ marginTop:28 }}>
              <SL>📈 Investment Snapshot</SL>
              <div style={{ display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:12,marginBottom:16 }}>
                {investments.map((inv,i)=>(
                  <div key={inv.id} style={{ background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:12,padding:"16px 18px",boxShadow:"0 1px 3px rgba(28,25,23,0.06)" }}>
                    <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10 }}>
                      <span style={{ fontSize:20 }}>{inv.icon}</span>
                      <span style={{ fontFamily:"DM Mono",fontSize:12,color:inv.gain>=0?"#166534":"#991B1B",fontWeight:500 }}>
                        {inv.gain>=0?"+":""}{inv.pct}%
                      </span>
                    </div>
                    <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:11,color:"#A8A29E",marginBottom:6 }}>{inv.name}</div>
                    <div style={{ fontFamily:"Syne",fontWeight:800,fontSize:19,color:"#1C1917",marginBottom:4 }}>
                      {inv.currency} {inv.value.toLocaleString()}
                    </div>
                    <div style={{ fontFamily:"DM Mono",fontSize:11,color:inv.gain>=0?"#166534":"#991B1B" }}>
                      {inv.gain>=0?"+":""}{inv.currency} {Math.abs(inv.gain)} gain/loss
                    </div>
                    <div style={{ display:"flex",gap:8,marginTop:10 }}>
                      {[["value","Value"],["gain","Gain/Loss"],["pct","% Change"]].map(([field,label])=>(
                        <input key={field} type="number" value={inv[field]} placeholder={label}
                          onChange={e=>setInvestments(prev=>prev.map(x=>x.id===inv.id?{...x,[field]:Number(e.target.value)}:x))}
                          style={{ flex:1,background:"transparent",border:"none",borderBottom:"1px solid #EAE7E1",color:"#78716C",fontFamily:"DM Mono",fontSize:9,padding:"2px 0",minWidth:0 }} />
                      ))}
                    </div>
                  </div>
                ))}
              </div>
              <div style={{ background:"#F0FDF4",border:"1px solid #BBF7D0",borderRadius:10,padding:"12px 16px",display:"flex",justifyContent:"space-between",alignItems:"center" }}>
                <span style={{ fontFamily:"Syne",fontWeight:700,fontSize:12,color:"#78716C" }}>Total Portfolio (approx CHF)</span>
                <span style={{ fontFamily:"DM Mono",fontSize:15,color:"#166534",fontWeight:500 }}>
                  CHF {investments.filter(i=>i.currency==="CHF").reduce((s,i)=>s+i.value,0).toLocaleString()}
                  <span style={{ fontSize:11,color:"#A8A29E",marginLeft:8 }}>+ ₪ {investments.filter(i=>i.currency==="₪").reduce((s,i)=>s+i.value,0).toLocaleString()}</span>
                </span>
              </div>
            </div>
          </>}

          {/* ═══ FAMILY ═══ */}
          {tab==="family" && <>

            {/* Header with child name edit */}
            <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20 }}>
              <div style={{ display:"flex",alignItems:"center",gap:12 }}>
                <div style={{ width:44,height:44,borderRadius:12,background:"#F3E8FF",border:"1px solid #DDD6FE",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22 }}>👶</div>
                <div>
                  <input value={childName} onChange={e=>setChildName(e.target.value)}
                    style={{ fontFamily:"Syne",fontWeight:800,fontSize:20,color:"#1C1917",background:"transparent",border:"none",borderBottom:"1px dashed #D9D5CE",padding:"2px 0",width:220 }} />
                  <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#A8A29E",marginTop:2 }}>click name to edit</div>
                </div>
              </div>
              <button className="br" onClick={()=>setShowAddChild(true)}
                style={{ background:"#C8102E",border:"none",color:"#fff",fontFamily:"Syne",fontWeight:700,fontSize:11,padding:"8px 14px",borderRadius:8,cursor:"pointer",transition:"background 0.18s" }}>
                + Add Expense
              </button>
            </div>

            {/* Summary cards */}
            <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:22 }}>
              {[
                { label:"This Month",    value:`CHF ${childThisMonth.toFixed(2)}`,  icon:"📅", col:"#8B5CF6" },
                { label:"This Year",     value:`CHF ${childTotal.toFixed(2)}`,       icon:"📊", col:"#C8102E" },
                { label:"Entries",       value:String(allChildExpenses.length),      icon:"📝", col:"#166534" },
              ].map((s,i)=>(
                <div key={i} style={{ background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 3px rgba(28,25,23,0.06)" }}>
                  <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start" }}>
                    <div style={{ fontFamily:"DM Mono",fontSize:9,color:"#A8A29E",textTransform:"uppercase",letterSpacing:1,marginBottom:8 }}>{s.label}</div>
                    <span style={{ fontSize:14 }}>{s.icon}</span>
                  </div>
                  <div style={{ fontFamily:"Syne",fontWeight:800,fontSize:17,color:s.col }}>{s.value}</div>
                </div>
              ))}
            </div>

            {/* By subcategory */}
            {childBySubcat.length > 0 && <>
              <SL>By Category</SL>
              <div style={{ background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:12,padding:"10px 16px",marginBottom:22,boxShadow:"0 1px 3px rgba(28,25,23,0.06)" }}>
                {childBySubcat.map(([cat,amt],i)=>{
                  const pct = childTotal>0 ? (amt/childTotal)*100 : 0;
                  const col = CAT_COLORS[cat]||"#8B5CF6";
                  return (
                    <div key={i} style={{ display:"flex",alignItems:"center",gap:12,padding:"8px 0",borderBottom:i<childBySubcat.length-1?"1px solid #F0EDE7":"none" }}>
                      <span style={{ fontSize:14 }}>{CAT_ICONS[cat]||"💳"}</span>
                      <span style={{ fontFamily:"Syne",fontSize:12,color:"#44403C",flex:1 }}>{cat.replace("Child/","")}</span>
                      <div style={{ width:80,height:5,background:"#EAE7E1",borderRadius:3,overflow:"hidden" }}>
                        <div style={{ height:"100%",background:col,borderRadius:3,width:`${Math.min(pct,100)}%` }} />
                      </div>
                      <span style={{ fontFamily:"DM Mono",fontSize:12,color:col,width:80,textAlign:"right" }}>CHF {amt.toFixed(2)}</span>
                    </div>
                  );
                })}
              </div>
            </>}

            {/* Child expense sub-categories quick buttons */}
            <SL>Expense Types</SL>
            <div style={{ display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginBottom:22 }}>
              {[
                { label:"🏫 School Fees",   cat:"Child/School", sub:"tuition, trips, materials" },
                { label:"📚 Books & Supplies",cat:"Child/School",sub:"stationery, folders, bags" },
                { label:"👕 Clothing",       cat:"Child/Other", sub:"seasonal wardrobe" },
                { label:"⚽ Activities",      cat:"Child/Other", sub:"sports, music, clubs" },
                { label:"🩺 Medical",         cat:"Child/Other", sub:"doctor, dentist, meds" },
                { label:"🎁 Gifts",           cat:"Child/Other", sub:"friends' birthdays" },
              ].map((item,i)=>(
                <div key={i} onClick={()=>{ setNewChild(p=>({...p,cat:item.cat,desc:item.label.split(" ").slice(1).join(" ")})); setShowAddChild(true); }}
                  style={{ background:"#FAFAF8",border:"1px solid #EAE7E1",borderRadius:10,padding:"10px 12px",cursor:"pointer",transition:"all 0.15s" }}
                  onMouseEnter={e=>e.currentTarget.style.background="#F0EDE7"}
                  onMouseLeave={e=>e.currentTarget.style.background="#FAFAF8"}>
                  <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:11,color:"#1C1917",marginBottom:3 }}>{item.label}</div>
                  <div style={{ fontFamily:"DM Mono",fontSize:9,color:"#A8A29E" }}>{item.sub}</div>
                </div>
              ))}
            </div>

            {/* Expense list */}
            <SL>All Child Expenses</SL>
            {allChildExpenses.length === 0
              ? <EmptyState icon="👶" title="No child expenses yet" desc={"Click '+ Add Expense' to log school fees, clothing, activities and more.\nImported transactions tagged as Child/School or Child/Other also appear here."} />
              : allChildExpenses.map((e,i)=>(
                <div key={i} className="hr" style={{ display:"flex",alignItems:"center",gap:12,padding:"10px 14px",background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:10,marginBottom:5,transition:"background 0.12s",boxShadow:"0 1px 2px rgba(28,25,23,0.04)" }}>
                  <div style={{ width:34,height:34,borderRadius:8,flexShrink:0,background:`${CAT_COLORS[e.cat]||"#8B5CF6"}18`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14 }}>
                    {CAT_ICONS[e.cat]||"📚"}
                  </div>
                  <div style={{ flex:1,minWidth:0 }}>
                    <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:12,color:"#1C1917",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis" }}>{e.desc}</div>
                    <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#A8A29E",marginTop:2 }}>
                      {e.date}
                      {e.note && <span style={{ marginLeft:8,color:"#78716C" }}>{e.note}</span>}
                      {e.source==="import" && <span style={{ marginLeft:8,color:"#A8A29E",fontSize:9 }}>from CSV</span>}
                    </div>
                  </div>
                  <div style={{ fontFamily:"DM Mono",fontSize:13,color:CAT_COLORS[e.cat]||"#8B5CF6",flexShrink:0,fontWeight:500 }}>
                    CHF {Number(e.amount).toFixed(2)}
                  </div>
                  {e.source!=="import" && (
                    <button onClick={async()=>{ try{ await api.deleteChildExpense(e.id); setChildExp(prev=>prev.filter(x=>x.id!==e.id)); }catch(e2){} }}
                      style={{ background:"none",border:"none",color:"#D9D5CE",cursor:"pointer",fontSize:13,padding:"2px 4px" }}>✕</button>
                  )}
                </div>
              ))
            }

            {/* Annual school calendar */}
            <div style={{ marginTop:24,background:"#EFF6FF",border:"1px solid #BFDBFE",borderRadius:12,padding:18 }}>
              <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:13,color:"#1E40AF",marginBottom:12 }}>📅 Swiss School Year — Upcoming Costs</div>
              {[
                { month:"Aug/Sep", label:"Back to School",       note:"Supplies, new backpack, gym kit",          est:"CHF 150–300" },
                { month:"Oct",     label:"Autumn school trip",   note:"Klassenlager or day excursion",            est:"CHF 50–120" },
                { month:"Dec",     label:"Winter clothing",      note:"Jacket, boots, snowsuit",                  est:"CHF 100–250" },
                { month:"Mar/Apr", label:"Spring school trip",   note:"Schulreise — typically 1–2 days",          est:"CHF 80–150" },
                { month:"Jun/Jul", label:"Summer activities",    note:"Camps, sports, holiday childcare",         est:"CHF 200–800" },
              ].map((item,i)=>(
                <div key={i} style={{ display:"flex",alignItems:"flex-start",gap:12,padding:"7px 0",borderBottom:i<4?"1px solid #DBEAFE":"none" }}>
                  <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#1E40AF",width:60,flexShrink:0,paddingTop:1 }}>{item.month}</div>
                  <div style={{ flex:1 }}>
                    <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:12,color:"#1C1917" }}>{item.label}</div>
                    <div style={{ fontFamily:"DM Mono",fontSize:10,color:"#78716C" }}>{item.note}</div>
                  </div>
                  <div style={{ fontFamily:"DM Mono",fontSize:11,color:"#1E40AF",flexShrink:0 }}>{item.est}</div>
                </div>
              ))}
            </div>
          </>}

        </div>

        <div style={{ padding:"14px 28px",borderTop:"1px solid #EAE7E1",display:"flex",justifyContent:"space-between" }}>
          <span style={{ fontFamily:"DM Mono",fontSize:10,color:"#C4BDB7" }}>MoneyLens · data stored in this artifact only · CH/IL</span>
          <span style={{ fontFamily:"DM Mono",fontSize:10,color:"#C4BDB7" }}>2025–2026</span>
        </div>
      </div>

      {/* ⚡ QUICK LOG floating button */}
      <button onClick={()=>setShowQL(true)} style={{
        position:"fixed", bottom:28, right:24, zIndex:500,
        width:54, height:54, borderRadius:"50%",
        background:"#C8102E", border:"none",
        boxShadow:"0 4px 20px rgba(200,16,46,0.4)",
        cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center",
        fontSize:26, color:"#fff", transition:"all 0.2s",
      }}
        onMouseEnter={e=>{ e.currentTarget.style.transform="scale(1.1)"; e.currentTarget.style.boxShadow="0 6px 28px rgba(200,16,46,0.5)"; }}
        onMouseLeave={e=>{ e.currentTarget.style.transform="scale(1)";   e.currentTarget.style.boxShadow="0 4px 20px rgba(200,16,46,0.4)"; }}>
        +
      </button>

      {/* Quick Log sheet */}
      {showQL && (
        <div onClick={()=>setShowQL(false)} style={{ position:"fixed",inset:0,background:"rgba(28,25,23,0.45)",display:"flex",alignItems:"flex-end",justifyContent:"center",zIndex:1001,padding:"0 0 0 0" }}>
          <div onClick={e=>e.stopPropagation()} style={{
            background:"#FFFFFF", borderRadius:"20px 20px 0 0",
            border:"1px solid #EAE7E1",
            boxShadow:"0 -8px 40px rgba(28,25,23,0.12)",
            padding:"24px 24px 36px", width:"100%", maxWidth:600,
          }}>
            {/* Handle bar */}
            <div style={{ width:40,height:4,background:"#EAE7E1",borderRadius:2,margin:"0 auto 20px" }} />

            <div style={{ fontFamily:"Syne",fontWeight:800,fontSize:16,color:"#1C1917",marginBottom:18 }}>⚡ Quick Log</div>

            {/* Big amount input */}
            <div style={{ display:"flex",alignItems:"center",gap:10,marginBottom:20 }}>
              <span style={{ fontFamily:"Syne",fontWeight:700,fontSize:24,color:"#A8A29E" }}>CHF</span>
              <input type="number" autoFocus placeholder="0.00"
                value={ql.amount} onChange={e=>setQl(p=>({...p,amount:e.target.value}))}
                onKeyDown={e=>{ if(e.key==="Enter") submitQL(); }}
                style={{ flex:1,fontFamily:"Syne",fontWeight:800,fontSize:32,color:"#1C1917",background:"transparent",border:"none",borderBottom:"2px solid #C8102E",padding:"4px 0",textAlign:"right" }} />
            </div>

            {/* Description */}
            <input type="text" placeholder="Description (optional)"
              value={ql.desc} onChange={e=>setQl(p=>({...p,desc:e.target.value}))}
              style={{ width:"100%",fontFamily:"DM Mono",fontSize:13,color:"#44403C",background:"#FAF8F5",border:"1px solid #EAE7E1",borderRadius:8,padding:"9px 12px",marginBottom:14 }} />

            {/* Category picker */}
            <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:10,color:"#A8A29E",letterSpacing:1,textTransform:"uppercase",marginBottom:8 }}>Category</div>
            <div style={{ display:"flex",flexWrap:"wrap",gap:6,marginBottom:16 }}>
              {Object.entries(CAT_ICONS).filter(([k])=>k!=="Insurance").map(([cat,icon])=>(
                <button key={cat} onClick={()=>setQl(p=>({...p,cat}))} style={{
                  background: ql.cat===cat ? CAT_COLORS[cat]||"#4B5563" : "#F5F3EF",
                  border: `1px solid ${ql.cat===cat ? CAT_COLORS[cat]||"#4B5563" : "#EAE7E1"}`,
                  color: ql.cat===cat ? "#fff" : "#44403C",
                  borderRadius:20, padding:"5px 10px", cursor:"pointer",
                  fontFamily:"Syne", fontWeight:600, fontSize:11, transition:"all 0.15s",
                  display:"flex", alignItems:"center", gap:4,
                }}>
                  <span>{icon}</span> {cat.replace("Child/","")}
                </button>
              ))}
            </div>

            {/* Account picker */}
            <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:10,color:"#A8A29E",letterSpacing:1,textTransform:"uppercase",marginBottom:8 }}>Account</div>
            <div style={{ display:"flex",gap:8,marginBottom:22 }}>
              {accounts.map(a=>(
                <button key={a.id} onClick={()=>setQl(p=>({...p,acc:a.id}))} style={{
                  flex:1, background: ql.acc===a.id ? a.color : "#F5F3EF",
                  border:`1px solid ${ql.acc===a.id ? a.color : "#EAE7E1"}`,
                  color: ql.acc===a.id ? "#fff" : "#78716C",
                  borderRadius:8, padding:"7px 6px", cursor:"pointer",
                  fontFamily:"Syne", fontWeight:600, fontSize:10, transition:"all 0.15s",
                }}>
                  {a.icon} {a.name}
                </button>
              ))}
            </div>

            <button onClick={submitQL} disabled={!ql.amount} style={{
              width:"100%", background: ql.amount ? "#C8102E" : "#EAE7E1",
              border:"none", color: ql.amount ? "#fff" : "#A8A29E",
              fontFamily:"Syne", fontWeight:700, fontSize:15,
              padding:"14px", borderRadius:12, cursor: ql.amount ? "pointer" : "default",
              transition:"all 0.18s",
            }}>
              {ql.amount ? `Save CHF ${ql.amount} · ${ql.cat}` : "Enter amount to save"}
            </button>
          </div>
        </div>
      )}

      {/* Add Child Expense Modal */}
      {showAddChild && (
        <div onClick={()=>setShowAddChild(false)} style={{ position:"fixed",inset:0,background:"rgba(28,25,23,0.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1002,padding:20 }}>
          <div onClick={e=>e.stopPropagation()} style={{ background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:18,padding:26,width:"100%",maxWidth:440,boxShadow:"0 20px 60px rgba(28,25,23,0.15)" }}>
            <div style={{ fontFamily:"Syne",fontWeight:800,fontSize:15,color:"#1C1917",marginBottom:18 }}>👶 Add Child Expense</div>

            {[
              ["Description","desc","text","e.g. School trip to Zürich"],
              ["Amount (CHF)","amount","number","0.00"],
              ["Date","date","date",""],
              ["Note (optional)","note","text","e.g. paid by cash"],
            ].map(([label,key,type,ph])=>(
              <div key={key} style={{ marginBottom:14 }}>
                <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:10,color:"#A8A29E",letterSpacing:1,textTransform:"uppercase",marginBottom:6 }}>{label}</div>
                <input type={type} placeholder={ph} value={newChild[key]}
                  onChange={e=>setNewChild(p=>({...p,[key]:e.target.value}))}
                  style={{ width:"100%",background:"#FAF8F5",border:"1px solid #D9D5CE",borderRadius:8,color:"#1C1917",fontFamily:"DM Mono",fontSize:12,padding:"9px 12px" }} />
              </div>
            ))}

            <div style={{ marginBottom:16 }}>
              <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:10,color:"#A8A29E",letterSpacing:1,textTransform:"uppercase",marginBottom:8 }}>Category</div>
              <div style={{ display:"flex",gap:8,flexWrap:"wrap" }}>
                {[
                  { cat:"Child/School", label:"📚 School" },
                  { cat:"Child/Other",  label:"🎒 Other" },
                ].map(opt=>(
                  <button key={opt.cat} onClick={()=>setNewChild(p=>({...p,cat:opt.cat}))} style={{
                    flex:1, background: newChild.cat===opt.cat ? "#8B5CF6" : "#F5F3EF",
                    border:`1px solid ${newChild.cat===opt.cat ? "#8B5CF6" : "#D9D5CE"}`,
                    color: newChild.cat===opt.cat ? "#fff" : "#44403C",
                    borderRadius:8,padding:"8px 12px",cursor:"pointer",
                    fontFamily:"Syne",fontWeight:600,fontSize:12,transition:"all 0.15s",
                  }}>{opt.label}</button>
                ))}
              </div>
            </div>

            <div style={{ display:"flex",gap:10 }}>
              <button className="br" onClick={saveChildEntry}
                style={{ flex:1,background:"#C8102E",border:"none",color:"#fff",fontFamily:"Syne",fontWeight:700,fontSize:13,padding:"11px",borderRadius:9,cursor:"pointer",transition:"background 0.18s" }}>
                Save
              </button>
              <button onClick={()=>setShowAddChild(false)}
                style={{ background:"#F5F3EF",border:"1px solid #D9D5CE",color:"#78716C",borderRadius:9,padding:"11px 16px",fontFamily:"Syne",fontSize:12,cursor:"pointer" }}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Add Subscription Modal */}
      {showAddSub && (
        <div onClick={()=>setShowAddSub(false)} style={{ position:"fixed",inset:0,background:"rgba(28,25,23,0.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,padding:20 }}>
          <div onClick={e=>e.stopPropagation()} style={{ background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:16,padding:26,width:"100%",maxWidth:420,boxShadow:"0 20px 60px rgba(28,25,23,0.15)" }}>
            <div style={{ fontFamily:"Syne",fontWeight:800,fontSize:15,color:"#1C1917",marginBottom:18 }}>Add Recurring Cost</div>
            {[
              ["Name / Description","name","text","e.g. Helsana Premium"],
              ["Amount (CHF)","amount","number","0.00"],
              ["Icon (emoji)","icon","text","🏥"],
            ].map(([label,key,type,ph])=>(
              <div key={key} style={{ marginBottom:14 }}>
                <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:10,color:"#A8A29E",letterSpacing:1,marginBottom:6,textTransform:"uppercase" }}>{label}</div>
                <input type={type} placeholder={ph} value={newSub[key]}
                  onChange={e=>setNewSub(p=>({...p,[key]:e.target.value}))}
                  style={{ width:"100%",background:"#FAF8F5",border:"1px solid #D9D5CE",borderRadius:8,color:"#1C1917",fontFamily:"DM Mono",fontSize:12,padding:"9px 12px" }} />
              </div>
            ))}
            <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:14 }}>
              <div>
                <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:10,color:"#A8A29E",letterSpacing:1,marginBottom:6,textTransform:"uppercase" }}>Frequency</div>
                <select value={newSub.freq} onChange={e=>setNewSub(p=>({...p,freq:e.target.value}))}
                  style={{ width:"100%",background:"#FAF8F5",border:"1px solid #D9D5CE",borderRadius:8,color:"#78716C",fontFamily:"DM Mono",fontSize:12,padding:"9px 10px",cursor:"pointer" }}>
                  {["weekly","monthly","quarterly","yearly"].map(f=><option key={f} value={f}>{f}</option>)}
                </select>
              </div>
              <div>
                <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:10,color:"#A8A29E",letterSpacing:1,marginBottom:6,textTransform:"uppercase" }}>Category</div>
                <select value={newSub.cat} onChange={e=>setNewSub(p=>({...p,cat:e.target.value}))}
                  style={{ width:"100%",background:"#FAF8F5",border:"1px solid #D9D5CE",borderRadius:8,color:"#78716C",fontFamily:"DM Mono",fontSize:12,padding:"9px 10px",cursor:"pointer" }}>
                  {Object.keys(CAT_COLORS).map(c=><option key={c} value={c}>{c}</option>)}
                </select>
              </div>
            </div>
            <label style={{ display:"flex",alignItems:"center",gap:10,marginBottom:18,cursor:"pointer" }}>
              <input type="checkbox" checked={newSub.taxDeductible} onChange={e=>setNewSub(p=>({...p,taxDeductible:e.target.checked}))}
                style={{ width:16,height:16,accentColor:"#10B981" }} />
              <span style={{ fontFamily:"DM Mono",fontSize:12,color:"#44403C" }}>Tax deductible in Switzerland</span>
            </label>
            <div style={{ display:"flex",gap:10 }}>
              <button className="br" onClick={()=>{
                if (!newSub.name||!newSub.amount) return;
                setSubs(prev=>[...prev,{...newSub,id:`s-${Date.now()}`,amount:Number(newSub.amount)}]);
                setNewSub({name:"",amount:"",freq:"monthly",cat:"Other",icon:"💳",taxDeductible:false});
                setShowAddSub(false);
              }} style={{ flex:1,background:"#C8102E",border:"none",color:"#fff",fontFamily:"Syne",fontWeight:700,fontSize:12,padding:"11px",borderRadius:9,cursor:"pointer",transition:"background 0.18s" }}>
                Save
              </button>
              <button className="bg" onClick={()=>setShowAddSub(false)}
                style={{ background:"#F5F3EF",border:"1px solid #D9D5CE",color:"#78716C",borderRadius:9,padding:"11px 16px",fontFamily:"Syne",fontSize:12,cursor:"pointer",transition:"background 0.15s" }}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* CSV Import Modal */}
      {showImport&&(
        <div onClick={()=>{setShowImport(false);setCsvPreview(null);}} style={{ position:"fixed",inset:0,background:"rgba(28,25,23,0.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,padding:20 }}>
          <div onClick={e=>e.stopPropagation()} style={{ background:"#FFFFFF",border:"1px solid #EAE7E1",borderRadius:18,padding:28,width:"100%",maxWidth:540,boxShadow:"0 20px 60px rgba(28,25,23,0.15)" }}>
            <div style={{ fontFamily:"Syne",fontWeight:800,fontSize:16,color:"#1C1917",marginBottom:20 }}>Import Bank Statement (CSV)</div>
            <div style={{ marginBottom:18 }}>
              <div style={{ fontFamily:"Syne",fontWeight:600,fontSize:10,color:"#303030",letterSpacing:1,marginBottom:8 }}>SELECT ACCOUNT</div>
              <div style={{ display:"flex",gap:8,flexWrap:"wrap" }}>
                {accounts.map(a=>(
                  <button key={a.id} onClick={()=>setCsvAccId(a.id)} style={{ flex:"1 1 auto",background:csvAccId===a.id?a.color:"#F5F3EF",border:`1px solid ${csvAccId===a.id?a.color:"#D9D5CE"}`,color:csvAccId===a.id?"#fff":"#44403C",borderRadius:8,padding:"7px 10px",cursor:"pointer",fontFamily:"Syne",fontWeight:600,fontSize:11,transition:"all 0.15s" }}>{a.icon} {a.name}</button>
                ))}
              </div>
            </div>
            {!csvPreview
              ? <UploadZone label="Drop CSV, Excel (.xlsx) or ODS file here" onFile={handleCSVFile} accept=".csv,.txt,.xlsx,.xls,.ods" />
              : <>
                <div style={{ fontFamily:"Syne",fontWeight:700,fontSize:13,color:"#166534",marginBottom:12 }}>✓ {csvPreview.length} transactions detected</div>
                <div style={{ maxHeight:240,overflowY:"auto",background:"#FAF8F5",borderRadius:8,border:"1px solid #EAE7E1" }}>
                  {csvPreview.slice(0,25).map((t,i)=>(
                    <div key={i} style={{ display:"flex",justifyContent:"space-between",padding:"8px 12px",borderBottom:i<24?"1px solid #F0EDE7":"none" }}>
                      <div>
                        <span style={{ fontFamily:"DM Mono",fontSize:11,color:"#1C1917" }}>{t.desc?.slice(0,34)}</span>
                        <span style={{ fontFamily:"DM Mono",fontSize:9,color:"#C4BDB7",marginLeft:8 }}>{t.date}</span>
                        <span style={{ fontFamily:"DM Mono",fontSize:9,color:CAT_COLORS[t.category]||"#444",marginLeft:8 }}>{t.category}</span>
                      </div>
                      <span style={{ fontFamily:"DM Mono",fontSize:11,flexShrink:0,marginLeft:8,color:t.amount<0?"#991B1B":"#166534" }}>
                        {t.amount<0?"-":"+"}CHF {Math.abs(t.amount).toFixed(2)}
                      </span>
                    </div>
                  ))}
                </div>
                {csvPreview.length>25&&<div style={{ fontFamily:"DM Mono",fontSize:10,color:"#252525",textAlign:"center",padding:"8px 0" }}>+{csvPreview.length-25} more</div>}
                <div style={{ display:"flex",gap:10,marginTop:16 }}>
                  <button className="br" onClick={confirmImport}
                    style={{ flex:1,background:"#C8102E",border:"none",color:"#fff",fontFamily:"Syne",fontWeight:700,fontSize:12,padding:"11px",borderRadius:9,cursor:"pointer",transition:"background 0.18s" }}>
                    Import {csvPreview.length} Transactions
                  </button>
                  <button className="bg" onClick={()=>setCsvPreview(null)}
                    style={{ background:"#F5F3EF",border:"1px solid #D9D5CE",color:"#78716C",borderRadius:9,padding:"11px 16px",fontFamily:"Syne",fontSize:12,cursor:"pointer",transition:"background 0.15s" }}>
                    Back
                  </button>
                </div>
              </>
            }
            <div style={{ marginTop:14,fontFamily:"DM Mono",fontSize:10,color:"#A8A29E",lineHeight:1.9 }}>
              <b style={{color:"#78716C"}}>Neon:</b> App → Profile → Export transactions (CSV)<br/>
              <b style={{color:"#78716C"}}>SBB:</b> my.sbb.ch → My account → Billing → Export (CSV)<br/>
              <b style={{color:"#78716C"}}>Max:</b> Max app → Account → Statements → Export (Excel/CSV)<br/>
              <b style={{color:"#78716C"}}>Any bank:</b> Excel (.xlsx), CSV, or ODS files all work
            </div>
          </div>
        </div>
      )}
    </>
  );
}

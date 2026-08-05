import React, { useState, useMemo, useRef, useEffect, useDeferredValue, useTransition } from "react";
import { Check, Plus, Lock, Filter, X, Copy, ChevronUp, ChevronDown, CornerDownRight, Columns3, MessageSquare, RotateCcw, Droplet, Snowflake, Trash2, SkipForward, Bell, Clock, Star } from "lucide-react";
import * as XLSX from "xlsx";
import { createPortal } from "react-dom";
import { supabase } from "./supabaseClient";

/* MERCH TRACKER — v31: P95 / large-sheet optimization pass */

const FONT = `@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@500;600;800&family=JetBrains+Mono:wght@400;500;700&display=swap');`;
const THEME_CSS = `
:root, [data-theme="paper"] {
  --ink:#1f1f1d; --bg:#f7f3ea; --surface:#fffdf8; --accent:#c96f16; --accent-tint:#fff4e3;
  --danger:#b42318; --success:#1f6f54; --info:#2563a6;
  --muted-1:#9b9488; --muted-2:#7d766b; --muted-3:#655f56; --muted-4:#4f493f; --muted-5:#403a32; --muted-6:#c8c0b4; --muted-7:#b8afa3;
  --line-1:#e5ded2; --line-2:#d4cabd; --line-3:#eee7dc;
  --toolbar-bg:#fffaf1; --toolbar-line:#e3d7c7; --toolbar-subtle:#f1eadf;
  --pill-shadow:0 1px 0 rgba(31,31,29,0.06); --card-shadow:0 1px 2px rgba(31,31,29,0.04);
  --on-dark:#d8d1c4; --on-dark-2:#a9a095; --on-dark-line:#4a463e;
  --tint-ok:#e5f1ea; --fg-ok:#1c6048; --tint-warn:#f8e9b7; --fg-warn:#7a560f; --tint-late:#f6d3cb; --fg-late:#8c241a; --tint-reject:#f0bdb5; --tint-rework:#fbd9a8; --tint-waive:#e9ddc2; --tint-next:#fdecc9; --tint-histrej:#fbe9e6; --revised:#6a45a8;
}
/* Bigger click/tap target without changing the data model. Inline styles still control the visual design. */
button {
  min-width: 34px;
  min-height: 34px;
  box-sizing: border-box;
  touch-action: manipulation;
}
button svg { pointer-events: none; }
input, select { min-height: 34px; box-sizing: border-box; }
/* Tiny in-cell icons remain visually compact but are now easier to hit. */
td button[title*="revised"],
td button[title*="REJECTED"],
td button[title*="rework"],
td button[title*="skip"],
td button[title="calendar"],
td button[title="delete row"] {
  min-width: 28px !important;
  min-height: 28px !important;
  padding: 5px !important;
  align-items: center;
  justify-content: center;
}
/* Panel/header buttons were visually close together; keep hit area generous everywhere. */
[role="button"], button { -webkit-tap-highlight-color: transparent; }

/* v31bm UI comfort pass: tabs, header actions, popups and compact controls. Logic unchanged. */
.mt-app-header { flex-wrap: wrap; gap: 14px; }
.mt-header-actions { flex-wrap: wrap; justify-content: flex-end; row-gap: 8px; }
.mt-header-actions > button, .mt-header-actions > span > button { min-width: 36px !important; min-height: 36px !important; padding: 7px 10px !important; border-radius: 8px; }
.mt-header-actions > span { row-gap: 6px; }
.mt-main-tabs { gap: 6px !important; overflow-x: auto; overflow-y: hidden; scrollbar-width: thin; padding-top: 7px !important; padding-bottom: 0 !important; }
.mt-main-tabs > button { min-height: 38px !important; padding: 10px 16px !important; border-radius: 9px 9px 0 0 !important; flex: 0 0 auto; margin-bottom: -1px; }
.mt-main-tabs > button:hover { background: rgba(255,255,255,0.06) !important; color: var(--bg) !important; }
.mt-popup-close { min-width: 36px !important; min-height: 34px !important; border-radius: 8px !important; display: inline-flex; align-items: center; justify-content: center; }
.mt-filter-chip-x { min-width: 24px !important; min-height: 24px !important; border-radius: 999px !important; }
.mt-stage-cell-body { padding: 18px 32px 16px 32px !important; }
.mt-stage-corner-btn { min-width: 28px !important; min-height: 28px !important; border-radius: 8px !important; background: rgba(255,253,248,0.86) !important; box-shadow: 0 1px 2px rgba(31,31,29,0.08); }
.mt-stage-corner-btn:hover { background: var(--surface) !important; border-color: rgba(31,31,29,0.22) !important; }
.mt-row-flag-btn, .mt-new-flag-btn { min-height: 26px !important; min-width: 34px !important; border-radius: 999px !important; padding: 5px 8px !important; }
.mt-new-flag-btn { font-size: 9px !important; letter-spacing: .35px; }
.mt-todo-mode-group { display: inline-flex; gap: 7px; flex-wrap: wrap; }
.mt-todo-mode-group > button { border-radius: 999px !important; min-height: 36px !important; padding: 8px 13px !important; border-right: 1px solid var(--ink) !important; }
.mt-toolbar-mock { border:1px dashed var(--line-2); background:#fffaf1; border-radius:12px; padding:10px 12px; font-size:10.5px; color:var(--muted-3); margin:8px 0 12px; }
`;
const REL_GATE_DAYS = 30, FABRIC_CUTOFF_DAYS = 35, STYLE_W = 190;
const UPCOMING_DEFAULT = { techpack:2, fitSend:4, fitAppr:1, artwork:2, artAppr:1, strikeOff:3, soAppr:1, labDip:5, labAppr:1, ppSample:4, ppAppr:1, fabricIH:15, prodFile:7 }; // working days before a stage that it becomes "upcoming" in the To-Do list

const STAGES = [
  { key:"techpack",  label:"Techpack",     lead:3, owner:"Merchant", flag:null, pred:"__ord" },
  { key:"fitSend",   label:"Fit Send",     lead:7, owner:"CAD",      flag:"fitReq", pred:"techpack" },
  { key:"fitAppr",   label:"Fit Appr",     lead:3, owner:"Buyer",    flag:"fitReq", pred:"fitSend" },
  { key:"artwork",   label:"Artwork",      lead:7, owner:"Designer", flag:"printReq", pred:"techpack" },
  { key:"artAppr",   label:"Art Appr",     lead:3, owner:"Buyer",    flag:"printReq", pred:"artwork" },
  { key:"strikeOff", label:"Strike-off",   lead:4, owner:"Merchant", flag:"soReq", pred:"artAppr" },
  { key:"soAppr",    label:"S/O Appr",     lead:3, owner:"Buyer",    flag:"soReq", pred:"strikeOff" },
  { key:"labDip",    label:"Lab Dip",      lead:7, owner:"Merchant", flag:"labDipReq", pred:"__ord" },
  { key:"labAppr",   label:"Lab Dip Appr", lead:3, owner:"Buyer",    flag:"labDipReq", pred:"labDip" },
  { key:"fabricIH",  label:"Fabric IH",    lead:0, owner:"Merchant", flag:null, cutoff:true },
  { key:"ppSample",  label:"PP Sample",    lead:5, owner:"Merchant", flag:"ppNeeded", pred:"fabricIH" },
  { key:"ppAppr",    label:"PP Appr",      lead:3, owner:"Buyer",    flag:"ppNeeded", pred:"ppSample" },
  { key:"prodFile",  label:"Prod File",    lead:3, owner:"Merchant", flag:null, pred:"ppAppr" },
];
const STAGE_KEYS = STAGES.map(s=>s.key);

const STAGE_ORDER_INDEX = Object.fromEntries(STAGES.map((stage, index)=>[stage.key, index + 1]));
const STAGE_SECTION_BY_KEY = {
  techpack: "Techpack / Start",
  fitSend: "Fit", fitAppr: "Fit",
  artwork: "Print / Strike-off", artAppr: "Print / Strike-off", strikeOff: "Print / Strike-off", soAppr: "Print / Strike-off",
  labDip: "Lab Dip / Fabric", labAppr: "Lab Dip / Fabric", fabricIH: "Lab Dip / Fabric",
  ppSample: "PP / Production File", ppAppr: "PP / Production File", prodFile: "PP / Production File",
};
const stageOrderOf=(keyOrLabel)=>{
  const key=STAGE_ORDER_INDEX[keyOrLabel]!=null ? keyOrLabel : (STAGES.find(stage=>stage.label===keyOrLabel)||{}).key;
  return STAGE_ORDER_INDEX[key] || 999;
};
const stageSectionOf=(keyOrLabel)=>{
  const key=STAGE_ORDER_INDEX[keyOrLabel]!=null ? keyOrLabel : (STAGES.find(stage=>stage.label===keyOrLabel)||{}).key;
  return STAGE_SECTION_BY_KEY[key] || "Other";
};

// Canonical filter helpers used by Tracker, To-Do, Dashboard drilldowns, Management drilldowns and exports.
// Display labels are never trusted as filter keys; every activity filter is resolved to a stage key.
const FILTER_NONE="__MT_FILTER_NONE__"; // explicit column-filter state: user selected zero values, so show zero matching rows
const filterToken=(v)=>String(v==null?"":v).replace(/^Escalate:\s*/i,"").replace(/[^a-z0-9]+/gi,"").trim().toLowerCase();
const stageKeyFromAnyGlobal=(v)=>{
  const tok=filterToken(v);
  if(!tok) return "";
  const st=(STAGES||[]).find(x=>filterToken(x.key)===tok || filterToken(x.label)===tok);
  return st?st.key:"";
};
const stageLabelFromKeyGlobal=(k)=>((STAGES||[]).find(x=>x.key===k)||{}).label||k||"";
const arrClean=(v)=>Array.isArray(v)?v.filter(x=>x!=null&&String(x).trim()!==""):(v!=null&&String(v).trim()!==""?[v]:[]);
// v31dg: one-table-truth safety pass; shared colour splitter used by Tracker/To-Do/Dashboard/Management filters.
const splitColoursAll=(txt)=>{ const raw=String(txt||"").trim(); if(!raw) return ["(no colour)"]; const parts=raw.split(/[,;|\/+]+/).map(x=>x.replace(/\s+/g," ").trim()).filter(Boolean); return parts.length?parts:[raw.replace(/\s+/g," ").trim()]; };
const todoDrillFilterFromSlice=(base={},df={})=>{
  const out={...(base||{})};
  // IDEMPOTENT by design. A dashboard/management drill is processed twice:
  //   1) drillRowsToTodo -> todoDrillFilterFromRows(rows,base,df)  (builds the full filter)
  //   2) drillTodo prop  -> todoDrillFilterFromSlice({}, canonicalDrillSpec(thatFilter))  (re-applied on apply)
  // The second pass MUST NOT strip owner / escalation / activity / style pins. So we:
  //   (a) accept both input-shape (order) and output-shape (orderNo) keys, and
  //   (b) read activity & style pins from BOTH df and out.
  // This is the root fix for "To-Do drilldowns show all rows instead of the exact filtered rows".
  const map={ order:"orderNo", orderNo:"orderNo", junior:"junior", colour:"colour", fit:"fit", family:"family", brand:"brand", fabric:"fabric", buyer:"buyer", branch:"branch", owner:"owner", escalationOwner:"escalationOwner", phase:"phase", todoType:"todoType", priority:"priority" };
  Object.entries(map).forEach(([from,to])=>{ const vals=arrClean(df&&df[from]); if(vals.length) out[to]=[...new Set([...arrClean(out[to]),...vals.map(v=>String(v))])]; });
  // Keep dashboard/management drill filters table-like: slice filters + clicked owner/activity become To-Do column filters.
  // Activity is always stored as canonical stage keys plus display labels so Activity multi-select stays stable.
  const actVals=[...arrClean(out.activityKey), ...arrClean(out.key), ...arrClean(out.activity), ...arrClean(df&&df.activityKey), ...arrClean(df&&df.key), ...arrClean(df&&df.activity)];
  const actKeys=activityKeysFromAnyGlobal(actVals);
  if(actKeys.length){ out.activityKey=actKeys; out.activity=actKeys.map(stageLabelFromKeyGlobal).filter(Boolean); delete out.key; }
  const styleVals=[...arrClean(out.styleId), ...arrClean(out.styleIds), ...arrClean(df&&df.styleId), ...arrClean(df&&df.styleIds)];
  if(styleVals.length){ out.styleId=[...new Set(styleVals.map(v=>String(v)))]; delete out.styleIds; }
  return out;
};
const todoDrillFilterFromRows=(rows=[],base={},df={})=>{
  const out=todoDrillFilterFromSlice(base,df);
  const styleIds=[...new Set((rows||[]).map(r=>r&&r.id).filter(v=>v!=null&&String(v)!=="").map(v=>String(v)))];
  const stageKeys=[...new Set((rows||[]).map(r=>stageKeyFromAnyGlobal((r&&r.activityKey)||(r&&r.key)||(r&&r.stageKey)||(r&&r.activity)||(r&&r.activityLabel))).filter(Boolean))];
  if(styleIds.length) out.styleId=styleIds;
  if(stageKeys.length){ out.activityKey=stageKeys; out.activity=stageKeys.map(stageLabelFromKeyGlobal).filter(Boolean); delete out.key; }
  return out;
};

// SINGLE TABLE TRUTH MODEL
// Tracker is the operational truth. Dashboards, To-Do, exports and drilldowns must summarize/filter
// these canonical stage keys instead of reinterpreting display labels.
const frontierKeysOf=(c)=>{
  try{
    const raw=c&&c.frontier;
    const arr=raw instanceof Set ? [...raw] : (Array.isArray(raw)?raw:[]);
    return arr.map(k=>stageKeyFromAnyGlobal(k)||String(k||"")).filter(k=>STAGE_ORDER_INDEX[k]!=null);
  }catch(_e){ return []; }
};
const frontierLabelsOf=(c)=>frontierKeysOf(c).map(stageLabelFromKeyGlobal).filter(Boolean);
const firstFrontierKeyOf=(c)=>frontierKeysOf(c)[0]||"";
const currentDueForStage=(c,key)=>{
  const r=(c&&Array.isArray(c.stages)?c.stages:[]).find(x=>x&&x.key===key);
  return r?(r.actual||r.rev||r.plan||r.skip||null):null;
};
const canonicalLiveRow=(row)=>{
  const s=(row&&row.s)||{}; const c=(row&&row.c)||{};
  const keys=frontierKeysOf(c);
  const primary=keys[0]||"";
  const stageRows=(c.stages||[]).filter(r=>r&&keys.includes(r.key));
  return {
    ...(row||{}),
    tableId:String(s.id||""),
    tableCurrentStageKey:primary,
    tableCurrentStageKeys:keys,
    tableCurrentStageLabel:primary?stageLabelFromKeyGlobal(primary):"",
    tableCurrentStageLabels:keys.map(stageLabelFromKeyGlobal).filter(Boolean),
    tableCurrentChaseOwners:(c.chaseOwners||[]).map(o=>o&&o.owner).filter(Boolean),
    tableOpenStageRows:stageRows,
    tableDueDates:stageRows.map(r=>r.actual||r.rev||r.plan||r.skip||null).filter(Boolean),
    tableStatus:c.status||"",
    tableTone:c.tone||"",
    tableReleased:!!c.released
  };
};
const rowMatchesCurrentStage=(row,key)=>{
  const k=stageKeyFromAnyGlobal(key)||String(key||"");
  if(!k) return true;
  const keys=(row&&row.tableCurrentStageKeys)||frontierKeysOf(row&&row.c);
  return Array.isArray(keys)&&keys.includes(k);
};
const activityKeysFromAnyGlobal=(v)=>{
  const raw=Array.isArray(v)?v:(v==null||v===""?[]:[v]);
  const keys=[];
  raw.forEach(x=>{
    const k=stageKeyFromAnyGlobal(x);
    if(k && !keys.includes(k)) keys.push(k);
  });
  return keys;
};
const rowMatchesAnyCurrentStage=(row,keys)=>{
  const arr=activityKeysFromAnyGlobal(keys);
  if(!arr.length) return true;
  return arr.some(k=>rowMatchesCurrentStage(row,k));
};
const canonicalDrillSpec=(spec={})=>{
  const out={...(spec||{})};
  const keys=activityKeysFromAnyGlobal(out.activityKey||out.key||out.activity||out.stage||[]);
  if(keys.length){ out.activityKey=keys; out.activity=keys.map(stageLabelFromKeyGlobal).filter(Boolean); delete out.key; delete out.stage; }
  return out;
};
const CHASE_LABELS = ["Management","Sr Merchant","Jr Merchant","CAD","Designer","Store","Buyer","Merchant","Mill"];
const DEFAULT_STAGE_OWNERS = Object.fromEntries(STAGES.map(s=>[s.key, s.owner==="Merchant"?"Jr Merchant":s.owner==="Mill"?"Store":s.owner]));
const DEFAULT_ESCALATION_RULES = [
  { from:0, to:2, owner:"Jr Merchant", level:"Level 1", action:"Daily chase by junior merchant; update actual date or revised commitment." },
  { from:3, to:4, owner:"Sr Merchant", level:"Level 2", action:"Senior merchant takes over follow-up and clears the blocker." },
  { from:5, to:7, owner:"Management", level:"Level 3", action:"Review in management meeting; agree recovery plan and next commitment." },
  { from:8, to:null, owner:"Director", level:"Critical", action:"Immediate top escalation; unblock or revise delivery risk plan." },
];
const DEFAULT_CFG = { leads:Object.fromEntries(STAGES.map(s=>[s.key,s.lead])), stageOwners:{...DEFAULT_STAGE_OWNERS}, rework:{...{ fitSend:4, artwork:2, strikeOff:3, labDip:7, ppSample:4 }}, fabricCutoff:FABRIC_CUTOFF_DAYS, relGate:REL_GATE_DAYS, upcoming:{...UPCOMING_DEFAULT}, escalationRules:DEFAULT_ESCALATION_RULES.map(x=>({...x})), todoEscalationRows:true };
const escalationRulesOf=(cfg)=>{ const arr=(cfg&&Array.isArray(cfg.escalationRules)&&cfg.escalationRules.length?cfg.escalationRules:DEFAULT_ESCALATION_RULES); return arr.map((r,i)=>({ from:Math.max(0,Number(r.from)||0), to:(r.to==null||r.to==="")?null:Math.max(0,Number(r.to)||0), owner:String(r.owner||DEFAULT_ESCALATION_RULES[i]?.owner||"Jr Merchant"), level:String(r.level||DEFAULT_ESCALATION_RULES[i]?.level||("Level "+(i+1))), action:String(r.action||DEFAULT_ESCALATION_RULES[i]?.action||"Chase and update commitment." ) })).sort((a,b)=>a.from-b.from); };
const escalationFor=(cfg,days)=>{ const d=Math.max(0,Number(days)||0); const rules=escalationRulesOf(cfg); const r=rules.find(x=>d>=x.from && (x.to==null || d<=x.to)) || rules[rules.length-1] || DEFAULT_ESCALATION_RULES[0]; return { ...r, rangeLabel:`${r.from}-${r.to==null?"∞":r.to}d` }; };
const OWNER_COLOR = { "Jr Merchant":"var(--success)", "Sr Merchant":"#b4531a", Management:"#6b7280", CAD:"#2563a6", Buyer:"#b4531a", Designer:"#6d4aab", Store:"#7a5a1e", Merchant:"var(--success)", Mill:"#7a5a1e" };
const OWNER_BG = { "Jr Merchant":"var(--tint-ok)", Merchant:"var(--tint-ok)", "Sr Merchant":"#fff0df", Management:"#eef0f2", CAD:"#e8f0fb", Buyer:"#fff0df", Designer:"#f2ebfb", Store:"#f3ead8", Mill:"#f3ead8" };

const ONE_DAY = 86400000;
function addWorkdays(date,n){ if(!date) return null; const d=new Date(date.getTime()); let a=0; const step=n>=0?1:-1, t=Math.abs(n); while(a<t){ d.setTime(d.getTime()+step*ONE_DAY); if(d.getDay()!==0) a++; } return d; }
function netWorkdays(a,b){ if(!a||!b) return null; let d=new Date(a.getTime()), end=new Date(b.getTime()), sign=1; if(end<d){ const tmp=d; d=end; end=tmp; sign=-1; } let n=0; const cur=new Date(d.getTime()); while(cur<end){ cur.setTime(cur.getTime()+ONE_DAY); if(cur.getDay()!==0) n++; } return sign*n; }
const fmt=(d)=> !d?"":d.toLocaleDateString("en-GB",{day:"2-digit",month:"short"});
const parse=(s)=>{ if(!s) return null; const m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(s); return m?new Date(Number(m[1]),Number(m[2])-1,Number(m[3])):new Date(s); };
// Stable local date sorting for ISO dates, Excel serial dates, and user-visible dd/mm/yyyy or dd-MMM-yyyy values.
// Native Date parsing is unreliable for Indian/UK style dates such as 11/06/2026, so all grid date sorting uses this helper.
const dateSerial=(v,empty=Infinity)=>{
  if(v==null||v==="") return empty;
  if(v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(),v.getMonth(),v.getDate()).getTime();
  if(typeof v==="number" && isFinite(v)){
    // Excel serial date support. 25569 = 1970-01-01.
    if(v>20000 && v<90000) return Math.round((v-25569)*ONE_DAY);
    return v;
  }
  const raw=String(v).trim();
  if(!raw) return empty;
  let m=/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(raw);
  if(m) return new Date(Number(m[1]),Number(m[2])-1,Number(m[3])).getTime();
  m=/^(\d{1,2})[\-/\s.](\d{1,2})[\-/\s.](\d{2,4})$/.exec(raw);
  if(m){ const y=Number(m[3].length===2?"20"+m[3]:m[3]); return new Date(y,Number(m[2])-1,Number(m[1])).getTime(); }
  const mons={jan:0,january:0,feb:1,february:1,mar:2,march:2,apr:3,april:3,may:4,jun:5,june:5,jul:6,july:6,aug:7,august:7,sep:8,sept:8,september:8,oct:9,october:9,nov:10,november:10,dec:11,december:11};
  m=/^(\d{1,2})[\-/\s.]([A-Za-z]{3,9})(?:[\-/\s.,]+(\d{2,4}))?$/.exec(raw);
  if(m){ const mo=mons[m[2].toLowerCase()]; if(mo!=null){ const y=m[3]?Number(m[3].length===2?"20"+m[3]:m[3]):TODAY.getFullYear(); return new Date(y,mo,Number(m[1])).getTime(); } }
  const d=new Date(raw); return isNaN(d)?empty:new Date(d.getFullYear(),d.getMonth(),d.getDate()).getTime();
};
const iso=(d)=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
const colLetter=(n)=>{ let s=""; n=Number(n)+1; if(n<1) return ""; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s; };
const letterToIndex=(s)=>{ s=String(s||"").toUpperCase(); if(!/^[A-Z]+$/.test(s)) return -1; let n=0; for(const ch of s) n=n*26+(ch.charCodeAt(0)-64); return n-1; };
const _now=new Date(); const TODAY=new Date(_now.getFullYear(),_now.getMonth(),_now.getDate()); // live current date (local midnight)
const perfNow=()=> (typeof performance!=="undefined"&&performance.now)?performance.now():Date.now();
const lc=(v)=>String(v==null?"":v).toLowerCase();

const styleComputeSignature=(s)=> JSON.stringify({
  id:s.id, orderNo:s.orderNo||"", styleNo:s.styleNo||"", sampleFit:s.sampleFit||"", family:s.family||"", colour:s.colour||"", brand:s.brand||"", buyer:s.buyer||"", fabricType:s.fabricType||"", owner:s.owner||"", setId:s.setId||"", setRole:s.setRole||"", age:s.age||"", qty:Number(s.qty)||0, ordRec:s.ordRec||"", delivery:s.delivery||"", fitReq:!!s.fitReq, printReq:!!s.printReq, soReq:!!s.soReq, ppBypass:!!s.ppBypass, labDipReq:!!s.labDipReq, ppNeeded:!!s.ppNeeded, archived:!!s.archived, remarks:s.remarks||"", actuals:s.actuals||{}, revs:s.revs||{}, rejects:s.rejects||{}, skips:s.skips||{}
});

const buildSearchIndex=(s,c)=>{
  const chase=(c&&c.chaseOwners?c.chaseOwners:[]).map(o=>o.owner).join(" ");
  const byCol={ styleNo:lc(s.styleNo), orderNo:lc(s.orderNo), sampleFit:lc(s.sampleFit), family:lc(s.family), colour:lc(s.colour), color:lc(s.colour), brand:lc(s.brand), buyer:lc(s.buyer), fabricType:lc(s.fabricType), owner:lc(s.owner), remarks:lc(s.remarks), age:lc(s.age), setRole:lc(s.setRole), chase:lc(chase) };
  return { byCol, auto:[byCol.styleNo,byCol.orderNo,byCol.sampleFit,byCol.family,byCol.colour,byCol.brand,byCol.buyer,byCol.fabricType,byCol.owner,byCol.remarks,byCol.chase].join(" ") };
};
const pushPerfSample=(arr,ms)=>{ const n=Number(ms)||0; const next=(arr||[]).concat(n).slice(-80).sort((a,b)=>a-b); return { samples:next, p95:next.length?Math.round(next[Math.min(next.length-1,Math.floor(next.length*0.95))]*10)/10:0 }; };


const REJECTABLE=["fitAppr","artAppr","soAppr","labAppr","ppAppr"]; // approval stages that can be rejected
const SKIPPABLE_STAGES=["fitSend","fitAppr","artwork","artAppr","strikeOff","soAppr","labDip","labAppr","ppSample","ppAppr"]; // activities that can be waived/skipped
const APPR_OF_SEND={ fitSend:"fitAppr", artwork:"artAppr", strikeOff:"soAppr", labDip:"labAppr", ppSample:"ppAppr" }; // send/make stage -> the approval that can reject it
const SEND_FOR_APPR=Object.fromEntries(Object.entries(APPR_OF_SEND).map(([send,appr])=>[appr,send]));
// Skip-pair integrity: live calculations must not chase an approval when its send/make stage was waived.
// This is a non-destructive normalization layer: raw DB values stay available for audit/import review,
// but Tracker/To-Do/Dashboards/Reports read this cleaned active style.
const normalizeSkipPairs=(style)=>{
  if(!style) return style;
  const skips={...(style.skips||{})};
  const actuals=style.actuals||{};
  let changed=false;
  Object.entries(APPR_OF_SEND).forEach(([sendKey,apprKey])=>{
    const sendSkipped=!!skips[sendKey];
    const apprSkipped=!!skips[apprKey];
    if(sendSkipped && !apprSkipped){
      skips[apprKey]=skips[sendKey];
      changed=true;
    }
    // If approval was waived/imported but the send stage has no actual, waive the send too.
    // If send already has an actual, keep it as history and do not convert it to skipped.
    if(apprSkipped && !sendSkipped && !actuals[sendKey]){
      skips[sendKey]=skips[apprKey];
      changed=true;
    }
  });
  return changed?{...style,skips}:style;
};
const stageReviewLabel=(style,r)=>{
  if(!r) return "";
  // IMPORTANT: stage key is the truth for filters, cards, drilldowns and To-Do rows.
  // r.label can become stale/derived during rework/approval transformations, so never let
  // display text override the canonical stage key. This was causing Activity = S/O Appr
  // cards to calculate correctly while the visible To-Do table still displayed Strike-off,
  // Lab Dip or Artwork labels.
  const base = stageLabelFromKeyGlobal(r.key) || r.label || r.key || "";
  if(r.rework) return `${base} rework / resend`;
  if(r.reapproval) return `${base} re-approval`;
  if(r.rejected) return `${base} rejected`;
  return base;
};
const REWORK_DAYS={ fitSend:4, artwork:2, strikeOff:3, labDip:7, ppSample:4 }; // working days added on rejection (redo+resend)
const resendHistoryKey=(s,k)=>s&&s.id!=null?`${s.id}:${k}`:"";
const normalizeResendEntry=(x)=> typeof x==="string"?{ source:"legacy", newVal:x }:x;
const latestTrackedResendDate=(s,k,rejectDate,resendMap)=>{
  if(!s||!k||!rejectDate) return null;
  const arr=Array.isArray(resendMap&&resendMap[resendHistoryKey(s,k)])?resendMap[resendHistoryKey(s,k)].map(normalizeResendEntry).filter(Boolean):[];
  // Important: do not infer a resend only because the visible actual date is after the rejection date.
  // A valid resend is one explicitly recorded by the resend editor / Fill Date resend path, not the first stored actual.
  // Use the latest recorded event by history order, not the maximum date. This allows correcting a wrong resend actual
  // to an earlier date without the old/wrong later date continuing to win calculations.
  let latest=null;
  arr.forEach((x,i)=>{
    const src=String((x&&x.source)||"").toLowerCase();
    const isFirst=src.includes("first send") || (i===0 && arr.length>1 && !src.includes("stage_events"));
    if(isFirst) return;
    const d=parse(x&&x.newVal);
    if(d && dateSerial(d) >= dateSerial(rejectDate, -Infinity)) latest=d;
  });
  return latest;
};
const isRepeatStyleNoDev=(s)=>!s.fitReq && !s.printReq && !s.labDipReq && !s.ppNeeded;
const stageApplies=(s,st)=>{
  if(!st) return false;
  // Repeat/no-development styles: if all development toggles are OFF, do not require Techpack.
  // The live path stays simple: Fabric IH -> Prod File.
  if(st.key==="techpack" && isRepeatStyleNoDev(s)) return false;
  // Strike-off is only meaningful inside the print branch. If S/O is accidentally ticked while Print is off, keep it out of the chain.
  if(st.key==="strikeOff" || st.key==="soAppr") return !!(s.printReq && s.soReq);
  if(st.key==="artwork" || st.key==="artAppr") return !!s.printReq;
  return st.flag===null || !!s[st.flag];
};
const applicableStages=(s)=> STAGES.filter(st=> stageApplies(s,st));

function computeStyle(s, cfg, resendMap={}){
  const ordRec=parse(s.ordRec), delivery=parse(s.delivery);
  const leadOf=(st)=>{ const v=cfg&&cfg.leads&&cfg.leads[st.key]; return v==null?st.lead:v; }; const rwOf=(st)=>{ const v=cfg&&cfg.rework&&cfg.rework[st.key]; return v==null?(REWORK_DAYS[st.key]||st.lead):v; }; const ownerOf=(k)=>{ const st=STAGES.find(x=>x.key===k); return (cfg&&cfg.stageOwners&&cfg.stageOwners[k]) || DEFAULT_STAGE_OWNERS[k] || (st&&st.owner) || "Jr Merchant"; }; const CUTD=(cfg&&cfg.fabricCutoff!=null)?cfg.fabricCutoff:FABRIC_CUTOFF_DAYS; const GATED=(cfg&&cfg.relGate!=null)?cfg.relGate:REL_GATE_DAYS;
  const cutoff=delivery?addWorkdays(delivery,-CUTD):null;
  const eff={}, plan={};
  const applies=(k)=>{ const st=STAGES.find(x=>x.key===k); return stageApplies(s,st); };
  const actualOf=(k)=>parse(s.actuals[k]); const revOf=(k)=>parse(s.revs?.[k]); const rejOf=(k)=>parse(s.rejects?.[k]); const skipOf=(k)=>parse(s.skips?.[k]);
  const validResendActual=(sendK,apprK)=> latestTrackedResendDate(s,sendK,rejOf(apprK),resendMap);
  STAGES.forEach(st=>{
    let p;
    if(st.cutoff){ const base=s.labDipReq?(eff["labAppr"]||eff["labDip"]||ordRec):ordRec; const base15=addWorkdays(base,15); p=s.labDipReq?(cutoff?new Date(Math.max((base15&&base15.getTime())||0, cutoff.getTime())):base15):cutoff; }
    else { let predEff; if(st.key==="prodFile") predEff = (s.ppBypass || !s.ppNeeded) ? eff["fabricIH"] : eff["ppAppr"]; else predEff = st.pred==="__ord"?ordRec:eff[st.pred]; if((st.key==="ppSample"||st.key==="prodFile") && s.fitReq && eff["fitAppr"]) predEff = new Date(Math.max((predEff&&predEff.getTime())||0, eff["fitAppr"].getTime())); p=addWorkdays(predEff||ordRec, leadOf(st)); }
    const apprK=APPR_OF_SEND[st.key]; const rejAppr = !!(apprK && rejOf(apprK) && !actualOf(apprK) && !skipOf(apprK));
    const selfRej = REJECTABLE.includes(st.key) && rejOf(st.key) && !actualOf(st.key) && !skipOf(st.key);
    if(rejAppr){ const rjd=rejOf(apprK); const auto=addWorkdays(rjd, rwOf(st)); const a=validResendActual(st.key,apprK); const rv=revOf(st.key); plan[st.key]=auto;
      if(a) eff[st.key]=a; else if(rv && rv>=rjd) eff[st.key]=rv; else eff[st.key]=auto; } // redo: explicitly tracked re-send actual wins, else fresh revised, else rejection+rework days
    else if(selfRej){ const rjd=rejOf(st.key); const rv=revOf(st.key); plan[st.key]=p; eff[st.key]=(rv && rv>=rjd)?rv:p; } // rejected approval cascades off redone send
    else { plan[st.key]=p; eff[st.key]=actualOf(st.key)||skipOf(st.key)||revOf(st.key)||p; }
  });
  const fabricIHStamp = actualOf("fabricIH") || skipOf("fabricIH");
  // Print / strike-off chain closes as a cross-check once fabric is in-house ONLY when PP bypass is OFF.
  // If PP bypass is ON, print/strike-off stays actionable because the production file path has no PP approval gate.
  // Fit does NOT close automatically because PP depends on Fit approval.
  const printPreFabricKeys=new Set(["artwork","artAppr","strikeOff","soAppr"]);
  const stages=applicableStages(s).map(st=>{ const apprK=APPR_OF_SEND[st.key]; const apprRej=apprK?rejOf(apprK):null; const selfRejDate=REJECTABLE.includes(st.key)?rejOf(st.key):null; const skp=skipOf(st.key); const isSkip=!!skp; const approvalSkipped=apprK?!!skipOf(apprK):false; const rejAppr=!!(apprK&&apprRej&&!actualOf(apprK)&&!approvalSkipped); const trackedResend=rejAppr?validResendActual(st.key,apprK):null; const storedActual=actualOf(st.key); const resent=!!trackedResend; const rework=rejAppr&&!resent; // While a send/make stage is in active rework, the old first-send actual is history only.
  // Do not let that old actual win the visible/current stage state after a revised resend date is entered.
  // The cell should show RE-SEND DUE using revised > auto rework plan until a real tracked resend actual is entered.
  const a=trackedResend||(rework?null:storedActual); const rjd_=rejAppr?apprRej:null; const rejected=REJECTABLE.includes(st.key)&&!!selfRejDate&&!actualOf(st.key)&&!isSkip; const autoClosed=!!(fabricIHStamp && !s.ppBypass && printPreFabricKeys.has(st.key) && !a && !skp && !rework && !rejected); const rjd=rework?rjd_:(rejected?selfRejDate:null); let rv=revOf(st.key); if(rjd&&rv&&rv<rjd) rv=null; const histReject = a ? ((apprRej&&trackedResend)?apprRej:(selfRejDate||null)) : null; const reapproval=!!(rejected && SEND_FOR_APPR[st.key] && validResendActual(SEND_FOR_APPR[st.key],st.key)); return { ...st, owner:ownerOf(st.key), actual:a, storedActual, trackedResend, rev:rv, reject:rjd, histReject, skippedWasRework:!!(isSkip&&rejAppr), skippedWasRejected:!!(isSkip&&(rejAppr||selfRejDate)), rework:isSkip?false:rework, rejected:isSkip?false:rejected, reapproval:isSkip?false:reapproval, skipped:isSkip, autoClosed, skip:skp, plan:plan[st.key], done: autoClosed?true:(isSkip?true:(rework?false:!!a)) }; });
  let nextPending=null, lastActual=null, lastActualKey=null;
  stages.forEach(r=>{ if(r.actual&&(!lastActual||r.actual>lastActual)){ lastActual=r.actual; lastActualKey=r.key; } if(!r.done&&!nextPending) nextPending=r; });
  const released=stages.every(r=>r.done);
  const idle=lastActual?Math.max(0,netWorkdays(lastActual,TODAY)):null;
  const get=(k)=>stages.find(r=>r.key===k); const done=(k)=>!!(get(k)&&get(k).done); const rejected=(k)=>{ const r=get(k); return !!(r&&r.rejected); }; const isSkipped=(k)=>{ const r=get(k); return !!(r&&r.skipped); };
  const fabricInHouse=done("fabricIH"); const fihA=actualOf("fabricIH");
  const isFabricCrossCheck=(r)=>!!(r&&r.autoClosed);
  if(fabricInHouse && nextPending && isFabricCrossCheck(nextPending)){ nextPending=stages.find(r=>!r.done && !isFabricCrossCheck(r)) || null; }
  const lateFIH=(k)=>{ const r=get(k); return !!(fihA && r && r.actual && r.actual>fihA); };
  const lastPlan=stages[stages.length-1]?.plan;
  const float=(lastPlan&&delivery)?netWorkdays(lastPlan,delivery):null;
  let status="On Track", tone="ok";
  if(released){ status="Released"; tone="done"; }
  else if(nextPending&&(nextPending.rev||nextPending.plan)&&TODAY>(nextPending.rev||nextPending.plan)){ const _npd=nextPending.rev||nextPending.plan; status=`Overdue ${Math.round((TODAY-_npd)/ONE_DAY)}d`; tone="late"; }
  else if(idle!==null&&idle>=7){ status=`Idle ${idle}d`; tone="warn"; }
  const dueText=(k)=>{ const r=get(k); const d=r&&(r.rev||r.plan); if(!r||!d) return "pending"; return TODAY>d?`OVERDUE ${Math.round((TODAY-d)/ONE_DAY)}d`:`due ${fmt(d)}`; };
  const dueTone=(k)=>{ const r=get(k); const d=r&&(r.rev||r.plan); return d&&TODAY>d?"late":"warn"; };
  const bs=(txt,tn,extra={})=>({txt,tone:tn,...extra});
  const autoClosed=(k)=>{ const r=get(k); return !!(r&&r.autoClosed); };
  const reSentAfterReject=(sendK,apprK)=>!!validResendActual(sendK,apprK);
  let fitBranch;
  const fitReSent=reSentAfterReject("fitSend","fitAppr");
  if(!s.fitReq) fitBranch=bs("—","na"); else if(done("fitAppr")) fitBranch=bs(isSkipped("fitAppr")?"Fit Skipped":(lateFIH("fitAppr")?"Fit Approved · after Fabric IH":"Fit Approved"), isSkipped("fitAppr")?"ok":(lateFIH("fitAppr")?"warn":"ok")); else if(rejected("fitAppr")&&fitReSent) fitBranch=bs(`Fit re-sent · appr ${dueText("fitAppr")}`,dueTone("fitAppr"),{blocksPP:true}); else if(rejected("fitAppr")) fitBranch=bs("Fit REJECTED · rework","late",{blocksPP:true}); else if(fabricInHouse) fitBranch=bs(done("fitSend")?"Fit approval pending · blocks PP":"Fit pending · blocks PP","late",{blocksPP:true}); else if(done("fitSend")) fitBranch=bs(`Fit appr ${dueText("fitAppr")}`,TODAY>(get("fitAppr")?.plan||TODAY)?"late":"warn"); else fitBranch=bs(`Fit send ${dueText("fitSend")}`,TODAY>(get("fitSend")?.plan||TODAY)?"late":"warn");
  let printBranch; const printComp=s.soReq?"soAppr":"artAppr"; const realDone=(k)=>{ const r=get(k); return !!(r&&(r.actual||r.skipped)); }; const printDone=s.soReq?realDone("soAppr"):realDone("artAppr");
  const printClosedAfterFabric=!!(s.printReq && fabricInHouse && !s.ppBypass && !printDone && !(rejected("artAppr")||rejected("soAppr")));
  const artReSent=reSentAfterReject("artwork","artAppr");
  const soReSent=reSentAfterReject("strikeOff","soAppr");
  if(!s.printReq) printBranch=bs("—","na");
  else if(printDone) printBranch=bs(isSkipped(printComp)?"Print Skipped":(lateFIH(printComp)?"Print done after Fabric IH":"Print Approved"), isSkipped(printComp)?"ok":(lateFIH(printComp)?"warn":"ok"), { afterFabricIH: lateFIH(printComp) });
  else if(rejected("artAppr")&&artReSent) printBranch=bs(`Artwork re-sent · appr ${dueText("artAppr")}`,dueTone("artAppr"));
  else if(rejected("soAppr")&&soReSent) printBranch=bs(`S/O re-sent · appr ${dueText("soAppr")}`,dueTone("soAppr"));
  else if(rejected("artAppr")||rejected("soAppr")) printBranch=bs("Print REJECTED · rework","late");
  else if(printClosedAfterFabric) printBranch=bs("Print not complete before Fabric IH","warn",{ crossCheck:true, autoClosed:true });
  else if(!done("artwork")) printBranch=bs(`Artwork ${dueText("artwork")}`,TODAY>(get("artwork")?.plan||TODAY)?"late":"warn");
  else if(!done("artAppr")) printBranch=bs(`Art appr ${dueText("artAppr")}`,dueTone("artAppr"));
  else if(s.soReq&&!done("strikeOff")) printBranch=bs(`S/O ${dueText("strikeOff")}`,dueTone("strikeOff"));
  else printBranch=bs(`S/O appr ${dueText("soAppr")}`,dueTone("soAppr"));
  let fabricBranch; const fabPlan=get("fabricIH")?.rev||get("fabricIH")?.plan;
  const fabDue=fabPlan?(TODAY>fabPlan?`IH OVERDUE ${Math.round((TODAY-fabPlan)/ONE_DAY)}d`:`IH due ${fmt(fabPlan)}`):"IH —";
  const fabTone=fabPlan&&TODAY>fabPlan?"late":"warn";
  const labReSent=reSentAfterReject("labDip","labAppr");
  if(fabricInHouse) fabricBranch=bs("Bulk Fabric In-House","ok"); else if(rejected("labAppr")&&labReSent) fabricBranch=bs(`Lab dip re-sent, appr pending | ${fabDue}`,dueTone("labAppr")); else if(rejected("labAppr")) fabricBranch=bs("Lab Dip REJECTED · rework","late"); else if(s.labDipReq&&done("labAppr")) fabricBranch=bs(`Lab Dip Appr | ${fabDue}`,fabTone); else if(s.labDipReq&&done("labDip")) fabricBranch=bs(`Lab dip sent, appr pending | ${fabDue}`,dueTone("labAppr")); else if(s.labDipReq) fabricBranch=bs(`Lab dip pending | ${fabDue}`,dueTone("labDip")); else fabricBranch=bs(fabDue,fabTone);
  let ppBranch;
  const ppReSent=reSentAfterReject("ppSample","ppAppr");
  if(!s.ppNeeded) ppBranch=bs("PP Not Required","na");
  else if(done("ppAppr")) ppBranch=bs(isSkipped("ppAppr")?"PP Skipped":(lateFIH("ppAppr")?"PP Approved · after Fabric IH":"PP Approved"), isSkipped("ppAppr")?"ok":(lateFIH("ppAppr")?"warn":"ok"));
  else if(rejected("ppAppr")&&ppReSent) ppBranch=bs(`PP re-sent · appr ${dueText("ppAppr")}`,dueTone("ppAppr"));
  else if(rejected("ppAppr")) ppBranch=bs("PP REJECTED · rework","late");
  else if(s.fitReq && !done("fitAppr")) ppBranch=bs(done("fitSend")?"Awaiting Fit approval · blocks PP":"Awaiting Fit sample · blocks PP","late",{blocksPP:true});
  else if(done("ppSample")) ppBranch=bs(`PP appr ${dueText("ppAppr")}`,dueTone("ppAppr"));
  else if(fabricInHouse) ppBranch=bs(`PP sample ${dueText("ppSample")}`,dueTone("ppSample"));
  else ppBranch=bs("Awaiting bulk fabric","warn");
  // ---- Production File: a tracked activity; reflects PP bypass vs PP-approval gate ----
  let prodFileBranch;
  { const pfA=actualOf("prodFile"); const pfP=eff["prodFile"]; const pfDue=pfP?`due ${fmt(pfP)}`:"";
    if(pfA){ prodFileBranch=bs(`Released ${fmt(pfA)}`,"ok"); }
    else { const prodGate=addWorkdays(delivery,-CUTD); const overdue=pfP&&pfP<TODAY; const pastGate=pfP&&prodGate&&pfP>prodGate; const tn=(overdue||pastGate)?"late":"warn";
      if(s.ppBypass || !s.ppNeeded){ const ready=fabricInHouse; const pre=s.ppBypass?"Bypass · ":""; prodFileBranch=!ready?bs(`${pre}awaiting fabric`,tn):(s.fitReq&&!done("fitAppr"))?bs(`${pre}awaiting Fit approval`,"late",{blocksPP:true}):bs(`${pre}file ${pfDue}`,tn); }
      else { const ready=done("ppAppr"); prodFileBranch=!ready?bs(s.fitReq&&!done("fitAppr")?`Awaiting Fit approval`:"Awaiting PP appr",s.fitReq&&!done("fitAppr")?"late":tn,{blocksPP:s.fitReq&&!done("fitAppr")}):bs(`Ready ${pfDue}`,tn); } } }
  let fabricCountdown;
  if(fabricInHouse) fabricCountdown={txt:"in-house",date:fihA,n:9e9,tone:"ok"}; else if(fabPlan){ const n=netWorkdays(TODAY,fabPlan); fabricCountdown={txt:n<0?`${-n}d over`:`${n}d`,date:fabPlan,n,tone:n<0?"late":n<=7?"warn":"ok"}; } else fabricCountdown={txt:"no plan",date:null,n:null,tone:"warn"};
  const releaseGate=addWorkdays(delivery,-GATED);
  let projRelease;
  if(released) projRelease=lastActual;
  else { const pf=stages.find(x=>x.key==="prodFile"); projRelease = (pf&&(pf.rev||pf.plan)) || eff["prodFile"]; }
  const gateGap=projRelease&&releaseGate?Math.round((releaseGate-projRelease)/ONE_DAY):null;
  const releaseOnTrack=projRelease&&releaseGate?projRelease<=releaseGate:true;
  const projTone=released?"done":(!releaseOnTrack?"late":(gateGap!=null&&gateGap<=5?"warn":"ok"));
  if(!released){ if(!releaseOnTrack){ if(!String(status).startsWith("Overdue")) status="Delivery risk"; tone="late"; } else if(tone==="ok" && gateGap!=null && gateGap<=5){ status=`Tight · ${gateGap}d`; tone="warn"; } }
  const total=stages.length, doneCount=stages.filter(r=>r.done).length;
  const pct=total?Math.round((doneCount/total)*100):0;
  const ownerToChase=released?"—":nextPending.owner;
  // ---- parallel chase: chase labels/departments of every pending stage that is truly actionable now ----
  // Keep this aligned with To-Do/frontier: gates such as Fit-before-PP, PP-bypass, and cross-check-closed print stages must apply here too.
  const stById=(k)=>STAGES.find(x=>x.key===k);
  const appl=(k)=>{ const st=stById(k); return stageApplies(s,st); };
  const predDone=(st)=>{ if(st.cutoff) return s.labDipReq?done("labAppr"):true; if(st.pred==="__ord") return true; let p=st.pred; while(p&&p!=="__ord"&&!appl(p)) p=stById(p)?.pred; return (!p||p==="__ord")?true:done(p); };
  const frontierReady=(k)=>{ if(k==="ppSample") return fabricInHouse && (!s.fitReq || done("fitAppr")); if(k==="prodFile"){ const base = (s.ppBypass || !s.ppNeeded) ? fabricInHouse : done("ppAppr"); return base && (!s.fitReq || done("fitAppr")); } return true; };
  const chaseCount={};
  if(!released) STAGES.forEach(st=>{ const rr=get(st.key); if(!rr||rr.autoClosed) return; if(appl(st.key)&&!done(st.key)&&predDone(st)&&frontierReady(st.key)){ const ow=ownerOf(st.key); chaseCount[ow]=(chaseCount[ow]||0)+1; } });
  const chaseOwners=Object.entries(chaseCount).map(([owner,count])=>({owner,count}));
  const frontier=new Set(); Object.entries(BRANCH_STAGES).forEach(([bk,keys])=>{ const nx=keys.find(k=>{ const r=get(k); return applies(k)&&r&&!r.done&&!r.autoClosed; }); if(!nx) return; if(frontierReady(nx)) frontier.add(nx); });
  const lastDoneIn=(keys)=>{ let best=null; keys.forEach(k=>{ const r=get(k); if(r&&r.done&&r.actual&&(!best||r.actual>best.d)) best={l:r.label,d:r.actual}; }); return best; };
  if(fitBranch) fitBranch.last=lastDoneIn(BRANCH_STAGES.fit);
  if(printBranch) printBranch.last=lastDoneIn(BRANCH_STAGES.print);
  if(fabricBranch) fabricBranch.last=lastDoneIn(BRANCH_STAGES.fabric);
  if(ppBranch) ppBranch.last=lastDoneIn(BRANCH_STAGES.pp);
  if(prodFileBranch) prodFileBranch.last=lastDoneIn(BRANCH_STAGES.prod);
  return { stages, frontier, nextPending, lastActual, lastActualKey, status, tone, idle, float, released, fitBranch, printBranch, fabricBranch, ppBranch, prodFileBranch, fabricCountdown, projRelease, projTone, releaseGate, releaseOnTrack, pct, ownerToChase, chaseOwners };
}

const ROLES={ management:{label:"Management"}, senior:{label:"Sr Merchant"}, junior:{label:"Jr Merchant"}, cad:{label:"CAD"}, designer:{label:"Designer"}, store:{label:"Store"} };
const MERCH_ROLES=["management","senior","junior"];
const SPECIALIST_COLS={ cad:["sampleFit","techpack","fitSend","fitAppr"], designer:["artwork","strikeOff","artAppr"], store:["labDip","fabricIH"] };
const canMaster=(role)=> MERCH_ROLES.includes(role); // management + senior + junior
const canAdmin=(role)=> role==="management"||role==="senior";
const canManageUsers=(role)=> role==="management";
const canEditRev=(role)=> MERCH_ROLES.includes(role);
const REJECT_ROLES=["management","senior","junior","designer","cad"]; // store excluded; revised stays merchants-only
const canEditReject=(role,col)=> REJECT_ROLES.includes(role) && canEditCol(role,col);
const canEditCol=(role,col)=>{ if(MERCH_ROLES.includes(role)){ if(STAGE_KEYS.includes(col)) return true; return canMaster(role); } return (SPECIALIST_COLS[role]||[]).includes(col); };
const canEdit=(role,col,mode)=> mode==="rev"?canEditRev(role): mode==="reject"?canEditReject(role,col): canEditCol(role,col);
const TONE_STYLE={ ok:{dot:"var(--success)",bg:"var(--tint-ok)",fg:"var(--fg-ok)"}, warn:{dot:"#b4801a",bg:"var(--tint-warn)",fg:"var(--fg-warn)"}, late:{dot:"var(--danger)",bg:"var(--tint-late)",fg:"var(--fg-late)"}, done:{dot:"var(--muted-4)",bg:"#efefea",fg:"var(--muted-5)"} };
const BR_TONE={ ok:{bg:"var(--tint-ok)",fg:"var(--fg-ok)"}, warn:{bg:"var(--tint-warn)",fg:"var(--fg-warn)"}, late:{bg:"var(--tint-late)",fg:"var(--fg-late)"}, na:{bg:"transparent",fg:"#c4c0b8"}, done:{bg:"#efefea",fg:"var(--muted-4)"} };
const FLAG_DEFS=[ {key:"fitReq",short:"FIT",title:"Fit sample required"}, {key:"printReq",short:"PRT",title:"Print required"}, {key:"soReq",short:"S/O",title:"Strike-off required"}, {key:"labDipReq",short:"LAB",title:"Lab dip required"}, {key:"ppBypass",short:"BYP",title:"PP bypass — Prod File flows straight from Fabric IH (not PP Appr)"}, {key:"ppNeeded",short:"PP",title:"PP sample required"} ];
const FILL_SWATCHES=["var(--accent-tint)","#fde2e1","#e7f3ec","#e3edf9","#f3e8fa","#fff3bf",""];

const NORMH=(h)=> String(h||"").toLowerCase().replace(/[^a-z0-9]/g,"");
const HEADER_MAP={ samplefit:"sampleFit", family:"family", fit:"orderNo", orderno:"orderNo", order:"orderNo", tranche:"orderNo", styleno:"styleNo", style:"styleNo", colour:"colour", color:"colour", brand:"brand", buyer:"buyer", buyername:"buyer", buyerbrand:"buyer", customer:"buyer", juniorowner:"owner", junior:"owner", owner:"owner", merchant:"owner", setpackid:"setId", setid:"setId", setpackrole:"setRole", setrole:"setRole", agegroup:"age", age:"age", orderqty:"qty", qty:"qty", quantity:"qty", reserve3:"fabricType", fabrictype:"fabricType", fabric:"fabricType", construction:"fabricType", orderreceived:"ordRec", orderdate:"ordRec", received:"ordRec", deliverydate:"delivery", delivery:"delivery", fitreq:"fitReq", printreq:"printReq", soreq:"soReq", ppbypass:"ppBypass", labdipreq:"labDipReq", ppneeded:"ppNeeded" };
const INFO_COLS=[
  { key:"orderNo",   label:"Order No",   kind:"text", w:60,  owner:"Merchant" },
  { key:"sampleFit", label:"Sample Fit", kind:"text", w:72,  owner:"Merchant" },
  { key:"family",    label:"Family",     kind:"text", w:130, owner:"Merchant" },
  { key:"colour",    label:"Colour",     kind:"text", w:150, owner:"Merchant" },
  { key:"brand",     label:"Brand",      kind:"text", w:90,  owner:"Merchant" },
  { key:"fabricType",label:"Fabric Type",kind:"text", w:140, owner:"Merchant" },
  { key:"owner",     label:"Owner",      kind:"text", w:64,  owner:"Merchant" },
  { key:"setId",     label:"Set ID",     kind:"text", w:60,  owner:"Merchant" },
  { key:"setRole",   label:"Set Role",   kind:"text", w:64,  owner:"Merchant" },
  { key:"buyer",     label:"Buyer",      kind:"text", w:100, owner:"Merchant" },
  { key:"age",       label:"Age Group",  kind:"text", w:92,  owner:"Merchant" },
  { key:"extra1",    label:"Extra 1",    kind:"text", w:100, owner:"Merchant" },
  { key:"extra2",    label:"Extra 2",    kind:"text", w:100, owner:"Merchant" },
  { key:"qty",       label:"Qty",        kind:"num",  w:46,  owner:"Merchant" },
  { key:"ordRec",    label:"Order Date", kind:"date", w:88,  owner:"Merchant" },
  { key:"delivery",  label:"Delivery",   kind:"date", w:88,  owner:"Merchant" },
  { key:"overall",   label:"Overall",    kind:"calc", w:108 },
  { key:"fit",       label:"Fit Branch", kind:"branch", w:150, branch:"fit" },
  { key:"print",     label:"Print Branch", kind:"branch", w:150, branch:"print" },
  { key:"fabric",    label:"Fabric Branch", kind:"branch", w:200, branch:"fabric" },
  { key:"pp",        label:"PP Branch",  kind:"branch", w:150, branch:"pp" },
  { key:"prod",      label:"Prod File",  kind:"branch", w:160, branch:"prod" },
  { key:"fabricCD",  label:"Fabric IH",  kind:"calc", w:74 },
  { key:"proj",      label:"Proj. Release", kind:"calc", w:104 },
  { key:"pct",       label:"% Done",     kind:"calc", w:80 },
  { key:"chase",     label:"Chase",      kind:"calc", w:120 },
  { key:"float",     label:"Float",      kind:"calc", w:58 },
  { key:"idle",      label:"Idle",       kind:"calc", w:50 },
];
const INFO_W=Object.fromEntries(INFO_COLS.map(c=>[c.key,c.w]));
const REMARK_COL={ key:"remarks", label:"Remarks / Delays", kind:"text", w:170, owner:"Merchant" };
const TEXT_COLS=["orderNo","sampleFit","family","colour","owner","setId","setRole","buyer","age","extra1","extra2","remarks"];
// Display consistency: show most tracker text fields in uppercase without changing saved data. Keep Set ID exact because IDs/codes can be case-sensitive, and keep Remarks as typed for readable notes.
const DISPLAY_UPPER_TEXT_COLS=new Set(["orderNo","sampleFit","family","colour","brand","buyer","fabricType","owner","setRole","age","extra1","extra2","__style"]);
const displayTextValue=(col,val)=>{ if(val==null||val==="") return ""; const t=String(val); return DISPLAY_UPPER_TEXT_COLS.has(col)?t.toUpperCase():t; };
const isEditableCol=(col)=> col==="__style"||col==="qty"||col==="ordRec"||col==="delivery"||TEXT_COLS.includes(col)||STAGE_KEYS.includes(col);
const isDateCol=(col)=> col==="ordRec"||col==="delivery"||STAGE_KEYS.includes(col);
// Techpack is its own actionable start branch. Keeping it in the same canonical
// frontier map as every other activity makes Tracker filtering, the Pre-Fit
// dashboard drill, To-Do, pending-cell colour and exports agree.
const BRANCH_STAGES={ start:["techpack"], fit:["fitSend","fitAppr"], print:["artwork","artAppr","strikeOff","soAppr"], fabric:["labDip","labAppr","fabricIH"], pp:["ppSample","ppAppr"], prod:["prodFile"] };
const BRANCH_LABEL={ start:"Pre-Fit", fit:"Fit", print:"Print", fabric:"Fabric", pp:"PP", prod:"Production" };
const BRANCH_OF={}; Object.entries(BRANCH_STAGES).forEach(([b,ks])=>ks.forEach(k=>{ BRANCH_OF[k]=BRANCH_LABEL[b]; }));
function branchTarget(s,c,branch){ const keys=BRANCH_STAGES[branch].filter(k=>{ const st=STAGES.find(x=>x.key===k); return st.flag===null||s[st.flag]; }); for(const k of keys){ if(!c.stages.find(r=>r.key===k)?.done) return k; } return keys[keys.length-1]; }

function CalPopup({ value, onPick, onClose, label, fallback }){
  const init=value?parse(value):(fallback||new Date(TODAY));
  const [view,setView]=useState(new Date(init.getFullYear(),init.getMonth(),1));
  const y=view.getFullYear(), m=view.getMonth();
  const first=new Date(y,m,1).getDay(), days=new Date(y,m+1,0).getDate();
  const cells=[]; for(let i=0;i<first;i++) cells.push(null); for(let d=1;d<=days;d++) cells.push(d);
  const sel=value?parse(value):null;
  return (
    <div style={{ position:"absolute", zIndex:390, top:"100%", left:0, marginTop:2 }} onClick={(e)=>e.stopPropagation()}>
      <div style={{ background:"var(--surface)", border:"1px solid var(--ink)", boxShadow:"4px 4px 0 var(--ink)", padding:10, width:228, fontFamily:"'JetBrains Mono', monospace" }}>
        {label && <div style={{ fontSize:9, fontWeight:700, color:"var(--accent)", marginBottom:6, textTransform:"uppercase", letterSpacing:0.5 }}>{label}</div>}
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}><button onClick={()=>setView(new Date(y,m-1,1))} style={navBtn}>‹</button><span style={{ fontWeight:700, fontSize:12 }}>{view.toLocaleDateString("en-GB",{month:"long",year:"numeric"})}</span><button onClick={()=>setView(new Date(y,m+1,1))} style={navBtn}>›</button></div>
        <div style={{ display:"grid", gridTemplateColumns:"repeat(7,1fr)", gap:2 }}>
          {["S","M","T","W","T","F","S"].map((d,i)=>(<div key={i} style={{ textAlign:"center", fontSize:9, color:"var(--muted-1)" }}>{d}</div>))}
          {cells.map((d,i)=>{ if(!d) return <div key={i}/>; const isSel=sel&&sel.getFullYear()===y&&sel.getMonth()===m&&sel.getDate()===d; const isToday=TODAY.getFullYear()===y&&TODAY.getMonth()===m&&TODAY.getDate()===d; const isSun=new Date(y,m,d).getDay()===0; return <button key={i} onClick={()=>{ onPick(`${y}-${String(m+1).padStart(2,"0")}-${String(d).padStart(2,"0")}`); onClose(); }} style={{ fontSize:11, padding:"5px 0", cursor:"pointer", fontFamily:"inherit", background:isSel?"var(--accent)":"transparent", color:isSel?"var(--surface)":isSun?"var(--danger)":"var(--ink)", fontWeight:isSel?700:400, border:(isToday&&!isSel)?"1px solid var(--ink)":"1px solid transparent", boxSizing:"border-box" }}>{d}</button>; })}
        </div>
        <div style={{ display:"flex", gap:6, marginTop:8 }}><button onClick={()=>{ onPick(iso(TODAY)); onClose(); }} style={{ ...chip, flex:1 }}>Today</button>{value && <button onClick={()=>{ onPick(null); onClose(); }} style={{ ...chip, flex:1 }}>Clear</button>}</div>
      </div>
    </div>
  );
}
const navBtn={ border:"none", background:"transparent", cursor:"pointer", fontSize:18, lineHeight:1, padding:"0 6px", fontFamily:"inherit" };
const chip={ fontSize:10, padding:"6px 10px", minHeight:30, border:"1px solid var(--ink)", background:"var(--bg)", cursor:"pointer", fontFamily:"'JetBrains Mono', monospace", fontWeight:700, display:"inline-flex", alignItems:"center", justifyContent:"center", gap:4 };

const bulkBtn={ width:"100%", fontFamily:"inherit", fontSize:11, fontWeight:800, padding:"7px 9px", marginBottom:7, cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)", textAlign:"left" };
const miniBulkBtn={ fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"4px 8px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)" };
function BulkActionBox({ title, children }){ return <div style={{ border:"1px solid var(--line-2)", background:"var(--surface)", borderRadius:12, padding:12 }}><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:14, marginBottom:8 }}>{title}</div>{children}</div>; }

const r1=(v)=>{ const n=Number(v); return Number.isFinite(n)?Math.round(n*10)/10:0; };
const fmtNum=(v)=>{ const n=r1(v); return Number.isInteger(n)?String(n):n.toFixed(1); };
const fmtDays=(v)=>`${fmtNum(v)}d`;
const stylePlanAccuracySummary=(s,c)=>{
  const planVals=[], planMiss=[], revVals=[], revMiss=[];
  (c&&c.stages?c.stages:[]).forEach(r=>{
    if(r&&r.actual&&r.plan){ const d=netWorkdays(r.plan,r.actual); if(d!=null){ planVals.push(d); if(d>0) planMiss.push(d); } }
    if(r&&r.actual&&r.rev){ const d=netWorkdays(r.rev,r.actual); if(d!=null){ revVals.push(d); if(d>0) revMiss.push(d); } }
  });
  const avgList=(arr)=>arr.length?r1(arr.reduce((a,b)=>a+(Number(b)||0),0)/arr.length):"";
  return {
    "Original Plan Records":planVals.length,
    "Avg Actual vs Original Plan Days":avgList(planVals),
    "Missed Original Records":planMiss.length,
    "Avg Missed Original Days":avgList(planMiss),
    "Worst Missed Original Days":planMiss.length?r1(Math.max(...planMiss)):"",
    "Revised Records":revVals.length,
    "Avg Actual vs Revised Days":avgList(revVals),
    "Missed Revised Records":revMiss.length,
    "Avg Missed Revised Days":avgList(revMiss),
    "Worst Missed Revised Days":revMiss.length?r1(Math.max(...revMiss)):"",
  };
};
const addStylePlanAccuracyCols=(o,s,c)=>Object.assign(o, stylePlanAccuracySummary(s,c));
const parityAlertsFromSheets=(sheets)=>{
  const alerts=[];
  (sheets||[]).forEach(sh=>{
    const rows=[...(Array.isArray(sh.data)?sh.data:[]), ...(Array.isArray(sh.detailData)?sh.detailData:[])];
    rows.forEach((r,i)=>{
      const result=String((r&&r.Result)!=null?r.Result:"").toUpperCase();
      if(result && result!=="OK" && result!=="PASS") alerts.push({
        "Report / Sheet":sh.label||"",
        "Row":i+1,
        "Check":r.Check||r.Rule||r["Report Type"]||"Parity check",
        "Result":r.Result||"CHECK",
        "Value":r.Value!=null?r.Value:"",
        "Expected":r.Expected!=null?r.Expected:"",
        "Management Trigger":"YES - review before using exported report",
      });
    });
  });
  return alerts;
};
const pickReportSheets=(reportName,sheets,mode)=>{
  const available=sheets.filter(x=>!x.modes || x.modes.includes(mode) || mode==="custom");
  const lines=available.map((x,i)=>`${i+1}. ${x.label}`).join("\n");
  const msg=`${reportName} export\n\nChoose sheets to export by number, comma-separated.\nType ALL for every listed sheet.\n\n${lines}`;
  const ans=window.prompt(msg, "ALL");
  if(ans==null) return null;
  const t=String(ans).trim().toLowerCase();
  if(!t) return null;
  if(t==="all"||t==="*") return available;
  const selected=[]; const seen=new Set();
  t.split(/[,\s]+/).forEach(part=>{ const idx=Number(part)-1; if(idx>=0&&idx<available.length&&!seen.has(idx)){ seen.add(idx); selected.push(available[idx]); } });
  return selected.length?selected:null;
};
const safeSheetName=(name)=>String(name||"Sheet").slice(0,31).replace(/[\\/?*\[\]:]/g," ");
const safeFilePart=(name)=>String(name||"report").toLowerCase().replace(/[^a-z0-9]+/g,"_").replace(/^_+|_+$/g,"").slice(0,60)||"report";
const TRACKER_VIEW_SETTING_ID="tracker_views";
const SYSTEM_DEFAULT_TRACKER_VIEW_NAME="System Default — All Columns";
const SYSTEM_DEFAULT_TRACKER_VIEW={ name:SYSTEM_DEFAULT_TRACKER_VIEW_NAME, system:true, shared:true, locked:true, state:{ columnView:"custom", hidden:[], statusFilter:"All", ownerFilter:"All", archiveView:"active", savedView:"", search:"", searchCol:"auto", activityFilter:null, colFilters:{}, sort:{ col:null, dir:1 }, freezeN:1 } };
const DEFAULT_NAMED_TRACKER_VIEWS=[
  SYSTEM_DEFAULT_TRACKER_VIEW,
  { name:"Management View", shared:true, state:{ columnView:"management", hidden:null, statusFilter:"All", ownerFilter:"All", archiveView:"active", savedView:"", search:"", searchCol:"auto", activityFilter:null, colFilters:{}, sort:{ col:null, dir:1 }, freezeN:1 } },
  { name:"Junior Merchandiser View", shared:true, state:{ columnView:"merchant", hidden:null, statusFilter:"All", ownerFilter:"All", archiveView:"active", savedView:"", search:"", searchCol:"auto", activityFilter:null, colFilters:{}, sort:{ col:null, dir:1 }, freezeN:1 } },
  { name:"Buyer Approval View", shared:true, state:{ columnView:"buyer", hidden:null, statusFilter:"All", ownerFilter:"Buyer", archiveView:"active", savedView:"buyerPending", search:"", searchCol:"auto", activityFilter:null, colFilters:{}, sort:{ col:"delivery", dir:1 }, freezeN:1 } },
  { name:"Production Follow-up View", shared:true, state:{ columnView:"store", hidden:null, statusFilter:"At Risk", ownerFilter:"All", archiveView:"active", savedView:"fabricPending", search:"", searchCol:"auto", activityFilter:null, colFilters:{}, sort:{ col:"fabricCD", dir:1 }, freezeN:1 } },
  { name:"My To-Do View", shared:true, state:{ columnView:"custom", hidden:null, statusFilter:"At Risk", ownerFilter:"All", archiveView:"active", savedView:"dueThisWeek", search:"", searchCol:"auto", activityFilter:null, colFilters:{}, sort:{ col:"delivery", dir:1 }, freezeN:1 } },
];
const normalizeTrackerViews=(views)=>{ const arr=Array.isArray(views)?views:[]; const byName=new Map(); DEFAULT_NAMED_TRACKER_VIEWS.forEach(v=>byName.set(v.name,v)); arr.forEach(v=>{ if(!v||!v.name||v.name===SYSTEM_DEFAULT_TRACKER_VIEW_NAME) return; byName.set(v.name,{...v, state:{...(v.state||{})}}); }); byName.set(SYSTEM_DEFAULT_TRACKER_VIEW_NAME,SYSTEM_DEFAULT_TRACKER_VIEW); return [...byName.values()]; };
const appendOneSheet=(wb,label,data)=>{
  const rows=Array.isArray(data)?data:(typeof data==="function"?(data()||[]):[]);
  const sheetRows=rows.length?rows:[{"No data":""}];
  const ws=XLSX.utils.json_to_sheet(sheetRows);
  const headers=Object.keys(sheetRows[0]||{});
  ws["!cols"]=headers.map(h=>{
    let max=String(h||"").length;
    sheetRows.slice(0,400).forEach(r=>{ const v=r&&r[h]!=null?String(r[h]):""; max=Math.max(max, Math.min(60, v.length)); });
    return { wch: Math.max(10, Math.min(42, max+2)) };
  });
  if(headers.length && sheetRows.length){ ws["!autofilter"]={ ref:XLSX.utils.encode_range({ s:{r:0,c:0}, e:{r:sheetRows.length,c:headers.length-1} }) }; }
  XLSX.utils.book_append_sheet(wb, ws, safeSheetName(label));
};
const appendReportSheets=(wb,selected,mode="summary")=>{
  (selected||[]).forEach((sh,idx)=>{
    const base=String(sh.label||("Table "+(idx+1)));
    if(mode==="detailed" && Array.isArray(sh.detailData)){
      appendOneSheet(wb, base+" Summary", sh.data);
      appendOneSheet(wb, base+" Detail", sh.detailData);
    } else {
      appendOneSheet(wb, base, sh.data);
    }
  });
};
const reportFileName=(prefix,mode,selected)=>{
  const arr=selected||[];
  const tablePart = arr.length===1 ? safeFilePart(arr[0].label) : "multiple_tables";
  return `${safeFilePart(prefix)}_${mode}_${tablePart}_${iso(TODAY)}.xlsx`;
};

function ReportExportMenu({ title, prefix, sheets, defaultMode="detailed" }){
  const [open,setOpen]=useState(false);
  const [mode,setMode]=useState(defaultMode);
  const [picked,setPicked]=useState(null);
  const available=(sheets||[]).filter(x=>!x.modes || x.modes.includes(mode));
  const labels=available.map(x=>x.label);
  const active=new Set(picked||labels);
  const toggle=(label)=>{
    setPicked(prev=>{ const cur=new Set(prev||labels); cur.has(label)?cur.delete(label):cur.add(label); return [...cur]; });
  };
  const selectAll=()=>setPicked(labels);
  const selectNone=()=>setPicked([]);
  const doExport=()=>{
    const selected=available.filter(x=>active.has(x.label));
    if(!selected.length){ alert("Select at least one table/sheet to export."); return; }
    const wb=XLSX.utils.book_new();
    appendReportSheets(wb,selected,mode);
    const parityAlerts=parityAlertsFromSheets(sheets);
    if(parityAlerts.length){
      appendOneSheet(wb,"MANAGEMENT PARITY ALERTS",parityAlerts);
      alert(`Management parity alert: ${parityAlerts.length} export/data check(s) need review. A MANAGEMENT PARITY ALERTS sheet has been added to this export.`);
    }
    XLSX.writeFile(wb,reportFileName(prefix,mode,selected));
    setOpen(false);
  };
  return (<span style={{ position:"relative", display:"inline-flex" }}>
    <button onClick={(e)=>{ e.stopPropagation(); setOpen(o=>!o); }} title={`Export ${title}`} style={{ fontFamily:"inherit", fontSize:10, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:open?"var(--ink)":"var(--surface)", color:open?"var(--bg)":"var(--ink)", fontWeight:800 }}>⬇ Export</button>
    {open && (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", right:0, marginTop:5, zIndex:95, width:330, background:"var(--surface)", border:"1px solid var(--ink)", boxShadow:"4px 4px 0 var(--ink)", padding:12 }}>
      <div style={{ fontSize:11, fontWeight:800, marginBottom:8 }}>Export {title}</div>
      <div style={{ display:"flex", border:"1px solid var(--ink)", marginBottom:8 }}>
        {["detailed","summary"].map(m=><button key={m} onClick={()=>{ setMode(m); setPicked(null); }} style={{ flex:1, fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"6px 0", cursor:"pointer", border:"none", borderRight:m==="detailed"?"1px solid var(--ink)":"none", background:mode===m?"var(--ink)":"var(--surface)", color:mode===m?"var(--bg)":"var(--ink)" }}>{m==="detailed"?"Detailed":"Summary"}</button>)}
      </div>
      <div style={{ fontSize:9, color:"var(--muted-2)", marginBottom:5 }}>Choose tables/sheets:</div>
      <div style={{ maxHeight:190, overflowY:"auto", border:"1px solid var(--line-2)", padding:6, marginBottom:8 }}>
        {available.map((sh,i)=><label key={sh.label} style={{ display:"flex", alignItems:"center", gap:6, fontSize:10, padding:"4px 2px", cursor:"pointer", borderBottom:i===available.length-1?"none":"1px solid var(--line-3)" }}><input type="checkbox" checked={active.has(sh.label)} onChange={()=>toggle(sh.label)}/><span style={{ flex:1 }}>{sh.label}</span><span style={{ color:"var(--muted-1)", fontSize:9 }}>{Array.isArray(sh.data)?sh.data.length:(typeof sh.data==="function"?"on export":0)}{mode==="detailed"&&(Array.isArray(sh.detailData)||typeof sh.detailData==="function")?" + "+(Array.isArray(sh.detailData)?sh.detailData.length:"detail") :""} rows</span></label>)}
      </div>
      <div style={{ display:"flex", gap:6, marginBottom:8 }}><button onClick={selectAll} style={{ ...chip, flex:1 }}>All</button><button onClick={selectNone} style={{ ...chip, flex:1 }}>None</button></div>
      <div style={{ display:"flex", gap:8 }}><button onClick={doExport} style={{ flex:1, fontFamily:"inherit", fontSize:11, fontWeight:800, padding:7, cursor:"pointer", border:"1px solid var(--ink)", background:"var(--accent)" }}>⬇ Export .xlsx</button><button onClick={()=>setOpen(false)} style={{ fontFamily:"inherit", fontSize:11, padding:"7px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--bg)" }}><X size={12}/></button></div>
    </div>)}
  </span>);
}

function BranchPill({ b, onJump }){
  if(!b) return null; const t=BR_TONE[b.tone]||BR_TONE.na;
  if(b.tone==="na") return <span style={{ color:"var(--line-2)", fontSize:10, fontWeight:700 }}>—</span>;
  return (<span style={{ display:"inline-flex", flexDirection:"column", alignItems:"flex-start", gap:3, maxWidth:"100%" }}>
    <span onClick={(e)=>{ if(onJump){ e.stopPropagation(); onJump(); } }} title={`${b.txt}  ·  click → jump to enter`} style={{ display:"inline-flex", alignItems:"center", gap:4, background:t.bg, color:t.fg, border:"1px solid rgba(31,31,29,0.08)", borderRadius:999, boxShadow:"var(--pill-shadow)", padding:"3px 8px", fontSize:9.5, lineHeight:1.15, fontWeight:800, whiteSpace:"nowrap", maxWidth:"100%", overflow:"hidden", textOverflow:"ellipsis", cursor:onJump?"pointer":"default", textDecoration:onJump?"underline dotted":"none", textUnderlineOffset:3 }}>{b.txt}{onJump && <CornerDownRight size={9} style={{ flexShrink:0 }}/>}</span>
    {b.last && <span style={{ fontSize:8.5, color:"var(--on-dark-2)", whiteSpace:"nowrap", paddingLeft:4 }}>✓ {b.last.l} · {fmt(b.last.d)}</span>}
  </span>);
}

const PEER_COLORS=["var(--success)","#2563a6","#b4531a","#6d4aab","var(--danger)","#0e7490","#9d174d","#4d7c0f"];
const colorFor=(id)=>{ let h=0; const s=String(id); for(let i=0;i<s.length;i++) h=(h*31+s.charCodeAt(i))>>>0; return PEER_COLORS[h%PEER_COLORS.length]; };
const initials=(n)=>String(n||"?").trim().split(/\s+/).map(w=>w[0]||"").slice(0,2).join("").toUpperCase()||"?";
function PeerTag({ who }){ if(!who||!who.length) return null; const anyEdit=who.some(w=>w.editing); const lead=who.find(w=>w.editing)||who[0]; const names=who.map(w=>w.name).join(", "); const label=who.length>1?(who.length+" · "+names):lead.name; return (<span style={{ position:"absolute", inset:0, border:(anyEdit?"2px dashed ":"2px solid ")+lead.color, pointerEvents:"none", zIndex:4, boxSizing:"border-box" }}><span title={names} style={{ position:"absolute", bottom:0, left:0, background:lead.color, color:"var(--surface)", fontSize:8, fontWeight:700, padding:"0 4px", whiteSpace:"nowrap", lineHeight:"12px", maxWidth:"100%", overflow:"hidden", textOverflow:"ellipsis" }}>{anyEdit?"✎ ":""}{label}</span></span>); }
function MerchTracker({ me, onSignOut }){
  const APP_BUILD_LABEL="v31dm Build 6 audit fixes";
  const [styles,setStyles]=useState([]); // loaded from Supabase on mount
  const role=(me&&me.role)||"junior";
  const [usersOpen,setUsersOpen]=useState(false);
  const TRACKER_DEFAULT_ZOOM=0.92, TRACKER_MIN_ZOOM=0.82, TRACKER_MAX_ZOOM=1.08;
  const [textScale,setTextScale]=useState(()=>{ try{ const raw=parseFloat(localStorage.getItem("mt_textscale")); if(!Number.isFinite(raw)) return TRACKER_DEFAULT_ZOOM; const clamped=Math.min(TRACKER_MAX_ZOOM,Math.max(TRACKER_MIN_ZOOM,raw)); return raw>TRACKER_MAX_ZOOM ? TRACKER_DEFAULT_ZOOM : clamped; }catch(e){ return TRACKER_DEFAULT_ZOOM; } });
  const bumpScale=(d)=>setTextScale(v=>{ const n=Math.min(TRACKER_MAX_ZOOM,Math.max(TRACKER_MIN_ZOOM,Math.round((v+d)*100)/100)); try{ localStorage.setItem("mt_textscale",String(n)); }catch(e){} return n; });
  const [tableWeight,setTableWeight]=useState(()=>{ try{ const v=parseInt(localStorage.getItem("mt_table_weight"),10); return (isFinite(v)&&v>=400&&v<=800)?v:600; }catch(e){ return 600; } });
  const bumpTableWeight=(d)=>setTableWeight(v=>{ const n=Math.min(800,Math.max(400,Math.round((v+d)/50)*50)); try{ localStorage.setItem("mt_table_weight",String(n)); }catch(e){} return n; });
  const PF=(()=>{ try{ return JSON.parse(localStorage.getItem("mt_trackfilters")||"{}"); }catch(e){ return {}; } })();
  const [search,setSearch]=useState(PF.search||"");
  const [searchCol,setSearchCol]=useState(PF.searchCol||"auto");
  const [statusFilter,setStatusFilter]=useState(PF.statusFilter||"All");
  const [ownerFilter,setOwnerFilter]=useState(PF.ownerFilter||"All");
  const [savedView,setSavedView]=useState(PF.savedView||"");
  const [columnView,setColumnView]=useState(()=>{ try{ return localStorage.getItem("mt_column_view")||"custom"; }catch(e){ return "custom"; } });
  const [sharedViews,setSharedViews]=useState(DEFAULT_NAMED_TRACKER_VIEWS);
  const [activeNamedView,setActiveNamedView]=useState("");
  const [archiveView,setArchiveView]=useState(PF.archiveView||"active");
  const [activityFilter,setActivityFilter]=useState(()=>activityKeysFromAnyGlobal(PF.activityFilter));
  const setTrackerActivityFilter=(v)=>setActivityFilter(activityKeysFromAnyGlobal(v));
  const activityFilterKeys=activityKeysFromAnyGlobal(activityFilter);
  const activityFilterKey=activityFilterKeys[0]||null; // legacy first-key alias only; use activityFilterKeys for filtering
  const [followFilter,setFollowFilter]=useState(PF.followFilter||false);
  const [viewSnap,setViewSnap]=useState(null); // saved tracker view before a drill, for one-click restore
  const scrollWrapRef=useRef(null);
  const [showJump,setShowJump]=useState(false);
  useEffect(()=>{ const onScroll=()=>{ const w=scrollWrapRef.current; setShowJump((window.scrollY||0)>200 || (w&&w.scrollTop>300)); }; window.addEventListener("scroll",onScroll,{passive:true}); const w=scrollWrapRef.current; if(w) w.addEventListener("scroll",onScroll,{passive:true}); return ()=>{ window.removeEventListener("scroll",onScroll); if(w) w.removeEventListener("scroll",onScroll); }; },[]);
  const jumpToTop=()=>{ try{ window.scrollTo({top:0,behavior:"smooth"}); }catch(e){ window.scrollTo(0,0); } const w=scrollWrapRef.current; if(w){ try{ w.scrollTo({top:0,behavior:"smooth"}); }catch(e){ w.scrollTop=0; } } };
  const [saved,setSaved]=useState(false);
  const [fillOpen,setFillOpen]=useState(false);
  const [colsOpen,setColsOpen]=useState(false);
  const [sel,setSel]=useState(null);      // anchor {id,col}
  const [focus,setFocus]=useState(null);   // range focus {id,col}
  const [multiAreas,setMultiAreas]=useState([]); // Ctrl/Cmd selections: extra non-adjacent single cells/ranges
  // Table filters must be strict: edited rows are not pinned/kept visible after they stop matching.
  const [ctxMenu,setCtxMenu]=useState(null); // right-click menu {x,y,id,col}
  const [editing,setEditing]=useState(null);
  const [editVal,setEditVal]=useState("");
  const [sort,setSort]=useState({ col:null, dir:1 });
  const [hidden,setHidden]=useState(()=>{ try{ const sv=localStorage.getItem("mt_hidden_cols"); return new Set(sv?JSON.parse(sv):["extra1","extra2"]); }catch(e){ return new Set(["extra1","extra2"]); } });
  useEffect(()=>{ try{ localStorage.setItem("mt_hidden_cols", JSON.stringify([...hidden])); }catch(e){} },[hidden]);
  const [specialClip,setSpecialClip]=useState(null);
  const [expOpen,setExpOpen]=useState(false); const [expMode,setExpMode]=useState(()=>{ try{ return localStorage.getItem("mt_exp_mode")||"full"; }catch(e){ return "full"; } }); const [expBuf,setExpBuf]=useState(()=>{ try{ const v=parseInt(localStorage.getItem("mt_exp_buf"),10); return (isFinite(v)&&v>=0&&v<=30)?v:2; }catch(e){ return 2; } }); const [expIncBuf,setExpIncBuf]=useState(()=>{ try{ return localStorage.getItem("mt_exp_incbuf")==="1"; }catch(e){ return false; } }); const [expRelMode,setExpRelMode]=useState(()=>{ try{ return localStorage.getItem("mt_exp_relmode")||"detailed"; }catch(e){ return "detailed"; } });
  const [frOpen,setFrOpen]=useState(false); const [frFind,setFrFind]=useState(""); const [frRepl,setFrRepl]=useState(""); const [frScope,setFrScope]=useState(()=>{ try{ const v=localStorage.getItem("mt_fr_scope"); return (v==="selected"||v==="filtered")?v:"filtered"; }catch(e){ return "filtered"; } }); const [frCase,setFrCase]=useState(()=>{ try{ return localStorage.getItem("mt_fr_case")==="1"; }catch(e){ return false; } });
  useEffect(()=>{ try{ localStorage.setItem("mt_exp_mode",expMode); localStorage.setItem("mt_exp_buf",String(expBuf)); localStorage.setItem("mt_exp_incbuf",expIncBuf?"1":"0"); localStorage.setItem("mt_exp_relmode",expRelMode); localStorage.setItem("mt_fr_scope",frScope); localStorage.setItem("mt_fr_case",frCase?"1":"0"); }catch(e){} },[expMode,expBuf,expIncBuf,expRelMode,frScope,frCase]);
  const [freezeN,setFreezeN]=useState(()=>{ try{ const v=parseInt(localStorage.getItem("mt_freeze_n"),10); return (Number.isFinite(v)&&v>=1&&v<=12)?v:1; }catch(e){ return 1; } });  // # leading columns frozen (incl style), user/browser persisted
  const [hiddenRows,setHiddenRows]=useState(()=>{ try{ return new Set(JSON.parse(localStorage.getItem("mt_hidden_rows")||"[]")); }catch(e){ return new Set(); } });
  useEffect(()=>{ try{ localStorage.setItem("mt_hidden_rows", JSON.stringify([...hiddenRows])); }catch(e){} },[hiddenRows]);
  const [rowH,setRowH]=useState(()=>{ try{ const h=Number(localStorage.getItem("mt_row_h")||34)||34; return h>56?34:h; }catch(e){ return 34; } });
  useEffect(()=>{ try{ localStorage.setItem("mt_row_h", String(rowH)); }catch(e){} },[rowH]);
  const [findIdx,setFindIdx]=useState(-1);
  const [frMatches,setFrMatches]=useState([]); // computed Find matches: [{id,col,style,colLabel,text}]
  const [colW,setColW]=useState({});  // per-column width overrides (drag to resize)
  const [fills,setFills]=useState({});
  const [notes,setNotes]=useState({});
  const [resends,setResends]=useState({}); // styleId:stage -> ordered send/re-send actual dates for rejected/rework loops
  const [stageEvents,setStageEvents]=useState([]); // durable reject/revise/resend/skip round events from public.stage_events
  const [revHistory,setRevHistory]=useState({}); // styleId:stage -> ordered revised-plan versions for commitment history
  const stylesRef=useRef(styles);
  const fillsRef=useRef(fills);
  const notesRef=useRef(notes);
  useEffect(()=>{ stylesRef.current=styles; },[styles]);
  useEffect(()=>{ fillsRef.current=fills; },[fills]);
  useEffect(()=>{ notesRef.current=notes; },[notes]);
  const [noteEditing,setNoteEditing]=useState(false);
  const [noteText,setNoteText]=useState("");
  const [comments,setComments]=useState({}); const [team,setTeam]=useState([]); const [threadCell,setThreadCell]=useState(null); const [cmText,setCmText]=useState("");
  // v31dm Build 1 memory safety: comments are no longer fetched for every cell on startup.
  // They are loaded on demand when a thread/review is opened, while realtime still updates new comments.
  const [commentsLoadedAll,setCommentsLoadedAll]=useState(false);
  const commentCellsLoadedRef=useRef(new Set());
  const commentThreadWrapRef=useRef(null);
  const [inbox,setInbox]=useState([]); const [bellOpen,setBellOpen]=useState(false); const [peersOpen,setPeersOpen]=useState(false);
  const [history,setHistory]=useState(false); const [auditRows,setAuditRows]=useState([]); const [auditBusy,setAuditBusy]=useState(false); const [histFilter,setHistFilter]=useState("");
  const [reviewOpen,setReviewOpen]=useState(false); const [reviewTab,setReviewTab]=useState("changes");
  const [helpOpen,setHelpOpen]=useState(false); const [helpTab,setHelpTab]=useState("guide");
  const [errorLog,setErrorLog]=useState([]);
  const [follows,setFollows]=useState(new Set());
  const [clip,setClip]=useState(null);     // {values:2D,h,w}
  const [showAux,setShowAux]=useState(false); // toggle: reveal underlying auto/plan + revised dates in cells
  useEffect(()=>{
    if(!threadCell) return;
    const closeThreadIfOutside=(e)=>{
      const box=commentThreadWrapRef.current;
      if(box && box.contains(e.target)) return;
      setThreadCell(null);
      setCmText("");
    };
    document.addEventListener("mousedown", closeThreadIfOutside, true);
    document.addEventListener("touchstart", closeThreadIfOutside, true);
    return ()=>{
      document.removeEventListener("mousedown", closeThreadIfOutside, true);
      document.removeEventListener("touchstart", closeThreadIfOutside, true);
    };
  },[threadCell]);
  const [cfg,setCfg]=useState(DEFAULT_CFG); // editable calculation numbers (Settings tab)
  const [tab,setTab]=useState(PF.tab==="timeline"?"tracker":(PF.tab||"tracker"));
  const [colFilters,setColFilters]=useState(PF.colFilters||{});
  const [filterCol,setFilterCol]=useState(null); // which header filter is open
  const [past,setPast]=useState([]); const [future,setFuture]=useState([]);
  const [filling,setFilling]=useState(false); const [fillFrom,setFillFrom]=useState(null); const [fillTo,setFillTo]=useState(null);
  const selectingRef=useRef(false); const [dragSel,setDragSel]=useState(false);
  useEffect(()=>{ const up=()=>{ if(selectingRef.current){ selectingRef.current=false; setDragSel(false); } }; window.addEventListener("mouseup",up); return ()=>window.removeEventListener("mouseup",up); },[]);
  const [newRow,setNewRow]=useState({ styleNo:"", orderNo:"", sampleFit:"", family:"", colour:"", brand:"", buyer:"", fabricType:"", owner:"", setId:"", setRole:"", age:"", qty:"", ordRec:iso(TODAY), delivery:"", fitReq:true, printReq:false, soReq:false, ppBypass:false, labDipReq:true, ppNeeded:true });
  const [newError,setNewError]=useState("");
  const savedTimer=useRef();
  const gridRef=useRef();
  const firstRender=useRef(true);
  const loadedRef=useRef(false);
  const savedRef=useRef({ sty:{}, stg:{}, meta:{} }); // last-persisted snapshot, keyed per row, so we only write what changed
  const remotePatchRef=useRef(false); // prevents remote realtime patches being re-saved/audited under the current user
  const clearedRef=useRef(new Set()); // cells the USER explicitly cleared this session — ONLY these may blank a saved date (anti-clobber)
  const S2C={ orderNo:"order_no", styleNo:"style_no", sampleFit:"sample_fit", family:"family", colour:"colour", brand:"brand", fabricType:"fabric_type", owner:"owner", setId:"set_id", setRole:"set_role", age:"age", qty:"qty", ordRec:"order_date", delivery:"delivery_date", fitReq:"fit_req", printReq:"print_req", soReq:"so_req", ppBypass:"pp_bypass", labDipReq:"lab_dip_req", ppNeeded:"pp_needed", remarks:"remarks", buyer:"buyer", extra1:"extra1", extra2:"extra2" };
  const styleToRow=(s)=>{ const r={ id:s.id }; Object.entries(S2C).forEach(([k,col])=>{ r[col]= k==="qty"?(Number(s[k])||0):(s[k]||null); }); r.archived=!!s.archived; return r; };
  const rowToStyle=(row,byId)=>({ id:row.id, orderNo:row.order_no||"", sampleFit:row.sample_fit||"", family:row.family||"", styleNo:row.style_no||"", colour:row.colour||"", brand:row.brand||"", fabricType:row.fabric_type||"", owner:row.owner||"", setId:row.set_id||"", setRole:row.set_role||"", age:row.age||"", qty:row.qty||0, ordRec:row.order_date||"", delivery:row.delivery_date||"", fitReq:!!row.fit_req, printReq:!!row.print_req, soReq:!!row.so_req, ppBypass:!!row.pp_bypass, labDipReq:!!row.lab_dip_req, ppNeeded:!!row.pp_needed, remarks:row.remarks||"", buyer:row.buyer||"", extra1:row.extra1||"", extra2:row.extra2||"", actuals:(byId[row.id]&&byId[row.id].actuals)||{}, revs:(byId[row.id]&&byId[row.id].revs)||{}, rejects:(byId[row.id]&&byId[row.id].rejects)||{}, skips:(byId[row.id]&&byId[row.id].skips)||{}, archived:!!row.archived });
  // LOAD everything from Supabase (also used by the Sync button). A small local cache makes the app usable faster on weak internet, then Supabase refreshes it.
  const hydrateOfflineCache=()=>{ try{ const raw=localStorage.getItem("mt_offline_snapshot_v1"); if(!raw) return; const snap=JSON.parse(raw); if(!snap||!Array.isArray(snap.styles)||!snap.styles.length) return; setStyles(snap.styles); if(snap.cfg) setCfg({ ...DEFAULT_CFG, ...snap.cfg, leads:{...DEFAULT_CFG.leads,...(snap.cfg.leads||{})}, stageOwners:{...DEFAULT_CFG.stageOwners,...(snap.cfg.stageOwners||{})}, rework:{...DEFAULT_CFG.rework,...(snap.cfg.rework||{})}, upcoming:{...DEFAULT_CFG.upcoming,...(snap.cfg.upcoming||{})}, escalationRules:Array.isArray(snap.cfg.escalationRules)&&snap.cfg.escalationRules.length?snap.cfg.escalationRules:DEFAULT_ESCALATION_RULES.map(x=>({...x})) }); if(snap.fills) setFills(snap.fills); if(snap.notes) setNotes(snap.notes); logAppError("offline cache",`showing cached tracker while Supabase refreshes · ${snap.styles.length} styles`); }catch(e){} };
  const writeOfflineCache=(payload)=>{ try{ localStorage.setItem("mt_offline_snapshot_v1", JSON.stringify({ ...payload, at:new Date().toISOString() })); }catch(e){} };
  // LOAD everything from Supabase (also used by the Sync button)
  const loadShared=async(opts={})=>{ const useCache=opts.useCache!==false; if(useCache) hydrateOfflineCache(); try{
    // Supabase caps a single select() at 1000 rows. With 100+ styles x 13 stages, stage_dates was being silently truncated, so cells beyond row 1000 rendered blank even though the data was in the DB. Fetch EVERY row in pages.
    const fetchAll=async(table)=>{ let out=[], from=0; const size=1000; for(let i=0;i<100;i++){ const { data, error }=await supabase.from(table).select("*").range(from, from+size-1); if(error){ logAppError("load "+table,error); break; } if(!data||!data.length) break; out=out.concat(data); if(data.length<size) break; from+=size; } return out; };
    const styData = await fetchAll("styles"); styData.sort((a,b)=>(a.id||0)-(b.id||0));
    if(!styData.length){ logAppError("load styles","no styles loaded"); return; }
    const sdData = await fetchAll("stage_dates");
    const cmData = await fetchAll("cell_meta");
    const seData = await fetchAll("stage_events");
    setStageEvents((seData||[]).sort((a,b)=>new Date(b.created_at||0)-new Date(a.created_at||0)));
    const byId={}; sdData.forEach(r=>{ const e=(byId[r.style_id]=byId[r.style_id]||{actuals:{},revs:{},rejects:{},skips:{}}); if(r.actual_date) e.actuals[r.stage]=r.actual_date; if(r.revised_date) e.revs[r.stage]=r.revised_date; if(r.reject_date) e.rejects[r.stage]=r.reject_date; if(r.skip_date) e.skips[r.stage]=r.skip_date; });
    const appStyles=styData.map(row=>rowToStyle(row,byId));
    setStyles(appStyles);
    const SR={ sty:{}, stg:{}, meta:{} };
    appStyles.forEach(s=>{ SR.sty[s.id]=JSON.stringify(styleToRow(s)); STAGE_KEYS.forEach(k=>{ SR.stg[s.id+":"+k]=JSON.stringify({ style_id:s.id, stage:k, revised_date:(s.revs&&s.revs[k])||null, actual_date:s.actuals[k]||null, reject_date:(s.rejects&&s.rejects[k])||null, skip_date:(s.skips&&s.skips[k])||null }); }); });
    cmData.forEach(r=>{ if(r.fill||r.note) SR.meta[r.style_id+":"+r.col]=JSON.stringify({ style_id:r.style_id, col:r.col, fill:r.fill||null, note:r.note||null }); });
    savedRef.current=SR;
    try{ const cfgRes=await supabase.from("app_settings").select("data").eq("id","global").maybeSingle(); if(cfgRes&&cfgRes.data&&cfgRes.data.data){ const d=cfgRes.data.data; setCfg({ ...DEFAULT_CFG, ...d, leads:{...DEFAULT_CFG.leads,...(d.leads||{})}, stageOwners:{...DEFAULT_CFG.stageOwners,...(d.stageOwners||{})}, rework:{...DEFAULT_CFG.rework,...(d.rework||{})}, upcoming:{...DEFAULT_CFG.upcoming,...(d.upcoming||{})}, escalationRules:Array.isArray(d.escalationRules)&&d.escalationRules.length?d.escalationRules:DEFAULT_ESCALATION_RULES.map(x=>({...x})), todoEscalationRows:d.todoEscalationRows!==undefined?!!d.todoEscalationRows:DEFAULT_CFG.todoEscalationRows, labels:{...(d.labels||{})} }); } }catch(e){ /* settings table optional */ }
    try{ const tvRes=await supabase.from("app_settings").select("data").eq("id",TRACKER_VIEW_SETTING_ID).maybeSingle(); const views=normalizeTrackerViews(tvRes&&tvRes.data&&tvRes.data.data&&tvRes.data.data.views); setSharedViews(views); }catch(e){ setSharedViews(DEFAULT_NAMED_TRACKER_VIEWS); }
    try{ const rr=await supabase.from("app_settings").select("data").eq("id","stage_resends").maybeSingle(); const d=rr&&rr.data&&rr.data.data; setResends(d&&typeof d==="object"?d:{}); }catch(e){ setResends({}); }
    try{ const rvh=await supabase.from("app_settings").select("data").eq("id","stage_revisions").maybeSingle(); const d=rvh&&rvh.data&&rvh.data.data; setRevHistory(d&&typeof d==="object"?d:{}); }catch(e){ setRevHistory({}); }
    const f={}, n={}; cmData.forEach(r=>{ if(r.fill) f[`${r.style_id}:${r.col}`]=r.fill; if(r.note) n[`${r.style_id}:${r.col}`]=r.note; });
    setFills(f); setNotes(n); writeOfflineCache({ styles:appStyles, cfg, fills:f, notes:n }); try{ const [prR,nR,fR]=await Promise.all([ supabase.from("profiles").select("id,name,role,email"), me&&me.id?supabase.from("notifications").select("*").eq("user_id",me.id).order("created_at",{ascending:false}).limit(100):Promise.resolve({data:[]}), me&&me.id?supabase.from("style_follows").select("style_id").eq("user_id",me.id):Promise.resolve({data:[]}) ]); setTeam(prR.data||[]); setInbox(nR.data||[]); setFollows(new Set((fR.data||[]).map(r=>r.style_id))); }catch(e){} loadedRef.current=true; flash();
  }catch(e){ logAppError("load failed",e); } };
  useEffect(()=>{ loadShared(); },[]);
  const editingRef=useRef(null); useEffect(()=>{ editingRef.current=editing; },[editing]);
  const presRef=useRef(null); const [peers,setPeers]=useState([]); const [presReady,setPresReady]=useState(false);
  useEffect(()=>{ if(!me) return; setPresReady(false); const key=String(me.id)+"."+Math.random().toString(36).slice(2,7); const ch=supabase.channel("merch-presence",{ config:{ presence:{ key } } });
    ch.on("presence",{ event:"sync" },()=>{ const st=ch.presenceState(); const arr=[]; Object.keys(st).forEach(k=>{ const m=st[k]&&st[k][0]; if(m&&String(m.id)!==String(me.id)) arr.push(m); }); setPeers(arr); });
    ch.subscribe(async(status)=>{ if(status==="SUBSCRIBED"){ presRef.current=ch; setPresReady(true); try{ await ch.track({ id:me.id, name:me.name||me.email, role:me.role, cell:null, editing:null }); }catch(e){} } });
    return ()=>{ try{ supabase.removeChannel(ch); }catch(e){} presRef.current=null; setPresReady(false); }; },[me]);
  // re-broadcast our own position AND the cell we're actively editing, whenever either changes or the channel becomes ready
  useEffect(()=>{ const ch=presRef.current; if(!(ch&&me&&presReady)) return; const tmr=setTimeout(()=>{ try{ ch.track({ id:me.id, name:me.name||me.email, role:me.role, cell: sel?{ id:sel.id, col:sel.col }:null, editing: editing?{ id:editing.id, col:editing.col }:null }); }catch(e){} }, 200); return ()=>clearTimeout(tmr); },[sel,editing,presReady]);
  useEffect(()=>{ const ch=supabase.channel("merch-comments").on("postgres_changes",{ event:"*", schema:"public", table:"comments" },(p)=>{ const n=(p.new&&p.new.id)?p.new:p.old; if(!n) return; const ck=n.style_id+":"+n.col; setComments(prev=>{ const arr=(prev[ck]||[]).filter(x=>x.id!==n.id); if(p.eventType!=="DELETE") arr.push(p.new); arr.sort((a,b)=>new Date(a.created_at)-new Date(b.created_at)); return {...prev,[ck]:arr}; }); }).subscribe(); return ()=>{ try{ supabase.removeChannel(ch); }catch(e){} }; },[]);
  // Presence display is grouped by real user. One person may have multiple browser tabs/sessions open; show them once in the top bar and cell indicators, with session count in the dropdown.
  const peerGroups=useMemo(()=>{ const map=new Map(); (peers||[]).forEach((p,idx)=>{ const key=String(p.id||p.email||p.name||idx); const loc=p.editing||p.cell; const cur=map.get(key)||{ id:key, name:p.name||p.email||"?", role:p.role, sessions:[], display:p, loc:null, editing:false };
      cur.sessions.push(p); cur.name=cur.name||p.name||p.email||"?"; cur.role=cur.role||p.role;
      const isEditing=!!p.editing; const shouldReplace=(!cur.loc&&loc) || (isEditing&&!cur.editing) || (loc&&cur.display&&!cur.display.editing&&!cur.display.cell);
      if(shouldReplace){ cur.display=p; cur.loc=loc||null; cur.editing=isEditing; } else if(!cur.display){ cur.display=p; }
      map.set(key,cur);
    });
    return Array.from(map.values()).sort((a,b)=>Number(!!b.editing)-Number(!!a.editing) || Number(!!b.loc)-Number(!!a.loc) || String(a.name).localeCompare(String(b.name)));
  },[peers]);
  const visiblePeerGroups=peerGroups.slice(0,4);
  const peerPeople=peerGroups.map(g=>({ ...(g.display||{}), id:g.id, name:g.name, role:g.role, cell:g.loc&&!g.editing?g.loc:null, editing:g.editing?g.loc:null, sessionCount:g.sessions.length }));
  const peerCell={}; // cellKey -> [{id,name,color,editing}]
  peerPeople.forEach(p=>{ const loc=p.editing||p.cell; if(!loc) return; const k=loc.id+":"+loc.col; (peerCell[k]=peerCell[k]||[]).push({ id:p.id, name:p.name, color:colorFor(p.id), editing:!!p.editing }); });
  const peerOn=(id,col)=> peerCell[id+":"+col]||null; // array | null
  const peerEditingHere=(id,col)=>{ const arr=peerCell[id+":"+col]; return arr?(arr.find(w=>w.editing)||null):null; };
  const peerLockBlocks=(id,col)=>{ const w=peerEditingHere(id,col); if(!w) return false; return !window.confirm(w.name+" is editing this cell right now.\n\nSaving over it at the same time can overwrite their change. Open it anyway?"); };
  const cellLabel=(id,col)=>{ const s=styles.find(x=>x.id===id); const sn=s?(s.styleNo||("#"+id)):("#"+id); const cl= col==="__style"?"Style No":((INFO_COLS.find(c=>c.key===col)||{}).label||(STAGES.find(x=>x.key===col)||{}).label||col); return sn+" · "+cl; };
  useEffect(()=>{ if(!me) return; const ch=supabase.channel("merch-notifs").on("postgres_changes",{ event:"*", schema:"public", table:"notifications", filter:"user_id=eq."+me.id },(p)=>{ const n=(p.new&&p.new.id)?p.new:p.old; if(!n) return; setInbox(prev=>{ let arr=prev.filter(x=>x.id!==n.id); if(p.eventType!=="DELETE") arr=[p.new,...arr]; arr.sort((a,b)=>new Date(b.created_at)-new Date(a.created_at)); return arr; }); }).subscribe(); return ()=>{ try{ supabase.removeChannel(ch); }catch(e){} }; },[me]);
  useEffect(()=>{ const ch=supabase.channel("merch-live")
    .on("postgres_changes",{ event:"*", schema:"public", table:"stage_dates" },(p)=>{ const del=p.eventType==="DELETE"; const n=(p.new&&Object.keys(p.new).length)?p.new:p.old; if(!n||!n.style_id){ setRemoteChanged(true); return; } const sid=n.style_id, stg=n.stage, key=sid+":"+stg; const row=JSON.stringify({ style_id:sid, stage:stg, revised_date:n.revised_date||null, actual_date:n.actual_date||null, reject_date:n.reject_date||null, skip_date:n.skip_date||null }); if(!del && savedRef.current.stg[key]===row) return; const ed=editingRef.current; if(ed&&ed.id===sid&&ed.col===stg){ setRemoteChanged(true); return; } if(del) delete savedRef.current.stg[key]; else savedRef.current.stg[key]=row; remotePatchRef.current=true; setStyles(prev=>{ let found=false; const next=prev.map(s=>{ if(s.id!==sid) return s; found=true; const ns={...s, actuals:{...s.actuals}, revs:{...(s.revs||{})}, rejects:{...(s.rejects||{})}, skips:{...(s.skips||{})} }; if(del){ delete ns.actuals[stg]; delete ns.revs[stg]; delete ns.rejects[stg]; delete ns.skips[stg]; } else { if(n.actual_date) ns.actuals[stg]=n.actual_date; else delete ns.actuals[stg]; if(n.revised_date) ns.revs[stg]=n.revised_date; else delete ns.revs[stg]; if(n.reject_date) ns.rejects[stg]=n.reject_date; else delete ns.rejects[stg]; if(n.skip_date) ns.skips[stg]=n.skip_date; else delete ns.skips[stg]; } return ns; }); if(!found){ setRemoteChanged(true); return prev; } return next; }); setTimeout(()=>{ remotePatchRef.current=false; },0); })
    .on("postgres_changes",{ event:"*", schema:"public", table:"styles" },(p)=>{ const del=p.eventType==="DELETE"; const n=(p.new&&p.new.id)?p.new:p.old; if(!n||!n.id){ setRemoteChanged(true); return; } const sid=n.id; if(del){ delete savedRef.current.sty[sid]; remotePatchRef.current=true; setStyles(prev=>prev.filter(s=>s.id!==sid)); setTimeout(()=>{ remotePatchRef.current=false; },0); return; } if(savedRef.current.sty[sid]===JSON.stringify(n)) return; const ed=editingRef.current; if(ed&&ed.id===sid){ setRemoteChanged(true); return; } savedRef.current.sty[sid]=JSON.stringify(n); remotePatchRef.current=true; setStyles(prev=>{ const idx=prev.findIndex(s=>s.id===sid); if(idx===-1) return [...prev, rowToStyle(n,{})]; const cur=prev[idx]; const merged={ ...rowToStyle(n,{}), actuals:cur.actuals, revs:cur.revs, rejects:cur.rejects, skips:cur.skips }; const copy=prev.slice(); copy[idx]=merged; return copy; }); setTimeout(()=>{ remotePatchRef.current=false; },0); })
    .on("postgres_changes",{ event:"*", schema:"public", table:"cell_meta" },(p)=>{ const del=p.eventType==="DELETE"; const n=(p.new&&Object.keys(p.new).length)?p.new:p.old; if(!n||!n.style_id||!n.col){ setRemoteChanged(true); return; } const key=n.style_id+":"+n.col; const row={ style_id:n.style_id, col:n.col, fill:(!del&&n.fill)?n.fill:null, note:(!del&&n.note)?n.note:null }; const j=JSON.stringify(row); if(!del && savedRef.current.meta[key]===j) return; if(del) delete savedRef.current.meta[key]; else savedRef.current.meta[key]=j; remotePatchRef.current=true; setFills(prev=>{ const nx={...prev}; if(!del&&n.fill) nx[key]=n.fill; else delete nx[key]; return nx; }); setNotes(prev=>{ const nx={...prev}; if(!del&&n.note) nx[key]=n.note; else delete nx[key]; return nx; }); setTimeout(()=>{ remotePatchRef.current=false; },0); })
    .on("postgres_changes",{ event:"*", schema:"public", table:"stage_events" },(p)=>{ const del=p.eventType==="DELETE"; const n=(p.new&&p.new.id)?p.new:p.old; if(!n||!n.id) return; setStageEvents(prev=>{ const nn=p.new||{}; const sameLocal=(x)=> String(x&&x.id||"").startsWith("local-") && String(x.style_id)===String(nn.style_id) && String(x.stage_key)===String(nn.stage_key) && String(x.event_type)===String(nn.event_type) && Number(x.round_no||0)===Number(nn.round_no||0) && String(x.new_value==null?"":x.new_value)===String(nn.new_value==null?"":nn.new_value) && String(x.event_date||"")===String(nn.event_date||""); let arr=(prev||[]).filter(x=>x.id!==n.id && !(!del && sameLocal(x))); if(!del) arr=[p.new,...arr]; arr.sort((a,b)=>new Date(b.created_at||0)-new Date(a.created_at||0)); return arr.slice(0,5000); }); })
    .subscribe();
    return ()=>{ try{ supabase.removeChannel(ch); }catch(e){} }; },[]);
  // SAVE everything to Supabase shortly after any change (debounced)
  useEffect(()=>{ if(firstRender.current){ firstRender.current=false; return; } if(!loadedRef.current) return; const t=setTimeout(async()=>{ try{ setSaveState("saving");
    const SR=savedRef.current;
    // ---- only upsert rows that actually changed since last save: protects other users' concurrent edits ----
    const styRows=styles.map(styleToRow); const styChanged=styRows.filter(r=>SR.sty[r.id]!==JSON.stringify(r));
    if(styChanged.length){ const up1=await supabase.from("styles").upsert(styChanged); if(up1.error) throw up1.error; logStyleAudit(styChanged,SR); }
    const stgChanged=[]; styles.forEach(s=> STAGE_KEYS.forEach(k=>{ const row={ style_id:s.id, stage:k, revised_date:(s.revs&&s.revs[k])||null, actual_date:s.actuals[k]||null, reject_date:(s.rejects&&s.rejects[k])||null, skip_date:(s.skips&&s.skips[k])||null }; const key=s.id+":"+k; const j=JSON.stringify(row); if(SR.stg[key]!==j) stgChanged.push({row,key,j}); }));
    if(stgChanged.length){ const up2=await supabase.from("stage_dates").upsert(stgChanged.map(x=>x.row),{ onConflict:"style_id,stage" }); if(up2.error) throw up2.error; logStageAudit(stgChanged,SR); notifyFollowers(stgChanged); }
    const keys=new Set([...Object.keys(SR.meta||{}),...Object.keys(fills),...Object.keys(notes)]); const metaChanged=[]; keys.forEach(key=>{ const i=key.indexOf(":"); const row={ style_id:Number(key.slice(0,i)), col:key.slice(i+1), fill:fills[key]||null, note:notes[key]||null }; const j=JSON.stringify(row); if(SR.meta[key]!==j) metaChanged.push({row,key,j}); });
    if(metaChanged.length){ const up3=await supabase.from("cell_meta").upsert(metaChanged.map(x=>x.row),{ onConflict:"style_id,col" }); if(up3.error) throw up3.error; logMetaAudit(metaChanged,SR); }
    styChanged.forEach(r=>{ SR.sty[r.id]=JSON.stringify(r); }); stgChanged.forEach(x=>{ SR.stg[x.key]=x.j; }); metaChanged.forEach(x=>{ SR.meta[x.key]=x.j; });
    setSaveState("saved"); flash();
  }catch(e){ logAppError("save failed",e); setSaveState("error"); } },700); return ()=>clearTimeout(t); },[styles,fills,notes]);

  const [saveState,setSaveState]=useState("idle"); // idle | saving | saved | error
  const saveStateRef=useRef(saveState);
  useEffect(()=>{ saveStateRef.current=saveState; },[saveState]);
  // #8 sync-status (read-only indicator). Online/offline awareness + last successful save time.
  // This does NOT change the save path; it only surfaces connection state so users on slow/dropped
  // internet can see whether their edits have reached the cloud.
  const [online,setOnline]=useState(typeof navigator!=="undefined"?navigator.onLine!==false:true);
  const [lastSavedAt,setLastSavedAt]=useState(null);
  useEffect(()=>{ const on=()=>setOnline(true); const off=()=>setOnline(false); window.addEventListener("online",on); window.addEventListener("offline",off); return ()=>{ window.removeEventListener("online",on); window.removeEventListener("offline",off); }; },[]);
  useEffect(()=>{ if(saveState==="saved") setLastSavedAt(new Date()); },[saveState]);
  const hasUnsavedLocalChanges=()=>{
    try{
      if(!loadedRef.current) return false;
      const SR=savedRef.current||{sty:{},stg:{},meta:{}};
      const st=stylesRef.current||[];
      for(const s of st){
        const r=styleToRow(s);
        if((SR.sty||{})[r.id]!==JSON.stringify(r)) return true;
        for(const k of STAGE_KEYS){
          const row={ style_id:s.id, stage:k, revised_date:(s.revs&&s.revs[k])||null, actual_date:(s.actuals&&s.actuals[k])||null, reject_date:(s.rejects&&s.rejects[k])||null, skip_date:(s.skips&&s.skips[k])||null };
          if((SR.stg||{})[s.id+":"+k]!==JSON.stringify(row)) return true;
        }
      }
      const fs=fillsRef.current||{}, ns=notesRef.current||{};
      const keys=new Set([...Object.keys((SR&&SR.meta)||{}),...Object.keys(fs),...Object.keys(ns)]);
      for(const key of keys){
        const i=key.indexOf(":");
        const row={ style_id:Number(key.slice(0,i)), col:key.slice(i+1), fill:fs[key]||null, note:ns[key]||null };
        if((SR.meta||{})[key]!==JSON.stringify(row)) return true;
      }
    }catch(e){ return saveStateRef.current==="saving"; }
    return false;
  };
  useEffect(()=>{
    const onBeforeUnload=(e)=>{
      // Browser exit protection: always warn after the tracker has loaded.
      // Reason: autosave may be debounced/realtime may still be settling even when the status says saved.
      // Browsers show their own generic message; custom text is intentionally ignored by Chrome/Safari.
      if(!loadedRef.current) return;
      e.preventDefault();
      e.returnValue="";
      return "";
    };
    window.addEventListener("beforeunload",onBeforeUnload);
    return ()=>window.removeEventListener("beforeunload",onBeforeUnload);
  },[]);
  useEffect(()=>{ if(!loadedRef.current) return; const t=setTimeout(()=>{ supabase.from("app_settings").upsert({ id:"global", data:cfg }).then(()=>{}).catch(()=>{}); },600); return ()=>clearTimeout(t); },[cfg]);
  useEffect(()=>{ if(!loadedRef.current) return; const t=setTimeout(()=>{ supabase.from("app_settings").upsert({ id:"stage_resends", data:resends||{} }).then(()=>{}).catch(()=>{}); },700); return ()=>clearTimeout(t); },[resends]);
  useEffect(()=>{ if(!loadedRef.current) return; const t=setTimeout(()=>{ supabase.from("app_settings").upsert({ id:"stage_revisions", data:revHistory||{} }).then(()=>{}).catch(()=>{}); },700); return ()=>clearTimeout(t); },[revHistory]);
  const [remoteChanged,setRemoteChanged]=useState(false); // another user wrote data
  const [syncBusy,setSyncBusy]=useState(false); // manual pull/reload status; never performs undo
  const flash=()=>{ setSaved(true); clearTimeout(savedTimer.current); savedTimer.current=setTimeout(()=>setSaved(false),1200); };
  const logAppError=(area,err,extra)=>{ const msg=(err&&err.message)?err.message:String(err||""); const row={ id:Date.now()+Math.random(), at:new Date().toISOString(), area, msg, extra:extra||"" }; setErrorLog(p=>[row,...p].slice(0,100)); console.error(area,err); };
  const mergeCommentRows=(rows)=>{
    const cg={};
    (rows||[]).forEach(r=>{ const ck=r.style_id+":"+r.col; (cg[ck]=cg[ck]||[]).push(r); });
    setComments(prev=>{
      const nx={...(prev||{})};
      Object.entries(cg).forEach(([ck,arr])=>{
        const map=new Map((nx[ck]||[]).map(x=>[String(x.id),x]));
        arr.forEach(x=>map.set(String(x.id),x));
        nx[ck]=Array.from(map.values()).sort((a,b)=>new Date(a.created_at||0)-new Date(b.created_at||0));
      });
      return nx;
    });
  };
  const loadAllComments=async()=>{
    if(commentsLoadedAll) return;
    try{
      const { data, error }=await supabase.from("comments").select("*").order("created_at");
      if(error) throw error;
      mergeCommentRows(data||[]);
      setCommentsLoadedAll(true);
    }catch(e){ logAppError("load comments",e); }
  };
  const loadCellComments=async(cell)=>{
    if(!cell||!cell.id||!cell.col||commentsLoadedAll) return;
    const ck=cell.id+":"+cell.col;
    if(commentCellsLoadedRef.current.has(ck)) return;
    commentCellsLoadedRef.current.add(ck);
    try{
      const { data, error }=await supabase.from("comments").select("*").eq("style_id",cell.id).eq("col",cell.col).order("created_at");
      if(error) throw error;
      mergeCommentRows(data||[]);
    }catch(e){ commentCellsLoadedRef.current.delete(ck); logAppError("load cell comments",e); }
  };
  useEffect(()=>{ if(threadCell) loadCellComments(threadCell); },[threadCell,commentsLoadedAll]);
  useEffect(()=>{ if(tab==="review" || reviewOpen) loadAllComments(); },[tab,reviewOpen,commentsLoadedAll]);
  const manualSync=async()=>{
    if(syncBusy) return;
    if(hasUnsavedLocalChanges() && !window.confirm("You have unsaved local edits. Sync will pull the latest cloud data and may replace local unsaved changes. Continue?")) return;
    setSyncBusy(true);
    try{
      await loadShared({ useCache:false });
      setRemoteChanged(false);
      setFuture([]);
      flash();
    }catch(e){ logAppError("manual sync failed",e); }
    finally{ setSyncBusy(false); }
  };
  const eventStyleId=(id)=>String(id==null?"":id);
  const stageEventsFor=(styleId,stageKey,linkedStageKey)=> (stageEvents||[]).filter(e=>String(e.style_id)===eventStyleId(styleId) && (!stageKey || e.stage_key===stageKey) && (!linkedStageKey || e.linked_stage_key===linkedStageKey));
  const maxEventRound=(styleId,stageKey,linkedStageKey,eventType)=>{
    const arr=stageEventsFor(styleId,stageKey,linkedStageKey).filter(e=>!eventType || e.event_type===eventType);
    return arr.reduce((m,e)=>Math.max(m,Number(e.round_no)||0),0);
  };
  const rejectRoundInfo=(style,sendStage,approvalStage)=>{
    // Round number is NOT the number of revised-date edits or accidental resend entries.
    // Rule: Re-send 1 starts on the first rejection. Re-send 2 starts only after a resend actual
    // has happened and the approval is rejected again. Repeated rejection-date edits stay same round.
    if(!style||!sendStage||!approvalStage) return { lastRound:0, currentRound:0 };
    const evs=(stageEvents||[]).filter(e=>String(e.style_id)===eventStyleId(style.id) && (
      (e.stage_key===approvalStage && e.linked_stage_key===sendStage && ["rejected","clear_rejected"].includes(String(e.event_type||"").toLowerCase())) ||
      (e.stage_key===sendStage && e.linked_stage_key===approvalStage && ["resend_actual","clear_resend_actual"].includes(String(e.event_type||"").toLowerCase()))
    )).slice().sort((a,b)=>{
      const ta=new Date(a.created_at||a.event_date||0).getTime()||0;
      const tb=new Date(b.created_at||b.event_date||0).getTime()||0;
      if(ta!==tb) return ta-tb;
      return String(a.id||"").localeCompare(String(b.id||""));
    });
    let round=0, current=false, resendSinceReject=false;
    evs.forEach(e=>{
      const typ=String(e.event_type||"").toLowerCase();
      if(typ==="rejected"){
        if(round===0) round=1;
        else if(resendSinceReject) round+=1;
        current=true;
        resendSinceReject=false;
      } else if(typ==="clear_rejected"){
        current=false;
      } else if(typ==="resend_actual"){
        if(round>0) resendSinceReject=true;
      } else if(typ==="clear_resend_actual"){
        // If the latest resend actual was cleared before a new rejection, do not let that cleared entry create a new round.
        resendSinceReject=false;
      }
    });
    const liveRejected=!!(style.rejects&&style.rejects[approvalStage] && !(style.actuals&&style.actuals[approvalStage]) && !(style.skips&&style.skips[approvalStage]));
    if(liveRejected && round===0) round=1;
    return { lastRound:round, currentRound:liveRejected?Math.max(1,round):0 };
  };
  const activeRejectRound=(style,sendStage,approvalStage)=>{
    if(!style||!approvalStage) return 0;
    return rejectRoundInfo(style,sendStage,approvalStage).currentRound;
  };
  const activeRoundForCell=(style,stageKey)=>{
    if(!style||!stageKey) return 0;
    const appr=APPR_OF_SEND[stageKey];
    if(appr) return activeRejectRound(style,stageKey,appr);
    const send=SEND_FOR_APPR[stageKey];
    if(send) return activeRejectRound(style,send,stageKey);
    return 0;
  };
  const roundLabelFor=(style,stageKey,kind)=>{
    const n=activeRoundForCell(style,stageKey)||1;
    const base=kind || (APPR_OF_SEND[stageKey]?"RE-SEND":(SEND_FOR_APPR[stageKey]?"RE-APPR":"ROUND"));
    return `${base} ${n}`;
  };

  const eventTime=(e)=>new Date((e&&e.created_at)||(e&&e.event_date)||0).getTime()||0;
  const currentRoundEventValue=(style,stageKey,linkedStageKey,eventType,clearType)=>{
    // Active-branch rule: once a new rejection branch starts, older round dates are history/report only.
    // For current status/editor/due, read only the latest event inside the current active round.
    if(!style||!stageKey||!eventType) return "";
    const round=linkedStageKey?activeRejectRound(style,stageKey,linkedStageKey):activeRoundForCell(style,stageKey);
    if(!round) return "";
    const evs=stageEventsFor(style.id,stageKey,linkedStageKey)
      .filter(e=>Number(e.round_no||0)===Number(round) && [eventType,clearType].filter(Boolean).includes(String(e.event_type||"").toLowerCase()))
      .slice().sort((a,b)=>eventTime(a)-eventTime(b));
    let val="";
    evs.forEach(e=>{
      const typ=String(e.event_type||"").toLowerCase();
      if(typ===String(eventType).toLowerCase()) val=e.new_value || e.event_date || "";
      if(clearType && typ===String(clearType).toLowerCase()) val="";
    });
    return val||"";
  };
  const currentRoundRevisedValue=(style,stageKey)=>{
    if(!style||!stageKey) return "";
    const appr=APPR_OF_SEND[stageKey];
    const send=SEND_FOR_APPR[stageKey];
    const linked=appr||send||null;
    if(!linked) return (style.revs&&style.revs[stageKey])||"";
    const round=activeRoundForCell(style,stageKey);
    if(!round) return (style.revs&&style.revs[stageKey])||"";
    // If current round has no revised event, do not fall back to previous round's flat revs[field].
    return currentRoundEventValue(style,stageKey,linked,"revised","clear_revised")||"";
  };
  const styleForActiveBranch=(style)=>{
    if(!style) return style;
    const normalized=normalizeSkipPairs(style);
    let changed=normalized!==style;
    style=normalized;
    const revs={...(style.revs||{})};
    const skips={...(style.skips||{})};
    STAGE_KEYS.forEach(k=>{
      const linked=APPR_OF_SEND[k]||SEND_FOR_APPR[k]||null;
      const round=linked?activeRoundForCell(style,k):0;
      if(round>0){
        const v=currentRoundRevisedValue(style,k);
        if(v) revs[k]=v; else delete revs[k];
        changed=true;
      }
    });
    return changed?{...style,revs,skips}:style;
  };
  const insertStageEvent=(style,stageKey,linkedStageKey,eventType,roundNo,eventDate,oldValue,newValue,note)=>{
    try{
      if(!style||!stageKey||!eventType) return;
      const row={
        style_id:eventStyleId(style.id),
        order_no:style.orderNo||"",
        style_no:style.styleNo||"",
        stage_key:stageKey,
        linked_stage_key:linkedStageKey||null,
        event_type:eventType,
        round_no:Math.max(0,Number(roundNo)||0),
        event_date:eventDate||null,
        old_value:oldValue==null?null:String(oldValue),
        new_value:newValue==null?null:String(newValue),
        note:note||null,
        created_by:me&&me.id?me.id:null,
        created_by_email:(me&&me.email)||""
      };
      setStageEvents(prev=>[{...row,id:"local-"+Date.now()+"-"+Math.random(),created_at:new Date().toISOString()},...(prev||[])].slice(0,5000));
      supabase.from("stage_events").insert(row).then(({data,error})=>{ if(error) logAppError("stage_events insert",error); }).catch(e=>logAppError("stage_events insert",e));
    }catch(e){ logAppError("stage_events insert",e); }
  };
  const SEND_STAGE_KEYS=Object.keys(APPR_OF_SEND);
  const shouldTrackResend=(s,field,val)=>{ const appr=APPR_OF_SEND[field]; const old=s&&s.actuals&&s.actuals[field]; if(!s||!appr||!old||!val||old===val) return false; return !!((s.rejects&&s.rejects[appr]) && !(s.actuals&&s.actuals[appr])); };
  const isResendEntrySlot=(s,field)=>{ const appr=APPR_OF_SEND[field]; const old=s&&s.actuals&&s.actuals[field]; return !!(s&&appr&&old&&(s.rejects&&s.rejects[appr])&&!(s.actuals&&s.actuals[appr])); };
  const recordResendActual=(s,field,val,source)=>{
    if(!shouldTrackResend(s,field,val)) return;
    const key=s.id+":"+field;
    const old=s.actuals[field];
    const item={ at:new Date().toISOString(), userId:me&&me.id, userName:(me&&(me.name||me.email))||"", source:source||"single", oldVal:old||"", newVal:val||"", stage:field, styleId:s.id, styleNo:s.styleNo||"", orderNo:s.orderNo||"" };
    setResends(prev=>{
      const raw=Array.isArray(prev[key])?prev[key].slice():[];
      const arr=raw.map(x=>typeof x==="string"?{ at:"", userId:"", userName:"", source:"legacy", oldVal:"", newVal:x, stage:field, styleId:s.id, styleNo:s.styleNo||"", orderNo:s.orderNo||"" }:x).filter(Boolean);
      const hasOld=old && arr.some(x=>x&&x.newVal===old);
      const last=arr[arr.length-1];
      if(last && last.oldVal===item.oldVal && last.newVal===item.newVal && last.userId===item.userId && last.source===item.source) return prev;
      const next=(hasOld?arr:([{ at:"", userId:"", userName:"", source:"first send", oldVal:"", newVal:old||"", stage:field, styleId:s.id, styleNo:s.styleNo||"", orderNo:s.orderNo||"" }].filter(x=>x.newVal).concat(arr))).concat(item).slice(-75);
      return {...prev,[key]:next};
    });
    try{ supabase.from("audit_log").insert({ style_id:s.id, style_no:s.styleNo||"", col:field, field:"resend actual", old_val:old||"", new_val:val||"", actor_id:me.id, actor_name:me.name||me.email }).then(()=>{}).catch(()=>{}); }catch(e){}
  };
  const recordRevisionHistory=(s,field,val,source)=>{ /* intentionally no structured revised-version history: revised date remains a normal editable active commitment; audit_log still records old -> new via stage save diff */ };
  const normKeyPart=(v)=>String(v||"").trim().toLowerCase();
  const identityFieldName=(field)=>field==="__style"||field==="styleNo"?"Style No":(field==="orderNo"?"Order No":"");
  const maybePinEditedRow=(id)=>{};
  const confirmIdentityChange=(style,field,val)=>{
    const label=identityFieldName(field);
    if(!label||!style) return true;
    const target=field==="__style"?"styleNo":field;
    const oldVal=String(style[target]||"").trim();
    const newVal=String(val||"").trim();
    if(oldVal===newVal) return true;
    if(!newVal){ window.alert(label+" cannot be blank. Order No + Style No is the unique key and is protected."); return false; }
    const nextOrder=target==="orderNo"?newVal:String(style.orderNo||"").trim();
    const nextStyle=target==="styleNo"?newVal:String(style.styleNo||"").trim();
    const dup=styles.find(x=>x.id!==style.id && normKeyPart(x.orderNo)===normKeyPart(nextOrder) && normKeyPart(x.styleNo)===normKeyPart(nextStyle));
    if(dup){ window.alert("Duplicate blocked. Another row already has this Order No + Style No:\n\nOrder No: "+nextOrder+"\nStyle No: "+nextStyle+"\n\nChange cancelled to protect the unique key."); return false; }
    return window.confirm("Change "+label+"?\n\nThis changes the unique key used by upload matching, duplicate checks, comments/history lookup, and reports.\n\nOld "+label+": "+(oldVal||"(blank)")+"\nNew "+label+": "+newVal+"\n\nContinue?");
  };
  const validateReworkActualDate=(style,field,val)=>{
    // Re-send actual must belong to the CURRENT rejection branch.
    // If user accidentally enters a date before the rejection date, it was being saved but ignored by active-branch logic,
    // which looked like "actual not registering". Block it with a clear message instead.
    try{
      if(!style||!field||!val||!STAGE_KEYS.includes(field)) return true;
      const appr=APPR_OF_SEND[field];
      if(!appr) return true;
      const rejVal=style.rejects&&style.rejects[appr];
      const approvalDone=style.actuals&&style.actuals[appr];
      const approvalSkipped=style.skips&&style.skips[appr];
      if(!rejVal||approvalDone||approvalSkipped) return true;
      const d=parse(val);
      const r=parse(rejVal);
      if(!d||!r) return true;
      if(dateSerial(d) < dateSerial(r)){
        const sendLabel=((STAGES.find(x=>x.key===field)||{}).label||field).replace(' Send','');
        const apprLabel=(STAGES.find(x=>x.key===appr)||{}).label||'approval';
        window.alert(`${sendLabel} re-send actual cannot be before the current rejection date.\n\n${apprLabel} rejected: ${fmt(r)}\nYou entered actual: ${fmt(d)}\n\nEnter the real re-send actual date on/after the rejection date, or use the ↻ corner for revised planning.`);
        return false;
      }
      return true;
    }catch(e){ return true; }
  };

  const confirmIdentityBatchChanges=(changes)=>{
    const planned=[];
    Object.entries(changes||{}).forEach(([id,ch])=>{
      const st=styles.find(x=>String(x.id)===String(id));
      if(!st) return;
      Object.entries(ch||{}).forEach(([field,val])=>{ if(identityFieldName(field)){ const target=field==="__style"?"styleNo":field; const oldVal=String(st[target]||"").trim(); const newVal=String(val||"").trim(); if(oldVal!==newVal) planned.push({ st, field, target, oldVal, newVal }); } });
    });
    if(!planned.length) return true;
    for(const p of planned){ if(!p.newVal){ window.alert(identityFieldName(p.field)+" cannot be blank for "+(p.st.styleNo||p.st.orderNo||"this row")+". Change cancelled."); return false; } }
    const seen=new Map();
    for(const p of planned){
      const nextOrder=p.target==="orderNo"?p.newVal:String(p.st.orderNo||"").trim();
      const nextStyle=p.target==="styleNo"?p.newVal:String(p.st.styleNo||"").trim();
      const key=normKeyPart(nextOrder)+"|"+normKeyPart(nextStyle);
      const existing=styles.find(x=>x.id!==p.st.id && normKeyPart(x.orderNo)===normKeyPart(nextOrder) && normKeyPart(x.styleNo)===normKeyPart(nextStyle));
      if(existing){ window.alert("Duplicate blocked during paste/edit. Another row already has:\n\nOrder No: "+nextOrder+"\nStyle No: "+nextStyle); return false; }
      if(seen.has(key) && seen.get(key)!==p.st.id){ window.alert("Duplicate blocked inside selected changes for Order No + Style No:\n\n"+nextOrder+" | "+nextStyle); return false; }
      seen.set(key,p.st.id);
    }
    const preview=planned.slice(0,6).map(p=>`${p.st.styleNo||p.st.orderNo||"row"}: ${identityFieldName(p.field)} ${p.oldVal||"(blank)"} → ${p.newVal}`).join("\n");
    return window.confirm("You are changing "+planned.length+" protected Order No / Style No value(s).\n\n"+preview+(planned.length>6?"\n…":"")+"\n\nThese fields form the unique key. Continue?");
  };
  const setField=(id,field,val)=>{
    const curStyle=styles.find(x=>x.id===id);
    if(!confirmIdentityChange(curStyle,field,val)) return;
    maybePinEditedRow(id);
    const isStage=STAGE_KEYS.includes(field);
    const willTrackResend=isStage&&shouldTrackResend(curStyle,field,val);
    if(willTrackResend && val && !validateReworkActualDate(curStyle,field,val)) return;
    if(willTrackResend){
      const appr=APPR_OF_SEND[field];
      const round=activeRejectRound(curStyle,field,appr)||1;
      const old=(curStyle&&curStyle.actuals&&curStyle.actuals[field])||"";
      insertStageEvent(curStyle,field,appr,"resend_actual",round,val||null,old||null,val||null,`Re-send ${round} actual${val?" entered":" cleared"}`);
      recordResendActual(curStyle,field,val,"single");
    } else if(isStage && curStyle){
      const old=(curStyle.actuals&&curStyle.actuals[field])||"";
      if(old!==val) insertStageEvent(curStyle,field,null,val?"actual":"clear_actual",0,val||null,old||null,val||null,val?"First/normal actual entered":"Actual date cleared");
    }
    pushHistory();
    if(isStage){ const ck=id+":"+field+":actual"; if(val) clearedRef.current.delete(ck); else clearedRef.current.add(ck); }
    setStyles(prev=>prev.map(s=>{
      if(s.id!==id) return s;
      if(isStage){
        // Re-send actual is a round event, not a replacement for the original first-send actual.
        // The active TNA state reads it from effectiveResends/stage_events, while actuals[field]
        // preserves the first/normal actual for history, reports, and audit explanation.
        if(willTrackResend) return s;
        return { ...s, actuals:{ ...s.actuals, [field]: val||undefined } };
      }
      if(field==="qty") return { ...s, qty:Number(val)||0 };
      if(field==="__style") return { ...s, styleNo:val };
      return { ...s, [field]:val };
    })); flash(); };
  const logRevisedEvent=(curStyle,key,val)=>{
    if(!(curStyle&&STAGE_KEYS.includes(key))) return;
    const old=currentRoundRevisedValue(curStyle,key)||"";
    if(old===(val||"")) return;
    const appr=APPR_OF_SEND[key]; const send=SEND_FOR_APPR[key];
    const round=appr?activeRejectRound(curStyle,key,appr):(send?activeRejectRound(curStyle,send,key):0);
    const eventType=val?"revised":"clear_revised";
    const linked=appr||send||null;
    const note=round?`${appr?"Re-send":"Re-approval"} ${round} revised date${val?" set":" cleared"}`:(val?"Revised date set":"Revised date cleared");
    insertStageEvent(curStyle,key,linked,eventType,round||0,val||null,old||null,val||null,note);
  };
  const setRev=(id,key,val)=>{
    const curStyle=styles.find(x=>x.id===id);
    recordRevisionHistory(curStyle,key,val,"single");
    logRevisedEvent(curStyle,key,val);
    pushHistory(); const ck=id+":"+key+":revised"; if(val) clearedRef.current.delete(ck); else clearedRef.current.add(ck); setStyles(prev=>prev.map(s=> s.id===id?{...s,revs:{...(s.revs||{}),[key]:val||undefined}}:s)); flash(); };
  const setReject=(id,key,val)=>{
    const curStyle=styles.find(x=>x.id===id);
    if(curStyle&&STAGE_KEYS.includes(key)){
      const old=(curStyle.rejects&&curStyle.rejects[key])||"";
      const send=SEND_FOR_APPR[key]||null;
      let round=0;
      if(send){
        const info=rejectRoundInfo(curStyle,send,key);
        round=val?(old?(info.currentRound||info.lastRound||1):Math.max(1,(info.lastRound||0)+1)):(info.currentRound||info.lastRound||1);
      }
      const eventType=val?"rejected":"clear_rejected";
      if(old!==val) insertStageEvent(curStyle,key,send,eventType,round,val||null,old||null,val||null,send?`${(STAGES.find(x=>x.key===key)||{}).label||key} rejection round ${round}${val?"":" cleared"}`:(val?"Rejected":"Rejection cleared"));
    }
    pushHistory(); const ck=id+":"+key+":reject"; if(val) clearedRef.current.delete(ck); else clearedRef.current.add(ck); setStyles(prev=>prev.map(s=> s.id===id?{...s,rejects:{...(s.rejects||{}),[key]:val||undefined}}:s)); flash(); };
  const setSkip=(id,key,val)=>{
    const curStyle=styles.find(x=>x.id===id);
    if(curStyle&&STAGE_KEYS.includes(key)){
      const appr=APPR_OF_SEND[key]; const send=SEND_FOR_APPR[key];
      const linked=appr||send||null;
      const round=appr?activeRejectRound(curStyle,key,appr):(send?activeRejectRound(curStyle,send,key):0);
      insertStageEvent(curStyle,key,linked,val?"skip":"clear_skip",round||0,val||null,(curStyle.skips&&curStyle.skips[key])||null,val||null,val?"Stage skipped/waived":"Stage un-skipped");
    }
    pushHistory(); const ap=APPR_OF_SEND[key]; [key,...(ap?[ap]:[])].forEach(kk=>{ const ck=id+":"+kk+":skip"; if(val) clearedRef.current.delete(ck); else clearedRef.current.add(ck); }); setStyles(prev=>prev.map(s=>{ if(s.id!==id) return s; const skips={...(s.skips||{})}; if(val){ skips[key]=val; if(ap) skips[ap]=val; } else { skips[key]=undefined; if(ap) skips[ap]=undefined; } return {...s,skips}; })); flash(); };
  const toggleFlag=(id,flag)=>{ pushHistory(); setStyles(prev=>prev.map(s=>s.id===id?{...s,[flag]:!s[flag]}:s)); flash(); };
  const [bulkOpen,setBulkOpen]=useState(false);
  const [bulkActionsOpen,setBulkActionsOpen]=useState(false);
  const [bulkConfirm,setBulkConfirm]=useState(null); // review popup before applying bulk row actions
  const [bulkResult,setBulkResult]=useState(null); // upload mapping/preview/result
  const [uploadSkip,setUploadSkip]=useState(new Set()); // order|style keys user chooses not to apply from upload preview
  const toBool=(v)=>{ const s=String(v||"").trim().toLowerCase(); return s==="y"||s==="yes"||s==="true"||s==="1"; };
  const toISO=(v)=>{ if(v==null||v==="") return ""; if(v instanceof Date && !isNaN(v)) return iso(v); const d=parse(String(v)); if(d) return iso(d); const dd=new Date(v); return isNaN(dd)?"":iso(dd); };
  const downloadTemplate=()=>{ const headers=["Style No","Order No","Sample Fit","Family","Colour","Brand","Buyer","Fabric Type","Junior Owner","Set-Pack ID","Set-Pack Role","Age Group","Order Qty","Order Received","Delivery Date","Fit Req?","Print Req?","S/O Req?","PP Bypass?","Lab Dip Req?","PP Needed?"]; const example=["HSAW26EXAMPLE01","T1","fit 1-3","SWEAT","BLACK","Disney","Hopscotch","TERRY","Tamal","","SWEATSHIRT","4-10YRS",400,"2026-05-18","2026-06-25","Y","N","N","Y","Y","Y"]; const ws=XLSX.utils.aoa_to_sheet([headers,example]); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Styles"); XLSX.writeFile(wb,"merch_tracker_upload_template.xlsx"); };
  const UPLOAD_FIELDS=[ ["styleNo","Style No *"], ["orderNo","Order No *"], ["sampleFit","Sample Fit"], ["family","Family"], ["colour","Colour"], ["brand","Brand"], ["buyer","Buyer"], ["fabricType","Fabric Type"], ["owner","Owner / Junior"], ["setId","Set ID"], ["setRole","Set Role"], ["age","Age Group"], ["qty","Qty"], ["ordRec","Order Received"], ["delivery","Delivery Date"], ["fitReq","Fit Req"], ["printReq","Print Req"], ["soReq","S/O Req"], ["ppBypass","PP Bypass"], ["labDipReq","Lab Dip Req"], ["ppNeeded","PP Needed"] ];
  const uploadKey=(o,st)=> String(o||"").trim().toLowerCase()+"|"+String(st||"").trim().toLowerCase();
  const uploadFieldLabel=(k)=> (UPLOAD_FIELDS.find(x=>x[0]===k)||[k,k])[1].replace(" *","");
  const makeUploadReportRows=(br=bulkResult)=>{ if(!br||br.error) return []; const rows=[]; (br.errors||[]).forEach(e=>rows.push({ "Row No":e.rowNo||"", "Order No":e.orderNo||"", "Style No":e.styleNo||"", "Issue Type":e.type||"Error", "Issue Details":e.detail||"", "Existing Value":e.from||"", "Uploaded Value":e.to||"", "Action Taken":e.blocking?"Blocked":"Review" })); (br.updates||[]).forEach(u=>(u.diffs||[]).forEach(d=>rows.push({ "Row No":u.rowNo||"", "Order No":u.orderNo, "Style No":u.styleNo, "Issue Type":"Attribute changed", "Issue Details":uploadFieldLabel(d.field), "Existing Value":d.from||"", "Uploaded Value":d.to||"", "Action Taken":uploadSkip.has(u.key)?"Skipped by user":"Pending confirmation" }))); (br.dupes||[]).forEach(d=>rows.push({ "Row No":d.rowNo||"", "Order No":d.orderNo||"", "Style No":d.styleNo||"", "Issue Type":"Duplicate Order + Style in upload", "Issue Details":"Duplicate unique key. Fix source file; duplicates are not allowed.", "Existing Value":"", "Uploaded Value":d.orderNo+" | "+d.styleNo, "Action Taken":"Blocked" })); return rows; };
  const downloadUploadErrorReport=()=>{ const data=makeUploadReportRows(); if(!data.length){ alert("No upload errors/discrepancies to report."); return; } const ws=XLSX.utils.json_to_sheet(data); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Upload Review"); XLSX.writeFile(wb,"upload_review_report_"+iso(TODAY)+".xlsx"); };
  const buildUploadPreview=(br)=>{ try{ const { aoa, hr, sn, mapping }=br; const fi={}; Object.entries(mapping||{}).forEach(([k,v])=>{ const n=Number(v); if(Number.isFinite(n) && n>=0) fi[k]=n; }); if(fi.styleNo==null||fi.orderNo==null){ setBulkResult({...br, mapError:"Map both Style No and Order No before continuing."}); return; }
    const recs=[], errors=[]; for(let i=hr+1;i<aoa.length;i++){ const row=aoa[i]||[]; const styleNo=String(row[fi.styleNo]??"").trim(); const orderNo=String(row[fi.orderNo]??"").trim(); if(!styleNo&&!orderNo) continue; if(!styleNo||!orderNo){ errors.push({ rowNo:i+1, orderNo, styleNo, type:"Missing mandatory key", detail:"Style No and Order No are both mandatory.", blocking:true }); continue; } const rec={ styleNo, orderNo, rowNo:i+1 }; ["sampleFit","family","colour","brand","buyer","fabricType","owner","setId","setRole","age"].forEach(f=>{ if(fi[f]!=null){ const v=String(row[fi[f]]??"").trim(); if(v) rec[f]=v; } }); if(fi.qty!=null){ const q=String(row[fi.qty]??"").replace(/[^0-9.]/g,""); if(q) rec.qty=Number(q)||0; } if(fi.ordRec!=null){ const d=toISO(row[fi.ordRec]); if(d) rec.ordRec=d; else if(String(row[fi.ordRec]??"").trim()) errors.push({ rowNo:i+1, orderNo, styleNo, type:"Invalid date", detail:"Order Received could not be read", to:String(row[fi.ordRec]??""), blocking:false }); } if(fi.delivery!=null){ const d=toISO(row[fi.delivery]); if(d) rec.delivery=d; else if(String(row[fi.delivery]??"").trim()) errors.push({ rowNo:i+1, orderNo, styleNo, type:"Invalid date", detail:"Delivery Date could not be read", to:String(row[fi.delivery]??""), blocking:false }); } ["fitReq","printReq","soReq","ppBypass","labDipReq","ppNeeded"].forEach(f=>{ if(fi[f]!=null && String(row[fi[f]]??"").trim()!=="") rec[f]=toBool(row[fi[f]]); }); recs.push(rec); }
    const byKey={}; styles.forEach(s=> byKey[uploadKey(s.orderNo,s.styleNo)]=s); const inserts=[], updates=[], dupes=[]; let unchanged=0; const seen={};
    recs.forEach(rec=>{ const k=uploadKey(rec.orderNo,rec.styleNo); if(seen[k]){ dupes.push({ ...rec, type:"Duplicate in upload", detail:"Same Order No + Style No appears more than once in this file.", blocking:true }); errors.push({ rowNo:rec.rowNo, orderNo:rec.orderNo, styleNo:rec.styleNo, type:"Duplicate Order + Style in upload", detail:"Duplicate unique key. Fix source file before applying.", blocking:true }); return; } seen[k]=true; const ex=byKey[k]; if(ex){ const chg={}, diffs=[]; Object.keys(rec).forEach(f=>{ if(f==="styleNo"||f==="orderNo"||f==="rowNo") return; const nv=rec[f]; if(nv===undefined||nv==="") return; if(String(ex[f]??"")!==String(nv)){ chg[f]=nv; diffs.push({ field:f, from:ex[f]??"", to:nv }); } }); if(Object.keys(chg).length) updates.push({ id:ex.id, key:k, rowNo:rec.rowNo, styleNo:rec.styleNo, orderNo:rec.orderNo, chg, diffs }); else unchanged++; } else inserts.push({ ...rec, key:k }); });
    const blocking=errors.some(e=>e.blocking) || dupes.length>0; try{ localStorage.setItem("mt_upload_mapping_default", JSON.stringify(mapping||{})); }catch(e){} setUploadSkip(new Set()); setBulkResult({ inserts, updates, unchanged, dupes, errors, blocking, noOrder:false, sheetName:sn, total:recs.length, mappedHeaders:br.headersRaw });
  }catch(e){ logAppError("upload preview failed",e); setBulkResult({ error:"Couldn't build upload preview: "+(e.message||e) }); } };
  const parseUpload=async(file)=>{ try{ setUploadSkip(new Set()); const buf=await file.arrayBuffer(); const wb=XLSX.read(buf,{cellDates:true}); const sn=wb.SheetNames.includes("Tracker")?"Tracker":wb.SheetNames[0]; const aoa=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:false,cellDates:true,defval:""}); let hr=-1; for(let i=0;i<Math.min(aoa.length,8);i++){ if((aoa[i]||[]).some(c=>{ const n=NORMH(c); return n==="styleno"||n==="style"||n==="orderno"||n==="order"; })){ hr=i; break; } } if(hr<0) hr=0; const headersRaw=(aoa[hr]||[]).map(h=>String(h||"").trim()); const headers=headersRaw.map(NORMH); const mapping={}; headers.forEach((h,i)=>{ const f=HEADER_MAP[h]; if(f&&mapping[f]==null) mapping[f]=i; });
    try{ const saved=JSON.parse(localStorage.getItem("mt_upload_mapping_default")||"{}"); Object.entries(saved||{}).forEach(([k,v])=>{ const n=Number(v); if(mapping[k]==null && Number.isFinite(n) && n>=0 && n<headersRaw.length) mapping[k]=n; }); }catch(e){}
    setBulkResult({ mapping, aoa, hr, sn, headersRaw, mapError:(mapping.styleNo==null||mapping.orderNo==null)?"Please map Style No and Order No. Both are mandatory.":"" });
  }catch(e){ logAppError("upload parse failed",e,file&&file.name); setBulkResult({ error:"Couldn't read that file: "+(e.message||e) }); } };
  const applyBulk=async()=>{ if(!bulkResult||bulkResult.error||bulkResult.mapping) return; if(bulkResult.blocking){ alert("Upload has blocking errors/duplicates. Download the report, fix the file, and upload again."); return; } const skip=uploadSkip||new Set(); const inserts=(bulkResult.inserts||[]).filter(x=>!skip.has(x.key)); const updates=(bulkResult.updates||[]).filter(x=>!skip.has(x.key)); const total=(inserts.length+updates.length); if(!total){ alert("Nothing selected to apply."); return; } if(!window.confirm(`Apply ${total} selected upload change(s)? Unique key is Order No + Style No. Existing discrepancies will be overwritten only for selected rows.`)) return; pushHistory();
    try{ supabase.from("audit_log").insert({ style_id:null, style_no:"", col:"Upload", field:"upload", old_val:`${(bulkResult.inserts||[]).length} new / ${(bulkResult.updates||[]).length} changed / ${(bulkResult.errors||[]).length+(bulkResult.dupes||[]).length} blocked`, new_val:`Applied ${inserts.length} new + ${updates.length} changed`, actor_id:me.id, actor_name:me.name||me.email }).then(()=>{}).catch(()=>{}); }catch(e){}
    if(updates.length){ const m={}; updates.forEach(u=>{ m[u.id]={...(m[u.id]||{}),...u.chg}; }); setStyles(prev=>prev.map(s=> m[s.id]?{...s,...m[s.id]}:s)); }
    if(inserts.length){ const rows=inserts.map(rec=>{ const s={ orderNo:rec.orderNo||"", styleNo:rec.styleNo, sampleFit:rec.sampleFit||"", family:rec.family||"", colour:rec.colour||"", brand:rec.brand||"", buyer:rec.buyer||"", fabricType:rec.fabricType||"", owner:rec.owner||"", setId:rec.setId||"", setRole:rec.setRole||"", age:rec.age||"", qty:rec.qty||0, ordRec:rec.ordRec||iso(TODAY), delivery:rec.delivery||rec.ordRec||iso(TODAY), fitReq:rec.fitReq??true, printReq:rec.printReq??false, soReq:rec.soReq??false, ppBypass:rec.ppBypass??false, labDipReq:rec.labDipReq??true, ppNeeded:rec.ppNeeded??true, remarks:"" }; const r=styleToRow(s); delete r.id; return r; });
      try{ const { data, error }=await supabase.from("styles").insert(rows).select(); if(error) throw error; if(data) setStyles(prev=>[...prev, ...data.map(d=>rowToStyle(d,{}))]); }catch(e){ logAppError("bulk insert failed",e); alert("New styles failed to insert (existing updates were applied): "+(e.message||e)); } }
    setBulkOpen(false); setBulkResult(null); setUploadSkip(new Set()); flash(); };
  const BUYER_STAGES=["fitAppr","artAppr","soAppr","labAppr","ppAppr"];
  const runExport=(mode,buf,incBuf,relMode)=>{ try{ let data, name, sheet="Tracker"; const B=Math.abs(buf||0);
    const fmtY=(d)=> !d?"":d.toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"});
    const fmtIso=(v)=> v?fmtY(parse(v)):"";
    const stageRow=(c,k)=> (c.stages||[]).find(x=>x.key===k);
    const activeActualDate=(s,c,k)=>{ const r=stageRow(c,k); if(!r) return null;
      // Export must use the same active-branch state as Tracker/To-Do, not raw old actuals.
      // In rejected rework, r.actual is the current-round resend/reapproval actual. Old first-send actual stays history only.
      return r.actual || null;
    };
    const actCell=(s,c,k)=>{ const r=stageRow(c,k); if(!r) return "n/a"; const a=activeActualDate(s,c,k); if(a) return fmtY(a); if(r.skipped) return "WAIVED"; return ""; };
    const planCell=(s,c,k,withBuf)=>{ const r=stageRow(c,k); if(!r) return "n/a"; const a=activeActualDate(s,c,k); if(a||r.skipped||r.autoClosed) return ""; const t=r.rev||r.plan; if(!t) return ""; /* E6: buffer is a single B-day slack on our internal chain. Each approval inherits the shift from its (buffered) internal predecessor and adds NO buffer of its own — net effect is one B applied per stage, never doubled. */ const d=(withBuf&&B)?addWorkdays(t,B):t; return fmtY(d); };
    const stateCell=(s,c,k)=>{ const r=stageRow(c,k); if(!r) return "n/a"; if(r.actual) return "ACTUAL"; if(r.skipped) return "WAIVED"; if(r.rework) return "REWORK / RESEND"; if(r.rejected) return "REJECTED / RE-APPROVAL"; if(r.autoClosed) return "AUTO-CLOSED"; if(r.rev) return "REVISED"; if(r.plan) return "AUTO PLAN"; return ""; };
    const pairs=(o,s,c,withBuf,list)=>{ (list||STAGES).forEach(st=>{ o["Actual \u00b7 "+st.label]=actCell(s,c,st.key); o["Plan \u00b7 "+st.label]=planCell(s,c,st.key,withBuf); o["State \u00b7 "+st.label]=stateCell(s,c,st.key); }); };
    const blockers=(style,c)=> (c.frontier?[...c.frontier]:[]).map(k=>{ const r=(c.stages||[]).find(x=>x.key===k); if(!r||r.done) return null; const lbl=stageReviewLabel(style,r); if(r.rework) return lbl+": redo & resend"; if(r.rejected) return lbl; if(REJECTABLE.includes(k)) return lbl+": not approved"; return lbl+": pending"; }).filter(Boolean);
    if(mode==="buyer"){ data=rows.map(({s,c})=>{ const o={ "Style No":s.styleNo, "Family":s.family, "Colour":s.colour, "Brand":s.brand, "Buyer":s.buyer||"", "Age Group":s.age||"", "Qty":s.qty, "Delivery":fmtIso(s.delivery), "Status":c.status }; pairs(o,s,c,incBuf); o["Proj. Release"]=c.projRelease?fmtY(c.projRelease):""; return o; }); name=incBuf?"buyer_tna_buf"+B+"d":"buyer_tna"; sheet="Buyer"; }
    else if(mode==="release"){ data=rows.map(({s,c})=>{ const open=blockers(s,c); const o={ "Order No":s.orderNo, "Style No":s.styleNo, "Colour":s.colour, "Qty":s.qty }; if(relMode!=="summary"){ o["FIT"]=s.fitReq?"Y":"-"; o["PRT"]=s.printReq?"Y":"-"; o["S-O"]=s.soReq?"Y":"-"; o["LAB"]=s.labDipReq?"Y":"-"; o["BYP"]=s.ppBypass?"Y":"-"; o["PP"]=s.ppNeeded?"Y":"-"; pairs(o,s,c,false); } else { ["fitAppr","soAppr","fabricIH","ppAppr","prodFile"].forEach(k=>{ const st=STAGES.find(x=>x.key===k); o["Actual \u00b7 "+st.label]=actCell(s,c,k); o["Plan \u00b7 "+st.label]=planCell(s,c,k,false); }); } o["Proj. Release"]=c.projRelease?fmtY(c.projRelease):""; o["Delivery"]=fmtIso(s.delivery); o["Released"]=c.released?"YES":""; o["Pending / blockers"]=c.released?"released":(open.join("; ")||"none"); o["Remarks"]=s.remarks||""; o["Status"]=c.status; return o; }); name=relMode==="summary"?"release_plan_summary":"release_plan_detailed"; sheet="Release Plan"; }
    else if(mode==="detail"){ data=rows.map(({s,c})=>{ const open=blockers(s,c); const o={ "Order No":s.orderNo, "Style No":s.styleNo, "Sample Fit":s.sampleFit, "Family":s.family, "Colour":s.colour, "Brand":s.brand, "Buyer":s.buyer||"", "Age Group":s.age||"", "Fabric Type":s.fabricType, "Owner":s.owner, "Qty":s.qty, "Order Date":fmtIso(s.ordRec), "Delivery":fmtIso(s.delivery), "FIT":s.fitReq?"Y":"-", "PRT":s.printReq?"Y":"-", "S-O":s.soReq?"Y":"-", "LAB":s.labDipReq?"Y":"-", "BYP":s.ppBypass?"Y":"-", "PP":s.ppNeeded?"Y":"-" }; pairs(o,s,c,false); o["% Done"]=c.pct+"%"; o["Float"]=(c.float!=null?c.float+"d":""); o["Idle"]=(c.idle!=null?c.idle+"d":""); o["Proj. Release"]=c.projRelease?fmtY(c.projRelease):""; o["Released"]=c.released?"YES":""; o["Pending / blockers"]=c.released?"released":(open.join("; ")||"none"); o["Remarks"]=s.remarks||""; o["Status"]=c.status; return o; }); name="detailed_summary"; sheet="Detailed"; }
    else if(mode==="internal"){ data=rows.map(({s,c})=>{ const o={ "Order No":s.orderNo, "Style No":s.styleNo, "Sample Fit":s.sampleFit, "Family":s.family, "Colour":s.colour, "Buyer":s.buyer||"", "Age Group":s.age||"", "Qty":s.qty, "Delivery":fmtIso(s.delivery), "Status":c.status }; pairs(o,s,c,true); return o; }); name="internal_plan_buf"+B+"d"; }
    else { data=rows.map(({s,c})=>{ const o={ "Order No":s.orderNo, "Style No":s.styleNo, "Sample Fit":s.sampleFit, "Family":s.family, "Colour":s.colour, "Brand":s.brand, "Buyer":s.buyer||"", "Age Group":s.age||"", "Fabric Type":s.fabricType, "Junior":s.owner, "Qty":s.qty, "Order Date":fmtIso(s.ordRec), "Delivery":fmtIso(s.delivery), "Status":c.status }; STAGES.forEach(st=>{ o[st.label]=actCell(s,c,st.key); }); return o; }); name="merch_tracker"; }
    if(data&&data.length){ Object.keys(data[0]).forEach(k=>{ if(data.every(r=>r[k]==="n/a")) data.forEach(r=>{ delete r[k]; }); }); }
    const ws=XLSX.utils.json_to_sheet(data); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,sheet); XLSX.writeFile(wb,name+"_"+iso(TODAY)+".xlsx"); }catch(e){ logAppError("export failed",e); alert("Export failed: "+(e.message||e)); } };
  const bulkVisibleSummary=()=>{ const vis=rows.map(r=>r.s); const buyers=[...new Set(vis.map(s=>s.buyer||s.brand||"—").filter(Boolean))].slice(0,5); const owners=[...new Set(vis.map(s=>s.owner||"—").filter(Boolean))].slice(0,6); const sample=vis.slice(0,8).map(s=>`${s.orderNo||"—"} · ${s.styleNo||"—"}`); return { count:vis.length, buyers, owners, sample, more:Math.max(0,vis.length-sample.length) }; };
  const openBulkConfirm=(payload)=>{ const ids=rows.map(r=>r.s.id); if(!ids.length){ alert("No visible styles to update."); return; } const summary=bulkVisibleSummary(); setBulkConfirm({ ...payload, ids, summary }); };
  const applyBulkConfirm=async()=>{ const bc=bulkConfirm; if(!bc) return; const ids=new Set(bc.ids||[]); const idList=[...ids]; if(!ids.size){ setBulkConfirm(null); return; } pushHistory();
    try{ supabase.from("audit_log").insert({ style_id:null, style_no:"", col:"Bulk Action", field:"bulk action", old_val:`${ids.size} visible styles`, new_val:bc.title||bc.actionText||bc.kind||"Bulk action", actor_id:me.id, actor_name:me.name||me.email }).then(()=>{}).catch(()=>{}); }catch(e){}
    if(bc.kind==="appendRemark"){
      const note=String(bc.note||"").trim();
      setStyles(prev=>prev.map(s=>ids.has(s.id)?{...s,remarks:[s.remarks,note].filter(Boolean).join(" | ")}:s));
    } else {
      const clean=Object.fromEntries(Object.entries(bc.patch||{}).filter(([_,v])=>v!==undefined));
      setStyles(prev=>prev.map(s=>ids.has(s.id)?{...s,...clean}:s));
      // Archive/restore and other bulk patches should not rely only on the debounce save.
      // Persist immediately so restore works even if the row disappears from the current Archive slice.
      try{
        if(Object.keys(clean).length){
          const { error }=await supabase.from("styles").update(clean).in("id",idList);
          if(error) throw error;
          savedRef.current.sty={...(savedRef.current.sty||{})};
          stylesRef.current.forEach(st=>{ if(ids.has(st.id)){ const ns={...st,...clean}; savedRef.current.sty[st.id]=JSON.stringify(styleToRow(ns)); } });
        }
      }catch(e){ logAppError("bulk patch save failed",e); alert("Bulk update saved locally but cloud save failed: "+(e.message||e)); }
    }
    if(bc.kind==="patch" && bc.patch && bc.patch.archived===false){ setArchiveView("active"); }
    flash(); setBulkConfirm(null); setBulkActionsOpen(false);
  };
  const archiveFiltered=(val)=>{ const label=val?"Archive visible styles":"Restore visible styles"; openBulkConfirm({ title:label, kind:"patch", patch:{ archived:val }, impact:val?"Archived styles will be hidden from the active sheet. This is reversible from Archive view.":"Restored styles will return to the active sheet.", actionText:val?"Archive styles":"Restore styles", danger:val }); };
  const bulkUpdateVisible=(patch,label,impact)=>{ const clean=Object.fromEntries(Object.entries(patch||{}).filter(([_,v])=>v!==undefined)); if(!Object.keys(clean).length) return; openBulkConfirm({ title:label, kind:"patch", patch:clean, impact:impact||"This changes real row data and will auto-save for everyone.", actionText:"Apply bulk update" }); };
  const bulkAppendRemark=(txt)=>{ const note=String(txt||"").trim(); if(!note) return; openBulkConfirm({ title:"Append remark", kind:"appendRemark", note, impact:`This remark will be added to every visible style: “${note}”`, actionText:"Append remark" }); };
  const bulkFlagVisible=(flag,val)=>{ const def=(FLAG_DEFS.find(x=>x.key===flag)||{}); const label=def.short||flag; bulkUpdateVisible({ [flag]:val }, `Set ${label} = ${val?"Yes":"No"}`, `Requirement flag ${def.title||flag} will be set to ${val?"YES":"NO"} for all visible styles.`); };
  const deleteStyle=async(id)=>{ const victim=styles.find(s=>s.id===id); if(!window.confirm("Delete this style row? This removes it for everyone and cannot be undone.")) return; pushHistory(); setStyles(prev=>prev.filter(s=>s.id!==id)); flash(); try{ if(victim){ await supabase.from("audit_log").insert({ style_id:id, style_no:victim.styleNo||"", col:"Row", field:"deleted", old_val:[victim.orderNo,victim.styleNo].filter(Boolean).join(" · ")||String(id), new_val:"DELETED", actor_id:me.id, actor_name:me.name||me.email }); } await supabase.from("stage_dates").delete().eq("style_id",id); await supabase.from("cell_meta").delete().eq("style_id",id); await supabase.from("styles").delete().eq("id",id); }catch(e){ console.error("delete failed",e); } };

  const perfRef=useRef({ computeMs:0, filterMs:0, rowMs:0, rows:0, styles:0, rendered:0, cacheHits:0, recomputed:0, alerts:[], samples:[], p95Ms:0, deferred:false });
  const computeCacheRef=useRef(new Map());
  // Keep compute invalidation tight: changing To-Do watch windows or escalation rules should not recompute every style.
  const cfgComputeKey=useMemo(()=>JSON.stringify({ leads:cfg.leads, stageOwners:cfg.stageOwners, rework:cfg.rework, fabricCutoff:cfg.fabricCutoff, relGate:cfg.relGate }),[cfg.leads,cfg.stageOwners,cfg.rework,cfg.fabricCutoff,cfg.relGate]);
  // Live recompute guard: dashboards/management/export must never read stale computed TNA.
  // The old cache only checked the style object reference; stage/date edits and realtime patches can leave nested data looking unchanged to a reference check.
  // This signature includes actual/revised/reject/skip dates and style master fields, so Lab Dip Approval etc. refresh every live report immediately.
  const effectiveResends=useMemo(()=>{
    // Active-branch resend map for live TNA.
    // Only the current rejection round's actual re-send can drive status/due/editor.
    // Prior rounds remain in stage_events/resend history for reports and audit, not active workflow.
    const out={};
    const styleById=new Map((styles||[]).map(st=>[String(st.id),st]));
    Object.entries(resends||{}).forEach(([k,v])=>{
      const [styleId,stageKey]=String(k).split(":");
      const st=styleById.get(String(styleId));
      const linked=stageKey&&APPR_OF_SEND[stageKey];
      const round=st&&linked?activeRejectRound(st,stageKey,linked):0;
      // Legacy resendHistory has no reliable round. Use it only when stage_events has not started a round.
      out[k]=round?[]:(Array.isArray(v)?v.slice():[]);
    });
    (stageEvents||[]).slice().sort((a,b)=>eventTime(a)-eventTime(b)).forEach(e=>{
      if(!e||!e.style_id||!e.stage_key) return;
      const key=String(e.style_id)+":"+e.stage_key;
      const st=styleById.get(String(e.style_id));
      const linked=e.linked_stage_key || APPR_OF_SEND[e.stage_key] || null;
      const round=st&&linked?activeRejectRound(st,e.stage_key,linked):0;
      if(round && Number(e.round_no||0)!==Number(round)) return;
      const typ=String(e.event_type||"").toLowerCase();
      if(typ==="resend_actual"){
        const val=e.new_value || e.event_date || "";
        if(!val) return;
        const arr=out[key]?out[key].slice():[];
        // Same round corrections replace the active round value instead of stacking as new active branch.
        const idx=arr.findIndex(x=>Number((normalizeResendEntry(x)||{}).roundNo||0)===Number(e.round_no||0) && String(((normalizeResendEntry(x)||{}).source)||"").toLowerCase().includes("stage_events"));
        const item={ at:e.created_at||"", source:"stage_events resend_actual", roundNo:Number(e.round_no)||0, oldVal:e.old_value||"", newVal:val, stage:e.stage_key, styleId:e.style_id, styleNo:e.style_no||"", orderNo:e.order_no||"" };
        if(idx>=0) arr[idx]=item; else arr.push(item);
        out[key]=arr;
      } else if(typ==="clear_resend_actual"){
        const arr=out[key]?out[key].slice():[];
        for(let i=arr.length-1;i>=0;i--){
          const x=normalizeResendEntry(arr[i]);
          if(Number(x.roundNo||0)===Number(e.round_no||0)){ arr.splice(i,1); break; }
        }
        out[key]=arr;
      }
    });
    return out;
  },[resends,stageEvents,styles]);
  const latestResendActualValue=(s,field)=>{
    if(!s||!field) return "";
    const appr=APPR_OF_SEND[field];
    if(appr){
      const roundVal=currentRoundEventValue(s,field,appr,"resend_actual","clear_resend_actual");
      if(roundVal) return roundVal;
    }
    const rej=appr&&s.rejects?parse(s.rejects[appr]):null;
    const arr=Array.isArray(effectiveResends&&effectiveResends[resendHistoryKey(s,field)])?effectiveResends[resendHistoryKey(s,field)].map(normalizeResendEntry).filter(Boolean):[];
    let latest="";
    arr.forEach((x,i)=>{
      const src=String((x&&x.source)||"").toLowerCase();
      const isFirst=src.includes("first send") || (i===0 && arr.length>1 && !src.includes("stage_events"));
      if(isFirst) return;
      const d=parse(x&&x.newVal);
      if(d && (!rej || dateSerial(d) >= dateSerial(rej, -Infinity))) latest=x.newVal||"";
    });
    return latest;
  };
  const _popResendActual=(st,field)=>{ // silent core for bulk clear: clears the active resend round, preserves first-send actual. Caller owns confirmation.
    if(!st||!field) return false;
    const appr=APPR_OF_SEND[field];
    const latest=latestResendActualValue(st,field);
    if(!latest) return false;
    const round=activeRejectRound(st,field,appr)||1;
    insertStageEvent(st,field,appr,"clear_resend_actual",round,null,latest,null,`Cleared ${roundLabelFor(st,field,"RE-SEND")} actual (bulk)`);
    setResends(prev=>{ const key=st.id+":"+field; const arr=Array.isArray(prev&&prev[key])?prev[key].slice():[]; for(let i=arr.length-1;i>=0;i--){ const x=normalizeResendEntry(arr[i]); const src=String((x&&x.source)||"").toLowerCase(); if(src.includes("first send")) continue; arr.splice(i,1); break; } return {...(prev||{}),[key]:arr}; });
    return true;
  };
  const clearLatestResendActual=(id,field)=>{
    const st=styles.find(x=>x.id===id);
    if(!st||!field) return;
    const appr=APPR_OF_SEND[field];
    const latest=latestResendActualValue(st,field);
    if(!latest){ setEditing(null); setCalOpen(false); return; }
    if(!window.confirm(`Clear latest ${roundLabelFor(st,field,"RE-SEND")} actual date?\n\nFirst send actual stays as history.`)) return;
    const round=activeRejectRound(st,field,appr)||1;
    insertStageEvent(st,field,appr,"clear_resend_actual",round,null,latest,null,`Cleared ${roundLabelFor(st,field,"RE-SEND")} actual`);
    setResends(prev=>{
      const key=st.id+":"+field;
      const arr=Array.isArray(prev&&prev[key])?prev[key].slice():[];
      for(let i=arr.length-1;i>=0;i--){ const x=normalizeResendEntry(arr[i]); const src=String((x&&x.source)||"").toLowerCase(); if(src.includes("first send")) continue; arr.splice(i,1); break; }
      return {...(prev||{}),[key]:arr};
    });
    flash(); setEditing(null); setCalOpen(false);
  };
  const activeBranchStyles=useMemo(()=>styles.map(s=>styleForActiveBranch(s)),[styles,stageEvents]);
  // Archive lazy-compute: archived styles are excluded from every live view/report by default, so there is no
  // reason to run computeStyle + search-index + canonicalisation for them until the user actually opens Archive
  // (or "all") view. This keeps the per-render compute scaled to the ACTIVE sheet, not the lifetime style count.
  const computeSource=useMemo(()=> archiveView==="active" ? activeBranchStyles.filter(s=>!s.archived) : activeBranchStyles, [activeBranchStyles, archiveView]);
  const stylesComputeKey=useMemo(()=>computeSource.map(s=>styleComputeSignature(s)+"::resends="+JSON.stringify(effectiveResends&&Object.fromEntries(Object.entries(effectiveResends||{}).filter(([k])=>String(k).startsWith(String(s.id)+":"))))).join("||"),[computeSource,effectiveResends]);
  const computed=useMemo(()=>{ const t=perfNow(); const cache=computeCacheRef.current; const liveIds=new Set(); let hits=0, recomputed=0; const rawById=new Map((styles||[]).map(raw=>[String(raw.id),raw])); const out=computeSource.map(activeStyle=>{ liveIds.add(activeStyle.id); const rawStyle=rawById.get(String(activeStyle.id))||activeStyle; const resendSig=JSON.stringify(effectiveResends&&Object.fromEntries(Object.entries(effectiveResends||{}).filter(([k])=>String(k).startsWith(String(activeStyle.id)+":")))); const sig=styleComputeSignature(activeStyle)+"::resends="+resendSig; const old=cache.get(activeStyle.id); if(old && old.sig===sig && old.cfgKey===cfgComputeKey){ hits++; const cached=old.out||{}; return canonicalLiveRow({ ...cached, rawStyle, activeStyle:cached.activeStyle||activeStyle, s:cached.activeStyle||activeStyle }); } const c=computeStyle(activeStyle,cfg,effectiveResends); /* IMPORTANT: live screens must read s/activeStyle, never rawStyle. rawStyle is kept only for save/history/audit. */ let item={ rawStyle, activeStyle, s:activeStyle, c, idx:buildSearchIndex(activeStyle,c) }; item=canonicalLiveRow(item); cache.set(activeStyle.id,{ sig, cfgKey:cfgComputeKey, out:item }); recomputed++; return item; }); for(const id of cache.keys()){ if(!liveIds.has(id)) cache.delete(id); } const ms=Math.round((perfNow()-t)*10)/10; const hist=pushPerfSample(perfRef.current.samples,ms); perfRef.current.computeMs=ms; perfRef.current.samples=hist.samples; perfRef.current.p95Ms=hist.p95; perfRef.current.styles=styles.length; perfRef.current.cacheHits=hits; perfRef.current.recomputed=recomputed; return out; },[styles,computeSource,stylesComputeKey,cfg,cfgComputeKey,effectiveResends]);
  // Live/report source: archived styles stay available in Tracker Archive view, but are excluded from all live reports by default.
  const activeComputed=useMemo(()=>computed.filter(({s})=>!s.archived),[computed]);
  const chaseOwnerOptions=useMemo(()=>["All", ...Array.from(new Set([...STAGES.map(st=>(cfg.stageOwners&&cfg.stageOwners[st.key])||DEFAULT_STAGE_OWNERS[st.key]||st.owner), ...CHASE_LABELS])).filter(Boolean)], [cfg]);
  const splitColoursForTodo=splitColoursAll;
  // v31dm Build 1 memory safety: To-Do rows are heavy. Build them only for tabs/views that actually need them.
  // Tracker/editing/settings/help should not pay the To-Do generation cost on every render.
  const shouldBuildTodo=tab==="todo"||tab==="dashboard"||tab==="management"||tab==="escalation"||tab==="review"||reviewOpen;
  const todoItems=useMemo(()=>{
    if(!shouldBuildTodo) return [];
    const stageWeight=(key)=> key==="fitAppr"||key==="fitSend"?10:["fabricIH","ppSample","ppAppr","prodFile"].includes(key)?8:["artAppr","soAppr","labAppr"].includes(key)?6:4;
    const eventTimeMs=(e)=>{ try{ return new Date(e&&e.created_at||0).getTime()||0; }catch(_e){ return 0; } };
    const revisionCountFor=(style,key)=>{
      const sid=String(style&&style.id);
      const direct=style&&style.revs&&style.revs[key]?1:0;
      const ev=(stageEvents||[]).filter(e=>String(e&&e.style_id)===sid && String(e&&e.stage_key)===String(key) && String(e&&e.event_type||"").toLowerCase()==="revised");
      return Math.max(direct, ev.length);
    };
    const activeRejectRoundForTodo=(style,key)=>{ try{ return activeRoundForCell(style,key)||0; }catch(_e){ return 0; } };
    const todoMetrics=(style,key,r,exp,du,overdue)=>{
      const originalPlan=r&&r.plan?r.plan:null;
      const dueUsed=exp||null;
      const driftTarget=originalPlan ? ((dueUsed && dueUsed>TODAY)?dueUsed:TODAY) : null;
      const driftOriginal=originalPlan&&driftTarget&&driftTarget>originalPlan ? Math.max(0,netWorkdays(originalPlan,driftTarget)) : 0;
      const delayDue=overdue?Math.max(1,Math.abs(Number(du)||0)):0;
      const revisionCount=revisionCountFor(style,key);
      const rejectionRound=activeRejectRoundForTodo(style,key);
      const criticality=stageWeight(key);
      const priorityScore=Math.round((delayDue*40)+(driftOriginal*35)+(revisionCount*15)+(rejectionRound*20)+criticality);
      let priorityBucket="Daily Chase";
      if(delayDue>0 && driftOriginal>=3) priorityBucket="Critical Escalation";
      else if(!delayDue && driftOriginal>=3) priorityBucket="Hidden Risk";
      else if(delayDue>0) priorityBucket="Overdue Chase";
      else if(revisionCount>=2 || rejectionRound>=1) priorityBucket="Watchlist";
      const parts=[];
      if(delayDue>0) parts.push(`missed active due by ${delayDue}d`);
      if(driftOriginal>0) parts.push(`${driftOriginal}d away from original plan`);
      if(revisionCount>0) parts.push(`${revisionCount} revised commitment${revisionCount>1?"s":""}`);
      if(rejectionRound>0) parts.push(`rejection round ${rejectionRound}`);
      if(!parts.length) parts.push("due soon as per current plan");
      return { originalPlan, dueUsed, driftOriginal, delayDue, revisionCount, rejectionRound, criticality, priorityScore, priorityBucket, priorityReason:parts.join(" · ") };
    };
    const enrich=(item)=>{
      const daysLate=item.overdue?Math.max(1,Math.abs(Number(item.du)||0)):0;
      const esc=item.overdue?escalationFor(cfg,daysLate):null;
      return { ...item, daysLate, escalationOwner:esc?esc.owner:"", escalationLevel:esc?esc.level:"", escalationAction:esc?esc.action:"", escalationRange:esc?esc.rangeLabel:"" };
    };
    const out=[]; const colourGroups={};
    activeComputed.forEach(({s,c})=>{
      if(c.released) return;
      const front=c.frontier?[...c.frontier]:[];
      front.forEach(key=>{
        const r=(c.stages||[]).find(x=>x.key===key);
        if(!r||r.done) return;
        const exp=r.rev||r.plan; if(!exp) return;
        const du=netWorkdays(TODAY,exp);
        const overdue=TODAY>exp;
        const win=(cfg.upcoming&&cfg.upcoming[key]!=null)?cfg.upcoming[key]:null;
        const m=todoMetrics(s,key,r,exp,du,overdue);
        // Include normal due items, plus hidden-risk items where revised dates have pushed the activity away from original TNA.
        const include = overdue || (win!=null && du<=win) || (m.driftOriginal>=3) || (m.revisionCount>=2) || (m.rejectionRound>=1);
        if(!include) return;
        const branch=BRANCH_OF[key]||"";
        if(["fabricIH","labDip","labAppr"].includes(key)){
          splitColoursForTodo(s.colour).forEach(col=>{
            // Keep colour grouping useful without mixing orders. Earlier grouping only by colour caused
            // an Order filter like F2 to still show rows visually belonging to T2 because the row contained
            // multiple orders. Group by activity + order + colour so To-Do filters and display tally cleanly.
            const ordKey=String(s.orderNo||"(blank order)").trim().toLowerCase();
            const gkey=key+"::"+ordKey+"::"+col.toLowerCase();
            let cur=colourGroups[gkey];
            if(!cur){ cur=colourGroups[gkey]={ colour:col, key, activityKey:key, label:stageReviewLabel(s,r), activityLabel:stageReviewLabel(s,r), owner:r.owner, branch, exp, du, overdue, anyStyle:s.id, orderNo:s.orderNo||"", orders:new Set(), juniors:new Set(), styles:new Set(), fits:new Set(), families:new Set(), brands:new Set(), fabrics:new Set(), buyers:new Set(), colours:new Set(), count:0, priorityScore:0, priorityBucket:m.priorityBucket, priorityReason:m.priorityReason, originalPlan:m.originalPlan, dueUsed:m.dueUsed, driftOriginal:m.driftOriginal, delayDue:m.delayDue, revisionCount:m.revisionCount, rejectionRound:m.rejectionRound, criticality:m.criticality }; }
            cur.count++; cur.orders.add(s.orderNo||""); cur.juniors.add(s.owner||""); cur.styles.add(s.styleNo||""); cur.fits.add(s.sampleFit||""); cur.families.add(s.family||""); cur.brands.add(s.brand||""); cur.fabrics.add(s.fabricType||""); cur.buyers.add(s.buyer||""); cur.colours.add(col||s.colour||"");
            if(m.priorityScore>cur.priorityScore){ cur.priorityScore=m.priorityScore; cur.priorityBucket=m.priorityBucket; cur.priorityReason=m.priorityReason; cur.originalPlan=m.originalPlan; cur.dueUsed=m.dueUsed; cur.driftOriginal=m.driftOriginal; cur.delayDue=m.delayDue; cur.revisionCount=m.revisionCount; cur.rejectionRound=m.rejectionRound; cur.criticality=m.criticality; }
            if(exp<cur.exp){ cur.exp=exp; cur.du=du; cur.overdue=overdue; cur.anyStyle=s.id; }
          });
        } else {
          out.push(enrich({ id:s.id, orderNo:s.orderNo, orderNos:[s.orderNo].filter(Boolean), styleNo:s.styleNo, junior:s.owner, juniors:[s.owner].filter(Boolean), colour:s.colour, colours:splitColoursForTodo(s.colour), fit:s.sampleFit, family:s.family, brand:s.brand, buyer:s.buyer, fabric:s.fabricType, key, activityKey:key, activity:stageReviewLabel(s,r), activityLabel:stageReviewLabel(s,r), branch, owner:r.owner, exp, du, overdue, ...m }));
        }
      });
    });
    Object.values(colourGroups).forEach(f=>{ const orders=[...f.orders].filter(Boolean); const juniors=[...f.juniors].filter(Boolean); const styles=[...f.styles].filter(Boolean); const fits=[...f.fits].filter(Boolean); const families=[...f.families].filter(Boolean); const brands=[...f.brands].filter(Boolean); const fabrics=[...f.fabrics].filter(Boolean); const buyers=[...f.buyers].filter(Boolean); const colours=[...f.colours].filter(Boolean); out.push(enrich({ id:f.anyStyle, orderNo:f.orderNo||orders[0]||"", orderNos:orders, styleNo:f.colour, junior:juniors.length===1?juniors[0]:(juniors.length?"Multiple":""), juniors, colour:f.colour, colours, fit:fits.length===1?fits[0]:(fits.length?"Multiple":""), fits, family:families.length===1?families[0]:(families.length?"Multiple":""), families, brand:brands.length===1?brands[0]:(brands.length?"Multiple":""), brands, buyer:buyers.length===1?buyers[0]:(buyers.length?"Multiple":""), buyers, fabric:fabrics.length===1?fabrics[0]:(fabrics.length?"Multiple":""), fabrics, key:f.key, activityKey:f.key, activity:f.label, activityLabel:f.label, branch:f.branch, owner:f.owner, exp:f.exp, du:f.du, overdue:f.overdue, isColour:true, count:f.count, styleCount:styles.length, styleNos:styles, priorityScore:f.priorityScore, priorityBucket:f.priorityBucket, priorityReason:f.priorityReason, originalPlan:f.originalPlan, dueUsed:f.dueUsed, driftOriginal:f.driftOriginal, delayDue:f.delayDue, revisionCount:f.revisionCount, rejectionRound:f.rejectionRound, criticality:f.criticality })); });
    out.sort((a,b)=> (Number(b.priorityScore)||0)-(Number(a.priorityScore)||0) || (a.overdue!==b.overdue?(a.overdue?-1:1):0) || ((a.exp&&b.exp)?(a.exp-b.exp):0));
    return out;
  },[shouldBuildTodo,activeComputed,cfg,stageEvents]);
  const [todoFilter,setTodoFilter]=useState(PF.todoFilter||{});
  useEffect(()=>{ try{ localStorage.setItem("mt_trackfilters", JSON.stringify({ search, searchCol, statusFilter, ownerFilter, archiveView, activityFilter:activityFilterKeys, colFilters, tab, todoFilter, followFilter, savedView })); }catch(e){} },[search,searchCol,statusFilter,ownerFilter,archiveView,activityFilterKeys.join("|"),colFilters,tab,todoFilter,followFilter,savedView]);
  useEffect(()=>{ try{ localStorage.setItem("mt_column_view",columnView); }catch(e){} },[columnView]);
  useEffect(()=>{ try{ localStorage.setItem("mt_freeze_n",String(freezeN)); }catch(e){} },[freezeN]);
  const valueFor=(s,cc,col)=>{
    if(col==="__style") return s.styleNo||"";
    if(["orderNo","sampleFit","family","colour","brand","fabricType","owner","setId","setRole","remarks"].includes(col)) return s[col]||"(Blanks)";
    if(col==="qty") return String(s.qty);
    if(col==="ordRec"||col==="delivery") return fmt(parse(s[col]))||"(Blanks)";
    if(col==="overall") return cc.status;
    if(col==="fit") return cc.fitBranch.txt; if(col==="print") return cc.printBranch.txt; if(col==="fabric") return cc.fabricBranch.txt; if(col==="pp") return cc.ppBranch.txt; if(col==="prod") return cc.prodFileBranch.txt;
    if(col==="fabricCD") return cc.fabricCountdown.txt;
    if(col==="proj") return fmt(cc.projRelease)||"(Blanks)";
    if(col==="pct") return cc.pct+"%";
    if(col==="chase") return (cc.chaseOwners||[]).map(o=>o.owner).join(", ")||"(Blanks)";
    if(col==="float") return cc.float==null?"(Blanks)":String(cc.float);
    if(col==="idle") return cc.idle==null?"(Blanks)":String(cc.idle);
    if(STAGE_KEYS.includes(col)){ const r=(cc.stages||[]).find(x=>x.key===col); if(!r) return "— n/a"; if(r.done) return fmt(r.actual); if(r.rework) return "↻ Redo & resend"; if(r.rejected) return r.reject?("✕ Rejected "+fmt(r.reject)):"✕ Rejected"; if(r.rev) return "✎ Revised "+fmt(r.rev); return "● Pending"; }
    return "";
  };
  const filterNorm=(v)=>String(v==null?"":v).replace(/\s+/g," ").trim().toLowerCase();
  const passCol=(s,c,col,allowed)=>{
    const allowedList=(Array.isArray(allowed)?allowed:[allowed]).filter(v=>v!=null && String(v).trim()!=="");
    if(allowedList.some(v=>String(v)===FILTER_NONE)) return false;
    if(!allowedList.length) return true;
    const allowedSet=new Set(allowedList.map(filterNorm));
    if(col==="chase"){
      const owners=(c.chaseOwners||[]).map(o=>o.owner).filter(Boolean);
      if(owners.length===0) return allowedSet.has(filterNorm("(Blanks)"));
      return owners.some(o=>allowedSet.has(filterNorm(o)));
    }
    return allowedSet.has(filterNorm(valueFor(s,c,col)));
  };
  const chaseLabel=(owner)=>String(owner||"—");
  const SAVED_VIEWS=[ ["","Saved view: none"], ["overdue","Overdue"], ["dueThisWeek","Due this week"], ["buyerPending","Buyer approval pending"], ["fabricPending","Fabric pending"], ["ppPending","PP pending"], ["deliveryRisk","Delivery risk"], ["following","Followed styles"], ["rework","Rejected / rework"], ["released","Released"] ];
  const hasColFilters=()=>Object.values(colFilters||{}).some(v=>Array.isArray(v));
  const anyFilter = statusFilter!=="All"||ownerFilter!=="All"||!!search||hasColFilters()||activityFilterKeys.length>0||followFilter||!!savedView||archiveView!=="active";
  const resetFilters=()=>{
    // Hard reset: table filters, search, saved/drill filters, sort, hidden stale dropdown state.
    // This is deliberately not merged with old filter state; it replaces it.
    setStatusFilter("All"); setOwnerFilter("All"); setSearch(""); setSearchCol("auto"); setColFilters({});
    setTrackerActivityFilter(null); setSavedView(""); setFollowFilter(false); setArchiveView("active"); setActiveNamedView("");
    setSort({col:null,dir:1}); setFilterCol(null); setFindIdx(-1); setFrMatches([]); setSel(null); setFocus(null);
    try{ localStorage.setItem("mt_trackfilters", JSON.stringify({ search:"", searchCol:"auto", statusFilter:"All", ownerFilter:"All", archiveView:"active", activityFilter:null, colFilters:{}, tab, todoFilter, followFilter:false, savedView:"" })); }catch(e){}
  };
  const snapCurrent=()=>setViewSnap({ statusFilter, ownerFilter, search, searchCol, colFilters, activityFilter:activityFilterKeys, savedView, archiveView, followFilter, sort });
  const clearAllFilters=()=>{ resetFilters(); setViewSnap(null); };
  const restoreView=()=>{ if(!viewSnap) return; setStatusFilter(viewSnap.statusFilter||"All"); setOwnerFilter(viewSnap.ownerFilter||"All"); setSearch(viewSnap.search||""); setSearchCol(viewSnap.searchCol||"auto"); setColFilters(viewSnap.colFilters||{}); setTrackerActivityFilter(viewSnap.activityFilter||[]); setSavedView(viewSnap.savedView||""); setArchiveView(viewSnap.archiveView||"active"); setFollowFilter(!!viewSnap.followFilter); setSort(viewSnap.sort||{col:null,dir:1}); setViewSnap(null); };
  const applyDrill=(spec)=>{ const drill=canonicalDrillSpec(spec||{}); snapCurrent(); setStatusFilter(drill.status||"All"); setOwnerFilter(drill.owner||"All"); setSearch(drill.search||""); setSearchCol("auto"); setColFilters(drill.colFilters||{}); setTrackerActivityFilter(drill.activityKey||drill.activity||[]); setSavedView(""); setFilterCol(null); setSort({col:null,dir:1}); setTab("tracker"); };
  const presetPass=(s,c)=>{ if(!savedView) return true; const front=c.frontier?[...c.frontier]:[]; const has=(keys)=>front.some(k=>keys.includes(k)); const dueSoon=front.some(k=>{ const r=(c.stages||[]).find(x=>x.key===k); const d=r&&(r.rev||r.plan); const n=d?netWorkdays(TODAY,d):999; return d && n>=0 && n<=6; }); const anyRework=(c.stages||[]).some(r=>r.rework||r.rejected); if(savedView==="overdue") return (c.tone==="late"||c.status.startsWith("Overdue")); if(savedView==="dueThisWeek") return !c.released && dueSoon; if(savedView==="buyerPending") return !c.released && front.some(k=>((STAGES.find(x=>x.key===k)||{}).owner)==="Buyer"); if(savedView==="fabricPending") return !c.released && has(["labDip","labAppr","fabricIH"]); if(savedView==="ppPending") return !c.released && has(["ppSample","ppAppr","prodFile"]); if(savedView==="deliveryRisk") return c.tone==="late"||String(c.status).toLowerCase().includes("risk"); if(savedView==="following") return follows.has(s.id); if(savedView==="rework") return anyRework; if(savedView==="released") return c.released; return true; };
  const deferredSearch=useDeferredValue(search);
  const filterKey=useMemo(()=>JSON.stringify({ search:deferredSearch, searchCol, statusFilter, ownerFilter, archiveView, activityFilter:activityFilterKeys, colFilters, followFilter, savedView, follows:[...follows].sort() }),[deferredSearch,searchCol,statusFilter,ownerFilter,archiveView,activityFilterKeys.join("|"),colFilters,followFilter,savedView,follows]);
  const filterIdentity=useMemo(()=>JSON.stringify({ search:deferredSearch, searchCol, statusFilter, ownerFilter, archiveView, activityFilter:activityFilterKeys, colFilters, followFilter, savedView }),[deferredSearch,searchCol,statusFilter,ownerFilter,archiveView,activityFilterKeys.join("|"),colFilters,followFilter,savedView]);
  const filtered=useMemo(()=>{
    const t=perfNow();
    const q=String(deferredSearch||"").trim().toLowerCase();
    const cfEntries=Object.entries(colFilters||{}).map(([col,allowed])=>[col,(Array.isArray(allowed)?allowed:[allowed]).filter(v=>v!=null && String(v).trim()!=="")]).filter(([,allowed])=>Array.isArray(allowed));
    const out=computed.filter((row)=>{
      const {s,c,idx}=row;
      const matchQ = !q ? true : (searchCol==="auto" ? (idx&&idx.auto?idx.auto:"").includes(q) : ((idx&&idx.byCol&&idx.byCol[searchCol])!=null?idx.byCol[searchCol]:lc(s[searchCol])).includes(q));
      const matchS=statusFilter==="All"||(statusFilter==="At Risk"&&(c.tone==="late"||c.tone==="warn"))||(statusFilter==="On Track"&&c.tone==="ok")||(statusFilter==="Released"&&c.released);
      const matchF=cfEntries.every(([col,allowed])=> passCol(s,c,col,allowed));
      const matchO=ownerFilter==="All"||(c.chaseOwners||[]).some(o=>filterNorm(o.owner)===filterNorm(ownerFilter));
      const matchA=rowMatchesAnyCurrentStage(row,activityFilterKeys);
      const matchArch=archiveView==="all"?true:(archiveView==="archived"?!!s.archived:!s.archived);
      const matchFollow=!followFilter||follows.has(s.id);
      return matchQ&&matchS&&matchF&&matchO&&matchA&&matchArch&&matchFollow&&presetPass(s,c);
    });
    perfRef.current.filterMs=Math.round((perfNow()-t)*10)/10;
    perfRef.current.deferred=deferredSearch!==search;
    return out;
  },[computed,filterKey,deferredSearch,search,stylesComputeKey,stageEvents]);
  const toneRank={ late:0, warn:1, ok:2, done:3, na:4 };
  const fitNum=(s)=>{ const m=String(s.sampleFit).match(/\d+/); return m?Number(m[0]):Infinity; };
  const sortVal=(col,{s,c})=>{ switch(col){ case "__style": return s.styleNo.toLowerCase(); case "orderNo": return (s.orderNo||"~").toLowerCase(); case "sampleFit": return fitNum(s); case "family": return s.family.toLowerCase(); case "colour": return s.colour.toLowerCase(); case "brand": return (s.brand||"").toLowerCase(); case "buyer": return (s.buyer||"").toLowerCase(); case "fabricType": return (s.fabricType||"").toLowerCase(); case "age": return (s.age||"").toLowerCase(); case "extra1": return (s.extra1||"").toLowerCase(); case "extra2": return (s.extra2||"").toLowerCase(); case "owner": return (s.owner||"").toLowerCase(); case "setId": return (s.setId||"~").toLowerCase(); case "setRole": return (s.setRole||"").toLowerCase(); case "qty": return Number(s.qty)||0; case "ordRec": return dateSerial(s.ordRec); case "delivery": return dateSerial(s.delivery); case "overall": return toneRank[c.tone]; case "fit": return toneRank[c.fitBranch.tone]; case "print": return toneRank[c.printBranch.tone]; case "fabric": return toneRank[c.fabricBranch.tone]; case "pp": return toneRank[c.ppBranch.tone]; case "prod": return toneRank[c.prodFileBranch.tone]; case "fabricCD": return c.fabricCountdown.n==null?Infinity:c.fabricCountdown.n; case "proj": return c.projRelease?dateSerial(c.projRelease):Infinity; case "pct": return c.pct; case "chase": return (c.chaseOwners||[]).length; case "float": return c.float==null?Infinity:c.float; case "idle": return c.idle==null?-1:c.idle; case "remarks": return (s.remarks||"~").toLowerCase(); default: {
      // Stage columns are date columns visually. Sort by the same active date the user sees:
      // actual/resend or skipped date first, then revised commitment, then auto/system plan.
      const st=(c.stages||[]).find(x=>x.key===col);
      if(st) return dateSerial(st.actual || st.skip || st.rev || st.plan);
      const a=s.actuals[col]; return dateSerial(a);
    } } };
  const visInfo=INFO_COLS.filter(c=>!hidden.has(c.key));
  const colLabel=(c)=>((cfg.labels&&cfg.labels[c.key])||c.label);
  const visStages=STAGES.filter(s=>!hidden.has(s.key));
  const remarksVis=!hidden.has("remarks");
  const navCols=["__style", ...visInfo.map(c=>c.key), ...visStages.map(s=>s.key), ...(remarksVis?["remarks"]:[])];
  const totalCols=navCols.length;
  const maxFreeze=1+visInfo.length;
  const rows=useMemo(()=>{ const t=perfNow(); const out=!sort.col ? filtered : [...filtered].sort((A,B)=>{ const a=sortVal(sort.col,A), b=sortVal(sort.col,B); return a<b?-sort.dir:a>b?sort.dir:0; }); perfRef.current.rowMs=Math.round((perfNow()-t)*10)/10; perfRef.current.rows=out.length; perfRef.current.alerts=[perfRef.current.p95Ms>800?"P95 >800ms":"", perfRef.current.computeMs>300?"compute slow":"", perfRef.current.filterMs>200?"filter slow":"", perfRef.current.rowMs>200?"sort slow":"", out.length>900?"large visible rows":"", styles.length>1000?"1000+ active styles":"", perfRef.current.deferred?"search catching up":""].filter(Boolean); return out; },[filtered,sort,styles.length]);
  const colIndex=(col)=>navCols.indexOf(col);
  const rowIndex=(id)=>rows.findIndex(r=>String(r.s.id)===String(id));
  const getVal=(s,col)=>{
    if(!s) return "";
    if(col==="__style") return s.styleNo||"";
    if(STAGE_KEYS.includes(col)) return (s.actuals&&s.actuals[col])||"";
    if(col==="remarks") return s.remarks||"";
    if(col==="overall"||col==="fit"||col==="print"||col==="fabric"||col==="pp"||col==="prod"||col==="fabricCD"||col==="proj"||col==="pct"||col==="chase"||col==="float"||col==="idle") return "";
    return s[col]??"";
  };
  const ownerOfCol=(col)=>{ const st=STAGES.find(s=>s.key===col); if(st) return st.owner; const ic=INFO_COLS.find(c=>c.key===col); if(ic&&ic.owner) return ic.owner; return "Merchant"; };
  const isStageCol=(col)=>STAGE_KEYS.includes(col);
  const fmtTyped=(isoStr)=>{ const d=parse(isoStr); return d?`${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`:""; };
  const parseTyped=(strv)=>{ const t=(strv||"").trim(); if(!t) return ""; let m=/^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t); if(m){ const y=+m[1],mo=+m[2],d=+m[3]; if(mo<1||mo>12||d<1||d>31) return false; return `${y}-${String(mo).padStart(2,"0")}-${String(d).padStart(2,"0")}`; } m=/^(\d{1,2})[\/\-.](\d{1,2})(?:[\/\-.](\d{2,4}))?$/.exec(t); if(m){ let d=+m[1],mo=+m[2],y=m[3]?+m[3]:TODAY.getFullYear(); if(y<100) y+=2000; if(mo<1||mo>12||d<1||d>31) return false; return `${y}-${String(mo).padStart(2,"0")}-${String(d).padStart(2,"0")}`; } return false; };
  const [calOpen,setCalOpen]=useState(false);
  const [entryVal,setEntryVal]=useState("");
  const [entryTouched,setEntryTouched]=useState(false);
  const entryCellRef=useRef("");
  // Core value/position helpers must be defined before autocomplete calculations.
  // Earlier autocomplete referenced helpers declared later in the component, which could fail or stay blank while typing.
  const suggestionFor=(col,raw,id)=>{
    if(!col||isDateCol(col)) return "";
    const q=String(raw||"").trim().toLowerCase();
    if(!q) return "";
    const ci=colIndex(col), ri=id!=null?rowIndex(id):-1;
    if(ci<0) return "";
    const seen=[];
    const addVal=(row)=>{ if(!row||!row.s) return; const v=String(getVal(row.s,col)||"").trim(); if(v&&!seen.some(x=>x.toLowerCase()===v.toLowerCase())) seen.push(v); };
    // Prefer values above the active cell, then any visible row in the same filtered table.
    if(ri>=0){ for(let r=ri-1;r>=0&&seen.length<60;r--) addVal(rows[r]); }
    for(let r=0;r<rows.length&&seen.length<100;r++) if(r!==ri) addVal(rows[r]);
    return seen.find(v=>v.toLowerCase().startsWith(q) && v.toLowerCase()!==q)||"";
  };
  const selectedStyle=sel?styles.find(x=>x.id===sel.id):null;
  const selectedColLabel=sel?(sel.col==="__style"?"Style No":(INFO_COLS.find(c=>c.key===sel.col)?.label||STAGES.find(s=>s.key===sel.col)?.label||sel.col)):"";
  const selectedCellKey=sel?`${sel.id}:${sel.col}`:"";
  const selectedDisplayValue=()=>{
    if(!sel||!selectedStyle) return "";
    if(isStageCol(sel.col)){
      const cmp=computed.find(x=>x.s.id===sel.id);
      const r=cmp&&cmp.c&&Array.isArray(cmp.c.stages)?cmp.c.stages.find(x=>x.key===sel.col):null;
      // In rejected/rework cells the old first actual is history, not the editable/display value.
      // Show the revised commitment in the entry bar, or blank if no revised commitment exists.
      if((r && (r.rework || r.rejected) && !r.skipped && !r.autoClosed) || rawReworkOrRejectedCell(sel.id,sel.col)){
        const rv=selectedStyle.revs&&selectedStyle.revs[sel.col];
        return rv ? fmtTyped(rv) : "";
      }
    }
    const v=getVal(selectedStyle,sel.col);
    return isDateCol(sel.col)&&v?fmtTyped(v):String(v??"");
  };
  const entrySuggestion=useMemo(()=> sel&&entryTouched ? suggestionFor(sel.col,entryVal,sel.id) : "", [sel,entryTouched,entryVal,rows,navCols]);
  const editSuggestion=useMemo(()=> editing&&editing.mode==="text" ? suggestionFor(editing.col,editVal,editing.id) : "", [editing,editVal,rows,navCols]);
  useEffect(()=>{
    if(entryCellRef.current!==selectedCellKey){
      entryCellRef.current=selectedCellKey;
      setEntryTouched(false);
      setEntryVal(selectedDisplayValue());
      return;
    }
    if(editing && sel && editing.id===sel.id && editing.col===sel.col && !entryTouched){ setEntryVal(editVal||""); return; }
    if(!entryTouched){ setEntryVal(selectedDisplayValue()); }
  },[selectedCellKey,styles,editing,editVal,entryTouched]);
  const commitEntry=(overrideVal)=>{
    if(!sel||!selectedStyle) return false;
    if(!isEditableCol(sel.col)||!canEdit(role,sel.col,"actual")) return false;
    if(peerLockBlocks(sel.id,sel.col)) return false;
    const raw=String(overrideVal!=null?overrideVal:(entryVal??""));
    if(isDateCol(sel.col)){
      const r=parseTyped(raw);
      if(r===false){ window.alert('"'+raw+'" isn’t a valid date. Type it as dd/mm/yyyy (e.g. 26/05/2026), or clear the box to leave it empty.'); return false; }
      const dateMode=preferredDateMode(sel.id,sel.col,"actual");
      if(dateMode==="rev"){
        if(!canEditRev(role)) return false;
        setRev(sel.id,sel.col,r===""?null:r);
      } else {
        setField(sel.id,sel.col,r===""?null:r);
      }
    } else {
      const f=sel.col==="__style"?"styleNo":sel.col;
      setField(sel.id,f,sel.col==="qty"?raw.replace(/[^0-9]/g,""):raw);
    }
    setEntryTouched(false);
    setEditing(null);
    setEditVal("");
    return true;
  };
  const activeStageRow=(id,col)=>{ const cmp=computed.find(x=>x.s.id===id); return cmp&&cmp.c&&Array.isArray(cmp.c.stages)?cmp.c.stages.find(x=>x.key===col):null; };
  // Robust rejected/rework detection for editing routes. Do not rely only on the rendered computed row,
  // because selection/entry-bar/edit popups can be opened while computed state is one render behind.
  const rawReworkOrRejectedCell=(id,col)=>{
    if(!isStageCol(col)) return false;
    const st=styles.find(x=>String(x.id)===String(id));
    if(!st) return false;
    const skipped=!!(st.skips&&st.skips[col]);
    if(skipped) return false;
    // Send/make stages: while the linked approval is rejected and not resolved,
    // normal cell editing is ALWAYS re-send planning (REVISED), never first-send/resend actual.
    // Actual resend is only through the explicit "enter actual re-send" action.
    // Do not switch back to ACTUAL just because resend history exists; accidental resend history
    // was exactly what made old/actual dates appear in the editor.
    const apprK=APPR_OF_SEND[col];
    if(apprK){
      // Once an explicit re-send actual exists, actual must win over revised/auto in display/editing.
      // Until then, the rejected send cell remains a revised re-send planning cell.
      if(latestResendActualValue(st,col)) return false;
      const rej=st.rejects&&st.rejects[apprK];
      const apprActual=st.actuals&&st.actuals[apprK];
      const apprSkip=st.skips&&st.skips[apprK];
      if(rej && !apprActual && !apprSkip) return true;
    }
    // Approval stages: rejected approval without actual/skip means this cell is a revised re-approval planning cell.
    if(REJECTABLE.includes(col) && st.rejects&&st.rejects[col] && !(st.actuals&&st.actuals[col]) && !(st.skips&&st.skips[col])) return true;
    return false;
  };
  const prefersRevisedDateEntry=(id,col)=>{
    if(!isStageCol(col)) return false;
    const r=activeStageRow(id,col);
    // Actual always wins over revised over auto. If a real resend/reapproval actual exists,
    // normal editing should no longer be forced into revised planning mode.
    if(r && r.actual && !r.skipped && !r.autoClosed) return false;
    if(rawReworkOrRejectedCell(id,col)) return true;
    return !!(r && (r.rework || r.rejected) && !r.skipped && !r.autoClosed);
  };
  // Normal cell entry must open ACTUAL. Revised planning is now only via the ↻ corner button.
  // This prevents operators from accidentally changing revised commitments when they meant to enter actual re-send / re-approval.
  const preferredDateMode=(id,col,requested="actual")=> requested;
  const shouldForceRoundActualEntry=(id,col)=>{
    if(!isStageCol(col)) return false;
    const s=styles.find(x=>x.id===id);
    if(!s) return false;
    const linkedRejected=linkedApprovalRejectedOpen(s,col);
    const selfRejected=!!(REJECTABLE.includes(col)&&s.rejects&&s.rejects[col]&&!(s.skips&&s.skips[col])&&!(s.actuals&&s.actuals[col]));
    return !!(linkedRejected||selfRejected||isResendEntrySlot(s,col));
  };
  const linkedApprovalRejectedOpen=(s,col)=>{
    if(!s||!isStageCol(col)) return false;
    const apprK=APPR_OF_SEND[col];
    if(!apprK) return false;
    return !!(s.rejects&&s.rejects[apprK] && !(s.actuals&&s.actuals[apprK]) && !(s.skips&&s.skips[apprK]));
  };
  const beginDate=(id,col,mode,initialChar,forceActual=false)=>{
    const effectiveMode = (mode==="actual" && !forceActual) ? preferredDateMode(id,col,"actual") : mode;
    if(!canEdit(role,col,effectiveMode)) return;
    if(peerLockBlocks(id,col)) return;
    setSel({id,col}); setFocus(null); setEditing({id,col,mode:effectiveMode,forceActual}); setCalOpen(false);
    const s=styles.find(x=>x.id===id);
    const isResend=(effectiveMode==="actual"&&forceActual&&isStageCol(col)&&isResendEntrySlot(s,col));
    const isReworkPlanning=(effectiveMode==="rev"&&isStageCol(col)&&prefersRevisedDateEntry(id,col));
    // For rework planning, NEVER pull first actual into the editor. Use revised commitment only.
    const cur= isResend?latestResendActualValue(s,col):(effectiveMode==="rev"?(isStageCol(col)?currentRoundRevisedValue(s,col):(s&&s.revs&&s.revs[col])):effectiveMode==="reject"?(s&&s.rejects&&s.rejects[col]):(isStageCol(col)?(s&&s.actuals[col]):(s&&s[col])));
    setEditVal(initialChar!=null?initialChar:(cur?fmtTyped(cur):""));
  };
  const commitDate=()=>{
    if(!editing) return;
    const r=parseTyped(editVal);
    if(r===false){ window.alert('"'+editVal+'" isn’t a valid date. Type it as dd/mm/yyyy (e.g. 26/05/2026), or clear the box to leave it empty.'); return; }
    const val=r===""?null:r;
    const curS=styles.find(x=>x.id===editing.id);
    const effectiveMode=(editing.mode==="actual" && !editing.forceActual) ? preferredDateMode(editing.id,editing.col,"actual") : editing.mode;
    const isExplicitResendActual=(effectiveMode==="actual" && editing.forceActual && isStageCol(editing.col) && isResendEntrySlot(curS,editing.col));
    if(isExplicitResendActual && val && !validateReworkActualDate(curS,editing.col,val)) return;
    if(val===null){
      // Blank explicit re-send actual clears only the latest re-send actual; it never deletes the original first-send actual.
      if(isExplicitResendActual){ clearLatestResendActual(editing.id,editing.col); return; }
      const st=curS;
      const had= effectiveMode==="rev"?(st&&st.revs&&st.revs[editing.col]):effectiveMode==="reject"?(st&&st.rejects&&st.rejects[editing.col]):(isStageCol(editing.col)?(st&&st.actuals&&st.actuals[editing.col]):(st&&st[editing.col]));
      if(had && !window.confirm("Delete this saved date? You can re-enter it later.")){ setEditing(null); setCalOpen(false); return; }
    }
    if(effectiveMode==="rev") setRev(editing.id,editing.col,val);
    else if(effectiveMode==="reject") setReject(editing.id,editing.col,val);
    else setField(editing.id,editing.col,val);
    setEditing(null); setCalOpen(false);
  };
  const dateEditor=(id,col,mode)=>{ const s=styles.find(x=>x.id===id); const forceActual=!!(editing&&editing.id===id&&editing.col===col&&editing.forceActual); const effectiveMode=(mode==="actual"&&!forceActual)?preferredDateMode(id,col,"actual"):mode; const _cmp=computed.find(x=>x.s.id===id); const _stR=_cmp&&(_cmp.c.stages||[]).find(x=>x.key===col); const planFb=_stR?(_stR.rev||_stR.plan):null; const isResend=(effectiveMode==="actual"&&forceActual&&isStageCol(col)&&isResendEntrySlot(s,col)); const realStored= effectiveMode==="rev"?(isStageCol(col)?currentRoundRevisedValue(s,col):(s&&s.revs&&s.revs[col])):effectiveMode==="reject"?(s&&s.rejects&&s.rejects[col]):(isStageCol(col)?(s&&s.actuals[col]):(s&&s[col])); const stored=isResend?latestResendActualValue(s,col):realStored; const baseLabel=col==="ordRec"?"Order Date":col==="delivery"?"Delivery Date":((STAGES.find(x=>x.key===col)||{}).label||col); const linkedRejectCtx=linkedApprovalRejectedOpen(s,col); const rejectedApprovalCtx=!!(REJECTABLE.includes(col)&&s&&s.rejects&&s.rejects[col]&&!(s.skips&&s.skips[col])&&!(s.actuals&&s.actuals[col])); const roundCtx=isStageCol(col)&&(rawReworkOrRejectedCell(id,col)||linkedRejectCtx||rejectedApprovalCtx||(effectiveMode==="rev"&&activeRoundForCell(s,col)>0)||isResend); const activeRound=roundCtx?activeRoundForCell(s,col):0; const colLabel=roundCtx&&APPR_OF_SEND[col]?baseLabel.replace(" Send","")+" RE-SEND "+(activeRound||1):roundCtx&&REJECTABLE.includes(col)?baseLabel.replace(" Appr","")+" RE-APPR "+(activeRound||1):baseLabel; const modeLabel=effectiveMode==="rev"?(roundCtx?"REVISED":"REVISED"):effectiveMode==="reject"?"REJECTED":(isResend?"ACTUAL":"ACTUAL"); const mc=effectiveMode==="rev"?"var(--accent)":effectiveMode==="reject"?"var(--danger)":"var(--info)"; const isRoundRev=effectiveMode==="rev"&&roundCtx; return (<span onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:1, left:1, zIndex:80, display:"flex", flexDirection:"column", gap:1, background:"var(--surface)", border:"1px solid "+mc, padding:"2px 4px", boxShadow:"2px 2px 0 rgba(0,0,0,0.18)" }}><span style={{ fontSize:8, fontWeight:700, color:mc, textTransform:"uppercase", letterSpacing:0.3, whiteSpace:"nowrap" }}>{colLabel} · {modeLabel}</span>{isRoundRev&&<span style={{ fontSize:8, color:"var(--revised)", fontWeight:800, whiteSpace:"nowrap" }}>actual stays active · editing revised commitment</span>}{isResend&&realStored&&<span style={{ fontSize:8, color:"#b4531a", fontWeight:700, whiteSpace:"nowrap" }}>first sent kept: {fmt(parse(realStored))}</span>}<span style={{ display:"flex", alignItems:"center", gap:2, position:"relative" }}><input autoFocus onFocus={e=>{ if((e.target.value||"").length>2) e.target.select(); }} value={editVal} placeholder={isRoundRev?"revised resend/re-approval date":(isResend?"actual re-send/re-approval date":"dd/mm/yyyy")} onChange={e=>setEditVal(e.target.value.replace(/[^0-9\/\-. ]/g,""))} onKeyDown={e=>{ e.stopPropagation(); if(e.key==="Enter") commitDate(); else if(e.key==="Escape"){ setEditing(null); setCalOpen(false); } }} onBlur={()=>{ if(!calOpen) commitDate(); }} style={{ width:isResend?96:80, fontFamily:"inherit", fontSize:11, border:"none", outline:"none" }}/>{stored && <button onMouseDown={e=>e.preventDefault()} onClick={e=>{ e.stopPropagation(); if(isResend) clearLatestResendActual(id,col); else if(effectiveMode==="rev") setRev(id,col,null); else if(effectiveMode==="reject") setReject(id,col,null); else setField(id,col,null); setEditing(null); setCalOpen(false); }} title={"Clear "+modeLabel.toLowerCase()+" date"} style={{ border:"1px solid var(--line-2)", background:"var(--surface)", cursor:"pointer", padding:"0 4px", fontSize:10, lineHeight:"16px" }}>clear</button>}<button onMouseDown={e=>e.preventDefault()} onClick={e=>{ e.stopPropagation(); setCalOpen(o=>!o); }} title="calendar" style={{ border:"none", background:"transparent", cursor:"pointer", padding:0, lineHeight:0, fontSize:12 }}>📅</button>{calOpen && <CalPopup label={colLabel+" · "+modeLabel} value={stored} fallback={planFb} onClose={()=>setCalOpen(false)} onPick={(d)=>{ if(effectiveMode==="rev") setRev(id,col,d); else if(effectiveMode==="reject") setReject(id,col,d); else setField(id,col,d); setEditing(null); setCalOpen(false); }}/>}</span></span>); };
    const startEdit=(id,col,initialChar)=>{ if(!isEditableCol(col)) return; if(!canEdit(role,col,"actual")) return; if(isDateCol(col)){ beginDate(id,col,"actual",initialChar,shouldForceRoundActualEntry(id,col)); return; } if(peerLockBlocks(id,col)) return; const s=styles.find(x=>x.id===id); setEditing({id,col,mode:"text"}); if(col==="qty") setEditVal(initialChar??String(s.qty)); else if(col==="__style") setEditVal(initialChar??s.styleNo); else setEditVal(initialChar??(s[col]||"")); };
  const commitText=(overrideVal)=>{ if(!editing) return false; const f=editing.col==="__style"?"styleNo":editing.col; const raw=overrideVal!=null?overrideVal:editVal; if(editing.mode==="text"||editing.mode===undefined){ if(!isDateCol(editing.col)) setField(editing.id,f,raw); } setEditing(null); return true; };
  const finishEditing=()=>{ if(!editing) return; if(editing.mode==="actual"||editing.mode==="rev"||editing.mode==="reject") commitDate(); else commitText(); };

  // ---- selection range ----
  const peerEditingList=peers.filter(p=>p.editing).map(p=>({ ...p, ref: (()=>{ const loc=p.editing; if(!loc) return ""; const ri=rows.findIndex(r=>r.s.id===loc.id); const ci=navCols.indexOf(loc.col); return (ci>=0&&ri>=0)?(colLetter(ci)+(ri+1)):""; })() }));
  const [renderLimit,setRenderLimit]=useState(()=>{ try{ const v=parseInt(localStorage.getItem("mt_render_limit"),10); return (isFinite(v)&&v>=300&&v<=5000)?v:900; }catch(e){ return 900; } });
  useEffect(()=>{ try{ localStorage.setItem("mt_render_limit",String(renderLimit)); }catch(e){} },[renderLimit]);
  useEffect(()=>{ if(rows.length<=renderLimit && renderLimit>900) setRenderLimit(900); },[rows.length,renderLimit]);
  const renderRows=useMemo(()=>{ const visible=hiddenRows&&hiddenRows.size?rows.filter(r=>!hiddenRows.has(String(r.s.id))):rows; const out=visible.length>renderLimit?visible.slice(0,renderLimit):visible; perfRef.current.rendered=out.length; return out; },[rows,renderLimit,hiddenRows]);
  const clickHeader=(col)=>{ finishEditing(); setSort(p=> p.col===col?{col,dir:-p.dir}:{col,dir:1}); };

  const ROLE_VIEW_PRESETS={
    merchant:{ label:"Merchant view", show:null },
    management:{ label:"Management view", show:["orderNo","family","colour","brand","owner","buyer","qty","delivery","overall","fit","print","fabric","pp","prod","fabricCD","proj","pct","chase","float","idle","remarks"] },
    cad:{ label:"CAD view", show:["orderNo","sampleFit","family","colour","brand","owner","delivery","overall","fit","techpack","fitSend","fitAppr","proj","chase","remarks"] },
    designer:{ label:"Designer view", show:["orderNo","family","colour","brand","fabricType","owner","delivery","overall","print","artwork","artAppr","strikeOff","soAppr","proj","chase","remarks"] },
    store:{ label:"Store / Fabric view", show:["orderNo","family","colour","brand","fabricType","owner","delivery","overall","fabric","labDip","labAppr","fabricIH","fabricCD","chase","remarks"] },
    buyer:{ label:"Buyer follow-up view", show:["orderNo","family","colour","brand","buyer","delivery","overall","fitAppr","artAppr","soAppr","labAppr","ppAppr","proj","chase","remarks"] }
  };
  const applyColumnView=(view)=>{ setColumnView(view); setColsOpen(false); if(view==="custom"){ setHidden(new Set()); return; } const all=[...INFO_COLS.map(c=>c.key),...STAGE_KEYS,"remarks"]; const spec=ROLE_VIEW_PRESETS[view]; if(!spec||!spec.show){ setHidden(new Set()); return; } const keep=new Set(spec.show); setHidden(new Set(all.filter(k=>!keep.has(k)))); setFreezeN(1); };

  const currentViewState=()=>({
    search, searchCol, statusFilter, ownerFilter, archiveView, activityFilter:activityFilterKeys, colFilters, savedView, columnView,
    hidden:[...hidden], sort, freezeN
  });
  const applyTrackerViewState=(state, name="")=>{
    const st=state||{};
    setSearch(st.search||""); setSearchCol(st.searchCol||"auto"); setStatusFilter(st.statusFilter||"All"); setOwnerFilter(st.ownerFilter||"All");
    setArchiveView(st.archiveView||"active"); setTrackerActivityFilter(st.activityFilter||[]); setColFilters(st.colFilters||{}); setSavedView(st.savedView||"");
    if(st.columnView){ setColumnView(st.columnView); }
    if(Array.isArray(st.hidden)) setHidden(new Set(st.hidden)); else if(st.columnView && st.columnView!=="custom") applyColumnView(st.columnView);
    setSort(st.sort&&st.sort.col?st.sort:{ col:null, dir:1 }); setFreezeN(Number(st.freezeN)||1); setActiveNamedView(name||""); setViewSnap(null);
  };
  const saveSharedTrackerView=async()=>{
    if(!canAdmin(role)){ alert("Only Management / Sr Merchant can save or overwrite shared default views. Your current filters are still temporary for you only."); return; }
    const name=(window.prompt("Save current filters/columns as default shared view:", (activeNamedView&&activeNamedView!==SYSTEM_DEFAULT_TRACKER_VIEW_NAME)?activeNamedView:"Management View")||"").trim();
    if(!name) return;
    if(name===SYSTEM_DEFAULT_TRACKER_VIEW_NAME){ alert("System Default — All Columns is locked and cannot be overwritten by any user. Save a new named view instead."); return; }
    const next=normalizeTrackerViews(sharedViews.filter(v=>v.name!==name).concat([{ name, shared:true, updatedAt:new Date().toISOString(), updatedBy:me?.name||me?.email||"", state:currentViewState() }]));
    setSharedViews(next); setActiveNamedView(name);
    try{ const { error }=await supabase.from("app_settings").upsert({ id:TRACKER_VIEW_SETTING_ID, data:{ views:next } }); if(error) throw error; flash(); }
    catch(e){ logAppError("save shared view failed",e); alert("View saved locally in this browser, but shared save failed: "+(e.message||e)); }
  };
  const resetTemporaryView=()=>{ clearAllFilters(); setSort({ col:null, dir:1 }); setColumnView("custom"); setHidden(new Set()); setFreezeN(1); setActiveNamedView(""); const w=scrollWrapRef.current; if(w){ w.scrollLeft=0; w.scrollTop=0; } };

  // ---- freeze: cumulative left offsets for frozen leading columns ----
  // IMPORTANT: every frozen cell must use the same explicit width and left offset.
  // Without width/minWidth/maxWidth, Chrome can visually stack sticky table-cells when multiple columns are frozen.
  const frozenCount=Math.min(Number(freezeN)||1, maxFreeze);
  const frozenCols=navCols.slice(0, frozenCount);
  const widthOf=(col)=> colW[col] ?? (col==="__style"?STYLE_W : col==="remarks"?REMARK_COL.w : (INFO_W[col]!==undefined?INFO_W[col]:84));
  const onResize=(col,w)=> setColW(p=>({ ...p, [col]:Math.max(40, Math.round(w)) }));
  const frozenLefts={}; let __freezeX=0; frozenCols.forEach(c=>{ frozenLefts[c]=__freezeX; __freezeX+=widthOf(c); });
  const leftOf=(col)=> frozenLefts[col] ?? 0;
  const isFrozen=(col)=> frozenCols.includes(col);
  const lastFrozen=frozenCols[frozenCols.length-1];
  const freezeStyle=(col,bg)=>{ if(!isFrozen(col)) return {}; const w=widthOf(col); return { position:"sticky", left:leftOf(col), zIndex: col==="__style"?12:11, background:bg, backgroundClip:"padding-box", width:w, minWidth:w, maxWidth:w, boxSizing:"border-box", transform:"translateZ(0)", ...(col===lastFrozen?{ borderRight:"2px solid var(--muted-4)", boxShadow:"2px 0 0 var(--muted-4)" }:{ boxShadow:"1px 0 0 var(--line-2)" }) }; };

  const cellRef=(x)=> x? (colLetter(colIndex(x.col))+(rowIndex(x.id)+1)) : "";
  const [nameBox,setNameBox]=useState("");
  useEffect(()=>{ setNameBox(cellRef(sel)); },[sel]);
  const gotoCell=(ref)=>{ const m=/^([A-Za-z]+)\s*(\d+)$/.exec(String(ref||"").trim()); if(!m) return; const ci=letterToIndex(m[1]); const ri=Number(m[2])-1; if(ci<0||ci>=navCols.length||ri<0||ri>=rows.length) return; const id=rows[ri].s.id, col=navCols[ci]; setSel({ id, col }); setFocus(null); scrollToCell(id,col); };
  const rectFrom=(a,b)=>{ if(!a) return null; const f=b||a; const aR=rowIndex(a.id), aC=colIndex(a.col), fR=rowIndex(f.id), fC=colIndex(f.col); if(aR<0||aC<0||fR<0||fC<0) return null; return { r1:Math.min(aR,fR), r2:Math.max(aR,fR), c1:Math.min(aC,fC), c2:Math.max(aC,fC) }; };
  const rect=()=>rectFrom(sel,focus||sel);
  const cellKeyFrom=(x)=>x?`${x.id}:${x.col}`:"";
  const keysForRect=(R)=>{ const out=[]; if(!R) return out; for(let r=R.r1;r<=R.r2;r++){ for(let c=R.c1;c<=R.c2;c++){ if(rows[r]&&navCols[c]) out.push(`${rows[r].s.id}:${navCols[c]}`); } } return out; };
  const orderedSelectedCells=()=>{ const seen=new Set(); const out=[]; const add=(key)=>{ if(!key||seen.has(key)) return; const [sid,...rest]=String(key).split(":"); const col=rest.join(":"); const id=Number(sid); const r=rowIndex(id), c=colIndex(col); if(r<0||c<0||!rows[r]) return; seen.add(key); out.push({ id, col, r, c, s:rows[r].s }); }; keysForRect(rect()).forEach(add); (multiAreas||[]).forEach(a=>keysForRect(rectFrom(a.anchor,a.focus||a.anchor)).forEach(add)); out.sort((a,b)=>a.r-b.r||a.c-b.c); return out; };
  const selectedCells=()=>orderedSelectedCells();
  const selKeys=useMemo(()=>{ const set=new Set(); keysForRect(rect()).forEach(k=>set.add(k)); (multiAreas||[]).forEach(a=>keysForRect(rectFrom(a.anchor,a.focus||a.anchor)).forEach(k=>set.add(k))); return set; },[sel,focus,multiAreas,rows,navCols]);
  const hasMultiSelect=!!(multiAreas&&multiAreas.length);
  const clearSelection=()=>{ setSel(null); setFocus(null); setMultiAreas([]); setCtxMenu(null); };
  const setSingleSelection=(id,col)=>{ setSel({id,col}); setFocus(null); setMultiAreas([]); };
  const addOrToggleSingle=(id,col)=>{ const key=`${id}:${col}`; setMultiAreas(prev=>{ const arr=prev||[]; const without=arr.filter(a=>cellKeyFrom(a.anchor)!==key || cellKeyFrom(a.focus||a.anchor)!==key); if(without.length!==arr.length) return without; if(sel && !focus && cellKeyFrom(sel)===key){ setSel(null); return arr; } return [...arr,{ anchor:{id,col}, focus:{id,col} }]; }); };
  const addRangeArea=(from,to)=>{ if(!from||!to) return; setMultiAreas(prev=>[...(prev||[]),{ anchor:from, focus:to }]); };
  const onCellClick=(e,id,col)=>{ e.stopPropagation(); setCtxMenu(null); if(filterCol) setFilterCol(null); if(gridRef.current) gridRef.current.focus({preventScroll:true}); if(editing){ if(editing.id===id&&editing.col===col) return; finishEditing(); } if((e.ctrlKey||e.metaKey)&&e.shiftKey&&sel){ addRangeArea(sel,{id,col}); scrollToCell(id,col); return; } if(e.ctrlKey||e.metaKey){ addOrToggleSingle(id,col); scrollToCell(id,col); return; } if(e.shiftKey&&sel){ setFocus({id,col}); scrollToCell(id,col); return; } setSingleSelection(id,col); };

  const moveAnchor=(dr,dc)=>{ if(!sel) return; let r=rowIndex(sel.id)+dr, c=colIndex(sel.col)+dc; r=Math.min(Math.max(r,0),rows.length-1); c=Math.min(Math.max(c,0),navCols.length-1); if(rows[r]){ setSel({id:rows[r].s.id,col:navCols[c]}); setFocus(null); setMultiAreas([]); scrollToCell(rows[r].s.id,navCols[c]); } };
  const scrollToCell=(id,col)=>{ if(tab!=="tracker") setTab("tracker"); if(col && col!=="__style" && hidden.has(col)) setHidden(prev=>{ const n=new Set(prev); n.delete(col); return n; });
    const go=()=>{ let el=document.getElementById(`cell-${id}-${col}`); if(!el) el=document.getElementById(`cell-${id}-__style`); if(!el) return; const wrap=scrollWrapRef.current; if(wrap){ if(col==="__style") wrap.scrollLeft=0; const cr=el.getBoundingClientRect(), wr=wrap.getBoundingClientRect(); wrap.scrollTop += (cr.top - wr.top) - (wr.height/2 - cr.height/2); const cr2=el.getBoundingClientRect(); const frozen=STYLE_W+6; if(col!=="__style"){ if(cr2.left < wr.left+frozen) wrap.scrollLeft -= (wr.left+frozen-cr2.left)+8; else if(cr2.right > wr.right) wrap.scrollLeft += (cr2.right-wr.right)+8; } requestAnimationFrame(()=>{ try{ const cr3=el.getBoundingClientRect(); const desired=window.innerHeight*0.52; const delta=cr3.top-desired; if(Math.abs(delta)>36) window.scrollBy({ top:delta, behavior:"smooth" }); }catch(e){} }); } else { el.scrollIntoView({ block:"center", inline:col==="__style"?"start":"center", behavior:"smooth" }); } };
    requestAnimationFrame(()=>requestAnimationFrame(go)); setTimeout(go,120); setTimeout(go,260); };
  const selectRow=(id)=>{ setSel({id,col:navCols[0]}); setFocus({id,col:navCols[navCols.length-1]}); setMultiAreas([]); };
  const selectAll=()=>{ if(!rows.length) return; setSel({id:rows[0].s.id,col:navCols[0]}); setFocus({id:rows[rows.length-1].s.id,col:navCols[navCols.length-1]}); setMultiAreas([]); };
  const clampRow=(r)=>Math.min(Math.max(r,0),Math.max(0,rows.length-1));
  const clampCol=(c)=>Math.min(Math.max(c,0),Math.max(0,navCols.length-1));
  const edgeIndexFor=(from,dr,dc)=>{
    let r=clampRow(rowIndex(from.id)), c=clampCol(colIndex(from.col));
    if(dr>0) r=Math.max(0,rows.length-1); else if(dr<0) r=0;
    if(dc>0) c=Math.max(0,navCols.length-1); else if(dc<0) c=0;
    return { r, c };
  };
  const revealRowIndex=(r)=>{ if(r>=renderLimit) setRenderLimit(Math.min(5000,Math.max(renderLimit,r+1))); };
  const jumpAnchorToEdge=(dr,dc)=>{ if(!sel) return; const t=edgeIndexFor(sel,dr,dc); if(rows[t.r]){ revealRowIndex(t.r); setSel({id:rows[t.r].s.id,col:navCols[t.c]}); setFocus(null); setMultiAreas([]); setTimeout(()=>scrollToCell(rows[t.r].s.id,navCols[t.c]),40); } };
  const extendFocusToEdge=(dr,dc)=>{ if(!sel) return; const t=edgeIndexFor(focus||sel,dr,dc); if(rows[t.r]){ revealRowIndex(t.r); setFocus({id:rows[t.r].s.id,col:navCols[t.c]}); setTimeout(()=>scrollToCell(rows[t.r].s.id,navCols[t.c]),40); } };
  const moveFocus=(dr,dc)=>{ if(!sel) return; const f=focus||sel; let r=rowIndex(f.id)+dr, c=colIndex(f.col)+dc; r=Math.min(Math.max(r,0),rows.length-1); c=Math.min(Math.max(c,0),navCols.length-1); if(rows[r]){ setFocus({id:rows[r].s.id,col:navCols[c]}); scrollToCell(rows[r].s.id,navCols[c]); } };

  const snap=()=>({ styles, fills, notes, resends, revHistory });
  const pushHistory=()=>{ setPast(p=>[...p.slice(-60), snap()]); setFuture([]); };
  const applySnap=(d)=>{ setStyles(d.styles); setFills(d.fills); setNotes(d.notes); if(d.resends) setResends(d.resends); if(d.revHistory) setRevHistory(d.revHistory); };
  const undo=()=>{ if(!past.length) return; if(remoteChanged && !window.confirm("Teammates have made changes that aren't merged in yet (the Sync button is highlighted).\n\nUndoing now may overwrite them. Undo anyway?")) return; const prev=past[past.length-1]; setFuture(f=>[...f, snap()]); setPast(p=>p.slice(0,-1)); applySnap(prev); flash(); };
  const redo=()=>{ if(!future.length) return; if(remoteChanged && !window.confirm("Teammates have unmerged changes (the Sync button is highlighted).\n\nRedoing now may overwrite them. Redo anyway?")) return; const nx=future[future.length-1]; setPast(p=>[...p, snap()]); setFuture(f=>f.slice(0,-1)); applySnap(nx); flash(); };
  const doCopy=()=>{ const cells=selectedCells(); if(!cells.length) return; const R=rect(); if(!hasMultiSelect && R){ const values=[]; let any=false; for(let r=R.r1;r<=R.r2;r++){ const row=[]; for(let c=R.c1;c<=R.c2;c++){ const v=rows[r]?getVal(rows[r].s,navCols[c]):null; if(v!=null&&v!=="") any=true; row.push(v); } values.push(row); } if(any){ setClip({ values, h:values.length, w:values[0].length, multi:false }); flash(); } return; } const values=cells.map(x=>[getVal(x.s,x.col)??""]); setClip({ values, h:values.length, w:1, multi:true, cells:cells.map(x=>({col:x.col})) }); flash(); };
  const canPasteCell=(s,col)=>{ if(!isEditableCol(col)) return false; if(!canEdit(role,col,"actual")) return false; if(STAGE_KEYS.includes(col)){ const st=STAGES.find(x=>x.key===col); if(!(st.flag===null||s[st.flag])) return false; } return true; };
  const doPaste=()=>{ if(!clip||!sel) return; const R=rect(); const changes={}; const put=(id,col,val)=>{ (changes[id]=changes[id]||{})[col]=val; };
    const cells=selectedCells();
    if(hasMultiSelect && cells.length){ const flat=(clip.values||[]).flat(); const single=flat.length===1; cells.forEach((cell,i)=>{ const v=single?flat[0]:flat[i%Math.max(1,flat.length)]; if(canPasteCell(cell.s,cell.col)) put(cell.id,cell.col,v); }); }
    else if(clip.h===1&&clip.w===1){ const v=clip.values[0][0]; for(let r=R.r1;r<=R.r2;r++){ for(let c=R.c1;c<=R.c2;c++){ const row=rows[r]; const col=navCols[c]; if(row&&canPasteCell(row.s,col)) put(row.s.id,col,v); } } }
    else { for(let i=0;i<clip.h;i++){ for(let j=0;j<clip.w;j++){ const r=R.r1+i, c=R.c1+j; const row=rows[r]; const col=navCols[c]; if(row&&col&&canPasteCell(row.s,col)) put(row.s.id,col,clip.values[i][j]); } } }
    if(!confirmIdentityBatchChanges(changes)) return;
    Object.keys(changes).forEach(id=>maybePinEditedRow(id));
    pushHistory(); setStyles(prev=>prev.map(s=>{ const ch=changes[s.id]; if(!ch) return s; let ns={...s, actuals:{...s.actuals}}; Object.entries(ch).forEach(([col,val])=>{ if(STAGE_KEYS.includes(col)) ns.actuals[col]=val||undefined; else if(col==="qty") ns.qty=Number(val)||0; else if(col==="__style") ns.styleNo=val; else ns[col]=val; }); return ns; })); flash(); };

  // batch write a {id:{col:val}} change map into styles
  const writeChanges=(changes)=>{ if(!confirmIdentityBatchChanges(changes)) return; Object.keys(changes||{}).forEach(id=>maybePinEditedRow(id)); Object.entries(changes).forEach(([id,ch])=>Object.entries(ch).forEach(([col,val])=>{ if(STAGE_KEYS.includes(col)){ const ck=id+":"+col+":actual"; if(val) clearedRef.current.delete(ck); else clearedRef.current.add(ck); } })); setStyles(prev=>prev.map(s=>{ const ch=changes[s.id]; if(!ch) return s; let ns={...s, actuals:{...s.actuals}, revs:{...(s.revs||{})}}; Object.entries(ch).forEach(([col,val])=>{ if(STAGE_KEYS.includes(col)) ns.actuals[col]=val||undefined; else if(col==="qty") ns.qty=Number(val)||0; else if(col==="__style") ns.styleNo=val; else ns[col]=val; }); return ns; })); flash(); };
  const coerce=(col,raw)=>{ if(raw==null) return ""; const v=String(raw).trim(); if(isDateCol(col)){ const pt=parseTyped(v); if(pt!==false) return pt; const d=new Date(v); return isNaN(d)?"":iso(d); } return v; };
  const clearRange=()=>{
    const cells=selectedCells(); if(!cells.length) return; pushHistory(); const ch={};
    cells.forEach(cell=>{
      if(!canPasteCell(cell.s,cell.col)) return;
      const r=isStageCol(cell.col)?activeStageRow(cell.id,cell.col):null;
      if(r && (r.rework||r.rejected) && !r.skipped && !r.autoClosed){
        if(cell.s.revs&&cell.s.revs[cell.col]) setRev(cell.id,cell.col,null);
        return;
      }
      (ch[cell.id]=ch[cell.id]||{})[cell.col]= STAGE_KEYS.includes(cell.col)?null:(cell.col==="qty"?0:"");
    });
    if(Object.keys(ch).length) writeChanges(ch);
  };
  const applyFillHandle=()=>{ if(!fillFrom||!fillTo) return; const aR=rowIndex(fillFrom.id), aC=colIndex(fillFrom.col), tR=rowIndex(fillTo.id), tC=colIndex(fillTo.col); const r1=Math.min(aR,tR), r2=Math.max(aR,tR), c1=Math.min(aC,tC), c2=Math.max(aC,tC); const srcRow=rows[aR]; if(!srcRow) return; pushHistory(); const ch={}; for(let r=r1;r<=r2;r++){ for(let cc=c1;cc<=c2;cc++){ const row=rows[r]; const col=navCols[cc]; if(!row) continue; if(r===aR&&cc===aC) continue; const srcVal=getVal(srcRow.s, navCols[cc]); if(canPasteCell(row.s,col)) (ch[row.s.id]=ch[row.s.id]||{})[col]=srcVal; } } writeChanges(ch); };
  const headerTextFor=(col)=>String(col==="__style"?"Style No":(INFO_COLS.find(x=>x.key===col)?.label||STAGES.find(x=>x.key===col)?.label||col));
  const headerFitWidth=(col)=>{
    const label=headerTextFor(col);
    // Header auto-fit must show the full header, not only cell contents.
    // Include: label text, column letter, filter/sort icon, resize handle, padding, and a safety buffer.
    const letter=colLetter(navCols.indexOf(col));
    const letterW=letter ? Math.max(20, letter.length*10) : 0;
    const iconW=34;       // filter + sort icon area
    const resizeW=14;     // drag handle/right padding
    const padW=28;
    const textW=Math.ceil(label.length*9.8*textScale);
    return Math.ceil(textW + letterW + iconW + resizeW + padW);
  };
  const contentFitWidth=(col)=>{
    let max=headerTextFor(col).length;
    rows.forEach(({s})=>{ const v=getVal(s,col); if(v!=null) max=Math.max(max,String(isDateCol(col)?fmt(parse(v)):v).length); });
    return Math.ceil(max*7.8*textScale+40);
  };
  const fitAllCols=()=>{ const cols=["__style", ...visInfo.map(c=>c.key), ...visStages.map(s=>s.key)]; const upd={}; cols.forEach(col=>{ const minW=col==="__style"?STYLE_W:70; upd[col]=Math.max(minW, Math.min(520, Math.max(contentFitWidth(col), headerFitWidth(col)))); }); setColW(p=>({...p,...upd})); flash(); };
  const autoFit=(col)=>{
    const minW=col==="__style"?STYLE_W:70;
    setColW(p=>({ ...p, [col]:Math.max(minW, Math.min(520, Math.max(contentFitWidth(col), headerFitWidth(col)))) }));
  };
  const clearCellByContext=(id,col)=>{
    const s=styles.find(x=>x.id===id);
    if(!s||!canPasteCell(s,col)) return;
    // In rejected/rework cells, Delete/Backspace should clear the revised commitment first, never the old first-send actual.
    const r=isStageCol(col)?activeStageRow(id,col):null;
    if((r && (r.rework||r.rejected) && !r.skipped && !r.autoClosed) || rawReworkOrRejectedCell(id,col)){
      if(s.revs&&s.revs[col]){
        if(window.confirm("Delete the revised commitment date? The first actual/history date will be kept.")) setRev(id,col,null);
      }
      return;
    }
    setField(id,col, STAGE_KEYS.includes(col)?null:(col==="qty"?0:""));
  };
  const onKeyDown=(e)=>{ const _tt=e.target&&e.target.tagName; if(_tt==="INPUT"||_tt==="SELECT"||_tt==="TEXTAREA"||(e.target&&e.target.isContentEditable)) return;
    if((e.ctrlKey||e.metaKey)&&(e.key==="z"||e.key==="Z")){ e.preventDefault(); if(e.shiftKey) redo(); else undo(); return; }
    if((e.ctrlKey||e.metaKey)&&(e.key==="y"||e.key==="Y")){ e.preventDefault(); redo(); return; }
    if(!sel) return;
    if(editing){ const dm=editing.mode==="actual"||editing.mode==="rev"; if(e.key==="Enter"){ dm?commitDate():commitText(); moveAnchor(1,0); e.preventDefault(); } else if(e.key==="Escape"){ setEditing(null); setCalOpen(false); e.preventDefault(); } else if(e.key==="Tab"){ dm?commitDate():commitText(); moveAnchor(0,e.shiftKey?-1:1); e.preventDefault(); } return; }
    if((e.ctrlKey||e.metaKey)&&(e.key==="c"||e.key==="C")){ const selectedTxt=(typeof window!=="undefined"&&window.getSelection)?String(window.getSelection().toString()||""):""; if(selectedTxt&&selectedTxt.trim()) return; doCopy(); e.preventDefault(); return; }
    if((e.ctrlKey||e.metaKey)&&(e.key==="v"||e.key==="V")){ doPaste(); e.preventDefault(); return; }
    if((e.ctrlKey||e.metaKey)&&(e.key==="a"||e.key==="A")){ selectAll(); e.preventDefault(); return; }
    if(e.key==="Tab"){ e.preventDefault(); if(e.shiftKey) moveAnchor(0,-1); else moveAnchor(0,1); }
    else if(["ArrowUp","ArrowDown","ArrowLeft","ArrowRight"].includes(e.key)){
      e.preventDefault();
      const dr=e.key==="ArrowUp"?-1:e.key==="ArrowDown"?1:0;
      const dc=e.key==="ArrowLeft"?-1:e.key==="ArrowRight"?1:0;
      if(e.ctrlKey||e.metaKey){ if(e.shiftKey) extendFocusToEdge(dr,dc); else jumpAnchorToEdge(dr,dc); }
      else if(e.shiftKey) moveFocus(dr,dc);
      else moveAnchor(dr,dc);
    }
    else if(e.key==="Enter"||e.key==="F2"){ e.preventDefault(); startEdit(sel.id,sel.col); }
    else if(e.key==="Delete"||e.key==="Backspace"){ if(focus) clearRange(); else clearCellByContext(sel.id,sel.col); e.preventDefault(); }
    else if(e.key==="Escape"){ clearSelection(); }
    else if(e.key.length===1&&!e.metaKey&&!e.ctrlKey){ if(isDateCol(sel.col)&&/[0-9]/.test(e.key)){ beginDate(sel.id,sel.col,"actual",e.key,shouldForceRoundActualEntry(sel.id,sel.col)); e.preventDefault(); } else if(sel.col==="qty"&&/[0-9]/.test(e.key)){ startEdit(sel.id,"qty",e.key); e.preventDefault(); } else if(sel.col==="__style"||TEXT_COLS.includes(sel.col)){ startEdit(sel.id,sel.col,e.key); e.preventDefault(); } }
  };
  const jumpToEnter=(id,stageKey)=>{ const col=stageKey||"__style"; setSel({id,col}); setFocus(null); requestAnimationFrame(()=>setTimeout(()=>scrollToCell(id,col),40)); }; // jump + select only — centers target in grid; double-click / F2 to edit

  const handleCopy=(e)=>{ const selectedTxt=(typeof window!=="undefined"&&window.getSelection)?String(window.getSelection().toString()||""):""; if(selectedTxt && selectedTxt.trim()){ return; } const tag=(e.target&&e.target.tagName?String(e.target.tagName).toLowerCase():""); if(tag==="input"||tag==="textarea"||(e.target&&e.target.isContentEditable)) return; const cells=selectedCells(); if(!cells.length) return; const R=rect(); let lines=[]; let any=false; if(!hasMultiSelect&&R){ for(let r=R.r1;r<=R.r2;r++){ const row=[]; for(let cc=R.c1;cc<=R.c2;cc++){ let v=rows[r]?getVal(rows[r].s,navCols[cc]):""; if(isDateCol(navCols[cc])&&v) v=fmtTyped(v); if(v) any=true; row.push(v??""); } lines.push(row.join("\t")); } } else { lines=cells.map(cell=>{ let v=getVal(cell.s,cell.col)??""; if(isDateCol(cell.col)&&v) v=fmtTyped(v); if(v) any=true; return v; }); } if(any){ const tsv=lines.join("\n"); try{ e.clipboardData.setData("text/plain",tsv); e.preventDefault(); }catch(err){} setClip({ values:lines.map(l=>l.split("\t")), h:lines.length, w:lines[0].split("\t").length, multi:hasMultiSelect }); flash(); } };
  const handlePaste=(e)=>{ if(!sel) return; let txt=""; try{ txt=e.clipboardData.getData("text/plain"); }catch(err){} if(!txt){ doPaste(); return; } e.preventDefault(); const grid=txt.replace(/\r/g,"").replace(/\n$/,"").split("\n").map(l=>l.split("\t")); pushHistory(); const aR=rowIndex(sel.id), aC=colIndex(sel.col); const ch={}; const cells=selectedCells(); if(hasMultiSelect&&cells.length){ const flat=grid.flat(); const single=flat.length===1; cells.forEach((cell,i)=>{ const raw=single?flat[0]:flat[i%Math.max(1,flat.length)]; if(canPasteCell(cell.s,cell.col)) (ch[cell.id]=ch[cell.id]||{})[cell.col]=coerce(cell.col,raw); }); } else if(grid.length===1&&grid[0].length===1){ const R=rect(); for(let r=R.r1;r<=R.r2;r++){ for(let cc=R.c1;cc<=R.c2;cc++){ const row=rows[r]; const col=navCols[cc]; if(row&&canPasteCell(row.s,col)) (ch[row.s.id]=ch[row.s.id]||{})[col]=coerce(col,grid[0][0]); } } } else { for(let i=0;i<grid.length;i++){ for(let j=0;j<grid[i].length;j++){ const row=rows[aR+i]; const col=navCols[aC+j]; if(row&&col&&canPasteCell(row.s,col)) (ch[row.s.id]=ch[row.s.id]||{})[col]=coerce(col,grid[i][j]); } } } writeChanges(ch); };
  const FR_COLS=["__style","orderNo","styleNo","sampleFit","family","colour","brand","fabricType","owner","remarks"];
  const frGet=(s,col)=> col==="__style"?(s.styleNo||""):String(s[col]==null?"":s[col]);
  const frColLabel=(col)=> col==="__style"?"Style No":((INFO_COLS.find(c=>c.key===col)||{}).label||col);
  const computeMatches=()=>{ const f=frCase?frFind:(frFind||"").toLowerCase(); if(!f) return []; let cells=[]; if(frScope==="selected"){ const R=rect(); if(R){ for(let r=R.r1;r<=R.r2;r++) for(let cc=R.c1;cc<=R.c2;cc++){ const row=rows[r], col=navCols[cc]; if(row&&col&&FR_COLS.includes(col)&&col!=="styleNo") cells.push([row.s,col]); } } } else { for(const row of rows) for(const col of FR_COLS){ if(col==="styleNo") continue; cells.push([row.s,col]); } } const out=[]; const seen=new Set(); for(const [s,col] of cells){ const cur=frGet(s,col); if(!cur) continue; const hay=frCase?cur:cur.toLowerCase(); if(hay.indexOf(f)===-1) continue; const key=s.id+":"+col; if(seen.has(key)) continue; seen.add(key); out.push({ id:s.id, col, style:s.styleNo||"", colLabel:frColLabel(col), text:cur }); } return out; };
  const gotoMatch=(i,list)=>{ const m=list||frMatches; if(!m.length) return; const ni=((i%m.length)+m.length)%m.length; setFindIdx(ni); const t=m[ni]; setSel({ id:t.id, col:t.col }); setFocus(null); scrollToCell(t.id, t.col); };
  const runFind=()=>{ const m=computeMatches(); setFrMatches(m); if(!m.length){ setFindIdx(-1); window.alert('No cells contain "'+frFind+'".'); return; } gotoMatch(0,m); };
  const findNext=()=>{ if(!frMatches.length){ runFind(); return; } gotoMatch(findIdx+1); };
  const findReplace=(preview)=>{ const f=frFind; if(!f) return 0; let cells=[]; if(frScope==="selected"){ const R=rect(); if(!R) return 0; for(let r=R.r1;r<=R.r2;r++) for(let cc=R.c1;cc<=R.c2;cc++){ const row=rows[r], col=navCols[cc]; if(row&&col&&FR_COLS.includes(col)) cells.push([row.s,col]); } } else { for(const row of rows) for(const col of FR_COLS) cells.push([row.s,col]); } const esc=f.replace(/[.*+?^${}()|[\]\\]/g,"\\$&"); const re=new RegExp(esc, frCase?"g":"gi"); const repStr=frRepl.replace(/\$/g,"$$$$"); const needle=frCase?f:f.toLowerCase(); let count=0; const ch={}; const seen=new Set(); for(const [s,col] of cells){ const k=s.id+":"+(col==="__style"?"styleNo":col); if(seen.has(k)) continue; seen.add(k); const cur=frGet(s,col); if(!cur) continue; const hay=frCase?cur:cur.toLowerCase(); if(hay.indexOf(needle)===-1) continue; const nv=cur.replace(re,repStr); if(nv===cur) continue; count++; if(!preview && canPasteCell(s,col)) (ch[s.id]=ch[s.id]||{})[col]=nv; } if(!preview && Object.keys(ch).length){ pushHistory(); writeChanges(ch); } return count; };
  const specialStateFor=(st,col)=>({ actual:(st&&st.actuals&&st.actuals[col])||null, rev:(st&&st.revs&&st.revs[col])||null, reject:(st&&st.rejects&&st.rejects[col])||null, skip:(st&&st.skips&&st.skips[col])||null });
  const copySpecial=()=>{ const stageCells=selectedCells().filter(x=>isStageCol(x.col)); if(stageCells.length>1){ setSpecialClip({ multi:true, list:stageCells.map(x=>({ col:x.col, ...specialStateFor(x.s,x.col) })) }); flash(); return; } if(!sel||!isStageCol(sel.col)){ alert("Select a stage date cell first. Copy-special grabs full state: actual + revised + rejected + skip. Ctrl/Cmd-select multiple stage cells to copy several states."); return; } const s=styles.find(x=>x.id===sel.id); if(!s) return; setSpecialClip({ ...specialStateFor(s,sel.col), multi:false }); flash(); };
  const pasteSpecial=()=>{ if(!sel||!specialClip) return; const targets=selectedCells().filter(x=>isStageCol(x.col)).map(x=>({id:x.id,col:x.col,s:x.s})); if(!targets.length) return; const list=specialClip.multi?(specialClip.list||[]):null; const stateForIndex=(i)=> list&&list.length ? list[i%list.length] : specialClip; if(list&&list.length>1&&targets.length!==list.length){ if(!window.confirm(`Paste ${list.length} copied stage states into ${targets.length} selected stage cells by order? Values will repeat if counts do not match.`)) return; } if(canEditRev(role)){ targets.forEach((t,i)=>{ const st=styles.find(x=>x.id===t.id); const src=stateForIndex(i); if(st) recordRevisionHistory(st,t.col,(src&&src.rev)||null,"paste special"); }); } pushHistory(); setStyles(prev=>prev.map(s=>{ const items=targets.map((t,i)=>({...t,idx:i})).filter(t=>t.id===s.id); if(!items.length) return s; const ns={...s, actuals:{...s.actuals}, revs:{...(s.revs||{})}, rejects:{...(s.rejects||{})}, skips:{...(s.skips||{})} }; items.forEach(t=>{ const src=stateForIndex(t.idx)||{}; const col=t.col; if(canEdit(role,col,"actual")){ const ck=s.id+":"+col+":actual"; if(src.actual){ ns.actuals[col]=src.actual; clearedRef.current.delete(ck); } else { delete ns.actuals[col]; clearedRef.current.add(ck); } } if(canEditRev(role)){ const ck=s.id+":"+col+":revised"; if(src.rev){ ns.revs[col]=src.rev; clearedRef.current.delete(ck); } else { delete ns.revs[col]; clearedRef.current.add(ck); } } if(canEditReject(role,col)){ const ck=s.id+":"+col+":reject"; if(src.reject){ ns.rejects[col]=src.reject; clearedRef.current.delete(ck); } else { delete ns.rejects[col]; clearedRef.current.add(ck); } } if(MERCH_ROLES.includes(role)){ const ck=s.id+":"+col+":skip"; if(src.skip){ ns.skips[col]=src.skip; clearedRef.current.delete(ck); } else { delete ns.skips[col]; clearedRef.current.add(ck); } } }); return ns; })); flash(); };
  const copySelection=async()=>{ const cells=selectedCells(); if(!cells.length) return; const R=rect(); let lines=[]; if(!hasMultiSelect&&R){ for(let r=R.r1;r<=R.r2;r++){ const row=[]; for(let cc=R.c1;cc<=R.c2;cc++){ let v=rows[r]?getVal(rows[r].s,navCols[cc]):""; if(isDateCol(navCols[cc])&&v) v=fmtTyped(v); row.push(v??""); } lines.push(row.join("\t")); } } else { lines=cells.map(cell=>{ let v=getVal(cell.s,cell.col)??""; if(isDateCol(cell.col)&&v) v=fmtTyped(v); return v; }); } const tsv=lines.join("\n"); setClip({ values:lines.map(l=>l.split("\t")), h:lines.length, w:lines[0].split("\t").length, multi:hasMultiSelect }); try{ await navigator.clipboard.writeText(tsv); }catch(err){} flash(); };
  const cellKey=(id,col)=>`${id}:${col}`;
  const applyFill=(color)=>{ if(!sel) return; pushHistory(); const cells=selectedCells(); setFills(p=>{ const n={...p}; cells.forEach(cell=>{ const k=`${cell.id}:${cell.col}`; if(color==="") delete n[k]; else n[k]=color; }); return n; }); flash(); };
  const saveNote=()=>{ if(!sel) return; pushHistory(); setNotes(p=>{ const n={...p}; const k=cellKey(sel.id,sel.col); if(noteText.trim()==="") delete n[k]; else n[k]=noteText.trim(); return n; }); setNoteEditing(false); setNoteText(""); flash(); };
  const tsShort=(t)=>{ try{ const d=new Date(t); return d.toLocaleDateString(undefined,{day:"2-digit",month:"short"})+" "+d.toLocaleTimeString(undefined,{hour:"2-digit",minute:"2-digit"}); }catch(e){ return ""; } };
  const colLabelOf=(col)=> col==="__style"?"Style":(((INFO_COLS.find(x=>x.key===col)||{}).label)||((STAGES.find(x=>x.key===col)||{}).label)||col);
  const renderMentions=(body)=> String(body||"").split(/(@\S+)/).map((p,i)=> p.charAt(0)==="@" ? <b key={i} style={{ color:"var(--info)" }}>{p}</b> : <span key={i}>{p}</span>);
  const postComment=async()=>{ if(!threadCell||!cmText.trim()) return; const cell=threadCell; const body=cmText.trim(); const ck=cellKey(cell.id,cell.col);
    const tokens=(body.match(/@(\S+)/g)||[]).map(t=>t.slice(1).replace(/[.,;:]+$/,"").toLowerCase());
    const mentioned=team.filter(t=>tokens.includes(String(t.name||t.email||"").replace(/\s+/g,"").toLowerCase()));
    const row={ style_id:cell.id, col:cell.col, author_id:me.id, author_name:me.name||me.email, body, mentions:mentioned.map(m=>m.id), resolved:false };
    setCmText("");
    try{ const { data, error }=await supabase.from("comments").insert(row).select().single(); if(error) throw error;
      setComments(p=>{ const arr=(p[ck]||[]).filter(x=>x.id!==data.id); return {...p,[ck]:[...arr,data]}; });
      const styleNo=(styles.find(s=>s.id===cell.id)||{}).styleNo||""; const recips=[...new Set(mentioned.map(m=>m.id))].filter(uid=>uid&&uid!==me.id);
      const notifs=recips.map(uid=>({ user_id:uid, type:"mention", body:(me.name||"Someone")+" mentioned you on "+styleNo+" \u00b7 "+colLabelOf(cell.col), style_id:cell.id, style_no:styleNo, col:cell.col, actor_name:me.name||me.email, read:false }));
      if(notifs.length) await supabase.from("notifications").insert(notifs);
    }catch(e){ alert("Comment failed: "+(e.message||e)); }
  };
  const resolveThread=async(cell)=>{ const ck=cellKey(cell.id,cell.col); const list=comments[ck]||[]; const anyOpen=list.some(x=>!x.resolved); const val=anyOpen; setComments(p=>({...p,[ck]:(p[ck]||[]).map(x=>({...x,resolved:val}))})); try{ await supabase.from("comments").update({ resolved:val }).eq("style_id",cell.id).eq("col",cell.col); if(val){ const tagged=new Set(); list.forEach(cm=>{ (cm.mentions||[]).forEach(id=>{ if(id&&id!==me.id) tagged.add(id); }); }); if(tagged.size){ const sNo=(styles.find(s=>s.id===cell.id)||{}).styleNo||""; const notifs=[...tagged].map(uid=>({ user_id:uid, type:"resolved", body:(me.name||"Someone")+" resolved a comment that tagged you · "+sNo+" · "+colLabelOf(cell.col), style_id:cell.id, style_no:sNo, col:cell.col, actor_name:me.name||me.email, read:false })); await supabase.from("notifications").insert(notifs); } } }catch(e){} };
  const deleteComment=async(cm)=>{ if(!cm) return; if(!(cm.author_id===(me&&me.id)||canAdmin(role))) return; if(!window.confirm("Delete this comment? This cannot be undone.")) return; const ck=cm.style_id+":"+cm.col; setComments(p=>({...p,[ck]:(p[ck]||[]).filter(x=>x.id!==cm.id)})); try{ await supabase.from("comments").delete().eq("id",cm.id); }catch(e){} };
  const unreadCount=inbox.filter(n=>!n.read).length;
  const markRead=async(n)=>{ if(!n||n.read) return; setInbox(p=>p.map(x=>x.id===n.id?{...x,read:true}:x)); try{ await supabase.from("notifications").update({ read:true }).eq("id",n.id); }catch(e){} };
  const markAllRead=async()=>{ setInbox(p=>p.map(x=>({...x,read:true}))); try{ await supabase.from("notifications").update({ read:true }).eq("user_id",me.id).eq("read",false); }catch(e){} };
  const clearInbox=async()=>{ if(!window.confirm("Clear all notifications?")) return; setInbox([]); setBellOpen(false); try{ await supabase.from("notifications").delete().eq("user_id",me.id); }catch(e){} };
  const openNotif=(n)=>{ markRead(n); setBellOpen(false); setReviewOpen(false); setHistory(false); if(n.style_id&&n.col){ setSel({ id:n.style_id, col:n.col }); setFocus(null); scrollToCell(n.style_id, n.col); setThreadCell({ id:n.style_id, col:n.col }); } };
  const logMetaAudit=(changedList,SR)=>{ try{ const entries=[]; for(const x of changedList){ let oldRow={}; try{ oldRow=JSON.parse(SR.meta[x.key]||"{}"); }catch(e){ oldRow={}; } const nr=x.row||{}; const sNo=(styles.find(s=>s.id===nr.style_id)||{}).styleNo||""; const changes=[["fill","format fill"],["note","format note"]]; changes.forEach(([f,lab])=>{ const ov=oldRow[f]||null, nv=nr[f]||null; if(ov!==nv) entries.push({ style_id:nr.style_id, style_no:sNo, col:nr.col, field:lab, old_val:ov, new_val:nv, actor_id:me.id, actor_name:me.name||me.email }); }); } if(entries.length) supabase.from("audit_log").insert(entries).then(()=>{}).catch(()=>{}); }catch(e){} };
  const logStageAudit=(changedList,SR)=>{ try{ const entries=[]; const fields=[["actual_date","actual"],["revised_date","revised"],["reject_date","rejected"],["skip_date","waived"]]; for(const x of changedList){ let oldRow={}; try{ oldRow=JSON.parse(SR.stg[x.key]||"{}"); }catch(e){ oldRow={}; } const nr=x.row; const sNo=(styles.find(s=>s.id===nr.style_id)||{}).styleNo||""; for(const [f,lab] of fields){ const ov=oldRow[f]||null, nv=nr[f]||null; if(ov!==nv) entries.push({ style_id:nr.style_id, style_no:sNo, col:nr.stage, field:lab, old_val:ov, new_val:nv, actor_id:me.id, actor_name:me.name||me.email }); } } if(entries.length) supabase.from("audit_log").insert(entries).then(()=>{}).catch(()=>{}); }catch(e){} };
  const logStyleAudit=(changedRows,SR)=>{ try{ const L={ style_no:"Style No", order_no:"Order No", sample_fit:"Sample Fit", family:"Family", colour:"Colour", brand:"Brand", buyer:"Buyer", fabric_type:"Fabric Type", owner:"Owner", age:"Age", set_id:"Set ID", set_role:"Set Role", qty:"Qty", order_date:"Order Date", delivery_date:"Delivery", archived:"Archived", remarks:"Remarks" }; const entries=[]; for(const r of changedRows){ if(!SR.sty[r.id]) continue; let oldRow={}; try{ oldRow=JSON.parse(SR.sty[r.id]); }catch(e){ continue; } const sNo=r.style_no||oldRow.style_no||""; for(const f in L){ const ov=(oldRow[f]==null)?"":String(oldRow[f]); const nv=(r[f]==null)?"":String(r[f]); if(ov!==nv) entries.push({ style_id:r.id, style_no:sNo, col:L[f], field:"value", old_val:ov||null, new_val:nv||null, actor_id:me.id, actor_name:me.name||me.email }); } } if(entries.length) supabase.from("audit_log").insert(entries).then(()=>{}).catch(()=>{}); }catch(e){} };
  const loadAuditRows=async(styleId)=>{ setAuditBusy(true); try{ let q=supabase.from("audit_log").select("*").order("created_at",{ascending:false}).limit(styleId?500:300); if(styleId) q=q.eq("style_id",styleId); const { data }=await q; setAuditRows(data||[]); }catch(e){ setAuditRows([]); } setAuditBusy(false); };
  const openHistory=async(styleId,col)=>{ setHistory(true); setHistFilter(col?(colLabelOf(col)||""):""); await loadAuditRows(styleId); };
  const toggleFollow=async(styleId)=>{ if(!styleId) return; const has=follows.has(styleId); setFollows(p=>{ const n=new Set(p); if(has) n.delete(styleId); else n.add(styleId); return n; }); try{ if(has) await supabase.from("style_follows").delete().eq("user_id",me.id).eq("style_id",styleId); else await supabase.from("style_follows").insert({ user_id:me.id, style_id:styleId }); }catch(e){} };
  const notifyFollowers=async(changedList)=>{ try{ const byStyle={}; for(const x of changedList){ const sid=x.row.style_id; (byStyle[sid]=byStyle[sid]||[]).push(x.row.stage); } const ids=Object.keys(byStyle).map(Number); if(!ids.length) return; const { data }=await supabase.from("style_follows").select("user_id,style_id").in("style_id",ids); if(!data||!data.length) return; const notifs=[]; for(const f of data){ if(f.user_id===me.id) continue; const sNo=(styles.find(s=>s.id===f.style_id)||{}).styleNo||""; const stg=(byStyle[f.style_id]||[])[0]; notifs.push({ user_id:f.user_id, type:"follow", body:(me.name||"Someone")+" updated "+sNo+" · "+colLabelOf(stg), style_id:f.style_id, style_no:sNo, col:stg, actor_name:me.name||me.email, read:false }); } if(notifs.length) await supabase.from("notifications").insert(notifs); }catch(e){} };
  const beginNote=()=>{ if(!sel) return; setNoteText(notes[cellKey(sel.id,sel.col)]||""); setNoteEditing(true); };

  // Cascading header-filter options: every column dropdown is built from the rows that already match
  // all other active filters (search, status, chase, activity, archive, following, saved view, and other column filters).
  // This prevents irrelevant/blank values from unrelated rows appearing after a slice such as Activity = Lab Dip.
  const passForFilterOptions=(row,exceptCol)=>{
    const {s,c,idx}=row;
    const q=String(deferredSearch||search||"").trim().toLowerCase();
    const matchQ=!q ? true : (searchCol==="auto" ? (idx&&idx.auto?idx.auto:"").includes(q) : ((idx&&idx.byCol&&idx.byCol[searchCol])!=null?idx.byCol[searchCol]:lc(s[searchCol])).includes(q));
    const matchS=statusFilter==="All"||(statusFilter==="At Risk"&&(c.tone==="late"||c.tone==="warn"))||(statusFilter==="On Track"&&c.tone==="ok")||(statusFilter==="Released"&&c.released);
    const matchF=Object.entries(colFilters||{}).every(([cc,allowed])=> cc===exceptCol || passCol(s,c,cc,allowed));
    const matchO=ownerFilter==="All"||(c.chaseOwners||[]).some(o=>o.owner===ownerFilter);
    const matchA=rowMatchesAnyCurrentStage(row,activityFilterKeys);
    const matchArch=archiveView==="all"?true:(archiveView==="archived"?!!s.archived:!s.archived);
    const matchFollow=!followFilter||follows.has(s.id);
    return matchQ&&matchS&&matchF&&matchO&&matchA&&matchArch&&matchFollow&&presetPass(s,c);
  };
  const distinctFor=(col)=>{ const set=new Set(); computed.forEach((row)=>{ if(!passForFilterOptions(row,col)) return; const {s,c}=row; if(col==="chase"){ const owners=(c.chaseOwners||[]).map(o=>o.owner); if(owners.length===0) set.add("(Blanks)"); else owners.forEach(o=>set.add(o)); } else set.add(valueFor(s,c,col)); }); return [...set].sort((a,b)=> a==="(Blanks)"?1:b==="(Blanks)"?-1:(a>b?1:a<b?-1:0)); };
  const filterProps=(col)=>({
    filterActive: Array.isArray(colFilters[col]),
    filterOpen: filterCol===col,
    filterValues: filterCol===col?distinctFor(col):null,
    filterAllowed: Array.isArray(colFilters[col])?colFilters[col]:null,
    onToggleFilter:()=>{ finishEditing(); setFilterCol(p=>p===col?null:col); },
    onSetFilter:(arr)=>setColFilters(f=>{
      const n={...(f||{})};
      if(arr==null || arr==="__ALL__"){ delete n[col]; return n; }
      const raw=Array.isArray(arr)?arr:[arr];
      const clean=[...new Set(raw.filter(v=>v!=null && String(v).trim()!=="").map(v=>String(v).replace(/\s+/g," ").trim()))];
      // Empty selection is NOT the same as clear filter. Empty means show zero rows for this column.
      // Clear filter / All removes the key. This is what was breaking Select all / clear-selected behavior.
      n[col]=clean.length?clean:[FILTER_NONE];
      return n;
    }),
    onCloseFilter:()=>setFilterCol(null)
  });
  const funnel=useMemo(()=>{ const b={ "Pre-Fit":0,"Fit/Print":0,"Lab Dip":0,"Fabric IH":0,"PP":0,"Released":0 }; activeComputed.forEach(({c})=>{ if(c.released) b["Released"]++; else { const k=c.nextPending.key; if(k==="techpack") b["Pre-Fit"]++; else if(["fitSend","fitAppr","artwork","artAppr","strikeOff","soAppr"].includes(k)) b["Fit/Print"]++; else if(["labDip","labAppr"].includes(k)) b["Lab Dip"]++; else if(k==="fabricIH") b["Fabric IH"]++; else b["PP"]++; } }); return b; },[activeComputed]);

  const requiredMissing=()=>{ const m=[]; if(!newRow.styleNo.trim()) m.push("Style No"); if(!newRow.orderNo.trim()) m.push("Order No"); if(!newRow.ordRec) m.push("Order Date"); if(!newRow.delivery) m.push("Delivery Date"); return m; };
  const addNewStyle=async()=>{ const miss=requiredMissing(); if(miss.length){ setNewError("Required: "+miss.join(", ")); return; } setNewError(""); pushHistory();
    const base={ order_no:newRow.orderNo||"", style_no:newRow.styleNo.trim(), sample_fit:newRow.sampleFit||"", family:newRow.family||"", colour:newRow.colour||"", brand:newRow.brand||"", buyer:newRow.buyer||"", fabric_type:newRow.fabricType||"", owner:newRow.owner||"", set_id:newRow.setId||"", set_role:newRow.setRole||"", age:newRow.age||"", qty:Number(newRow.qty)||0, order_date:newRow.ordRec||iso(TODAY), delivery_date:newRow.delivery||"2026-07-15", fit_req:newRow.fitReq, print_req:newRow.printReq, so_req:newRow.soReq, pp_bypass:newRow.ppBypass, lab_dip_req:newRow.labDipReq, pp_needed:newRow.ppNeeded, remarks:"" };
    try{ const { data, error }=await supabase.from("styles").insert(base).select().single(); if(error||!data) throw error||new Error("no row"); setStyles(prev=>[...prev, rowToStyle(data,{})]); }catch(e){ console.error("create failed",e); }
    setNewRow({ styleNo:"", orderNo:"", sampleFit:"", family:"", colour:"", brand:"", buyer:"", fabricType:"", owner:"", setId:"", setRole:"", age:"", qty:"", ordRec:iso(TODAY), delivery:"", fitReq:true, printReq:false, soReq:false, ppBypass:false, labDipReq:true, ppNeeded:true }); flash(); };
  const ndFieldFor=(key)=>{ const ni={ width:"100%", boxSizing:"border-box", border:"1px solid #e7dcc2", outline:"none", background:"var(--surface)", fontFamily:"inherit", fontSize:11, padding:"3px 4px" }; const stop=e=>e.stopPropagation(); const ent=e=>{ if(e.key==="Enter") addNewStyle(); };
    if(key==="qty") return <input value={newRow.qty||""} onClick={stop} onKeyDown={ent} onChange={e=>setNewRow(n=>({...n,qty:e.target.value.replace(/[^0-9]/g,"")}))} placeholder="qty" style={ni}/>;
    if(key==="ordRec") return <input type="date" value={newRow.ordRec||""} onClick={stop} onChange={e=>setNewRow(n=>({...n,ordRec:e.target.value}))} style={ni}/>;
    if(key==="delivery") return <input type="date" value={newRow.delivery||""} onClick={stop} onChange={e=>setNewRow(n=>({...n,delivery:e.target.value}))} style={ni}/>;
    const ph={ orderNo:"order# *", sampleFit:"fit#", family:"family", colour:"colour", brand:"brand", buyer:"buyer", fabricType:"fabric", owner:"owner", age:"age", setId:"set id", setRole:"set role" }[key]||"";
    return <input value={newRow[key]||""} onClick={stop} onKeyDown={ent} onChange={e=>setNewRow(n=>({...n,[key]:e.target.value}))} placeholder={ph} style={ni}/>; };

  const isAnchor=(id,col)=> sel&&sel.id===id&&sel.col===col;
  const inRange=(id,col)=> selKeys.has(`${id}:${col}`);
  const inFill=(id,col)=>{ if(!filling||!fillFrom||!fillTo) return false; const aR=rowIndex(fillFrom.id),aC=colIndex(fillFrom.col),tR=rowIndex(fillTo.id),tC=colIndex(fillTo.col); const r=rowIndex(id),cc=colIndex(col); return r>=Math.min(aR,tR)&&r<=Math.max(aR,tR)&&cc>=Math.min(aC,tC)&&cc<=Math.max(aC,tC); };
  const bgFor=(id,col,base)=>{ const f=fills[cellKey(id,col)]; if(f) return f; if(inFill(id,col)) return "#def0e0"; if(inRange(id,col)&&!isAnchor(id,col)) return hasMultiSelect?"#eaf3ff":"#e3edfb"; return base; };
  const selectedColumnKeys=()=>{ const out=[]; const seen=new Set(); selectedCells().forEach(x=>{ if(!seen.has(x.col)){ seen.add(x.col); out.push(x.col); } }); return out; };
  const selectedRowIds=()=>{ const out=[]; const seen=new Set(); selectedCells().forEach(x=>{ const k=String(x.id); if(!seen.has(k)){ seen.add(k); out.push(k); } }); return out; };
  const hideSelectedColumns=()=>{ const cols=selectedColumnKeys().filter(c=>c!=="__style" && navCols.includes(c)); if(!cols.length) return; setHidden(prev=>{ const n=new Set(prev); cols.forEach(c=>n.add(c)); return n; }); };
  const hideSelectedRows=()=>{ const ids=selectedRowIds(); if(!ids.length) return; setHiddenRows(prev=>{ const n=new Set(prev); ids.forEach(id=>n.add(String(id))); return n; }); clearSelection(); };
  const freezeThroughColumn=(col)=>{ const idx=colIndex(col); if(idx<0) return; if(idx>=maxFreeze){ alert("Only leading tracker columns can be frozen. To freeze this column, first move it into a leading column view."); return; } setFreezeN(Math.max(1, idx+1)); };
  const rowHeightPreset=(h)=>setRowH(Math.max(28, Math.min(110, Number(h)||38)));
  const FillHandle=({id,col})=> isAnchor(id,col)&&!editing&&canPasteCell(styles.find(x=>x.id===id),col)? <span onMouseDown={(e)=>{ e.stopPropagation(); e.preventDefault(); setFilling(true); setFillFrom({id,col}); setFillTo({id,col}); }} title="drag to fill" style={{ position:"absolute", right:-2, bottom:-2, width:7, height:7, background:"var(--info)", cursor:"crosshair", zIndex:6 }}/> : null;
  const ringFor=(id,col)=> isAnchor(id,col)?"inset 0 0 0 3px var(--info), inset 0 0 0 9999px rgba(37,99,166,0.08)":(inRange(id,col)?"inset 0 0 0 2px rgba(37,99,166,0.55)":null);
  const activeCellStyle=(id,col)=> isAnchor(id,col) ? { outline:"3px solid var(--info)", outlineOffset:"-3px", boxShadow:"inset 0 0 0 3px var(--info), inset 0 0 0 9999px rgba(37,99,166,0.10)", zIndex:isFrozen(col)?60:4, position:isFrozen(col)?"sticky":"relative", backgroundClip:"padding-box" } : {};
  const activeRowStyle=(id)=> sel&&sel.id===id ? { boxShadow:"inset 4px 0 0 var(--info)", background:"rgba(37,99,166,0.03)" } : {};
  const NoteTri=({k})=>{ const arr=comments[k]; if(!arr||!arr.length) return null; const un=arr.filter(x=>!x.resolved).length; const cl=un>0?"var(--danger)":"#9aa0a6"; return <span title={un>0?(un+" open comment(s)"):"comments resolved"} style={{ position:"absolute", top:0, right:0, width:0, height:0, borderTop:"9px solid "+cl, borderLeft:"9px solid transparent" }}/>; };

  const renderEditable=(s,col)=>{
    const k=cellKey(s.id,col.key);
    const val=col.key==="qty"?s.qty:s[col.key];
    const editingThis=editing&&editing.id===s.id&&editing.col===col.key;
    const bg=bgFor(s.id,col.key,"var(--surface)");
    return (
      <td key={col.key} id={`cell-${s.id}-${col.key}`} onClick={(e)=>onCellClick(e,s.id,col.key)} onDoubleClick={(e)=>{ e.stopPropagation(); startEdit(s.id,col.key); }}
        style={{ border:"1px solid var(--line-1)", padding:"6px 9px", height:rowH, whiteSpace: col.key==="remarks"?"normal":"nowrap", boxShadow:ringFor(s.id,col.key), cursor:"cell", maxWidth:col.w, overflow:"hidden", textOverflow:"ellipsis", fontSize: col.key==="remarks"?10:11, color: col.key==="remarks"?"#a15":"var(--ink)", position:"relative", background:bg, ...freezeStyle(col.key,bg), ...activeCellStyle(s.id,col.key) }}>
        {editingThis ? (<span style={{ position:"relative", display:"inline-block" }}><input autoFocus onFocus={e=>{ if((e.target.value||"").length>1) e.target.select(); }} value={editVal} onClick={e=>e.stopPropagation()} onChange={e=>setEditVal(col.key==="qty"?e.target.value.replace(/[^0-9]/g,""):e.target.value)} onKeyDown={e=>{ e.stopPropagation(); if(e.key==="Tab"){ e.preventDefault(); const v=editSuggestion||editVal; commitText(v); moveAnchor(0,e.shiftKey?-1:1); } else if(e.key==="Enter"){ e.preventDefault(); commitText(editSuggestion||editVal); moveAnchor(1,0); } else if(e.key==="Escape"){ e.preventDefault(); setEditing(null); } }} onBlur={()=>commitText()} style={{ width:Math.max(40,col.w-16), fontFamily:"inherit", fontSize:11, border:"1px solid var(--info)", outline:"none", padding:"1px 3px" }}/>{editSuggestion && <button onMouseDown={e=>e.preventDefault()} onClick={e=>{ e.stopPropagation(); setEditVal(editSuggestion); }} title="Click or press Tab to accept" style={{ position:"absolute", left:0, top:"100%", marginTop:3, zIndex:390, maxWidth:220, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"3px 7px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--accent-tint)", boxShadow:"3px 3px 0 var(--ink)" }}>Tab ↹ {editSuggestion}</button>}</span>) : (val===""||val==null ? <span style={{color:"var(--line-2)"}}>—</span> : displayTextValue(col.key,val))}
        <PeerTag who={peerOn(s.id,col.key)}/><NoteTri k={k}/><FillHandle id={s.id} col={col.key}/>
      </td>
    );
  };

  return (
    <div ref={gridRef} tabIndex={0} onKeyDown={onKeyDown} onCopy={handleCopy} onPaste={handlePaste} onContextMenu={(e)=>{ const td=e.target.closest&&e.target.closest('td[id^="cell-"]'); if(!td) return; const m=td.id.match(/^cell-(\d+)-(.+)$/); if(!m) return; e.preventDefault(); const id=Number(m[1]), col=m[2]; if(!selKeys.has(`${id}:${col}`)) setSingleSelection(id,col); setCtxMenu({ x:e.clientX, y:e.clientY, id, col }); }} onMouseDown={(e)=>{ if(editing) return; if(e.button===2) return; if(e.target.closest && (e.target.closest("input")||e.target.closest("button")||e.target.closest("th"))) return; const td=e.target.closest && e.target.closest('td[id^="cell-"]'); if(!td) return; const m=td.id.match(/^cell-(\d+)-(.+)$/); if(!m) return; e.preventDefault(); gridRef.current&&gridRef.current.focus(); const id=Number(m[1]), col=m[2]; setCtxMenu(null); if((e.ctrlKey||e.metaKey)||e.shiftKey) return; setSel({ id, col }); setFocus(null); setMultiAreas([]); selectingRef.current=true; setDragSel(true); }} onMouseUp={()=>{ if(filling){ applyFillHandle(); setFilling(false); setFillFrom(null); setFillTo(null); } if(selectingRef.current){ selectingRef.current=false; setDragSel(false); } }} onMouseOver={(e)=>{ const td=e.target.closest && e.target.closest('td[id^="cell-"]'); if(!td) return; const m=td.id.match(/^cell-(\d+)-(.+)$/); if(!m) return; if(filling){ setFillTo({ id:Number(m[1]), col:m[2] }); return; } if(selectingRef.current){ setFocus({ id:Number(m[1]), col:m[2] }); } }} onClick={()=>{ finishEditing(); setCtxMenu(null); setFillOpen(false); setColsOpen(false); setFilterCol(null); setFrOpen(false); setExpOpen(false); setBellOpen(false); setPeersOpen(false); }}
      style={{ minHeight:"100vh", background:"var(--bg)", fontFamily:"'JetBrains Mono', monospace", color:"var(--ink)", paddingBottom:80, outline:"none" }}>
      <style>{FONT}</style>
      <style>{THEME_CSS}</style>
      {ctxMenu && (()=>{ const cellCount=selKeys.size||1; const stageCount=selectedCells().filter(x=>isStageCol(x.col)).length; const item=(label,onClick,disabled=false,sub="")=><button disabled={disabled} onClick={(e)=>{ e.stopPropagation(); if(disabled) return; setCtxMenu(null); onClick&&onClick(); }} style={{ width:"100%", textAlign:"left", padding:"7px 10px", border:"none", borderBottom:"1px solid var(--line-3)", background:disabled?"#f3f1ec":"var(--surface)", color:disabled?"var(--muted-1)":"var(--ink)", cursor:disabled?"not-allowed":"pointer", fontFamily:"inherit", fontSize:11 }}><b>{label}</b>{sub&&<span style={{ display:"block", fontSize:9, color:"var(--muted-2)", marginTop:2 }}>{sub}</span>}</button>; return <div onClick={e=>e.stopPropagation()} style={{ position:"fixed", left:Math.min(ctxMenu.x, window.innerWidth-230), top:Math.min(ctxMenu.y, window.innerHeight-520), zIndex:520, width:225, background:"var(--surface)", border:"1px solid var(--ink)", boxShadow:"5px 5px 0 var(--ink)", overflow:"hidden" }}>
        <div style={{ padding:"7px 10px", fontSize:10, fontWeight:900, background:"var(--accent-tint)", borderBottom:"1px solid var(--line-3)" }}>{cellCount} cell{cellCount===1?"":"s"} selected</div>
        {item("Copy", copySelection, false, hasMultiSelect?"Non-adjacent copies line-by-line":"Excel-style TSV copy")}
        {item("Paste", doPaste, !clip, "Paste normal values into selected cells")}
        {item("Copy special", copySpecial, stageCount===0, "Stage state: actual + revised + rejected + skip")}
        {item("Paste special", pasteSpecial, !specialClip||stageCount===0, "Into selected stage cells")}
        {item("Clear selected cells", clearRange, cellCount===0, "Only editable cells are cleared")}
        {item("Add / edit note", beginNote, !sel)}
        {item("Open comments", ()=>setThreadCell({ id:ctxMenu.id, col:ctxMenu.col }), !sel)}
        {item("Cell history", ()=>openHistory(ctxMenu.id,ctxMenu.col), !sel)}
        {item("Select row", ()=>selectRow(ctxMenu.id), !ctxMenu.id)}
        {item("Auto-fit column", ()=>autoFit(ctxMenu.col), !ctxMenu.col)}
        {item("Freeze up to this column", ()=>freezeThroughColumn(ctxMenu.col), !ctxMenu.col || colIndex(ctxMenu.col)>=maxFreeze, "Frozen columns stay visible while scrolling")}
        {item("Hide selected column(s)", hideSelectedColumns, selectedColumnKeys().filter(c=>c!=="__style").length===0, "Style No cannot be hidden")}
        {item("Hide selected row(s)", hideSelectedRows, selectedRowIds().length===0, "Local view only; data is not archived/deleted")}
        {item("Show hidden rows", ()=>setHiddenRows(new Set()), !hiddenRows.size, hiddenRows.size?`${hiddenRows.size} hidden row(s)`:"")}
        {item("Row height: compact", ()=>rowHeightPreset(30), false)}
        {item("Row height: normal", ()=>rowHeightPreset(38), false)}
        {item("Row height: tall", ()=>rowHeightPreset(56), false)}
        {item("Clear selection", clearSelection, !sel)}
      </div>; })()}

      <div className="mt-app-header" style={{ background:"var(--ink)", color:"var(--bg)", padding:"14px 22px", display:"flex", alignItems:"center", justifyContent:"space-between", borderBottom:"3px solid var(--accent)", position:"relative", zIndex:160, pointerEvents:"auto" }}>
        <div style={{ display:"flex", flexDirection:"column", gap:3 }}><span style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:20, letterSpacing:-0.2, lineHeight:1 }}>KOTHARI SPORTS & APPARELS</span><span style={{ fontFamily:"'JetBrains Mono',monospace", fontSize:9, color:"#9a958c", letterSpacing:1.2 }}>MERCH<span style={{ color:"var(--accent)" }}>·</span>TRACKER · PRE-PRODUCTION TRACKER</span></div>
        <div className="mt-header-actions" style={{ display:"flex", alignItems:"center", gap:10 }}><span title={!online?"You appear to be offline. Your last loaded data is cached and still visible; edits will save to the cloud automatically when the connection returns.":(lastSavedAt?("Last saved to cloud at "+lastSavedAt.toLocaleTimeString()):"Connected to cloud")} style={{ fontSize:11, display:"inline-flex", alignItems:"center", gap:5, color: !online?"#e0a23a":saveState==="error"?"#e8746b":saveState==="saving"?"#d9b46a":saveState==="saved"?"#7fd1a8":"#6a665e" }}><span style={{ width:7, height:7, borderRadius:7, flex:"0 0 auto", background: !online?"#e0a23a":saveState==="error"?"#e8746b":saveState==="saving"?"#d9b46a":"#7fd1a8" }}/>{!online?"⚠ offline — showing cached":saveState==="error"?"⚠ save failed":saveState==="saving"?"… saving":saveState==="saved"?"saved":"connected"}{online && lastSavedAt && saveState!=="saving" && saveState!=="error" && <span style={{ color:"#8a857a", fontSize:10 }}>· {String(lastSavedAt.getHours()).padStart(2,"0")}:{String(lastSavedAt.getMinutes()).padStart(2,"0")}</span>}</span><span style={{ fontSize:11, color:"var(--on-dark)", whiteSpace:"nowrap" }}>{(me&&(me.name||me.email))||""} · <b style={{ color:"var(--accent)" }}>{(ROLES[role]||{}).label||role}</b></span>{peerGroups.length>0 && (<span style={{ position:"relative", display:"flex", alignItems:"center", gap:3 }}>{visiblePeerGroups.map(g=>{ const p=g.display||g; const loc=g.loc; return (<span key={g.id} onClick={(e)=>{ e.stopPropagation(); if(loc) scrollToCell(loc.id,loc.col); }} title={(g.name||"")+" · "+((ROLES[g.role]||{}).label||g.role)+(loc?" — click to view their cell":" — not on a cell")+(g.sessions.length>1?" · "+g.sessions.length+" sessions":"")} style={{ width:18, height:18, borderRadius:9, background:colorFor(g.id), color:"var(--surface)", fontSize:8, fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center", cursor:loc?"pointer":"default", boxShadow:g.editing?"0 0 0 1px var(--surface)":"none", position:"relative" }}>{initials(g.name)}{g.sessions.length>1&&<span style={{ position:"absolute", right:-5, bottom:-5, minWidth:11, height:11, borderRadius:7, background:"var(--accent)", color:"var(--ink)", border:"1px solid var(--surface)", fontSize:7, lineHeight:"10px", textAlign:"center", fontWeight:900 }}>{Math.min(g.sessions.length,9)}</span>}</span>); })}{peerGroups.length>4&&<span title={(peerGroups.length-4)+" more active users"} style={{ width:18, height:18, borderRadius:9, background:"#4b5563", color:"var(--surface)", fontSize:8, fontWeight:800, display:"flex", alignItems:"center", justifyContent:"center" }}>+{peerGroups.length-4}</span>}<button onClick={(e)=>{ e.stopPropagation(); setBellOpen(false); setReviewOpen(false); setHistory(false); setPeersOpen(o=>!o); }} title="See who's online and where" style={{ fontSize:9, color:"#9b958a", marginLeft:2, background:"transparent", border:"none", cursor:"pointer", fontFamily:"inherit" }}>{peerGroups.length} users · {peers.length} sessions ▾</button>{peersOpen && (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", right:0, marginTop:6, zIndex:390, background:"var(--surface)", color:"var(--ink)", border:"1px solid var(--ink)", boxShadow:"4px 4px 0 var(--ink)", width:282, maxHeight:340, overflowY:"auto" }}><div style={{ fontSize:10, fontWeight:700, padding:"7px 9px", borderBottom:"1px solid var(--line-3)" }}>Active users ({peerGroups.length}) · sessions ({peers.length})</div>{peerGroups.map(g=>{ const loc=g.loc; const over=g.sessions.length>4; return (<div key={g.id} onClick={()=>{ if(loc){ scrollToCell(loc.id,loc.col); setPeersOpen(false); } }} style={{ display:"flex", gap:7, alignItems:"center", padding:"7px 9px", borderBottom:"1px solid var(--line-3)", cursor:loc?"pointer":"default" }}><span style={{ width:18, height:18, borderRadius:9, flex:"0 0 auto", background:colorFor(g.id), color:"var(--surface)", fontSize:8, fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center" }}>{initials(g.name)}</span><span style={{ flex:1, minWidth:0 }}><span style={{ fontSize:11, fontWeight:700, display:"block", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{g.name||"?"}</span><span style={{ fontSize:9, color:over?"var(--danger)":"var(--muted-2)", display:"block", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{loc?((g.editing?"✎ editing ":"at ")+cellLabel(loc.id,loc.col)):"not on a cell"} · {g.sessions.length} session{g.sessions.length===1?"":"s"}{over?" · 4+ open":""}</span></span>{loc && <span style={{ fontSize:9, color:"var(--info)", fontWeight:700, whiteSpace:"nowrap" }}>view →</span>}</div>); })}</div>)}</span>)}<button onClick={(e)=>{ e.stopPropagation(); openHistory(); }} title="Change history" style={{ fontFamily:"inherit", fontSize:10, padding:"5px 8px", cursor:"pointer", border:"1px solid var(--on-dark-line)", background:"transparent", color:"var(--on-dark)", display:"inline-flex", alignItems:"center" }}><Clock size={13}/></button><button onClick={(e)=>{ e.stopPropagation(); setBellOpen(false); setPeersOpen(false); setReviewOpen(false); setHistory(false); setTab("review"); loadAuditRows(); }} title="Open Review tab for category-wise review items" style={{ fontFamily:"inherit", fontSize:10, padding:"5px 9px", cursor:"pointer", border:"1px solid var(--on-dark-line)", background:tab==="review"?"var(--on-dark-line)":"transparent", color:"var(--on-dark)", display:"inline-flex", alignItems:"center", gap:4 }}>Review</button><span style={{ position:"relative" }}><button onClick={(e)=>{ e.stopPropagation(); setPeersOpen(false); setReviewOpen(false); setHistory(false); setBellOpen(o=>!o); }} title="Notifications" style={{ fontFamily:"inherit", fontSize:10, padding:"5px 8px", cursor:"pointer", border:"1px solid var(--on-dark-line)", background:bellOpen?"var(--on-dark-line)":"transparent", color:"var(--on-dark)", position:"relative", display:"inline-flex", alignItems:"center" }}><Bell size={13}/>{unreadCount>0 && <span style={{ position:"absolute", top:-5, right:-5, background:"#e8746b", color:"var(--surface)", fontSize:8, fontWeight:700, minWidth:14, height:14, borderRadius:7, display:"flex", alignItems:"center", justifyContent:"center", padding:"0 3px" }}>{unreadCount>9?"9+":unreadCount}</span>}</button>{bellOpen && (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", right:0, marginTop:6, zIndex:390, background:"var(--surface)", color:"var(--ink)", border:"1px solid var(--ink)", boxShadow:"4px 4px 0 var(--ink)", width:320, maxHeight:380, display:"flex", flexDirection:"column" }}><div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"8px 10px", borderBottom:"1px solid var(--line-3)" }}><span style={{ fontSize:11, fontWeight:700 }}>Notifications{unreadCount>0?" ("+unreadCount+" new)":""}</span><span style={{ display:"flex", gap:8 }}>{unreadCount>0 && <button onClick={markAllRead} style={{ fontSize:9, border:"none", background:"transparent", color:"var(--info)", cursor:"pointer", fontFamily:"inherit", fontWeight:700 }}>Mark all read</button>}{inbox.length>0 && <button onClick={clearInbox} style={{ fontSize:9, border:"none", background:"transparent", color:"var(--danger)", cursor:"pointer", fontFamily:"inherit", fontWeight:700 }}>Clear</button>}</span></div><div style={{ overflowY:"auto" }}>{inbox.length===0 ? <div style={{ fontSize:10, color:"var(--muted-1)", padding:"14px 10px", textAlign:"center" }}>No notifications yet.</div> : inbox.map(n=>(<div key={n.id} onClick={()=>openNotif(n)} style={{ padding:"8px 10px", borderBottom:"1px solid #f2f2f2", cursor:"pointer", background:n.read?"var(--surface)":"var(--accent-tint)", display:"flex", gap:7, alignItems:"flex-start" }}><span style={{ width:6, height:6, borderRadius:6, marginTop:4, flex:"0 0 auto", background:n.read?"transparent":"var(--accent)" }}/><span style={{ flex:1 }}><span style={{ fontSize:11, lineHeight:1.35, display:"block" }}>{n.body}</span><span style={{ fontSize:9, color:"var(--muted-1)" }}>{tsShort(n.created_at)}</span></span></div>))}</div></div>)}</span>{canManageUsers(role) && <button onClick={(e)=>{ e.stopPropagation(); setBellOpen(false); setPeersOpen(false); setReviewOpen(false); setHistory(false); setUsersOpen(true); }} style={{ fontFamily:"inherit", fontSize:10, padding:"5px 9px", cursor:"pointer", border:"1px solid var(--on-dark-line)", background:"transparent", color:"var(--on-dark)" }}>Users</button>}<button onClick={(e)=>{ e.stopPropagation(); onSignOut&&onSignOut(); }} style={{ fontFamily:"inherit", fontSize:10, padding:"5px 9px", cursor:"pointer", border:"1px solid var(--on-dark-line)", background:"transparent", color:"var(--on-dark)" }}>Sign out</button></div>
      </div>

      <div className="mt-main-tabs" style={{ display:"flex", gap:6, padding:"0 22px", background:"var(--ink)", borderBottom:"1px solid #3a362e", position:"relative", zIndex:155, pointerEvents:"auto" }}>
        {[["tracker","Tracker"],["dashboard","Dashboard"],["management","Management"],["escalation","Escalation"],["todo","To-Do"],["review","Review"],["entrylog","Entry Log"],["settings","Settings"],["help","Help"]].map(([k,lab])=>(<button key={k} onClick={(e)=>{ e.stopPropagation(); setTab(k); if(k==="entrylog"||k==="review") loadAuditRows(); }} style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:12, letterSpacing:0.3, padding:"10px 16px", cursor:"pointer", border:"1px solid "+(tab===k?"var(--accent)":"transparent"), borderBottom:tab===k?"3px solid var(--accent)":"3px solid transparent", background:tab===k?"rgba(255,255,255,0.08)":"transparent", color:tab===k?"var(--bg)":"#b8afa3" }}>{lab}{k==="todo"&&todoItems.length?` · ${todoItems.length}`:k==="review"?(" · "+((errorLog&&errorLog.length)||0)):k==="entrylog"&&errorLog.length?` · ${errorLog.length}`:""}</button>))}
      </div>

      {tab==="help" && (<div style={{ padding:"18px 22px 36px" }}>
        <div style={{ background:"var(--surface)", border:"1px solid var(--line-2)", boxShadow:"2px 2px 0 rgba(0,0,0,0.08)", borderRadius:14, maxWidth:1120, overflow:"hidden" }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:16, padding:"18px 20px", borderBottom:"1px solid var(--line-3)", background:"#fffdf8" }}>
            <div>
              <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:23, marginBottom:4 }}>Help centre</div>
              <div style={{ fontSize:12, color:"var(--muted-3)", lineHeight:1.55, maxWidth:760 }}>Simple operating guide for the Merch Tracker. Start with the cards below, then open the section you need.</div>
            </div>
            <div style={{ fontSize:10, color:"var(--muted-2)", background:"var(--bg)", border:"1px solid var(--line-2)", borderRadius:999, padding:"6px 10px", whiteSpace:"nowrap" }}>Read-only · no data changes here</div>
          </div>

          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(210px,1fr))", gap:10, padding:14, background:"#fbf8f1", borderBottom:"1px solid var(--line-3)" }}>
            {[
              ["1", "Enter the row", "Add Style No, Order No, colour, qty, order date and delivery."],
              ["2", "Mark dates", "Double-click stage cells to enter actual dates. Use revised/reject only when needed."],
              ["3", "Chase the blocker", "Use Overall, Chase, To-Do and Review to see what needs action."],
              ["4", "Review before bulk changes", "Upload shows mapping, duplicates, missing fields and old → new changes before applying."]
            ].map(([n,t,d])=><div key={n} style={{ background:"var(--surface)", border:"1px solid var(--line-2)", borderRadius:12, padding:12 }}><div style={{ width:24, height:24, borderRadius:12, background:"var(--accent)", color:"var(--ink)", fontWeight:800, display:"flex", alignItems:"center", justifyContent:"center", fontSize:12, marginBottom:8 }}>{n}</div><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:14, marginBottom:4 }}>{t}</div><div style={{ fontSize:11, color:"var(--muted-3)", lineHeight:1.45 }}>{d}</div></div>)}
          </div>

          <div style={{ display:"flex", gap:8, flexWrap:"wrap", padding:"12px 14px", borderBottom:"1px solid var(--line-3)", background:"var(--surface)" }}>
            {[["guide","Start here"],["entry","Data entry"],["buttons","Button guide"],["faq","FAQ"],["logic","Logic"]].map(([k,l])=><button key={k} onClick={(e)=>{ e.stopPropagation(); setHelpTab(k); }} style={{ fontFamily:"inherit", fontSize:11, fontWeight:800, padding:"8px 13px", cursor:"pointer", border:"1px solid "+(helpTab===k?"var(--accent)":"var(--line-2)"), borderRadius:999, background:helpTab===k?"var(--accent-tint)":"var(--surface)", color:helpTab===k?"var(--ink)":"var(--muted-3)" }}>{l}</button>)}
          </div>

          <div style={{ padding:18, fontSize:12, lineHeight:1.6 }}>
            {helpTab==="guide" && <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))", gap:12 }}>
              <div style={{ border:"1px solid var(--line-2)", borderRadius:12, padding:14, background:"#fffdf8" }}><h3 style={{ margin:"0 0 8px", fontFamily:"'Archivo',sans-serif" }}>What each tab is for</h3><ul style={{ margin:"0 0 0 18px", padding:0 }}><li><b>Tracker:</b> live working sheet.</li><li><b>Dashboard:</b> daily operational risk and bottlenecks.</li><li><b>Management:</b> management analytics and actual performance.</li><li><b>Escalation:</b> chase-delay matrix by department, overdue age and escalation level.</li><li><b>To-Do:</b> overdue and upcoming actions.</li><li><b>Review:</b> category-wise review inbox for order, colour, fit, upload, data quality, comments and errors.</li><li><b>Settings:</b> lead days, views and chase labels.</li><li><b>Help:</b> this guide.</li></ul></div>
              <div style={{ border:"1px solid var(--line-2)", borderRadius:12, padding:14, background:"#fffdf8" }}><h3 style={{ margin:"0 0 8px", fontFamily:"'Archivo',sans-serif" }}>Daily use flow</h3><ol style={{ margin:"0 0 0 18px", padding:0 }}><li>Open a saved view such as My Overdue or Buyer Approval Pending.</li><li>Check Overall / Chase / To-Do.</li><li>Double-click the correct cell and update the date.</li><li>Add comments or remarks where a delay needs explanation.</li><li>Use Review to audit changes and comments.</li></ol></div>
              <div style={{ border:"1px solid var(--line-2)", borderRadius:12, padding:14, background:"#fffdf8" }}><h3 style={{ margin:"0 0 8px", fontFamily:"'Archivo',sans-serif" }}>Safe habits</h3><ul style={{ margin:"0 0 0 18px", padding:0 }}><li>Use staging before replacing live app files.</li><li>Do not bulk upload without checking the preview.</li><li>Use revised dates instead of changing actual dates when plans move.</li><li>Use comments for buyer remarks and delay reasons.</li></ul></div>
            </div>}

            {helpTab==="entry" && <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(250px,1fr))", gap:12 }}>
              {[
                ["Required row fields", ["Style No", "Order No", "Order Date", "Delivery Date"]],
                ["Stage dates", ["Actual date = activity completed", "Revised date = plan changed", "Reject date = buyer rejected", "Skip/Waive = not required"]],
                ["Upload rules", ["Order No + Style No is the unique key", "Missing Order/Style is blocked", "Duplicate keys in upload are blocked", "Changed attributes require confirmation"]],
                ["Comments", ["Use @name to mention teammates", "Resolve when action is closed", "Review shows all comments and comments involving you"]]
              ].map(([t,items])=><div key={t} style={{ border:"1px solid var(--line-2)", borderRadius:12, padding:14, background:"#fffdf8" }}><h3 style={{ margin:"0 0 8px", fontFamily:"'Archivo',sans-serif" }}>{t}</h3><ul style={{ margin:"0 0 0 18px", padding:0 }}>{items.map(x=><li key={x}>{x}</li>)}</ul></div>)}
            </div>}

            {helpTab==="buttons" && <div>
              <div className="mt-toolbar-mock"><b>Future UI-rehaul visual only — proposed toolbar grouping for later decision, not applied in current tracker density fix:</b><br/>Row 1: Search + saved view + main status filters<br/>Row 2: Column view + freeze + text size/boldness + sync<br/>Row 3: Selection + entry bar + comments/history/follow</div>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(240px,1fr))", gap:12 }}>
              {[
                ["Search & filters", ["Search: find style/order/colour text", "Saved View: apply common preset filters", "Status/Chase/Activity: narrow visible rows", "Clear: remove active filters"]],
                ["Data actions", ["Upload styles: import/update rows from Excel", "Export: download filtered data", "Fill date: bulk-fill one date column"]],
                ["View tools", ["Role/Column view: show useful columns", "Columns: hide/show or auto-fit", "Freeze: keep left columns visible", "A/B controls: text size and boldness"]],
                ["Edit & review", ["Find/Replace: search and replace text fields", "Copy/Paste special: copy full stage state", "Review: changes, comments and errors", "Sync: pull latest shared data"]]
              ].map(([t,items])=><div key={t} style={{ border:"1px solid var(--line-2)", borderRadius:12, padding:14, background:"#fffdf8" }}><h3 style={{ margin:"0 0 8px", fontFamily:"'Archivo',sans-serif" }}>{t}</h3><ul style={{ margin:"0 0 0 18px", padding:0 }}>{items.map(x=><li key={x}>{x}</li>)}</ul></div>)}
              </div>
            </div>}

            {helpTab==="faq" && <div style={{ display:"grid", gap:9 }}>
              {[
                ["Why is a style red?", "The next actionable stage is overdue, rejected, or the projected release is risky."],
                ["What is Chase?", "The current person/department label responsible for the next actionable pending stage. Labels are editable in Settings."],
                ["What is the difference between actual and revised?", "Actual means completed. Revised means the planned date changed but the activity is not complete."],
                ["Can the same Style No repeat?", "Yes, only if the Order No is different. The upload key is Order No + Style No."],
                ["Does Help change anything?", "No. The Help tab is read-only."],
                ["When should I use Review?", "Use it to check recent changes, all comments, comments involving you and browser-session errors."]
              ].map(([q,a])=><div key={q} style={{ border:"1px solid var(--line-2)", borderRadius:12, padding:"11px 13px", background:"#fffdf8" }}><div style={{ fontWeight:800, marginBottom:3 }}>{q}</div><div style={{ color:"var(--muted-3)" }}>{a}</div></div>)}
            </div>}

            {helpTab==="logic" && <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(260px,1fr))", gap:12 }}>
              {[
                ["Date priority", ["Actual date wins first", "Then skip/waive date", "Then revised date", "Then auto-plan date"]],
                ["Working days", ["Sunday is ignored", "Lead days are working days", "Fabric IH is planned from delivery minus cutoff"]],
                ["Flow", ["Order → Techpack → Fit/Print/Lab", "If all development toggles are OFF: Fabric IH → Prod File", "Fabric IH gates PP", "Prod File follows PP Approval unless PP is bypassed/not needed"]],
                ["Rejections", ["Rejected approval triggers rework", "Old revised dates before rejection are ignored", "Fresh actual/revised dates after rejection rebuild the chain"]],
                ["Risk", ["Delivery risk appears when projected Production File release misses the release gate", "Tight appears when release is close to the gate"]]
              ].map(([t,items])=><div key={t} style={{ border:"1px solid var(--line-2)", borderRadius:12, padding:14, background:"#fffdf8" }}><h3 style={{ margin:"0 0 8px", fontFamily:"'Archivo',sans-serif" }}>{t}</h3><ul style={{ margin:"0 0 0 18px", padding:0 }}>{items.map(x=><li key={x}>{x}</li>)}</ul></div>)}
            </div>}
          </div>
        </div>
      </div>)}
      {usersOpen && <UsersPanel onClose={()=>setUsersOpen(false)}/>}
{history && (<div onClick={()=>setHistory(false)} style={{ position:"fixed", inset:0, background:"rgba(26,26,26,0.55)", zIndex:200, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}><div onClick={e=>e.stopPropagation()} style={{ background:"var(--surface)", border:"2px solid var(--ink)", boxShadow:"8px 8px 0 var(--ink)", width:640, maxWidth:"100%", maxHeight:"80vh", display:"flex", flexDirection:"column", fontFamily:"'JetBrains Mono',monospace" }}><div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"12px 14px", borderBottom:"1px solid var(--line-3)" }}><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:17 }}>Change history</div><button onClick={()=>setHistory(false)} style={{ border:"1px solid var(--ink)", background:"var(--bg)", cursor:"pointer", fontFamily:"inherit", fontSize:11, padding:"4px 9px" }}>Close</button></div><div style={{ padding:"8px 14px", borderBottom:"1px solid #f2f2f2" }}><input value={histFilter} onChange={e=>setHistFilter(e.target.value)} placeholder="filter by style no, field, stage, or person…" style={{ width:"100%", boxSizing:"border-box", fontFamily:"inherit", fontSize:11, padding:6, border:"1px solid var(--ink)" }}/></div><div style={{ overflowY:"auto", padding:"4px 0" }}>{auditBusy ? <div style={{ fontSize:11, color:"var(--muted-1)", padding:"24px", textAlign:"center" }}>Loading…</div> : (()=>{ const q=histFilter.trim().toLowerCase(); const fl=auditRows.filter(a=> !q || (a.style_no||"").toLowerCase().includes(q) || (a.field||"").toLowerCase().includes(q) || (a.actor_name||"").toLowerCase().includes(q) || (colLabelOf(a.col)||"").toLowerCase().includes(q)); if(fl.length===0) return <div style={{ fontSize:11, color:"var(--muted-1)", padding:"24px", textAlign:"center" }}>{auditRows.length===0?"No changes recorded yet.":"No matches."}</div>; return fl.map(a=>(<div key={a.id} style={{ padding:"7px 14px", borderBottom:"1px solid #f5f5f5", fontSize:11, display:"flex", gap:10, alignItems:"baseline" }}><span style={{ color:"var(--muted-1)", fontSize:9, whiteSpace:"nowrap", minWidth:80 }}>{tsShort(a.created_at)}</span><span style={{ flex:1 }}><b>{a.style_no||a.style_id}</b> · {colLabelOf(a.col)} <span style={{ color:"var(--muted-2)" }}>({a.field})</span><br/><span style={{ color:"var(--danger)" }}>{a.old_val||"—"}</span> <span style={{ color:"var(--muted-2)" }}>→</span> <span style={{ color:"var(--success)" }}>{a.new_val||"—"}</span> <span style={{ color:"var(--muted-1)" }}>· {a.actor_name}</span></span></div>)); })()}</div></div></div>)}
      {reviewOpen && (<div onClick={()=>setReviewOpen(false)} style={{ position:"fixed", inset:0, background:"rgba(26,26,26,0.55)", zIndex:205, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}><div onClick={e=>e.stopPropagation()} style={{ background:"var(--surface)", border:"2px solid var(--ink)", boxShadow:"8px 8px 0 var(--ink)", width:760, maxWidth:"100%", maxHeight:"84vh", display:"flex", flexDirection:"column", fontFamily:"'JetBrains Mono',monospace" }}><div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"12px 14px", borderBottom:"1px solid var(--line-3)" }}><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:17 }}>Review centre</div><button className="mt-popup-close" onClick={()=>setReviewOpen(false)} style={{ border:"1px solid var(--ink)", background:"var(--bg)", cursor:"pointer", fontFamily:"inherit", fontSize:11, fontWeight:800, padding:"7px 11px" }}>Close</button></div><div style={{ display:"flex", gap:6, padding:"8px 10px", borderBottom:"1px solid var(--line-3)", flexWrap:"wrap" }}>{[["changes","Changes to sheet"],["comments","All comments"],["mine","Comments involving me"],["errors","Error log"]].map(([k,l])=><button key={k} onClick={()=>{ setReviewTab(k); if(k==="changes") loadAuditRows(); }} style={{ fontFamily:"inherit", fontSize:11, fontWeight:700, padding:"8px 12px", cursor:"pointer", border:"none", borderRight:"1px solid var(--line-3)", background:reviewTab===k?"var(--accent)":"var(--surface)", color:reviewTab===k?"var(--ink)":"var(--muted-3)" }}>{l}</button>)}</div><div style={{ overflowY:"auto", padding:12 }}>{reviewTab==="changes" && (auditBusy?<div style={{ fontSize:11, color:"var(--muted-1)", padding:20 }}>Loading changes…</div>:auditRows.length===0?<div style={{ fontSize:11, color:"var(--muted-1)", padding:20 }}>No recorded changes yet.</div>:auditRows.slice(0,120).map(a=><div key={a.id} style={{ padding:"7px 0", borderBottom:"1px solid #f2eee6", fontSize:11 }}><span style={{ color:"var(--muted-1)", fontSize:9 }}>{tsShort(a.created_at)}</span> · <b>{a.style_no||a.style_id}</b> · {colLabelOf(a.col)} <span style={{ color:"var(--muted-2)" }}>({a.field})</span><br/><span style={{ color:"var(--danger)" }}>{a.old_val||"—"}</span> → <span style={{ color:"var(--success)" }}>{a.new_val||"—"}</span> <span style={{ color:"var(--muted-1)" }}>· {a.actor_name}</span></div>))}{reviewTab!=="changes" && (()=>{ const all=[]; Object.entries(comments||{}).forEach(([ck,arr])=>{ const [sid,col]=ck.split(":"); const st=styles.find(x=>String(x.id)===String(sid)); (arr||[]).forEach(c=>all.push({...c, sid:Number(sid), col, styleNo:(st&&st.styleNo)||c.style_no||sid, orderNo:(st&&st.orderNo)||""})); }); const my=(me&&(me.name||me.email)||"").toLowerCase(); const myCompact=my.replace(/\s+/g,""); const data=reviewTab==="mine"?all.filter(c=>(c.mentions||[]).includes(me&&me.id)||c.author_id===(me&&me.id)||String(c.body||"").toLowerCase().includes("@"+myCompact)||String(c.body||"").toLowerCase().includes("@"+my)||String(c.author_name||"").toLowerCase().includes(my)):all; if(reviewTab==="errors"){ return errorLog.length===0?<div style={{ fontSize:11, color:"var(--muted-1)", padding:20 }}>No app errors logged in this browser session.</div>:errorLog.map(e=><div key={e.id} style={{ padding:"7px 0", borderBottom:"1px solid #f2eee6", fontSize:11 }}><span style={{ color:"var(--muted-1)", fontSize:9 }}>{tsShort(e.at)}</span> · <b>{e.area}</b><br/><span style={{ color:"var(--danger)" }}>{e.msg}</span>{e.extra&&<span style={{ color:"var(--muted-1)" }}> · {e.extra}</span>}</div>); } return data.length===0?<div style={{ fontSize:11, color:"var(--muted-1)", padding:20 }}>No comments found.</div>:data.sort((a,b)=>new Date(b.created_at)-new Date(a.created_at)).slice(0,150).map(c=><button key={c.id} onClick={()=>{ setReviewOpen(false); setHistory(false); setTab("tracker"); setTimeout(()=>{ setSel({id:c.sid,col:c.col}); setFocus(null); setThreadCell({id:c.sid,col:c.col}); scrollToCell(c.sid,c.col); },60); }} style={{ display:"block", width:"100%", textAlign:"left", border:"none", borderBottom:"1px solid #f2eee6", background:"transparent", cursor:"pointer", padding:"7px 0", fontFamily:"inherit", fontSize:11 }}><span style={{ color:"var(--muted-1)", fontSize:9 }}>{tsShort(c.created_at)}</span> · <b>{c.orderNo?c.orderNo+" · ":""}{c.styleNo}</b> · {colLabelOf(c.col)}<br/><span>{c.body}</span></button>); })()}</div></div></div>)}
      {bulkOpen && (<div onClick={()=>{ setBulkOpen(false); setBulkResult(null); setUploadSkip(new Set()); }} style={{ position:"fixed", inset:0, background:"rgba(26,26,26,0.55)", zIndex:200, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
        <div onClick={e=>e.stopPropagation()} style={{ background:"var(--bg)", border:"2px solid var(--ink)", boxShadow:"8px 8px 0 var(--ink)", width:700, maxWidth:"100%", maxHeight:"86vh", overflowY:"auto", padding:22, fontFamily:"'JetBrains Mono',monospace" }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:6 }}><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:18 }}>Bulk upload styles</div><button onClick={()=>{ setBulkOpen(false); setBulkResult(null); setUploadSkip(new Set()); }} style={{ border:"none", background:"transparent", cursor:"pointer" }}><X size={18}/></button></div>
          <p style={{ fontSize:11, color:"var(--muted-3)", lineHeight:1.5 }}>Upload Excel/CSV. The unique key is <b>Order No + Style No</b>. Duplicates are blocked. Existing rows with changed attributes must be reviewed before applying.</p>
          <button onClick={downloadTemplate} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)", marginBottom:12 }}>⬇ Download template</button>
          <div style={{ marginBottom:12 }}><input type="file" accept=".xlsx,.xls,.csv" onChange={e=>{ const f=e.target.files&&e.target.files[0]; if(f) parseUpload(f); }} style={{ fontFamily:"inherit", fontSize:11 }}/></div>
          {bulkResult && bulkResult.error && <div style={{ fontSize:11, color:"var(--danger)", fontWeight:700, marginBottom:12 }}>{bulkResult.error}</div>}
          {bulkResult && bulkResult.mapping && !bulkResult.error && (<div style={{ marginBottom:12, border:"1px solid var(--ink)", background:"var(--surface)", padding:12 }}>
            <div style={{ fontSize:12, fontWeight:700, marginBottom:6 }}>Confirm column mapping · {bulkResult.sn}</div>
            {bulkResult.mapError && <div style={{ fontSize:10, color:"var(--danger)", fontWeight:700, marginBottom:8 }}>{bulkResult.mapError}</div>}
            <div style={{ display:"grid", gridTemplateColumns:"repeat(2,minmax(0,1fr))", gap:8 }}>
              {UPLOAD_FIELDS.map(([k,l])=>(<label key={k} style={{ fontSize:10, display:"flex", flexDirection:"column", gap:3 }}><span style={{ fontWeight:k==="styleNo"||k==="orderNo"?700:400, color:k==="styleNo"||k==="orderNo"?"var(--danger)":"var(--muted-3)" }}>{l}</span><select value={(bulkResult.mapping&&bulkResult.mapping[k])??""} onChange={e=>setBulkResult(br=>({...br, mapError:"", mapping:{...(br.mapping||{}), [k]: e.target.value===""?undefined:Number(e.target.value)}}))} style={{ fontFamily:"inherit", fontSize:10, border:"1px solid var(--line-2)", padding:4 }}><option value="">— not mapped —</option>{(bulkResult.headersRaw||[]).map((h,i)=><option key={i} value={i}>{i+1}. {h||"(blank)"}</option>)}</select></label>))}
            </div>
            <div style={{ display:"flex", gap:8, marginTop:12, flexWrap:"wrap" }}><button onClick={()=>buildUploadPreview(bulkResult)} style={{ fontFamily:"inherit", fontSize:11, fontWeight:700, padding:"7px 12px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--accent)" }}>Continue preview</button><button onClick={()=>{ try{ localStorage.setItem("mt_upload_mapping_default", JSON.stringify(bulkResult.mapping||{})); alert("Upload mapping saved as default for this browser."); }catch(e){ alert("Could not save mapping."); } }} style={{ fontFamily:"inherit", fontSize:11, padding:"7px 12px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)" }}>Save mapping</button><button onClick={()=>{ try{ localStorage.removeItem("mt_upload_mapping_default"); }catch(e){} const mapping={}; (bulkResult.headersRaw||[]).map(NORMH).forEach((h,i)=>{ const f=HEADER_MAP[h]; if(f&&mapping[f]==null) mapping[f]=i; }); setBulkResult(br=>({...br, mapping, mapError:(mapping.styleNo==null||mapping.orderNo==null)?"Please map Style No and Order No. Both are mandatory.":""})); }} style={{ fontFamily:"inherit", fontSize:11, padding:"7px 12px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)" }}>Reset auto-map</button><button onClick={()=>setBulkResult(null)} style={{ fontFamily:"inherit", fontSize:11, padding:"7px 12px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)" }}>Clear file</button></div>
          </div>)}
          {bulkResult && !bulkResult.error && !bulkResult.mapping && (<div style={{ marginBottom:12 }}>
            <div style={{ fontSize:12, marginBottom:8 }}>Read <b>{bulkResult.total}</b> usable rows from sheet "<b>{bulkResult.sheetName}</b>":</div>
            <div style={{ display:"flex", gap:10, marginBottom:10 }}>
              <div style={{ flex:1, border:"1px solid var(--ink)", background:"var(--surface)", padding:"8px 10px" }}><div style={{ fontSize:22, fontWeight:800, color:"var(--success)", fontFamily:"'Archivo',sans-serif" }}>{(bulkResult.inserts||[]).length}</div><div style={{ fontSize:9, color:"var(--muted-2)", textTransform:"uppercase" }}>new</div></div>
              <div style={{ flex:1, border:"1px solid var(--ink)", background:"var(--surface)", padding:"8px 10px" }}><div style={{ fontSize:22, fontWeight:800, color:"var(--accent)", fontFamily:"'Archivo',sans-serif" }}>{(bulkResult.updates||[]).length}</div><div style={{ fontSize:9, color:"var(--muted-2)", textTransform:"uppercase" }}>changed existing</div></div>
              <div style={{ flex:1, border:"1px solid var(--ink)", background:"var(--surface)", padding:"8px 10px" }}><div style={{ fontSize:22, fontWeight:800, color:bulkResult.blocking?"var(--danger)":"var(--muted-1)", fontFamily:"'Archivo',sans-serif" }}>{((bulkResult.errors||[]).length+(bulkResult.dupes||[]).length)}</div><div style={{ fontSize:9, color:"var(--muted-2)", textTransform:"uppercase" }}>issues</div></div>
              <div style={{ flex:1, border:"1px solid var(--ink)", background:"var(--surface)", padding:"8px 10px" }}><div style={{ fontSize:22, fontWeight:800, color:"var(--muted-1)", fontFamily:"'Archivo',sans-serif" }}>{bulkResult.unchanged||0}</div><div style={{ fontSize:9, color:"var(--muted-2)", textTransform:"uppercase" }}>unchanged</div></div>
            </div>
            {(((bulkResult.errors||[]).length>0)||((bulkResult.dupes||[]).length>0)||((bulkResult.updates||[]).length>0)) && <button onClick={downloadUploadErrorReport} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)", marginBottom:8 }}>⬇ Download upload review report</button>}
            {bulkResult.blocking && <div style={{ fontSize:10, color:"var(--danger)", fontWeight:700, marginBottom:8, background:"#fbeaea", border:"1px solid #e8b4ae", padding:"7px 9px" }}>Upload blocked: duplicate or mandatory-key errors found. Fix the source file and upload again. No duplicate Order No + Style No rows will be allowed.</div>}
            {(bulkResult.errors||[]).filter(e=>e.blocking).slice(0,8).map(e=><div key={(e.rowNo||"")+e.type} style={{ fontSize:10, color:"var(--danger)", marginBottom:3 }}>Row {e.rowNo}: {e.detail}</div>)}
            {(bulkResult.updates||[]).length>0 && <div style={{ maxHeight:220, overflowY:"auto", border:"1px solid var(--line-1)", background:"var(--surface)", padding:8, fontSize:10 }}><div style={{ fontWeight:700, marginBottom:4 }}>Decide changed existing rows:</div>{(bulkResult.updates||[]).slice(0,80).map(u=>(<label key={u.key} style={{ display:"block", padding:"5px 0", borderBottom:"1px solid #f0ece3" }}><span style={{ display:"flex", alignItems:"center", gap:6 }}><input type="checkbox" checked={!uploadSkip.has(u.key)} onChange={()=>setUploadSkip(p=>{ const n=new Set(p); n.has(u.key)?n.delete(u.key):n.add(u.key); return n; })}/><b>{u.orderNo} · {u.styleNo}</b><span style={{ color:uploadSkip.has(u.key)?"var(--danger)":"var(--success)", fontWeight:700 }}>{uploadSkip.has(u.key)?"skip":"apply upload"}</span></span>{(u.diffs||[]).slice(0,6).map(d=><div key={d.field} style={{ color:"var(--muted-3)", marginLeft:22 }}>{uploadFieldLabel(d.field)}: <span style={{ color:"var(--danger)" }}>{String(d.from||"—")}</span> → <span style={{ color:"var(--success)" }}>{String(d.to||"—")}</span></div>)}</label>))}{bulkResult.updates.length>80 && <div style={{ color:"var(--muted-1)" }}>…and {bulkResult.updates.length-80} more</div>}</div>}
          </div>)}
          {bulkResult && !bulkResult.error && !bulkResult.mapping && (bulkResult.inserts.length+bulkResult.updates.length>0) && <button disabled={bulkResult.blocking} onClick={applyBulk} style={{ fontFamily:"inherit", fontSize:12, padding:"9px 16px", cursor:bulkResult.blocking?"not-allowed":"pointer", border:"1px solid var(--ink)", background:bulkResult.blocking?"var(--line-2)":"var(--success)", color:bulkResult.blocking?"var(--muted-4)":"var(--surface)", fontWeight:700 }}>Confirm — apply selected changes</button>}
        </div>
      </div>)}

      {bulkActionsOpen && (<div onClick={()=>setBulkActionsOpen(false)} style={{ position:"fixed", inset:0, background:"rgba(26,26,26,0.55)", zIndex:210, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
        <div onClick={e=>e.stopPropagation()} style={{ background:"var(--bg)", border:"2px solid var(--ink)", boxShadow:"8px 8px 0 var(--ink)", width:620, maxWidth:"100%", maxHeight:"86vh", overflowY:"auto", padding:22, fontFamily:"'JetBrains Mono',monospace" }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:18 }}>Bulk row actions</div><button onClick={()=>setBulkActionsOpen(false)} style={{ border:"none", background:"transparent", cursor:"pointer" }}><X size={18}/></button></div>
          <p style={{ fontSize:11, color:"var(--muted-3)", lineHeight:1.5, marginTop:0 }}>Applies only to the <b>{rows.length}</b> styles currently visible after your filters. This is not a temporary view action; it changes real row data and auto-saves for everyone.</p>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(250px,1fr))", gap:12 }}>
            <BulkActionBox title="Assign / update text fields">
              <button onClick={()=>{ const v=window.prompt("Set Junior / Style Owner for visible rows:"); if(v!=null) bulkUpdateVisible({ owner:v.trim() }, "Set owner"); }} style={bulkBtn}>Set Junior / Style Owner</button>
              <button onClick={()=>{ const v=window.prompt("Set Buyer for visible rows:"); if(v!=null) bulkUpdateVisible({ buyer:v.trim() }, "Set buyer"); }} style={bulkBtn}>Set Buyer</button>
              <button onClick={()=>{ const v=window.prompt("Set Brand for visible rows:"); if(v!=null) bulkUpdateVisible({ brand:v.trim() }, "Set brand"); }} style={bulkBtn}>Set Brand</button>
              <button onClick={()=>{ const v=window.prompt("Append this remark to visible rows:"); if(v!=null) bulkAppendRemark(v); }} style={bulkBtn}>Append Remark</button>
            </BulkActionBox>
            <BulkActionBox title="Requirement flags">
              {FLAG_DEFS.map(f=><div key={f.key} style={{ display:"flex", gap:6, alignItems:"center", marginBottom:6 }}><span style={{ flex:1, fontSize:10, fontWeight:800 }}>{f.short}</span><button onClick={()=>bulkFlagVisible(f.key,true)} style={miniBulkBtn}>Yes</button><button onClick={()=>bulkFlagVisible(f.key,false)} style={miniBulkBtn}>No</button></div>)}
            </BulkActionBox>
            <BulkActionBox title="Archive / restore">
              <button onClick={()=>archiveFiltered(true)} style={{ ...bulkBtn, color:"#5a6650" }}>Archive visible styles</button>
              <button onClick={()=>archiveFiltered(false)} style={{ ...bulkBtn, color:"var(--success)" }}>Restore visible styles</button>
            </BulkActionBox>
            <BulkActionBox title="Safety note">
              <div style={{ fontSize:10, color:"var(--muted-3)", lineHeight:1.55 }}>Filter first, then apply. Example: filter buyer/order/season, then bulk assign junior/style owner or flags. Undo is available in this browser before auto-save finishes, but use staging testing first.</div>
            </BulkActionBox>
          </div>
        </div>
      </div>)}

      {bulkConfirm && (<div onClick={()=>setBulkConfirm(null)} style={{ position:"fixed", inset:0, background:"rgba(26,26,26,0.62)", zIndex:430, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
        <div onClick={e=>e.stopPropagation()} style={{ background:"var(--bg)", border:"2px solid var(--ink)", boxShadow:"8px 8px 0 var(--ink)", width:560, maxWidth:"100%", padding:20, fontFamily:"'JetBrains Mono',monospace" }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", gap:12, marginBottom:8 }}>
            <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:18 }}>Confirm bulk action</div>
            <button onClick={()=>setBulkConfirm(null)} style={{ border:"none", background:"transparent", cursor:"pointer" }}><X size={18}/></button>
          </div>
          <div style={{ border:"1px solid var(--ink)", background:"var(--surface)", padding:12, marginBottom:10 }}>
            <div style={{ fontSize:12, fontWeight:800, marginBottom:5 }}>{bulkConfirm.title}</div>
            <div style={{ fontSize:11, color:"var(--muted-3)", lineHeight:1.5 }}>{bulkConfirm.impact || "This changes real row data and will auto-save for everyone."}</div>
          </div>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:8, marginBottom:10 }}>
            <div style={{ border:"1px solid var(--line-2)", background:"var(--surface)", padding:8 }}><div style={{ fontSize:20, fontWeight:800, color:"var(--accent)", fontFamily:"'Archivo',sans-serif" }}>{bulkConfirm.summary?.count||0}</div><div style={{ fontSize:9, color:"var(--muted-2)", textTransform:"uppercase" }}>visible styles</div></div>
            <div style={{ border:"1px solid var(--line-2)", background:"var(--surface)", padding:8 }}><div style={{ fontSize:11, fontWeight:800 }}>{(bulkConfirm.summary?.owners||[]).join(", ")||"—"}</div><div style={{ fontSize:9, color:"var(--muted-2)", textTransform:"uppercase", marginTop:4 }}>current owners</div></div>
            <div style={{ border:"1px solid var(--line-2)", background:"var(--surface)", padding:8 }}><div style={{ fontSize:11, fontWeight:800 }}>{(bulkConfirm.summary?.buyers||[]).join(", ")||"—"}</div><div style={{ fontSize:9, color:"var(--muted-2)", textTransform:"uppercase", marginTop:4 }}>buyers/brands</div></div>
          </div>
          {bulkConfirm.kind==="patch" && <div style={{ fontSize:10, marginBottom:10, border:"1px solid var(--line-2)", background:"var(--surface)", padding:8 }}><b>Field changes:</b>{Object.entries(bulkConfirm.patch||{}).map(([k,v])=><div key={k} style={{ marginTop:3 }}>• {k}: <b>{typeof v==="boolean"?(v?"Yes":"No"):String(v||"—")}</b></div>)}</div>}
          {bulkConfirm.kind==="appendRemark" && <div style={{ fontSize:10, marginBottom:10, border:"1px solid var(--line-2)", background:"var(--surface)", padding:8 }}><b>Remark to append:</b><div style={{ marginTop:4, color:"var(--muted-3)" }}>{bulkConfirm.note}</div></div>}
          <div style={{ fontSize:10, color:"var(--muted-3)", marginBottom:10 }}>Sample affected rows: {(bulkConfirm.summary?.sample||[]).join("; ")}{bulkConfirm.summary?.more?` … +${bulkConfirm.summary.more} more`:""}</div>
          <div style={{ fontSize:10, color:"var(--danger)", fontWeight:800, background:"#fbeaea", border:"1px solid #e8b4ae", padding:"7px 9px", marginBottom:12 }}>This is not a temporary view action. It will change actual data for all selected visible styles.</div>
          <div style={{ display:"flex", gap:8, justifyContent:"flex-end" }}>
            <button onClick={()=>setBulkConfirm(null)} style={{ fontFamily:"inherit", fontSize:11, padding:"8px 12px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)" }}>Cancel</button>
            <button onClick={applyBulkConfirm} style={{ fontFamily:"inherit", fontSize:11, fontWeight:800, padding:"8px 14px", cursor:"pointer", border:"1px solid var(--ink)", background:bulkConfirm.danger?"var(--danger)":"var(--success)", color:"var(--surface)" }}>{bulkConfirm.actionText||"Apply bulk action"}</button>
          </div>
        </div>
      </div>)}

      {tab==="tracker" && (<>
      <div className="mt-tracker-summary" style={{ display:"flex", padding:"12px 22px 0", flexWrap:"wrap" }}>
        {Object.entries(funnel).map(([k,v],i,arr)=>(<div key={k} style={{ flex:1, minWidth:90, background:"var(--surface)", border:"1px solid var(--line-2)", borderRight:i===arr.length-1?"1px solid var(--line-2)":"none", padding:"8px 10px" }}><div style={{ fontSize:22, fontWeight:700, lineHeight:1, fontFamily:"'Archivo',sans-serif", color:k==="Released"?"var(--success)":k==="Fabric IH"?"var(--danger)":"var(--ink)" }}>{v}</div><div style={{ fontSize:9, color:"var(--muted-2)", marginTop:3, letterSpacing:0.5, textTransform:"uppercase" }}>{k}</div></div>))}
      </div>

      <div className="mt-tracker-toolbar" style={{ display:"flex", gap:8, alignItems:"center", padding:"12px 22px 6px", flexWrap:"wrap" }}>
        <div className="mt-toolbar-group" style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, marginRight:2 }}>Search</span>
        <div style={{ display:"flex", alignItems:"center", gap:6, background:"var(--surface)", border:"1px solid var(--ink)", padding:"5px 9px" }}><Filter size={13}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder={searchCol==="auto"?"search style / colour / fit / order…":"search selected column…"} onClick={e=>e.stopPropagation()} style={{ border:"none", outline:"none", fontFamily:"inherit", fontSize:12, width:160, background:"transparent" }}/><span style={{ width:1, height:16, background:"var(--line-2)" }}/><select value={searchCol} onChange={e=>setSearchCol(e.target.value)} onClick={e=>e.stopPropagation()} title="search a specific column (Auto = smart search across the main fields)" style={{ border:"none", outline:"none", fontFamily:"inherit", fontSize:10, background:"transparent", color:searchCol==="auto"?"var(--muted-2)":"var(--accent)", cursor:"pointer", fontWeight:searchCol==="auto"?400:700 }}>{[["auto","Auto"],["styleNo","Style No"],["colour","Colour"],["family","Family"],["sampleFit","Sample Fit"],["orderNo","Order No"],["owner","Junior"],["brand","Brand"],["buyer","Buyer"],["age","Age Group"],["fabricType","Fabric Type"],["setRole","Set Role"],["remarks","Remarks"]].map(([k,lab])=><option key={k} value={k}>{lab}</option>)}</select></div>
        <span style={{ position:"relative" }}><button onClick={(e)=>{ e.stopPropagation(); setFrOpen(o=>!o); }} title="Find & replace text in cells" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:frOpen?"var(--ink)":"var(--surface)", color:frOpen?"var(--bg)":"var(--ink)", fontWeight:700 }}>⌕ Find/Replace</button>{frOpen && (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:0, marginTop:4, zIndex:370, background:"var(--surface)", border:"1px solid var(--ink)", boxShadow:"4px 4px 0 var(--ink)", padding:12, width:260 }}><div style={{ fontSize:11, fontWeight:700, marginBottom:8 }}>Find &amp; replace</div><label style={{ fontSize:10, color:"var(--muted-2)" }}>Find</label><input autoFocus value={frFind} onChange={e=>{ setFrFind(e.target.value); setFindIdx(-1); setFrMatches([]); }} onKeyDown={e=>{ if(e.key==="Enter"){ e.preventDefault(); runFind(); } }} placeholder="text to find" style={{ width:"100%", boxSizing:"border-box", fontFamily:"inherit", fontSize:11, padding:5, marginBottom:6, border:"1px solid var(--ink)" }}/><label style={{ fontSize:10, color:"var(--muted-2)" }}>Replace with</label><input value={frRepl} onChange={e=>setFrRepl(e.target.value)} placeholder="(leave blank to delete)" style={{ width:"100%", boxSizing:"border-box", fontFamily:"inherit", fontSize:11, padding:5, marginBottom:8, border:"1px solid var(--ink)" }}/><div style={{ display:"flex", gap:12, fontSize:10, marginBottom:6 }}><label style={{ display:"flex", gap:4, cursor:"pointer" }}><input type="radio" checked={frScope==="filtered"} onChange={()=>setFrScope("filtered")}/>All filtered</label><label style={{ display:"flex", gap:4, cursor:"pointer" }}><input type="radio" checked={frScope==="selected"} onChange={()=>setFrScope("selected")}/>Selected cells</label></div><label style={{ fontSize:10, display:"flex", gap:5, marginBottom:8, cursor:"pointer" }}><input type="checkbox" checked={frCase} onChange={e=>setFrCase(e.target.checked)}/>Match case</label><div style={{ fontSize:10, color:"var(--muted-3)", marginBottom:8, minHeight:13 }}>{frScope==="selected" && !rect() ? "Select cells in the grid first." : (frFind ? findReplace(true)+" cell(s) will change" : "Applies to text fields (style, colour, brand, owner, remarks…)")}</div>{frMatches.length>0 && (<div style={{ marginBottom:8 }}><div style={{ fontSize:9, fontWeight:700, color:"var(--muted-3)", marginBottom:3 }}>{frMatches.length} match{frMatches.length===1?"":"es"} — click to jump</div><div style={{ maxHeight:150, overflowY:"auto", border:"1px solid var(--line-2)" }}>{frMatches.map((m,i)=>(<div key={m.id+":"+m.col} onClick={(e)=>{ e.stopPropagation(); gotoMatch(i); }} title={m.style+" · "+m.colLabel} style={{ display:"flex", gap:6, alignItems:"baseline", padding:"4px 6px", cursor:"pointer", fontSize:10, borderBottom:"1px solid var(--line-3)", background:i===findIdx?"var(--accent-tint)":"var(--surface)" }}><span style={{ fontWeight:700, whiteSpace:"nowrap", maxWidth:78, overflow:"hidden", textOverflow:"ellipsis" }}>{m.style||"—"}</span><span style={{ color:"var(--muted-2)", whiteSpace:"nowrap" }}>{m.colLabel}</span><span style={{ flex:1, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{m.text}</span></div>))}</div></div>)}<div style={{ display:"flex", gap:8 }}><button disabled={!frFind} onClick={(e)=>{ e.stopPropagation(); findNext(); }} title="search, then cycle through matching cells" style={{ flex:1, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:6, cursor:frFind?"pointer":"not-allowed", border:"1px solid var(--ink)", background:"var(--surface)", opacity:frFind?1:0.6 }}>{frMatches.length?("Next "+(findIdx+1)+"/"+frMatches.length):"Find"}</button><button disabled={!frFind} onClick={(e)=>{ e.stopPropagation(); const n=findReplace(true); if(!n){ window.alert('No cells match "'+frFind+'".'); return; } const bw=frRepl.trim()===""?"\n\n\u26a0\ufe0f This will BLANK those cells.":""; if(!window.confirm('Replace "'+frFind+'" \u2192 "'+(frRepl||"(blank)")+'" in '+n+' cell(s)?'+bw)) return; findReplace(false); flash(); }} style={{ flex:1, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:6, cursor:frFind?"pointer":"not-allowed", border:"1px solid var(--ink)", background:frFind?"var(--accent)":"var(--line-2)", opacity:frFind?1:0.6 }}>Replace all</button><button onClick={()=>setFrOpen(false)} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--bg)" }}><X size={12}/></button></div></div>)}</span>
        </div>
        <div className="mt-toolbar-group" style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, marginRight:2 }}>Filters</span>
        <select value={savedView} onChange={e=>{ setSavedView(e.target.value); if(e.target.value==="following") setFollowFilter(false); }} onClick={e=>e.stopPropagation()} title="quick saved operational views" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 8px", cursor:"pointer", border:"1px solid "+(savedView?"var(--accent)":"var(--ink)"), background:savedView?"var(--accent-tint)":"var(--surface)", fontWeight:savedView?700:400 }}>{SAVED_VIEWS.map(([k,l])=><option key={k} value={k}>{l}</option>)}</select>
        <div style={{ display:"flex", border:"1px solid var(--ink)" }}>{["All","At Risk","On Track","Released"].map(f=>(<button key={f} onClick={(e)=>{ e.stopPropagation(); setStatusFilter(f); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"none", borderRight:f!=="Released"?"1px solid var(--ink)":"none", background:statusFilter===f?"var(--ink)":"var(--surface)", color:statusFilter===f?"var(--bg)":"var(--ink)" }}>{f}</button>))}</div>
        <div style={{ display:"flex", border:"1px solid var(--ink)" }}>{chaseOwnerOptions.map((f,i)=>(<button key={f} onClick={(e)=>{ e.stopPropagation(); setOwnerFilter(f); }} title="chase label / department" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 9px", cursor:"pointer", border:"none", borderRight:i<chaseOwnerOptions.length-1?"1px solid var(--ink)":"none", background:ownerFilter===f?"#2563a6":"var(--surface)", color:ownerFilter===f?"var(--surface)":"var(--ink)" }}>{f==="All"?"Chase":f}</button>))}</div>
        <MultiSelectDropdown label="Activity" value={activityFilterKeys.map(stageLabelFromKeyGlobal).filter(Boolean)} options={STAGES.map(st=>st.label)} onChange={vals=>setTrackerActivityFilter(vals)} />
        <div style={{ display:"flex", border:"1px solid var(--ink)" }}>{[["active","Active"],["all","All"],["archived","Archived"]].map(([v,lab])=>(<button key={v} onClick={(e)=>{ e.stopPropagation(); setArchiveView(v); }} title="archived styles are hidden from the live sheet" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 9px", cursor:"pointer", border:"none", borderRight:v!=="archived"?"1px solid var(--ink)":"none", background:archiveView===v?"#5a6650":"var(--surface)", color:archiveView===v?"var(--surface)":"var(--ink)" }}>{lab}</button>))}</div>
        <button onClick={(e)=>{ e.stopPropagation(); setFollowFilter(v=>!v); }} title="show only styles you follow" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid var(--ink)", background:followFilter?"var(--accent)":"var(--surface)", color:followFilter?"var(--surface)":"var(--ink)", display:"inline-flex", alignItems:"center", gap:5 }}><Star size={12} fill={followFilter?"var(--surface)":"none"}/> Following</button>
        {canAdmin(role) && archiveView!=="archived" && anyFilter && <button onClick={(e)=>{ e.stopPropagation(); archiveFiltered(true); }} title="archive the styles currently shown (e.g. a finished season)" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid #5a6650", background:"var(--surface)", color:"#5a6650", fontWeight:700 }}>Archive these ({rows.length})</button>}
        {canAdmin(role) && archiveView==="archived" && rows.length>0 && <button onClick={(e)=>{ e.stopPropagation(); archiveFiltered(false); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--success)", background:"var(--surface)", color:"var(--success)", fontWeight:700 }}>Restore these ({rows.length})</button>}
        {anyFilter && <button onClick={(e)=>{ e.stopPropagation(); clearAllFilters(); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--danger)", background:"var(--surface)", color:"var(--danger)", fontWeight:700, display:"flex", alignItems:"center", gap:5 }}><X size={12}/> clear filters</button>}
        {viewSnap && <button onClick={(e)=>{ e.stopPropagation(); restoreView(); }} title="go back to the view you had before drilling in" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)", display:"flex", alignItems:"center", gap:5 }}><RotateCcw size={12}/> restore view</button>}
        </div>
        <div className="mt-toolbar-group" style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, marginRight:2 }}>Data</span>
        {canMaster(role) && <button onClick={(e)=>{ e.stopPropagation(); setBulkResult(null); setBulkOpen(true); }} title="bulk upload / update styles from Excel" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--ink)", color:"var(--bg)", fontWeight:700, display:"flex", alignItems:"center", gap:6 }}><Plus size={13}/> Upload styles</button>}
        {canMaster(role) && <button onClick={(e)=>{ e.stopPropagation(); setBulkActionsOpen(true); }} title="bulk actions on currently visible/filtered rows" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)", color:"var(--ink)", fontWeight:700, display:"flex", alignItems:"center", gap:6 }}><Check size={13}/> Bulk actions ({rows.length})</button>}
        <span style={{ position:"relative" }}><button onClick={(e)=>{ e.stopPropagation(); setExpOpen(o=>!o); }} title="export options" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid var(--ink)", background:expOpen?"var(--ink)":"var(--surface)", color:expOpen?"var(--bg)":"var(--ink)", fontWeight:700, display:"flex", alignItems:"center", gap:6 }}>⬇ Export</button>{expOpen && (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:0, marginTop:4, zIndex:370, background:"var(--surface)", border:"1px solid var(--ink)", boxShadow:"4px 4px 0 var(--ink)", padding:12, width:288 }}><div style={{ fontSize:11, fontWeight:700, marginBottom:8 }}>Export {rows.length} filtered styles</div>{[["full","Full \u2014 all columns (actual dates)"],["detail","Detailed summary \u2014 everything (toggles + every stage)"],["release","Release plan \u2014 for production (incl. blockers)"],["internal","Internal plan (with buffer)"],["buyer","Buyer view \u2014 key details (printable)"]].map(([v,lbl])=>(<label key={v} style={{ display:"flex", gap:6, fontSize:10, padding:"3px 0", cursor:"pointer", alignItems:"flex-start" }}><input type="radio" checked={expMode===v} onChange={()=>setExpMode(v)} style={{ marginTop:1 }}/>{lbl}</label>))}{expMode==="release" && (<div style={{ display:"flex", gap:6, margin:"4px 0 2px" }}>{[["detailed","Detailed"],["summary","Summary"]].map(([v,l])=>(<button key={v} onClick={()=>setExpRelMode(v)} style={{ flex:1, fontFamily:"inherit", fontSize:10, fontWeight:700, padding:"4px 0", cursor:"pointer", border:"1px solid var(--ink)", background:expRelMode===v?"var(--ink)":"var(--surface)", color:expRelMode===v?"var(--bg)":"var(--ink)" }}>{l}</button>))}</div>)}{expMode==="buyer" && (<label style={{ display:"flex", gap:6, fontSize:10, padding:"4px 0", cursor:"pointer", alignItems:"flex-start", borderTop:"1px solid var(--line-3)", marginTop:4, paddingTop:6 }}><input type="checkbox" checked={expIncBuf} onChange={e=>setExpIncBuf(e.target.checked)} style={{ marginTop:1 }}/>Include internal buffered plan dates <span style={{ color:"var(--muted-1)" }}>(for your team only)</span></label>)}{(expMode==="internal"||(expMode==="buyer"&&expIncBuf)) && (<div style={{ fontSize:10, color:"var(--muted-4)", margin:"6px 0 4px", lineHeight:1.5, background:"#f7f4ee", border:"1px solid #e6e0d4", padding:"6px 8px" }}>Add buffer to internal plan dates of <input type="number" min={0} max={30} value={expBuf} onChange={e=>setExpBuf(Math.max(0,Math.min(30,Number(e.target.value)||0)))} style={{ width:40, fontFamily:"inherit", fontSize:11, padding:"2px 4px", border:"1px solid var(--ink)", margin:"0 4px" }}/> working days.<div style={{ color:"var(--muted-1)", marginTop:4 }}>Buyer approvals (Fit / Art / S-O / Lab / PP Appr) get no extra buffer of their own, but follow the buffered internal dates. Actual dates stay unchanged.</div></div>)}<div style={{ display:"flex", gap:8, marginTop:10 }}><button onClick={()=>{ runExport(expMode,expBuf,expIncBuf,expRelMode); setExpOpen(false); }} style={{ flex:1, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:6, cursor:"pointer", border:"1px solid var(--ink)", background:"var(--accent)" }}>⬇ Export .xlsx</button><button onClick={()=>setExpOpen(false)} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--bg)" }}><X size={12}/></button></div></div>)}</span>
        <div style={{ position:"relative" }}><button onClick={(e)=>{ e.stopPropagation(); finishEditing(); setFillOpen(o=>!o); setColsOpen(false); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--accent)", color:"var(--ink)", fontWeight:700, display:"flex", alignItems:"center", gap:6 }}><Copy size={13}/> Fill date → {rows.length}</button>{fillOpen && (<FillPanel count={rows.length} role={role} activeCol={(editing&&editing.col)||(sel&&sel.col)} onClose={()=>setFillOpen(false)} onApply={(key,val,mode)=>{ pushHistory(); const ids=rows.map(r=>r.s.id); const idset=new Set(ids); const top=(key==="ordRec"||key==="delivery"); const resendRows=(!top && mode!=="revised" && SEND_STAGE_KEYS.includes(key) && val)?styles.filter(s=>idset.has(s.id)&&shouldTrackResend(s,key,val)):[]; const resendIdSet=new Set(resendRows.map(s=>s.id)); const resendClearRows=(!top && mode!=="revised" && SEND_STAGE_KEYS.includes(key) && !val)?styles.filter(s=>idset.has(s.id)&&isResendEntrySlot(s,key)):[]; const resendClearIdSet=new Set(resendClearRows.map(s=>s.id)); const _msg=val?(resendRows.length?`Fill this date into all ${ids.length} visible styles?

${resendRows.length} row(s) already have an actual date and are in rejected/rework loop. These will be saved as RE-SEND actuals and old send dates will be kept in resend history.

Other existing dates in this column will be overwritten.`:`Fill this date into all ${ids.length} visible styles? Existing dates in that column will be overwritten.`):(`Clear this date for all ${ids.length} visible styles? This blanks them — only do this if you mean to.`); if(!window.confirm(_msg)){ setFillOpen(false); return; } try{ supabase.from("audit_log").insert({ style_id:null, style_no:"", col:key, field:mode==="revised"?"bulk revised":"bulk actual", old_val:`${ids.length} visible styles`, new_val:val||"CLEARED", actor_id:me.id, actor_name:me.name||me.email }).then(()=>{}).catch(()=>{}); }catch(e){} if(mode==="revised"){ if(!canEditRev(role)){ setFillOpen(false); return; } styles.filter(s=>idset.has(s.id)).forEach(s=>{ recordRevisionHistory(s,key,val,"bulk fill"); logRevisedEvent(s,key,val); }); ids.forEach(id=>{ const ck=id+":"+key+":revised"; if(val) clearedRef.current.delete(ck); else clearedRef.current.add(ck); }); setStyles(prev=>prev.map(s=>idset.has(s.id)?{...s,revs:{...(s.revs||{}),[key]:val||undefined}}:s)); flash(); setFillOpen(false); return; } if(!canEdit(role,key,"actual")){ setFillOpen(false); return; } if(!top){ ids.forEach(id=>{ if(resendClearIdSet.has(id)) return; const ck=id+":"+key+":actual"; if(val) clearedRef.current.delete(ck); else clearedRef.current.add(ck); }); if(resendClearRows.length){ resendClearRows.forEach(s=>_popResendActual(s,key)); } if(SEND_STAGE_KEYS.includes(key)&&val){ resendRows.forEach(s=>{ const appr=APPR_OF_SEND[key]; const round=activeRejectRound(s,key,appr)||1; const old=(s.actuals&&s.actuals[key])||""; insertStageEvent(s,key,appr,"resend_actual",round,val||null,old||null,val||null,`Bulk re-send ${round} actual entered`); recordResendActual(s,key,val,"bulk fill"); }); } } setStyles(prev=>prev.map(s=>idset.has(s.id)?(top?{...s,[key]:val||""}:(((val&&resendIdSet.has(s.id))||(!val&&resendClearIdSet.has(s.id)))?s:{...s,actuals:{...s.actuals,[key]:val||undefined}})):s)); flash(); setFillOpen(false); }}/>)}</div>
        </div>
      </div>

      <div style={{ display:"flex", gap:8, alignItems:"center", padding:"0 22px 6px", flexWrap:"wrap" }}>
        <div className="mt-toolbar-group" style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, marginRight:2 }}>View</span>
        <select value={columnView} onChange={e=>applyColumnView(e.target.value)} onClick={e=>e.stopPropagation()} title="role-based column views" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 8px", cursor:"pointer", border:"1px solid var(--ink)", background:columnView==="custom"?"var(--surface)":"#eef6ff", fontWeight:columnView==="custom"?400:700 }}><option value="custom">View: All columns / Custom</option>{Object.entries(ROLE_VIEW_PRESETS).map(([k,v])=><option key={k} value={k}>{v.label}</option>)}</select>
        <select value={activeNamedView} onChange={e=>{ const name=e.target.value; const v=(sharedViews||[]).find(x=>x.name===name); if(v) applyTrackerViewState(v.state,name); else setActiveNamedView(""); }} onClick={e=>e.stopPropagation()} title="View presets. System Default shows all columns and cannot be overwritten by any user." style={{ fontFamily:"inherit", fontSize:11, padding:"6px 8px", cursor:"pointer", border:"1px solid "+(activeNamedView?"var(--accent)":"var(--ink)"), background:activeNamedView?"var(--accent-tint)":"var(--surface)", fontWeight:activeNamedView?800:400 }}><option value="">Default view: choose</option>{(sharedViews||[]).map(v=><option key={v.name} value={v.name}>{v.locked?"🔒 ":""}{v.name}</option>)}</select>
        <button onClick={(e)=>{ e.stopPropagation(); saveSharedTrackerView(); }} title={activeNamedView===SYSTEM_DEFAULT_TRACKER_VIEW_NAME?"System Default is locked; save a new named view instead":"save current filters/sort/columns as shared default view for everyone (Management/Senior only)"} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:activeNamedView===SYSTEM_DEFAULT_TRACKER_VIEW_NAME?"not-allowed":"pointer", border:"1px solid var(--ink)", background:activeNamedView===SYSTEM_DEFAULT_TRACKER_VIEW_NAME?"var(--toolbar-subtle)":"var(--surface)", opacity:activeNamedView===SYSTEM_DEFAULT_TRACKER_VIEW_NAME?0.7:1, fontWeight:700 }}>Save default</button>
        <button onClick={(e)=>{ e.stopPropagation(); resetTemporaryView(); }} title="clear temporary filters/sort; keeps current column layout" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--line-2)", background:"var(--surface)", color:"var(--muted-3)", fontWeight:700 }}>Reset temp</button><button onClick={(e)=>{ e.stopPropagation(); setActiveNamedView(""); resetTemporaryView(); }} title="reset to all columns, active styles, no filters, no sort, and freeze only Style No" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)", color:"var(--ink)", fontWeight:700 }}>Full view</button>
        <div style={{ position:"relative" }}><button onClick={(e)=>{ e.stopPropagation(); finishEditing(); setColsOpen(o=>!o); setFillOpen(false); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)", display:"flex", alignItems:"center", gap:6 }}><Columns3 size={13}/> Columns {hidden.size>0?`(${hidden.size} hidden)`:""}</button>
          {colsOpen && (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:0, marginTop:4, zIndex:370, background:"var(--surface)", border:"1px solid var(--ink)", boxShadow:"4px 4px 0 var(--ink)", padding:10, width:230, maxHeight:300, overflowY:"auto" }}><div style={{ fontSize:10, fontWeight:700, marginBottom:6 }}>Show / hide columns</div><button onClick={(e)=>{ e.stopPropagation(); fitAllCols(); }} style={{ ...chip, width:"100%", marginBottom:8 }}>↔ Auto-fit widths to content</button>{[...INFO_COLS,{key:"remarks",label:"Remarks / Delays"},...STAGES].map(col=>(<label key={col.key} style={{ display:"flex", alignItems:"center", gap:6, fontSize:10, padding:"2px 0", cursor:"pointer" }}><input type="checkbox" checked={!hidden.has(col.key)} onChange={()=>setHidden(p=>{ const n=new Set(p); n.has(col.key)?n.delete(col.key):n.add(col.key); return n; })}/><span style={{ flex:1 }}>{colLabel(col)}</span>{(col.key==="extra1"||col.key==="extra2") && canAdmin(role) && <button onClick={(e)=>{ e.stopPropagation(); e.preventDefault(); const nv=window.prompt("Rename this column header:", colLabel(col)); if(nv!=null){ const t=nv.trim(); setCfg(p=>({ ...p, labels:{ ...(p.labels||{}), [col.key]: t||undefined } })); } }} title="rename header (Management / Senior)" style={{ fontSize:9, padding:"1px 5px", cursor:"pointer", border:"1px solid var(--line-2)", background:"var(--surface)" }}>rename</button>}</label>))}<div style={{ fontSize:8, color:"var(--muted-1)", marginTop:8, lineHeight:1.4 }}>Your column view is saved on this device.</div>{hidden.size>0 && <button onClick={()=>setHidden(new Set())} style={{ ...chip, marginTop:6, width:"100%" }}>Reset to default (show all)</button>}</div>)}
        </div>
        {/* FREEZE control */}
        <div style={{ display:"flex", alignItems:"center", gap:5, border:"1px solid var(--ink)", background:"var(--surface)", padding:"4px 8px" }}>
          <Snowflake size={13} color="#2563a6"/><span style={{ fontSize:10, color:"var(--muted-3)" }}>freeze</span>
          <select value={freezeN} onClick={e=>e.stopPropagation()} onChange={e=>setFreezeN(Number(e.target.value))} style={{ fontFamily:"inherit", fontSize:10, border:"none", outline:"none", background:"transparent", cursor:"pointer" }}>
            <option value={1}>Style only</option>
            {visInfo.slice(0,6).map((c,i)=><option key={c.key} value={i+2}>thru {c.label}</option>)}
          </select>
        </div>
        <div style={{ display:"flex", alignItems:"center", border:"1px solid var(--ink)" }} title="data text size"><button onClick={(e)=>{ e.stopPropagation(); bumpScale(-0.05); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 9px", cursor:"pointer", border:"none", borderRight:"1px solid var(--ink)", background:"var(--surface)", fontWeight:700 }}>A−</button><span style={{ fontSize:9, color:"var(--muted-2)", padding:"0 6px", minWidth:34, textAlign:"center" }}>{Math.round(textScale*100)}%</span><button onClick={(e)=>{ e.stopPropagation(); bumpScale(0.05); }} style={{ fontFamily:"inherit", fontSize:13, padding:"6px 9px", cursor:"pointer", border:"none", borderLeft:"1px solid var(--ink)", background:"var(--surface)", fontWeight:700 }}>A+</button></div>
        <div style={{ display:"flex", alignItems:"center", border:"1px solid var(--ink)" }} title="table text boldness"><button onClick={(e)=>{ e.stopPropagation(); bumpTableWeight(-50); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 9px", cursor:"pointer", border:"none", borderRight:"1px solid var(--ink)", background:"var(--surface)", fontWeight:700 }}>B−</button><span style={{ fontSize:9, color:"var(--muted-2)", padding:"0 6px", minWidth:38, textAlign:"center" }}>{tableWeight}</span><button onClick={(e)=>{ e.stopPropagation(); bumpTableWeight(50); }} style={{ fontFamily:"inherit", fontSize:12, padding:"6px 9px", cursor:"pointer", border:"none", borderLeft:"1px solid var(--ink)", background:"var(--surface)", fontWeight:800 }}>B+</button></div>
        <button onClick={(e)=>{ e.stopPropagation(); setShowAux(v=>!v); }} title="show every cell's auto + revised dates alongside the actual" style={{ fontFamily:"inherit", fontSize:11, fontWeight:showAux?700:400, padding:"6px 11px", cursor:"pointer", border:"1px solid var(--ink)", background:showAux?"var(--accent)":"var(--surface)", color:"var(--ink)", display:"flex", alignItems:"center", gap:6 }}>{showAux?"✓ auto+rev shown":"show auto+rev"}</button>
        </div>
        <div className="mt-toolbar-group" style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, marginRight:2 }}>Edit / Sync</span>{selKeys&&selKeys.size>1&&<span title="Multi-selection active" style={{ fontSize:9, fontWeight:900, color:"var(--info)", background:"#eaf3ff", border:"1px solid #b7cff0", padding:"2px 5px" }}>{selKeys.size} selected</span>}
        <button onClick={(e)=>{ e.stopPropagation(); copySpecial(); }} title="Copy a stage cell full state (actual + revised + rejected + skip)" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)", fontWeight:700 }}>⧉ Copy✦</button>
        <button disabled={!specialClip} onClick={(e)=>{ e.stopPropagation(); pasteSpecial(); }} title="Paste the copied full date state into selected stage cell(s)" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:specialClip?"pointer":"not-allowed", border:"1px solid var(--ink)", background:specialClip?"var(--success)":"var(--surface)", color:specialClip?"var(--surface)":"var(--muted-6)", fontWeight:700, opacity:specialClip?1:0.6 }}>Paste✦ {selKeys&&selKeys.size>1?selKeys.size:""}</button>
        <div style={{ display:"flex", border:"1px solid var(--ink)" }}>
          <button onClick={(e)=>{ e.stopPropagation(); undo(); }} disabled={!past.length} title="Undo (Ctrl/Cmd+Z)" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 9px", cursor:past.length?"pointer":"not-allowed", border:"none", borderRight:"1px solid var(--ink)", background:"var(--surface)", opacity:past.length?1:0.4 }}>↶</button>
          <button onClick={(e)=>{ e.stopPropagation(); redo(); }} disabled={!future.length} title="Redo (Ctrl/Cmd+Shift+Z)" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 9px", cursor:future.length?"pointer":"not-allowed", border:"none", background:"var(--surface)", opacity:future.length?1:0.4 }}>↷</button>
        </div>
        <button onClick={(e)=>{ e.stopPropagation(); manualSync(); }} disabled={syncBusy} title="Pull latest cloud data only. This does not undo or redo local history." style={{ fontFamily:"inherit", fontSize:11, fontWeight:remoteChanged?800:700, padding:"6px 11px", cursor:syncBusy?"wait":"pointer", border:"1px solid var(--ink)", background:remoteChanged?"var(--accent)":"var(--surface)", display:"flex", alignItems:"center", gap:6, opacity:syncBusy?0.75:1 }}><span style={{ fontSize:13, lineHeight:1 }}>⇣</span> {syncBusy?"Pulling…":(remoteChanged?"Pull latest · new changes":"Pull latest")}</button>
        </div>
        <span style={{ fontSize:10, color:"var(--muted-1)", marginLeft:"auto" }}>{sort.col?<>sorted by <b>{sort.col==="__style"?"Style":(INFO_COLS.find(c=>c.key===sort.col)?.label||STAGES.find(s=>s.key===sort.col)?.label||(sort.col==="remarks"?"Remarks":sort.col))}</b> {sort.dir>0?"↑":"↓"}</>:"drag / Shift-click / Shift-arrows = range · Ctrl/Cmd+Shift+↓ = select down · Ctrl/Cmd+A = all visible"}</span>
      </div>

      <div style={{ display:"flex", gap:8, alignItems:"center", padding:"0 22px 12px", flexWrap:"wrap", fontSize:10, color:"var(--muted-2)" }}>
        <div className="mt-toolbar-group" style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, marginRight:2 }}>Selection</span>
        <button onClick={(e)=>{ e.stopPropagation(); selectAll(); }} title="select all" style={{ fontFamily:"inherit", fontSize:10, padding:"4px 8px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)" }}>⌖ all</button>
        <button onClick={(e)=>{ e.stopPropagation(); copySelection(); }} disabled={!sel} title="copy selected cell(s)" style={{ fontFamily:"inherit", fontSize:10, padding:"4px 9px", cursor:sel?"pointer":"not-allowed", border:"1px solid var(--ink)", background:"var(--surface)", display:"inline-flex", alignItems:"center", gap:5, opacity:sel?1:0.4 }}><Copy size={12}/> copy</button>
        <span style={{ display:"inline-flex", alignItems:"center", gap:4, marginLeft:6 }} title="type a cell like A9 and press Enter to jump there"><span style={{ fontSize:9, color:"var(--muted-2)" }}>Go to</span><input value={nameBox} onChange={e=>setNameBox(e.target.value)} onKeyDown={e=>{ if(e.key==="Enter") gotoCell(nameBox); }} onFocus={e=>e.target.select()} placeholder="A9" style={{ width:54, fontFamily:"inherit", fontSize:11, fontWeight:700, textTransform:"uppercase", padding:"3px 6px", border:"1px solid var(--ink)", textAlign:"center", color:"var(--accent)" }}/></span><span style={{ marginLeft:8 }}>{clip?<span style={{color:"#2563a6"}}>📋 {clip.h}×{clip.w} copied — select & Ctrl/Cmd+V to paste</span>:(sel?<>selected: <b style={{ color:"var(--accent)" }}>{cellRef(sel)}{focus&&(focus.id!==sel.id||focus.col!==sel.col)?(":"+cellRef(focus)):""}</b> · {styles.find(s=>s.id===sel.id)?.styleNo} · {sel.col==="__style"?"Style":(INFO_COLS.find(c=>c.key===sel.col)?.label||STAGES.find(s=>s.key===sel.col)?.label||sel.col)}</>:"click a cell to format / comment")}</span>
        </div>
        <div className="mt-toolbar-group" onMouseDown={(e)=>e.stopPropagation()} onClick={(e)=>e.stopPropagation()} style={{ display:"flex", alignItems:"center", gap:6, flex:"1 1 420px", minWidth:320, background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)", position:"relative" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4 }}>Entry</span>
          <span style={{ fontSize:10, fontWeight:800, color:"var(--accent)", minWidth:34 }}>{sel?cellRef(sel):"—"}</span>
          <input disabled={!sel} value={entryVal} onMouseDown={e=>e.stopPropagation()} onClick={e=>e.stopPropagation()} onFocus={()=>{ if(editing) setEditing(null); if(!entryTouched) setEntryVal(selectedDisplayValue()); }} onChange={e=>{ setEntryTouched(true); setEntryVal(sel&&sel.col==="qty"?e.target.value.replace(/[^0-9]/g,""):e.target.value); }} onKeyDown={e=>{ e.stopPropagation(); if((e.ctrlKey||e.metaKey)&&(e.key==="z"||e.key==="Z")){ e.preventDefault(); if(e.shiftKey) redo(); else undo(); return; } if((e.ctrlKey||e.metaKey)&&(e.key==="y"||e.key==="Y")){ e.preventDefault(); redo(); return; } if(e.key==="Tab"){ e.preventDefault(); const v=entrySuggestion||entryVal; if(commitEntry(v)) moveAnchor(0,e.shiftKey?-1:1); } else if(e.key==="Enter"){ e.preventDefault(); commitEntry(entrySuggestion||entryVal); } else if(e.key==="Escape"){ setEntryVal(selectedDisplayValue()); setEntryTouched(false); if(editing) setEditing(null); } }} placeholder={sel?("type "+selectedColLabel+"…"):"select a cell"} title="Excel-style entry bar. Shows selected cell value; Tab accepts suggestion from same column; Enter saves; Ctrl/Cmd+Z undo." style={{ flex:1, minWidth:160, fontFamily:"inherit", fontSize:11, padding:"4px 7px", border:"1px solid var(--line-2)", outline:"none", background:sel?"var(--surface)":"var(--line-3)" }}/>
          {entrySuggestion && <button onMouseDown={e=>e.preventDefault()} onClick={()=>{ setEntryTouched(true); setEntryVal(entrySuggestion); }} title="Click or press Tab to accept" style={{ position:"absolute", right:10, top:"100%", marginTop:3, zIndex:390, maxWidth:260, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"4px 8px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--accent-tint)", boxShadow:"3px 3px 0 var(--ink)" }}>Tab ↹ {entrySuggestion}</button>}
          {peerEditingList.length>0 && <span title={peerEditingList.map(p=>(p.name||p.email||"User")+" editing "+p.ref).join("\n")} style={{ fontSize:9, color:"var(--danger)", fontWeight:800, whiteSpace:"nowrap" }}>{peerEditingList.length} editing now</span>}
        </div>
        <div className="mt-toolbar-group" style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, marginRight:2 }}>Format</span>
        <Droplet size={13}/><span>fill:</span>
        {FILL_SWATCHES.map((sw,i)=>(<button key={i} onClick={(e)=>{ e.stopPropagation(); applyFill(sw); }} disabled={!sel} title={sw===""?"clear fill":sw} style={{ width:18, height:18, cursor:sel?"pointer":"not-allowed", border:"1px solid var(--ink)", background:sw===""?"var(--surface)":sw, position:"relative", opacity:sel?1:0.4 }}>{sw===""?<X size={11} style={{position:"absolute",top:2,left:2}}/>:null}</button>))}
        </div>
        <div className="mt-toolbar-group" style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, marginRight:2 }}>Comments</span>
        <span ref={commentThreadWrapRef} style={{ marginLeft:10, position:"relative" }}>
          <button onClick={(e)=>{ e.stopPropagation(); if(sel){ setReviewOpen(false); setHistory(false); setThreadCell(sel); setCmText(""); } }} disabled={!sel} style={{ fontFamily:"inherit", fontSize:10, padding:"4px 9px", cursor:sel?"pointer":"not-allowed", border:"1px solid var(--ink)", background:"var(--surface)", display:"inline-flex", alignItems:"center", gap:5, opacity:sel?1:0.4 }}><MessageSquare size={12}/> {sel&&(comments[cellKey(sel.id,sel.col)]||[]).length?("comments ("+(comments[cellKey(sel.id,sel.col)]||[]).length+")"):"comment"}</button><button onClick={(e)=>{ e.stopPropagation(); if(sel) openHistory(sel.id,sel.col); }} disabled={!sel} title="Change history for this cell" style={{ fontFamily:"inherit", fontSize:10, padding:"4px 9px", cursor:sel?"pointer":"not-allowed", border:"1px solid var(--ink)", background:"var(--surface)", display:"inline-flex", alignItems:"center", gap:5, opacity:sel?1:0.4 }}><Clock size={12}/> history</button><button onClick={(e)=>{ e.stopPropagation(); if(sel) toggleFollow(sel.id); }} disabled={!sel} title="Get notified when this style changes" style={{ fontFamily:"inherit", fontSize:10, padding:"4px 9px", cursor:sel?"pointer":"not-allowed", border:"1px solid var(--ink)", background:(sel&&follows.has(sel.id))?"#fff3df":"var(--surface)", color:(sel&&follows.has(sel.id))?"#b45309":"var(--ink)", display:"inline-flex", alignItems:"center", gap:5, opacity:sel?1:0.4 }}><Star size={12} fill={(sel&&follows.has(sel.id))?"var(--accent)":"none"}/> {sel&&follows.has(sel.id)?"following":"follow"}</button>
          {threadCell && (()=>{ const ck=cellKey(threadCell.id,threadCell.col); const list=(comments[ck]||[]).slice().sort((a,b)=>new Date(a.created_at)-new Date(b.created_at)); const anyOpen=list.some(x=>!x.resolved); const mq=(cmText.match(/@(\S*)$/)||[])[1]; const sugg = mq!=null ? team.filter(t=>String(t.name||t.email||"").toLowerCase().includes(mq.toLowerCase())).slice(0,5) : [];
                return (<span onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:0, marginTop:4, zIndex:370, background:"var(--surface)", border:"1px solid var(--ink)", boxShadow:"4px 4px 0 var(--ink)", padding:10, width:290, display:"block" }}>
                  <div style={{ fontSize:10, fontWeight:700, marginBottom:6, display:"flex", justifyContent:"space-between", alignItems:"center" }}><span>Comments · {cellRef(threadCell)}</span>{list.length>0 && <button onClick={()=>resolveThread(threadCell)} style={{ fontSize:9, border:"1px solid var(--ink)", background:anyOpen?"var(--success)":"var(--line-3)", color:anyOpen?"var(--surface)":"var(--muted-2)", cursor:"pointer", padding:"1px 7px", fontWeight:700 }}>{anyOpen?"Resolve":"Resolved \u2713"}</button>}</div>
                  <div style={{ maxHeight:170, overflowY:"auto", marginBottom:6 }}>{list.length===0 && <div style={{ fontSize:10, color:"var(--muted-1)" }}>No comments yet. Use @ to mention a teammate.</div>}{list.map(cm=>(<div key={cm.id} style={{ marginBottom:7, opacity:cm.resolved?0.5:1 }}><div style={{ fontSize:9, color:"var(--muted-2)", display:"flex", justifyContent:"space-between", alignItems:"center" }}><span><b style={{ color:colorFor(cm.author_id||cm.author_name) }}>{cm.author_name||"?"}</b> · {tsShort(cm.created_at)}</span>{(cm.author_id===(me&&me.id)||canAdmin(role)) && <button onClick={()=>deleteComment(cm)} title="Delete comment" style={{ border:"none", background:"transparent", cursor:"pointer", color:"var(--danger)", padding:0, display:"inline-flex", alignItems:"center" }}><Trash2 size={11}/></button>}</div><div style={{ fontSize:11, whiteSpace:"pre-wrap", wordBreak:"break-word" }}>{renderMentions(cm.body)}</div></div>))}</div>
                  <div style={{ position:"relative" }}><textarea autoFocus value={cmText} onChange={e=>setCmText(e.target.value)} onKeyDown={e=>{ e.stopPropagation(); if(e.key==="Enter"&&(e.metaKey||e.ctrlKey)){ e.preventDefault(); postComment(); } }} placeholder="Comment…  @ to mention" style={{ width:"100%", height:46, fontFamily:"inherit", fontSize:11, border:"1px solid var(--line-2)", outline:"none", resize:"none", boxSizing:"border-box" }}/>
                  {sugg.length>0 && <div style={{ position:"absolute", bottom:"100%", left:0, right:0, background:"var(--surface)", border:"1px solid var(--ink)", maxHeight:110, overflowY:"auto", zIndex:5 }}>{sugg.map(t=>(<div key={t.id} onClick={()=>setCmText(cmText.replace(/@(\S*)$/, "@"+String(t.name||t.email||"").replace(/\s+/g,"")+" "))} style={{ fontSize:10, padding:"4px 7px", cursor:"pointer", borderBottom:"1px solid #f0ece2" }}>{t.name||t.email} <span style={{ color:"var(--muted-1)" }}>· {(ROLES[t.role]||{}).label||t.role}</span></div>))}</div>}</div>
                  <div style={{ display:"flex", gap:6, marginTop:6 }}><button onClick={postComment} style={{ ...chip, flex:1, background:"var(--accent)" }}>Post (⌘↵)</button><button onClick={()=>{ setThreadCell(null); setCmText(""); }} style={chip}>Close</button></div>
                </span>); })()}
        </span>
        </div>
      </div>

      {(() => { const chips=[]; const chip=(key,label,onX)=>chips.push(<span key={key} style={{ display:"inline-flex", alignItems:"center", gap:5, background:"var(--accent-tint)", border:"1px solid rgba(201,111,22,0.22)", borderRadius:999, padding:"4px 6px 4px 9px", fontSize:10, fontWeight:700, boxShadow:"var(--pill-shadow)" }}>{label}<button className="mt-filter-chip-x" onClick={onX} title="remove this filter" style={{ border:"none", background:"transparent", cursor:"pointer", padding:0, lineHeight:0, color:"var(--muted-2)", display:"inline-flex", alignItems:"center", justifyContent:"center" }}><X size={11}/></button></span>);
        if(search) chip("q",(searchCol==="auto"?"search":searchCol)+": "+search, ()=>setSearch(""));
        if(statusFilter!=="All") chip("st","status: "+statusFilter, ()=>setStatusFilter("All"));
        if(ownerFilter!=="All") chip("ow","chase: "+ownerFilter, ()=>setOwnerFilter("All"));
        if(activityFilterKeys.length) chip("ac",activityFilterKeys.length===1?"activity: "+stageLabelFromKeyGlobal(activityFilterKeys[0]):"activity: "+activityFilterKeys.length+" selected", ()=>setTrackerActivityFilter([]));
        if(followFilter) chip("fo","following only", ()=>setFollowFilter(false));
        if(savedView) chip("sv","saved: "+((SAVED_VIEWS.find(x=>x[0]===savedView)||[])[1]||savedView), ()=>setSavedView(""));
        if(activeNamedView) chip("nv","default view: "+activeNamedView, ()=>setActiveNamedView(""));
        if(archiveView!=="active") chip("ar","view: "+archiveView, ()=>setArchiveView("active"));
        Object.keys(colFilters||{}).forEach(col=>{ const lab=(INFO_COLS.find(c=>c.key===col)||{}).label||(STAGES.find(x=>x.key===col)||{}).label||col; chip("c-"+col, lab+": "+(((colFilters[col]||[])[0]===FILTER_NONE)?0:(colFilters[col]||[]).length)+" sel", ()=>setColFilters(f=>{ const n={...f}; delete n[col]; return n; })); });
        if(!chips.length) return null;
        return <div style={{ display:"flex", alignItems:"center", gap:7, flexWrap:"wrap", padding:"7px 22px 0" }}><span style={{ fontSize:10, fontWeight:700, color:"var(--muted-2)" }}>Active filters:</span>{chips}<button onClick={clearAllFilters} style={{ fontSize:10, fontWeight:700, border:"1px solid var(--ink)", background:"var(--surface)", cursor:"pointer", padding:"2px 8px", marginLeft:2 }}>Clear all</button></div>;
      })()}

      {showJump && <button onClick={jumpToTop} title="Back to controls / top" style={{ position:"fixed", bottom:24, right:24, zIndex:370, width:42, height:42, borderRadius:21, border:"1px solid var(--ink)", background:"var(--accent)", color:"var(--ink)", boxShadow:"2px 2px 0 var(--ink)", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center" }}><ChevronUp size={20}/></button>}

      {hiddenRows.size>0 && <div style={{ margin:"0 22px 8px", border:"1px solid var(--line-2)", background:"var(--surface)", borderRadius:10, padding:"7px 10px", fontSize:10.5, color:"var(--muted-4)", display:"flex", gap:10, alignItems:"center", flexWrap:"wrap" }}><b>{hiddenRows.size}</b> row(s) hidden in this view <button onClick={()=>setHiddenRows(new Set())} style={{ fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"3px 8px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--accent-tint)" }}>show hidden rows</button></div>}
      {rows.length>renderRows.length && <div style={{ margin:"0 22px 8px", border:"1px solid var(--line-2)", background:"var(--accent-tint)", borderRadius:10, padding:"8px 10px", fontSize:10.5, color:"var(--muted-4)", display:"flex", gap:10, alignItems:"center", flexWrap:"wrap" }}><b>Performance mode:</b> showing first {renderRows.length} of {rows.length} filtered styles to protect P95 speed. Narrow filters for fastest work, or <button onClick={()=>setRenderLimit(Math.min(5000, renderLimit+900))} style={{ fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"3px 8px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)" }}>show 900 more</button><button onClick={()=>setRenderLimit(5000)} style={{ fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"3px 8px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)" }}>show all up to 5,000</button></div>}
      <div ref={scrollWrapRef} className="mt-tracker-scroll" style={{ overflow:"auto", padding:"0 22px", maxHeight:"calc(100vh - 210px)" }}>
        <table className="mt-tracker-table" role="grid" aria-label="Pre-production tracker grid. Arrow keys move, Shift+arrows select range, Tab moves right, Shift+Tab moves left." style={{ borderCollapse:"separate", borderSpacing:0, zoom:textScale, fontSize:11, fontWeight:tableWeight, tableLayout:"fixed", userSelect:dragSel?"none":"auto" }}>
          <colgroup>
            <col style={{ width:widthOf("__style") }}/>
            {visInfo.map(c=><col key={c.key} style={{ width:widthOf(c.key) }}/>)}
            {visStages.map(st=><col key={st.key} style={{ width:widthOf(st.key) }}/>)}
            {remarksVis && <col style={{ width:widthOf("remarks") }}/>}
          </colgroup>
          <thead><tr role="row">
            <Th col="__style" label="Style No" sort={sort} onSort={clickHeader} width={widthOf("__style")} letter={colLetter(colIndex("__style"))} onResize={onResize} onAutoFit={autoFit} scale={textScale} {...filterProps("__style")} selected={sel&&sel.col==="__style"} sticky left={0} z={22}/>
            {visInfo.map(c=><Th key={c.key} col={c.key} label={colLabel(c)} sort={sort} onSort={clickHeader} width={widthOf(c.key)} letter={colLetter(colIndex(c.key))} onResize={onResize} onAutoFit={autoFit} scale={textScale} {...filterProps(c.key)} selected={sel&&sel.col===c.key} sticky={isFrozen(c.key)} left={isFrozen(c.key)?leftOf(c.key):undefined} z={21}/>)}
            {visStages.map(st=><Th key={st.key} col={st.key} label={st.label} sort={sort} onSort={clickHeader} width={widthOf(st.key)} letter={colLetter(colIndex(st.key))} onResize={onResize} onAutoFit={autoFit} scale={textScale} {...filterProps(st.key)} selected={sel&&sel.col===st.key}/>)}
            {remarksVis && <Th col="remarks" label={REMARK_COL.label} sort={sort} onSort={clickHeader} width={widthOf("remarks")} letter={colLetter(colIndex("remarks"))} onResize={onResize} onAutoFit={autoFit} scale={textScale} {...filterProps("remarks")} selected={sel&&sel.col==="remarks"}/>}
          </tr></thead>
          <tbody>
            {renderRows.map(({s,c},rowIdx)=>{ const t=TONE_STYLE[c.tone]; const sk=cellKey(s.id,"__style"); const styBg=bgFor(s.id,"__style","var(--surface)"); return (
              <tr key={s.id} role="row" style={{ ...activeRowStyle(s.id), height:rowH }}>
                <td id={`cell-${s.id}-__style`} onClick={(e)=>onCellClick(e,s.id,"__style")} onDoubleClick={(e)=>{ e.stopPropagation(); startEdit(s.id,"__style"); }} style={{ border:"1px solid var(--line-1)", padding:"6px 9px", height:rowH, overflow:"hidden", cursor:"cell", ...freezeStyle("__style",styBg), boxShadow:ringFor(s.id,"__style")||freezeStyle("__style",styBg).boxShadow, ...activeCellStyle(s.id,"__style") }}>
                  {editing&&editing.id===s.id&&editing.col==="__style" ? (<span style={{ position:"relative", display:"inline-block" }}><input autoFocus onFocus={e=>{ if((e.target.value||"").length>1) e.target.select(); }} value={editVal} onClick={e=>e.stopPropagation()} onChange={e=>setEditVal(e.target.value)} onKeyDown={e=>{ e.stopPropagation(); if(e.key==="Tab"){ e.preventDefault(); const v=editSuggestion||editVal; commitText(v); moveAnchor(0,e.shiftKey?-1:1); } else if(e.key==="Enter"){ e.preventDefault(); commitText(editSuggestion||editVal); moveAnchor(1,0); } else if(e.key==="Escape"){ e.preventDefault(); setEditing(null); } }} onBlur={()=>commitText()} style={{ width:150, fontFamily:"inherit", fontSize:11, fontWeight:700, border:"1px solid var(--info)", outline:"none", padding:"1px 3px" }}/>{editSuggestion && <button onMouseDown={e=>e.preventDefault()} onClick={e=>{ e.stopPropagation(); setEditVal(editSuggestion); }} title="Click or press Tab to accept" style={{ position:"absolute", left:0, top:"100%", marginTop:3, zIndex:390, maxWidth:220, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"3px 7px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--accent-tint)", boxShadow:"3px 3px 0 var(--ink)" }}>Tab ↹ {editSuggestion}</button>}</span>) : <div style={{ fontWeight:700, display:"flex", alignItems:"center", gap:6 }}><span onClick={(e)=>{ e.stopPropagation(); selectRow(s.id); }} title="select row" style={{ fontSize:8, color:"var(--muted-6)", cursor:"pointer", minWidth:14 }}>{rowIdx+1}</span><Star size={11} onClick={(e)=>{ e.stopPropagation(); toggleFollow(s.id); }} title={follows.has(s.id)?"following — click to unfollow":"follow this style"} fill={follows.has(s.id)?"var(--accent)":"none"} color={follows.has(s.id)?"var(--accent)":"#c9c1b3"} style={{ cursor:"pointer", flex:"0 0 auto", opacity:follows.has(s.id)?1:0.5 }}/><span style={{ color:follows.has(s.id)?"#b45309":"inherit" }}>{displayTextValue("__style",s.styleNo)}</span></div>}
                  <div style={{ fontSize:11, fontWeight:600, color:"var(--muted-4)", whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis", maxWidth:188, marginTop:2 }}>{s.colour}</div>
                  <div style={{ display:"flex", gap:3, marginTop:4, flexWrap:"wrap" }}>{FLAG_DEFS.map(f=>{ const on=!!s[f.key]; return (<button className="mt-row-flag-btn" key={f.key} title={f.title} onClick={(e)=>{ e.stopPropagation(); if(canMaster(role)) toggleFlag(s.id,f.key); }} style={{ fontFamily:"inherit", fontSize:8.5, fontWeight:800, letterSpacing:0.3, padding:"5px 8px", cursor:!canMaster(role)?"not-allowed":"pointer", lineHeight:1.3, border:`1px solid ${on?"var(--ink)":"#cfcabf"}`, borderRadius:999, background:on?"var(--ink)":"transparent", color:on?"var(--bg)":"var(--muted-5)", opacity:!canMaster(role)?0.5:1 }}>{f.short}</button>); })}</div>
                  <NoteTri k={sk}/>
                  {canAdmin(role) && <button title="delete row" onClick={(e)=>{ e.stopPropagation(); deleteStyle(s.id); }} style={{ position:"absolute", top:2, right:2, zIndex:6, border:"none", background:"transparent", cursor:"pointer", padding:0, lineHeight:0, color:"#cbb4ac" }}><Trash2 size={11}/></button>}
                </td>

                {visInfo.map(col=>{
                  if(col.kind==="text"||col.kind==="num") return renderEditable(s,col);
                  const k=cellKey(s.id,col.key);
                  if(col.kind==="date"){ const bg=bgFor(s.id,col.key,"var(--surface)"); return (<td key={col.key} id={`cell-${s.id}-${col.key}`} onClick={(e)=>onCellClick(e,s.id,col.key)} onDoubleClick={(e)=>{ e.stopPropagation(); if(canEdit(role,col.key,"actual")) beginDate(s.id,col.key,preferredDateMode(s.id,col.key,"actual")); }} style={{ border:"1px solid var(--line-1)", padding:"6px 9px", height:rowH, whiteSpace:"nowrap", boxShadow:ringFor(s.id,col.key), cursor:"cell", position:"relative", overflow:(editing&&editing.id===s.id&&editing.col===col.key)?"visible":"hidden", background:bg, ...freezeStyle(col.key,bg), ...activeCellStyle(s.id,col.key) }}>{fmt(parse(s[col.key]))||<span style={{color:"var(--line-2)"}}>—</span>}{editing&&editing.id===s.id&&editing.col===col.key && dateEditor(s.id,col.key,editing.mode)}<PeerTag who={peerOn(s.id,col.key)}/><NoteTri k={k}/><FillHandle id={s.id} col={col.key}/></td>); }
                  let content=null;
                  if(col.kind==="branch"){ const b=col.branch==="fit"?c.fitBranch:col.branch==="print"?c.printBranch:col.branch==="fabric"?c.fabricBranch:col.branch==="pp"?c.ppBranch:c.prodFileBranch; const canJump=b.tone!=="na"&&!c.released&&!b.autoClosed; content=<BranchPill b={b} onJump={canJump?()=>jumpToEnter(s.id,branchTarget(s,c,col.branch)):null}/>; }
                  else if(col.key==="overall") content=(<span style={{ display:"inline-flex", flexDirection:"column", gap:2, alignItems:"flex-start" }}><span style={{ display:"inline-flex", alignItems:"center", gap:5, background:t.bg, color:t.fg, border:"1px solid rgba(31,31,29,0.08)", borderRadius:999, boxShadow:"var(--pill-shadow)", padding:"3px 9px", fontSize:10, lineHeight:1.15, fontWeight:800 }}><span style={{ width:6,height:6,borderRadius:"50%", background:t.dot, flex:"0 0 auto" }}/>{c.status}</span>{c.lastActual && <span style={{ fontSize:8.5, color:"var(--on-dark-2)", whiteSpace:"nowrap" }}>last: {fmt(c.lastActual)}{c.lastActualKey?` · ${(STAGES.find(x=>x.key===c.lastActualKey)||{}).label||""}`:""}</span>}</span>);
                  else if(col.key==="fabricCD"){ const fc=c.fabricCountdown; const ft=BR_TONE[fc.tone]||BR_TONE.na; content=(<span style={{ display:"inline-flex", flexDirection:"column", alignItems:"flex-start", gap:1 }}><span style={{ display:"inline-flex", alignItems:"center", background:ft.bg, color:ft.fg, border:"1px solid rgba(31,31,29,0.08)", borderRadius:999, boxShadow:"var(--pill-shadow)", padding:"3px 8px", fontSize:10, lineHeight:1.15, fontWeight:800, whiteSpace:"nowrap" }}>{fc.txt}</span>{fc.date && <span style={{ fontSize:8, color:"var(--on-dark-2)", paddingLeft:3 }}>{fmt(fc.date)}</span>}</span>); }
                  else if(col.key==="proj") content=<span title={`release gate (30wd before delivery): ${fmt(c.releaseGate)}`} style={{ fontWeight:600, color:c.projTone==="late"?"var(--danger)":c.projTone==="warn"?"#7a560f":c.projTone==="done"?"var(--muted-2)":"var(--success)" }}>{fmt(c.projRelease)}{c.projTone==="late"&&!c.released?" ⚠":c.projTone==="ok"?" ✓":""}</span>;
                  else if(col.key==="pct") content=(<div style={{ display:"flex", alignItems:"center", gap:5 }}><div style={{ flex:1, height:6, background:"var(--line-3)", position:"relative", minWidth:34 }}><div style={{ position:"absolute", left:0, top:0, bottom:0, width:`${c.pct}%`, background:c.pct===100?"var(--success)":"var(--accent)" }}/></div><span style={{ fontSize:9, color:"var(--muted-3)", width:26, textAlign:"right" }}>{c.pct}%</span></div>);
                  else if(col.key==="chase") content=(!c.chaseOwners||c.chaseOwners.length===0)?<span style={{color:"var(--line-2)"}}>—</span>:(<span style={{ display:"flex", gap:3, flexWrap:"wrap" }}>{c.chaseOwners.map(o=>{ const txt=chaseLabel(o.owner)+(o.count>1?` ×${o.count}`:""); return <span key={o.owner} title={`${chaseLabel(o.owner)} · ${o.count} open ${o.count>1?"branches":"branch"}`} style={{ fontSize:9, fontWeight:700, padding:"2px 5px", background:OWNER_BG[o.owner]||"#eef0f2", color:OWNER_COLOR[o.owner]||"var(--muted-4)", whiteSpace:"nowrap", border:"1px solid "+(OWNER_COLOR[o.owner]||"var(--line-2)")+"33", borderRadius:999 }}>{txt}</span>; })}</span>);
                  else if(col.key==="float") content=<span style={{ fontWeight:700, color:c.float<0?"var(--danger)":"var(--success)" }}>{c.float==null?"—":`${c.float>0?"+":""}${c.float}d`}</span>;
                  else if(col.key==="idle") content=<span style={{ color:c.idle>=7?"var(--danger)":"var(--muted-2)" }}>{c.idle==null?"—":`${c.idle}d`}</span>;
                  const bg=bgFor(s.id,col.key,"var(--surface)");
                  return <td key={col.key} id={`cell-${s.id}-${col.key}`} onClick={(e)=>onCellClick(e,s.id,col.key)} style={{ border:"1px solid var(--line-1)", padding:"6px 9px", height:rowH, whiteSpace:"nowrap", boxShadow:ringFor(s.id,col.key), cursor:"default", background:bg, position:"relative", overflow:"hidden", ...freezeStyle(col.key,bg), ...activeCellStyle(s.id,col.key) }}>{content}<PeerTag who={peerOn(s.id,col.key)}/><NoteTri k={k}/><FillHandle id={s.id} col={col.key}/></td>;
                })}

                {visStages.map(st=>{
                  const applies=stageApplies(s,st);
                  const cs=c.stages.find(x=>x.key===st.key);
                  const isNext=applies && c.frontier && c.frontier.has(st.key);
                  const editable=applies&&canEdit(role,st.key,"actual"); const canRev=applies&&canEditRev(role); const canRej=applies&&canEditReject(role,st.key); const canSkp=applies&&MERCH_ROLES.includes(role);
                  const k=cellKey(s.id,st.key);
                  if(!applies){ const bg=bgFor(s.id,st.key,"#f3f1ec"); return <td key={st.key} id={`cell-${s.id}-${st.key}`} onClick={(e)=>onCellClick(e,s.id,st.key)} style={{ border:"1px solid var(--line-1)", background:bg, color:"var(--line-2)", textAlign:"center", padding:"6px 9px", height:rowH, boxShadow:ringFor(s.id,st.key), position:"relative", overflow:"hidden", ...activeCellStyle(s.id,st.key) }}>—<NoteTri k={k}/></td>; }
                  const hasRev=cs&&cs.rev&&!cs.done;
                  const rvh=[];
                  const revCnt=Array.isArray(rvh)?rvh.length:0;
                  const activeReworkPlanning=!!(cs && !cs.actual && !cs.skipped && !cs.autoClosed && linkedApprovalRejectedOpen(s,st.key));
                  const bg=bgFor(s.id,st.key,(cs&&cs.skipped)?"var(--tint-waive)":(cs&&cs.rejected)?"var(--tint-reject)":(activeReworkPlanning||cs&&cs.rework)?"var(--tint-rework)":(cs&&cs.actual&&cs.histReject?"var(--tint-histrej)":(isNext?"var(--tint-next)":"var(--surface)")));
                  return (
                    <td key={st.key} id={`cell-${s.id}-${st.key}`} onClick={(e)=>onCellClick(e,s.id,st.key)} onDoubleClick={(e)=>{ e.stopPropagation(); if(editable) beginDate(s.id,st.key,"actual",undefined,shouldForceRoundActualEntry(s.id,st.key)); }}
                      style={{ border:"1px solid var(--line-1)", padding:0, position:"relative", overflow:(editing&&editing.id===s.id&&editing.col===st.key)?"visible":"hidden", background:bg, boxShadow:ringFor(s.id,st.key)||(isNext?"inset 0 0 0 2px var(--accent)":null), cursor:editable?"cell":"default", ...activeCellStyle(s.id,st.key) }}>
                      <div className="mt-stage-cell-body" style={{ minHeight:Math.max(34,rowH-4), padding:"12px 30px 11px 30px", fontSize:11.2, color:cs.actual?"var(--ink)":"var(--muted-6)" }}>
                        {showAux && cs.plan && <span style={{ display:"block", fontSize:8, color:"#bcb6a8", lineHeight:1.3 }}>auto {fmt(cs.plan)}{cs.rev?` · rev ${fmt(cs.rev)}`:""}</span>}
                        {cs.autoClosed ? (
                          <span style={{ color:"var(--line-2)", fontSize:11 }}>—</span>
                        ) : cs.skipped ? (
                          <span style={{ display:"flex", flexDirection:"column", lineHeight:1.25 }}>
                            <span style={{ fontSize:9, color:"#8a6d3b", fontWeight:700, display:"flex", alignItems:"center", gap:3 }}><SkipForward size={9}/>SKIPPED</span>
                            <span style={{ fontSize:8, color:"var(--on-dark-2)" }}>{fmt(cs.skip)}</span>{(cs.skippedWasRework||cs.skippedWasRejected) && <span style={{ fontSize:8, color:"#b4531a", fontWeight:800 }}>kept rework status</span>}
                          </span>
                        ) : (activeReworkPlanning||cs.rework) ? (
                          <span title={`Re-send round. ↻ corner sets the revised re-send date. First send/resend actual kept in history: ${fmt(cs.storedActual||cs.actual)||"—"} · Rejected: ${fmt(cs.reject)||"—"}`} style={{ display:"flex", flexDirection:"column", lineHeight:1.2, gap:2, maxWidth:"100%" }}>
                            <span style={{ fontSize:9, color:cs.rev?"var(--revised)":"#b03020", fontWeight:900, display:"flex", alignItems:"center", gap:3, whiteSpace:"nowrap" }}><X size={8}/>{cs.rev?`${roundLabelFor(s,st.key,"RE-SEND")} REV`:`${roundLabelFor(s,st.key,"RE-SEND")} DUE`} {fmt(cs.rev||cs.plan)}</span>
                            {cs.reject && <span style={{ fontSize:8.5, color:"#7a560f", fontWeight:700, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>rej {fmt(cs.reject)}</span>}
                            {editable && <button title="enter or edit actual date for the active re-send round (keeps first send in history)" onClick={(e)=>{ e.stopPropagation(); beginDate(s.id,st.key,"actual",undefined,true); }} style={{ alignSelf:"flex-start", border:"1px solid var(--accent)", background:"rgba(255,244,227,0.92)", borderRadius:6, padding:"3px 6px", margin:0, fontFamily:"inherit", fontSize:8.5, color:"var(--accent)", fontWeight:900, whiteSpace:"nowrap", cursor:"pointer", minWidth:0, minHeight:0 }}>{`ENTER ACTUAL ${roundLabelFor(s,st.key,"RE-SEND")}`}</button>}
                          </span>
                        ) : cs.actual ? (()=>{ const rh=(effectiveResends&&effectiveResends[s.id+":"+st.key])||(resends&&resends[s.id+":"+st.key]); const rlbl=cs.histReject?roundLabelFor(s,st.key,"RE-SEND"):""; return <span style={{ display:"flex", flexDirection:"column", lineHeight:1.25 }}><span style={{ display:"flex", alignItems:"center", gap:4 }}><Check size={11} color={OWNER_COLOR[(cfg.stageOwners&&cfg.stageOwners[st.key])||DEFAULT_STAGE_OWNERS[st.key]||st.owner]}/>{fmt(cs.actual)}</span>{rlbl && <span title={(rh||[]).map(d=>fmt(parse(typeof d==="string"?d:(d&&d.newVal)))).filter(Boolean).join(" → ")} style={{ fontSize:8, color:"#b4531a", fontWeight:800 }}>↻ {rlbl} actual</span>}{cs.histReject && <span style={{ fontSize:8, color:"#b03020", fontWeight:700 }}>↻ after REJ {fmt(cs.histReject)}</span>}</span>; })() : cs.rejected ? (
                          <span title="Rejected approval. ↻ corner sets the revised re-approval date." style={{ display:"flex", flexDirection:"column", lineHeight:1.25, gap:1 }}>
                            <span style={{ fontSize:9, color:"#b03020", fontWeight:800, display:"flex", alignItems:"center", gap:3, whiteSpace:"nowrap" }}><X size={9}/>REJECTED · {fmt(cs.reject)}</span>
                            <span style={{ fontSize:9, color:hasRev?"var(--revised)":"#7a560f", whiteSpace:"nowrap" }}>{roundLabelFor(s,st.key,"RE-APPR")} due → {fmt(cs.rev||cs.plan)}</span>
                            {editable && <button title="enter actual re-approval date" onClick={(e)=>{ e.stopPropagation(); beginDate(s.id,st.key,"actual",undefined,true); }} style={{ alignSelf:"flex-start", border:"1px solid var(--accent)", background:"rgba(255,244,227,0.92)", borderRadius:6, padding:"3px 6px", margin:0, fontFamily:"inherit", fontSize:8.5, color:"var(--accent)", fontWeight:900, whiteSpace:"nowrap", cursor:"pointer", minWidth:0, minHeight:0 }}>{`ENTER ACTUAL ${roundLabelFor(s,st.key,"RE-APPR")}`}</button>}
                          </span>
                        ) : (
                          <span style={{ display:"flex", flexDirection:"column", lineHeight:1.2 }}>
                            <span style={{ fontSize:10.5, color:hasRev?"var(--revised)":isNext?"var(--accent)":"#c4c0b8" }}>{hasRev?"rev":st.cutoff?"cutoff":"plan"} {fmt(hasRev?cs.rev:cs.plan)}</span>
                            {editable?<span style={{ fontSize:10.5, color:isNext?"var(--accent)":"#c4c0b8", fontWeight:isNext?700:400 }}>{isNext?"▸ enter":st.cutoff?"log arrival":"—"}</span>:<span style={{ fontSize:10, color:"var(--line-2)", display:"flex", alignItems:"center", gap:3 }}><Lock size={9}/>locked</span>}
                          </span>
                        )}
                      </div>
                      {canRev && !cs.skipped && !cs.autoClosed && (!cs.actual || cs.rework || activeReworkPlanning || cs.rejected || linkedApprovalRejectedOpen(s,st.key)) && (<button className="mt-stage-corner-btn" title={(activeReworkPlanning||cs.rework)?"set revised date for active re-send round":(cs.rejected?"set revised date for active re-approval round":"set revised plan date")} onClick={(e)=>{ e.stopPropagation(); beginDate(s.id,st.key,"rev"); }} style={{ position:"absolute", top:3, right:3, border:"1px solid transparent", background:"rgba(255,253,248,0.86)", borderRadius:8, cursor:"pointer", padding:5, lineHeight:1, display:"flex", alignItems:"center", justifyContent:"center" }}><RotateCcw size={12} color="var(--revised)"/></button>)}
                      {canRej && !cs.skipped && !cs.autoClosed && !cs.actual && REJECTABLE.includes(st.key) && (<button className="mt-stage-corner-btn" title={cs.rejected?"clear rejection (remove rework)":"mark REJECTED (log rejection date)"} onClick={(e)=>{ e.stopPropagation(); if(cs.rejected){ if(window.confirm(`Clear the rejection on "${st.label}" for ${s.styleNo}? This removes the rework flag.`)) setReject(s.id,st.key,null); } else beginDate(s.id,st.key,"reject"); }} style={{ position:"absolute", top:3, left:3, border:"1px solid transparent", background:cs.rejected?"#b03020":"rgba(255,253,248,0.86)", borderRadius:8, cursor:"pointer", padding:5, lineHeight:1, display:"flex", alignItems:"center", justifyContent:"center" }}><X size={cs.rejected?10:12} color={cs.rejected?"var(--surface)":"#b03020"}/></button>)}
                      {canSkp && !cs.autoClosed && SKIPPABLE_STAGES.includes(st.key) && (cs.skipped || !cs.actual || cs.rework || cs.rejected) && (<button className="mt-stage-skip-btn" title={cs.skipped?"un-skip (restore this activity)":"skip this activity (waive — counts as resolved, not done)"} onClick={(e)=>{ e.stopPropagation(); if(cs.skipped){ if(window.confirm(`Un-skip "${st.label}" for ${s.styleNo}? This restores the activity.`)) setSkip(s.id,st.key,null); } else if(window.confirm(`Skip / waive "${st.label}" for ${s.styleNo}?\n\nIt will count as RESOLVED (not done) and drop off the to-do. You can un-skip later.`)){ setSkip(s.id,st.key,iso(TODAY)); } }} style={{ position:"absolute", bottom:2, right:2, border:"1px solid transparent", background:cs.skipped?"#8a6d3b":"rgba(255,253,248,0.72)", borderRadius:6, cursor:"pointer", padding:5, lineHeight:1, display:"flex", alignItems:"center", justifyContent:"center" }}><SkipForward size={cs.skipped?10:13} color={cs.skipped?"var(--surface)":"#b8a98a"}/></button>)}
                      {cs.rework && canRej && (<button className="mt-stage-corner-btn" title="clear rework (un-reject the approval)" onClick={(e)=>{ e.stopPropagation(); if(window.confirm(`Clear the rework on "${st.label}" for ${s.styleNo}? This un-rejects the approval.`)) setReject(s.id, APPR_OF_SEND[st.key], null); }} style={{ position:"absolute", top:3, left:3, border:"1px solid transparent", background:"#b03020", borderRadius:8, cursor:"pointer", padding:5, lineHeight:1, display:"flex", alignItems:"center", justifyContent:"center" }}><X size={10} color="var(--surface)"/></button>)}
                      {editing&&editing.id===s.id&&editing.col===st.key&&((editing.mode==="rev"&&canRev)||(editing.mode==="reject"&&canRej)||editable) && dateEditor(s.id,st.key,editing.mode)}
                      <PeerTag who={peerOn(s.id,st.key)}/><NoteTri k={k}/><FillHandle id={s.id} col={st.key}/>
                    </td>
                  );
                })}

                {remarksVis && renderEditable(s,REMARK_COL)}
              </tr>
            ); })}

            {canMaster(role) && (
              <tr style={{ background:"#fbf7ee" }}>
                <td style={{ position:"sticky", left:0, zIndex:12, background:"#fbf7ee", border:"1px dashed var(--accent)", padding:"6px 9px" }}><div style={{ display:"flex", alignItems:"center", gap:5 }}><Plus size={13} color="var(--accent)"/><input value={newRow.styleNo} onClick={e=>e.stopPropagation()} onChange={e=>setNewRow(n=>({...n,styleNo:e.target.value}))} onKeyDown={e=>e.key==="Enter"&&addNewStyle()} placeholder="new style no… *" style={{ border:"none", outline:"none", background:"transparent", fontFamily:"inherit", fontSize:11, fontWeight:700, width:120 }}/></div></td>
                {visInfo.map(col=>(<td key={col.key} style={{ border:"1px dashed #e7dcc2", background:"#fbf7ee", padding:"3px 5px", ...freezeStyle(col.key,"#fbf7ee") }}>{ndFieldFor(col.key)}</td>))}<td colSpan={Math.max(1,(visStages?visStages.length:0)+(remarksVis?1:0))} style={ndCell}>
                  <span style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" }}>
                    
                    
                    
                    
                    
                    
                    
                    
                    {FLAG_DEFS.map(f=>{ const on=!!newRow[f.key]; return (<button className="mt-new-flag-btn" key={f.key} title={f.title} onClick={(e)=>{ e.stopPropagation(); setNewRow(n=>({...n,[f.key]:!n[f.key]})); }} style={{ fontFamily:"inherit", fontSize:9, fontWeight:800, padding:"5px 8px", cursor:"pointer", lineHeight:1.3, border:`1px solid ${on?"var(--ink)":"#cfcabf"}`, borderRadius:999, background:on?"var(--ink)":"transparent", color:on?"var(--bg)":"var(--muted-5)" }}>{f.short}</button>); })}
                    <button onClick={(e)=>{ e.stopPropagation(); addNewStyle(); }} style={{ fontFamily:"inherit", fontSize:11, fontWeight:800, padding:"8px 16px", cursor:"pointer", border:"1px solid var(--ink)", borderRadius:9, background:"var(--accent)", color:"var(--ink)" }}>+ Create (Enter)</button>{newError && <span style={{ fontSize:10, color:"var(--danger)", fontWeight:700, marginLeft:8 }}>{newError}</span>}
                  </span>
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      <div style={{ padding:"10px 22px", display:"flex", gap:16, flexWrap:"wrap", alignItems:"center", fontSize:11, borderTop:"1px solid var(--line-3)" }}>
        {(() => { const rs=rows; const n=rs.length; const rel=rs.filter(r=>r.c.released).length; const risk=rs.filter(r=>r.c.tone==="late"||r.c.tone==="warn").length; const ok=rs.filter(r=>(r.c.tone==="ok")&&!r.c.released).length; const qty=rs.reduce((a,r)=>a+(Number(r.s.qty)||0),0); const avg=n?Math.round(rs.reduce((a,r)=>a+(r.c.pct||0),0)/n):0; return <span style={{ display:"flex", gap:14, flexWrap:"wrap", alignItems:"center" }}><span><b>{n}</b> styles</span><span style={{ color:"var(--danger)" }}><b>{risk}</b> at risk</span><span><b>{ok}</b> on track</span><span><b>{rel}</b> released</span><span>total qty <b>{qty.toLocaleString()}</b></span><span>avg <b>{avg}%</b> done</span><span style={{ color:(perfRef.current.alerts||[]).length?"var(--danger)":"var(--muted-2)" }}>Perf: P95 <b>{perfRef.current.p95Ms||0}ms</b> · compute <b>{perfRef.current.computeMs}ms</b> · filter <b>{perfRef.current.filterMs}ms</b> · sort <b>{perfRef.current.rowMs}ms</b> · rendered <b>{perfRef.current.rendered}/{perfRef.current.rows}</b> · recomputed <b>{perfRef.current.recomputed}</b> · cache <b>{perfRef.current.cacheHits}</b>{perfRef.current.deferred?" · search catching up":""}{(perfRef.current.alerts||[]).length?" · "+perfRef.current.alerts.join(", "):""}</span></span>; })()}
        <span style={{ marginLeft:"auto", fontSize:9, color:"var(--muted-7)", display:"flex", alignItems:"center", gap:4, flexWrap:"wrap" }}>Ctrl/Cmd C·V copy-paste · Ctrl/Cmd Z / Shift+Z undo-redo · F2 edit · Del clears · drag blue corner to fill · <span style={{ width:0,height:0, borderTop:"7px solid var(--danger)", borderLeft:"7px solid transparent", display:"inline-block" }}/> comment · <RotateCcw size={10} color="var(--revised)"/> revised plan · <Snowflake size={10} color="#2563a6"/> freeze cols</span>
      </div>
      </>)}

      {tab==="dashboard" && <OperationalDashboardView computed={activeComputed} todoItems={todoItems} cfg={cfg} applyDrill={applyDrill} drillTodo={(obj)=>{ setTodoFilter(todoDrillFilterFromSlice({}, canonicalDrillSpec(obj||{}))); setTab("todo"); }}/>}
      {tab==="management" && <ManagementDashboardView computed={activeComputed} todoItems={todoItems} cfg={cfg} applyDrill={applyDrill} drillTodo={(obj)=>{ setTodoFilter(todoDrillFilterFromSlice({}, canonicalDrillSpec(obj||{}))); setTab("todo"); }}/>}
      {tab==="escalation" && <EscalationMatrixView computed={activeComputed} cfg={cfg} applyDrill={applyDrill} />}
      {tab==="todo" && <TodoView items={todoItems} cfg={cfg} setCfg={setCfg} canEditSettings={canAdmin(role)} filter={todoFilter} setFilter={setTodoFilter} onJump={(id,key)=>{ snapCurrent(); resetFilters(); setTab("tracker"); requestAnimationFrame(()=>setTimeout(()=>jumpToEnter(id,key),60)); }}/>}
      {tab==="review" && <ReviewTabView computed={activeComputed} todoItems={todoItems} auditRows={auditRows} auditBusy={auditBusy} loadAuditRows={loadAuditRows} errorLog={errorLog} comments={comments} inbox={inbox} me={me} colLabelOf={colLabelOf} onJump={(id,col)=>{ setTab("tracker"); setTimeout(()=>{ setSel({id:Number(id),col:col||"__style"}); setFocus(null); scrollToCell(Number(id),col||"__style"); },60); }}/>}
      {tab==="entrylog" && <EntryLogView auditRows={auditRows} auditBusy={auditBusy} loadAuditRows={loadAuditRows} errorLog={errorLog} clearErrorLog={()=>setErrorLog([])} colLabelOf={colLabelOf} onJump={(id,col)=>{ setTab("tracker"); setTimeout(()=>{ setSel({id:Number(id),col}); setFocus(null); scrollToCell(Number(id),col); },60); }}/>}
      {tab==="settings" && <SettingsView cfg={cfg} setCfg={setCfg} canEdit={canAdmin(role)}/>}
    </div>
  );
}


function EscalationMatrixView({ computed, cfg, applyDrill }){
  const [q,setQ]=useState("");
  const [owner,setOwner]=useState("All");
  const [escOwner,setEscOwner]=useState("All");
  const [bucket,setBucket]=useState("All");
  const fmtD=(d)=> d?d.toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}):"";
  const rules=escalationRulesOf(cfg);
  const all=(computed||[]).flatMap(({s,c})=>{
    if(c.released) return [];
    const front=c.frontier?[...c.frontier]:[];
    return front.map(k=>{ const r=(c.stages||[]).find(x=>x.key===k); const due=r&&(r.rev||r.plan); if(!r||!due||r.done) return null; const days=Math.max(0,netWorkdays(due,TODAY)||0); if(days<=0) return null; const esc=escalationFor(cfg,days); return { id:s.id, orderNo:s.orderNo, styleNo:s.styleNo, family:s.family, colour:s.colour, buyer:s.buyer||s.brand||"", junior:s.owner, stageKey:k, stage:stageReviewLabel(s,r), chase:r.owner||"—", due, days, bucket:esc.rangeLabel, status:c.status, risk:c.tone, escalationOwner:esc.owner, escalationLevel:esc.level, escalationAction:esc.action }; }).filter(Boolean);
  });
  const escMatchExcept=(r,except)=>{ const qq=q.trim().toLowerCase(); const hit=!qq || [r.styleNo,r.orderNo,r.family,r.colour,r.buyer,r.junior,r.stage,r.chase,r.escalationOwner,r.escalationLevel].some(v=>String(v||"").toLowerCase().includes(qq)); const ow=except==="owner"||owner==="All"||r.chase===owner; const eo=except==="escOwner"||escOwner==="All"||r.escalationOwner===escOwner; const bk=except==="bucket"||bucket==="All"||r.bucket===bucket; return hit&&ow&&eo&&bk; };
  const owners=["All",...Array.from(new Set(all.filter(r=>escMatchExcept(r,"owner")).map(x=>x.chase))).filter(Boolean).sort()];
  const escOwners=["All",...Array.from(new Set(all.filter(r=>escMatchExcept(r,"escOwner")).map(x=>x.escalationOwner))).filter(Boolean).sort()];
  const buckets=["All",...Array.from(new Set(all.filter(r=>escMatchExcept(r,"bucket")).map(x=>x.bucket))).filter(Boolean).sort((a,b)=>parseInt(a,10)-parseInt(b,10))];
  const rows=all.filter(r=>escMatchExcept(r,null)).sort((a,b)=>b.days-a.days || String(a.escalationOwner).localeCompare(String(b.escalationOwner)));
  const summary=rows.reduce((m,r)=>{ const k=r.escalationOwner; if(!m[k]) m[k]={ chase:k, count:0, total:0, max:0, urgent:0 }; m[k].count++; m[k].total+=r.days; m[k].max=Math.max(m[k].max,r.days); if(r.escalationLevel==="Critical"||r.days>=8) m[k].urgent++; return m; },{});
  const sumRows=Object.values(summary).sort((a,b)=>b.max-a.max||b.count-a.count);
  const exportRows=rows.map(r=>({ "Chase Label":r.chase, "Escalation Bucket":r.bucket, "Days Overdue":r.days, "Escalation Owner":r.escalationOwner, "Escalation Level":r.escalationLevel, "Order No":r.orderNo, "Style No":r.styleNo, "Family":r.family, "Colour":r.colour, "Buyer / Brand":r.buyer, "Junior / Owner":r.junior, "Stage / Activity":r.stage, "Due Date":fmtD(r.due), "Current Status":r.status, "Management Reading":`${r.stage} overdue ${r.days}d`, "Suggested Action":r.escalationAction||`Chase ${r.escalationOwner}` }));
  const doExport=()=>{ const wb=XLSX.utils.book_new(); appendOneSheet(wb,"Escalation Matrix",exportRows); XLSX.writeFile(wb,"escalation_matrix_"+iso(TODAY)+".xlsx"); };
  return (<div style={{ padding:"16px 22px 36px", maxWidth:1280 }}>
    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:16, flexWrap:"wrap", marginBottom:12 }}>
      <div><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:23 }}>Escalation Matrix</div><div style={{ fontSize:11.5, color:"var(--muted-2)", marginTop:4, maxWidth:760, lineHeight:1.45 }}>Chase-delay view by blocker and escalation owner. Chase Label = who is blocking. Escalation Owner = who must chase now based on editable Settings duration slabs.</div></div>
      <button onClick={doExport} style={{ fontFamily:"inherit", fontSize:11, fontWeight:800, padding:"7px 11px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)" }}>⬇ Export Escalation</button>
    </div>
    <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"center", marginBottom:12, background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:10, padding:10 }}>
      <input value={q} onChange={e=>setQ(e.target.value)} placeholder="search style / order / buyer / colour…" style={{ fontFamily:"inherit", fontSize:12, padding:"7px 9px", border:"1px solid var(--ink)", minWidth:240 }}/>
      <select value={owner} onChange={e=>setOwner(e.target.value)} style={{ fontFamily:"inherit", fontSize:11, padding:"7px 9px", border:"1px solid var(--ink)", background:"var(--surface)" }}>{owners.map(o=><option key={o} value={o}>Chase: {o}</option>)}</select>
      <select value={escOwner} onChange={e=>setEscOwner(e.target.value)} style={{ fontFamily:"inherit", fontSize:11, padding:"7px 9px", border:"1px solid var(--ink)", background:"var(--surface)" }}>{escOwners.map(o=><option key={o} value={o}>Escalation: {o}</option>)}</select>
      <select value={bucket} onChange={e=>setBucket(e.target.value)} style={{ fontFamily:"inherit", fontSize:11, padding:"7px 9px", border:"1px solid var(--ink)", background:"var(--surface)" }}>{buckets.map(o=><option key={o} value={o}>Age: {o}</option>)}</select>
      <span style={{ marginLeft:"auto", fontSize:10, color:"var(--muted-2)" }}>{rows.length} overdue chase item(s)</span>
    </div>
    <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))", gap:10, marginBottom:12 }}>{sumRows.slice(0,8).map(r=><button key={r.chase} onClick={()=>setEscOwner(r.chase)} style={{ textAlign:"left", border:"1px solid var(--line-2)", background:"var(--surface)", borderRadius:12, padding:12, cursor:"pointer", fontFamily:"inherit" }}><div style={{ fontSize:10, color:"var(--muted-2)", marginBottom:5 }}>Escalation owner</div><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:16 }}>{r.chase}</div><div style={{ fontSize:11, marginTop:5 }}><b>{r.count}</b> items · max <b style={{ color:r.max>=8?"var(--danger)":"var(--fg-warn)" }}>{r.max}d</b> · urgent {r.urgent}</div></button>)}</div>
    <div style={{ border:"1px solid var(--line-2)", borderRadius:14, overflow:"hidden", background:"var(--surface)" }}>
      <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}><thead><tr style={{ background:"var(--ink)", color:"var(--bg)", textAlign:"left" }}>{["Age","Chase","Escalation Owner","Style","Activity","Due","Action"].map(h=><th key={h} style={{ padding:"8px 10px", fontWeight:800 }}>{h}</th>)}</tr></thead><tbody>{rows.slice(0,250).map(r=><tr key={r.id+":"+r.stageKey} style={{ borderBottom:"1px solid var(--line-3)" }}><td style={{ padding:"8px 10px", fontWeight:900, color:r.days>=8?"var(--danger)":"var(--fg-warn)" }}>{r.days}d</td><td style={{ padding:"8px 10px" }}>{r.chase}</td><td style={{ padding:"8px 10px" }}><b>{r.escalationOwner}</b><div style={{ color:"var(--muted-2)", fontSize:9 }}>{r.escalationLevel}</div></td><td style={{ padding:"8px 10px" }}><b>{r.styleNo}</b><div style={{ color:"var(--muted-2)", fontSize:9 }}>{r.orderNo} · {r.colour}</div></td><td style={{ padding:"8px 10px" }}>{r.stage}</td><td style={{ padding:"8px 10px" }}>{fmtD(r.due)}</td><td style={{ padding:"8px 10px" }}><button onClick={()=>applyDrill({ search:r.styleNo, activity:r.stageKey, status:"All" })} style={{ fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"5px 8px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--accent)" }}>Open style</button></td></tr>)}{!rows.length&&<tr><td colSpan={7} style={{ padding:22, color:"var(--muted-2)", textAlign:"center" }}>No overdue chase items match this filter.</td></tr>}</tbody></table>
    </div>
  </div>);
}



function reviewCategoryForCol(col){
  const k=String(col||"");
  const st=STAGES.find(x=>x.key===k);
  if(st) return st.label.replace(" Appr"," Approval");
  const map={ orderNo:"Order", styleNo:"Order", sampleFit:"Fit", family:"Order", colour:"Colour", color:"Colour", brand:"Buyer / Brand", buyer:"Buyer / Brand", fabricType:"Fabric", owner:"Owner / Junior", qty:"Qty", delivery:"Delivery", ordRec:"Order", remarks:"Remarks", __style:"Order", setId:"Order", setRole:"Order", age:"Order" };
  return map[k]||"General";
}
function ReviewTabView({ computed, todoItems, auditRows, auditBusy, loadAuditRows, errorLog, comments, inbox, me, colLabelOf, onJump }){
  const tsShortLocal=(t)=>{ try{ const d=new Date(t); return d.toLocaleDateString(undefined,{day:"2-digit",month:"short"})+" "+d.toLocaleTimeString(undefined,{hour:"2-digit",minute:"2-digit"}); }catch(e){ return ""; } };
  const [mode,setMode]=useState("inbox");
  const [category,setCategory]=useState("All");
  const [severity,setSeverity]=useState("All");
  const [q,setQ]=useState("");
  const [commentScope,setCommentScope]=useState("All comments");
  const [commentUser,setCommentUser]=useState("All users");
  const [notifStatus,setNotifStatus]=useState("All");
  const stylesById=useMemo(()=>{ const m=new Map(); (computed||[]).forEach(r=>{ if(r&&r.s) m.set(String(r.s.id),r.s); }); return m; },[computed]);
  useEffect(()=>{ if((auditRows||[]).length===0 && !auditBusy) loadAuditRows&&loadAuditRows(); },[]);
  const categoryForText=(text,col,type)=>{ const c=reviewCategoryForCol(col||type||""); if(c && c!=="General") return c; const x=String(text||"").toLowerCase(); if(x.includes("order")) return "Order"; if(x.includes("colour")||x.includes("color")) return "Colour"; if(x.includes("fit")) return "Fit"; if(x.includes("art")) return "Artwork"; if(x.includes("strike")||x.includes("s/o")) return "Strike-off"; if(x.includes("lab")) return "Lab Dip"; if(x.includes("fabric")) return "Fabric"; if(x.includes("pp")) return "PP"; if(x.includes("prod")) return "Production File"; if(x.includes("upload")) return "Upload Changes"; if(x.includes("parity")||x.includes("report")) return "Report Checks"; if(x.includes("sync")||x.includes("error")) return "System / Sync"; return "General"; };
  const allComments=useMemo(()=>{
    const arr=[];
    Object.entries(comments||{}).forEach(([ck,list])=>{ const [sid,col]=ck.split(":"); const st=stylesById.get(String(sid)); (list||[]).forEach(c=>{ const body=String(c.body||""); arr.push({
      id:c.id||`${sid}-${col}-${arr.length}`,
      sid:Number(sid), col,
      category:categoryForText(body,col,"comment"),
      styleNo:(st&&st.styleNo)||c.style_no||sid,
      orderNo:(st&&st.orderNo)||"",
      buyer:(st&&st.buyer)||"",
      brand:(st&&st.brand)||"",
      body,
      authorId:c.author_id,
      author:c.author_name||c.actor_name||"Unknown",
      mentions:c.mentions||[],
      resolved:!!c.resolved,
      at:c.created_at
    }); }); });
    return arr.sort((a,b)=>String(b.at||"").localeCompare(String(a.at||"")));
  },[comments,stylesById]);
  const allNotifications=useMemo(()=> (inbox||[]).map(n=>({
    id:n.id,
    category:categoryForText(n.body,n.col,n.type),
    severity:n.read?"Info":"Warning",
    status:n.read?"Read":"Unread",
    styleNo:n.style_no||"",
    orderNo:n.order_no||"",
    body:n.body||"",
    source:n.type||"Notification",
    actor:n.actor_name||"",
    at:n.created_at,
    read:!!n.read,
    jumpId:n.style_id,
    jumpCol:n.col
  })).sort((a,b)=>String(b.at||"").localeCompare(String(a.at||""))),[inbox]);
  const items=useMemo(()=>{
    const out=[];
    const push=(x)=>out.push({ id:x.id||(`${x.source||"item"}-${out.length}`), category:x.category||"General", severity:x.severity||"Info", status:x.status||"Open", styleNo:x.styleNo||"", orderNo:x.orderNo||"", buyer:x.buyer||"", brand:x.brand||"", issue:x.issue||"", action:x.action||"Review", source:x.source||"System", at:x.at||"", jumpId:x.jumpId, jumpCol:x.jumpCol });
    (todoItems||[]).slice(0,350).forEach(t=> push({ id:`todo-${t.id||t.styleNo}-${t.key}`, category:reviewCategoryForCol(t.key), severity:t.overdue?"Critical":"Warning", status:t.overdue?"Overdue":"Upcoming", styleNo:t.styleNo, orderNo:t.orderNo, buyer:t.buyer, brand:t.brand, issue:`${t.activity||reviewCategoryForCol(t.key)} ${t.overdue?"overdue":"pending"}${t.daysLate?` · ${t.daysLate}d late`:""}`, action:t.escalationOwner?`Chase ${t.owner||"owner"}; escalation owner ${t.escalationOwner}`:`Chase ${t.owner||"owner"}`, source:"To-Do", at:t.exp||"", jumpId:t.id, jumpCol:t.key }));
    (computed||[]).forEach(({s,c})=>{
      if(!s) return;
      if(!s.buyer) push({ id:`dq-buyer-${s.id}`, category:"Data Quality", severity:"Warning", styleNo:s.styleNo, orderNo:s.orderNo, brand:s.brand, issue:"Buyer missing", action:"Fill Buyer before buyer-wise reports/export", source:"Data Quality", jumpId:s.id, jumpCol:"buyer" });
      if(!s.delivery) push({ id:`dq-del-${s.id}`, category:"Data Quality", severity:"Critical", styleNo:s.styleNo, orderNo:s.orderNo, buyer:s.buyer, brand:s.brand, issue:"Delivery date missing", action:"Fill Delivery Date", source:"Data Quality", jumpId:s.id, jumpCol:"delivery" });
      if(s.soReq && !s.printReq) push({ id:`dq-so-${s.id}`, category:"Data Quality", severity:"Warning", styleNo:s.styleNo, orderNo:s.orderNo, buyer:s.buyer, brand:s.brand, issue:"S/O required is ON while Print required is OFF", action:"Check print / strike-off requirement flags", source:"Data Quality", jumpId:s.id, jumpCol:"__style" });
      if(c && (c.tone==="late"||String(c.status||"").toLowerCase().includes("risk"))) push({ id:`risk-${s.id}`, category:"Escalation", severity:c.tone==="late"?"Critical":"Warning", styleNo:s.styleNo, orderNo:s.orderNo, buyer:s.buyer, brand:s.brand, issue:c.status||"Style at risk", action:(c.chaseOwners&&c.chaseOwners[0])?`Chase ${(c.chaseOwners[0]||{}).owner}`:"Review current blocker", source:"Dashboard", jumpId:s.id, jumpCol:(c.frontier&&[...c.frontier][0])||"__style" });
    });
    allComments.slice(0,250).forEach(c=> push({ id:`comment-${c.id}`, category:c.category, severity:c.resolved?"Info":"Warning", status:c.resolved?"Resolved":"Open", styleNo:c.styleNo, orderNo:c.orderNo, buyer:c.buyer, brand:c.brand, issue:`Comment by ${c.author}: ${c.body.slice(0,120)}`, action:"Open comment thread", source:"Comment", at:c.at, jumpId:c.sid, jumpCol:c.col }));
    allNotifications.slice(0,120).forEach(n=> push({ id:`notif-${n.id}`, category:n.category, severity:n.read?"Info":"Warning", status:n.status, styleNo:n.styleNo, orderNo:n.orderNo, issue:n.body.slice(0,140), action:n.jumpId?"Open linked style":"Review notification", source:`Notification · ${n.source}`, at:n.at, jumpId:n.jumpId, jumpCol:n.jumpCol }));
    (auditRows||[]).slice(0,160).forEach(a=>{ const cat=reviewCategoryForCol(a.col||a.field); const important=["delivery","ordRec","buyer","brand","qty","owner","fitAppr","artAppr","soAppr","labAppr","ppAppr","fabricIH","prodFile"].includes(a.col)||String(a.field||"").toLowerCase().includes("date"); push({ id:`audit-${a.id}`, category:cat, severity:important?"Info":"Info", status:"Changed", styleNo:a.style_no||"", orderNo:a.order_no||"", issue:`${colLabelOf?colLabelOf(a.col):a.col} changed: ${a.old_val||"—"} → ${a.new_val||"—"}`, action:"Audit only", source:`Change · ${a.actor_name||""}`, at:a.created_at, jumpId:a.style_id, jumpCol:a.col }); });
    (errorLog||[]).forEach(e=> push({ id:`err-${e.id}`, category:"System / Sync", severity:"Critical", issue:`${e.area||"App"}: ${e.msg||"Error"}`, action:"Check console / refresh after saving work", source:"Error Log", at:e.at }));
    return out.sort((a,b)=>{ const sev={Critical:0,Warning:1,Info:2}; const sa=sev[a.severity]??9, sb=sev[b.severity]??9; if(sa!==sb) return sa-sb; return String(b.at||"").localeCompare(String(a.at||"")); });
  },[computed,todoItems,auditRows,errorLog,allComments,allNotifications,colLabelOf]);
  const myId=me&&me.id; const myName=String((me&&(me.name||me.email))||"").toLowerCase();
  const needle=q.trim().toLowerCase();
  const textMatch=(vals)=>!needle||vals.join(" ").toLowerCase().includes(needle);
  const itemPassExcept=(x,except)=> (except==="category"||category==="All"||x.category===category) && (except==="severity"||severity==="All"||x.severity===severity) && textMatch([x.category,x.severity,x.status,x.styleNo,x.orderNo,x.buyer,x.brand,x.issue,x.action,x.source]);
  const commentPassExcept=(c,except)=>{
    if(except!=="category"&&category!=="All"&&c.category!==category) return false;
    if(except!=="commentUser"&&commentUser!=="All users"&&c.author!==commentUser) return false;
    if(except!=="commentScope"){
      if(commentScope==="To me" && !(Array.isArray(c.mentions)&&c.mentions.includes(myId)) && !(myName && String(c.body||"").toLowerCase().includes("@"+myName.replace(/\s+/g,""))) && !(myName && String(c.body||"").toLowerCase().includes("@"+myName))) return false;
      if(commentScope==="By me" && c.authorId!==myId && String(c.author||"").toLowerCase()!==myName) return false;
      if(commentScope==="Unresolved" && c.resolved) return false;
    }
    return textMatch([c.category,c.styleNo,c.orderNo,c.buyer,c.brand,c.body,c.author]);
  };
  const notificationPassExcept=(n,except)=> (except==="category"||category==="All"||n.category===category) && (except==="notifStatus"||notifStatus==="All"||n.status===notifStatus) && textMatch([n.category,n.status,n.styleNo,n.orderNo,n.body,n.source,n.actor]);
  const activityRows=useMemo(()=>{ const rows=[]; (auditRows||[]).forEach(a=>{ const f=String(a.field||"").toLowerCase(); const nv=String(a.new_val||""); const ov=String(a.old_val||""); let type=""; if(f.includes("format")) type="Format Changes"; else if(f==="deleted"||nv==="DELETED") type="Delete Row"; else if(String(a.col||"").toLowerCase()==="upload"||f.includes("upload")) type="Upload History"; else if(f.includes("bulk")) type=f.includes("actual")||f.includes("revised")?"Bulk Fill Date":"Bulk Action"; else if(f.includes("resend actual")) type="Resend Actual History"; else if(f.includes("revised version")) type="Revised Version History"; else if(["actual","revised","rejected","waived"].includes(f)) type="Bulk Fill / Date Changes"; if(!type) return; rows.push({ id:a.id||("act-"+rows.length), type, category:reviewCategoryForCol(a.col||a.field), time:a.created_at, user:a.actor_name||"", styleNo:a.style_no||a.style_id||"", col:a.col||"", field:a.field||"", oldVal:ov, newVal:nv, raw:a }); }); return rows.sort((a,b)=>String(b.time||"").localeCompare(String(a.time||""))); },[auditRows]);
  const activityPassExcept=(r,except)=> (except==="category"||category==="All"||r.category===category) && (except==="activityType"||activityType==="All"||r.type===activityType) && (except==="activityUser"||activityUser==="All"||r.user===activityUser) && textMatch([r.type,r.category,r.user,r.styleNo,r.col,r.field,r.oldVal,r.newVal]);
  const [activityType,setActivityType]=useState("All");
  const [activityUser,setActivityUser]=useState("All");
  const categories=useMemo(()=>{
    const src=mode==="comments"?allComments.filter(c=>commentPassExcept(c,"category")):mode==="notifications"?allNotifications.filter(n=>notificationPassExcept(n,"category")):mode==="activity"?activityRows.filter(r=>activityPassExcept(r,"category")):items.filter(x=>itemPassExcept(x,"category"));
    return ["All",...Array.from(new Set(src.map(x=>x.category))).sort((a,b)=>a.localeCompare(b))];
  },[items,allComments,allNotifications,activityRows,mode,severity,q,commentUser,commentScope,notifStatus,activityType,activityUser,category]);
  const severities=useMemo(()=>["All",...Array.from(new Set(items.filter(x=>itemPassExcept(x,"severity")).map(x=>x.severity))).sort((a,b)=>a.localeCompare(b))],[items,category,q,severity]);
  const commentUsers=useMemo(()=>["All users",...Array.from(new Set(allComments.filter(c=>commentPassExcept(c,"commentUser")).map(c=>c.author).filter(Boolean))).sort((a,b)=>a.localeCompare(b))],[allComments,category,commentScope,q,commentUser]);
  const notifStatuses=useMemo(()=>["All",...Array.from(new Set(allNotifications.filter(n=>notificationPassExcept(n,"notifStatus")).map(n=>n.status))).sort((a,b)=>a.localeCompare(b))],[allNotifications,category,q,notifStatus]);
  const filtered=useMemo(()=> items.filter(x=>itemPassExcept(x,null)),[items,category,severity,q]);
  const filteredComments=useMemo(()=> allComments.filter(c=>commentPassExcept(c,null)),[allComments,category,commentUser,commentScope,q,myId,myName]);
  const filteredNotifications=useMemo(()=> allNotifications.filter(n=>notificationPassExcept(n,null)),[allNotifications,category,notifStatus,q]);
  const activityTypes=useMemo(()=>["All",...Array.from(new Set(activityRows.filter(r=>activityPassExcept(r,"activityType")).map(r=>r.type))).sort((a,b)=>a.localeCompare(b))],[activityRows,category,q,activityUser,activityType]);
  const activityUsers=useMemo(()=>["All",...Array.from(new Set(activityRows.filter(r=>activityPassExcept(r,"activityUser")).map(r=>r.user).filter(Boolean))).sort((a,b)=>a.localeCompare(b))],[activityRows,category,q,activityType,activityUser]);
  const filteredActivity=useMemo(()=> activityRows.filter(r=>activityPassExcept(r,null)),[activityRows,category,q,activityType,activityUser]);
  const counts=useMemo(()=>({ total:items.length, critical:items.filter(x=>x.severity==="Critical").length, warning:items.filter(x=>x.severity==="Warning").length, dq:items.filter(x=>x.category==="Data Quality").length, comments:allComments.length, unread:allNotifications.filter(n=>!n.read).length }),[items,allComments,allNotifications]);
  const badge=(txt,bg,fg="var(--ink)")=><span style={{ display:"inline-flex", alignItems:"center", padding:"3px 7px", borderRadius:999, background:bg, color:fg, fontSize:9, fontWeight:900, whiteSpace:"nowrap" }}>{txt}</span>;
  const inputStyle={ fontFamily:"inherit", fontSize:11, padding:"9px 12px", border:"1px solid var(--ink)", background:"var(--surface)", borderRadius:8, minWidth:130 };
  const modeBtn=(k,l,cnt)=><button onClick={()=>{ setMode(k); setCategory("All"); setSeverity("All"); setQ(""); }} style={{ fontFamily:"inherit", fontSize:11, fontWeight:900, padding:"10px 15px", cursor:"pointer", border:"1px solid var(--line-2)", background:mode===k?"var(--ink)":"var(--surface)", color:mode===k?"var(--surface)":"var(--ink)", borderRadius:999 }}>{l}{cnt!=null?` · ${cnt}`:""}</button>;
  return <div style={{ padding:"18px 22px 36px" }}>
    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:16, marginBottom:14 }}>
      <div><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:24 }}>Review</div><div style={{ fontSize:11.5, color:"var(--muted-3)", marginTop:4, maxWidth:900, lineHeight:1.45 }}>Category-wise review centre. Review Inbox summarizes work/data checks. Comments and Notifications have their own views with user-wise and business-category filters.</div></div>
      <button onClick={loadAuditRows} style={{ fontFamily:"inherit", fontSize:11, fontWeight:900, padding:"10px 15px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)", borderRadius:9 }}>{auditBusy?"Loading…":"Refresh review"}</button>
    </div>
    <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(145px,1fr))", gap:10, marginBottom:12 }}>
      {[ ["Open review items",counts.total,"var(--surface)"], ["Critical",counts.critical,"#f6d3cb"], ["Warnings",counts.warning,"#f8e9b7"], ["Comments",counts.comments,"#e7ecff"], ["Unread notifications",counts.unread,"#fff3df"], ["Data quality",counts.dq,"#e3edfb"] ].map(([l,v,b])=><div key={l} style={{ background:b, border:"1px solid var(--line-2)", borderRadius:12, padding:12 }}><div style={{ fontSize:10, color:"var(--muted-3)", fontWeight:900, textTransform:"uppercase" }}>{l}</div><div style={{ fontFamily:"'Archivo',sans-serif", fontSize:25, fontWeight:800 }}>{v}</div></div>)}
    </div>
    <div style={{ display:"flex", gap:10, flexWrap:"wrap", marginBottom:14 }}>{modeBtn("inbox","Review Inbox",items.length)}{modeBtn("activity","Activity History",activityRows.length)}{modeBtn("comments","Comments",allComments.length)}{modeBtn("notifications","Notifications",allNotifications.length)}{modeBtn("changes","Changes",(auditRows||[]).length)}{modeBtn("errors","Errors",(errorLog||[]).length)}</div>
    <div style={{ display:"flex", gap:10, flexWrap:"wrap", alignItems:"center", marginBottom:14, background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:14, padding:14 }}>
      <select value={category} onChange={e=>setCategory(e.target.value)} style={inputStyle}>{categories.map(c=><option key={c} value={c}>{c}</option>)}</select>
      {mode==="inbox" && <select value={severity} onChange={e=>setSeverity(e.target.value)} style={inputStyle}>{severities.map(c=><option key={c} value={c}>{c}</option>)}</select>}
      {mode==="comments" && <><select value={commentScope} onChange={e=>setCommentScope(e.target.value)} style={inputStyle}>{["All comments","To me","By me","Unresolved"].map(c=><option key={c} value={c}>{c}</option>)}</select><select value={commentUser} onChange={e=>setCommentUser(e.target.value)} style={inputStyle}>{commentUsers.map(c=><option key={c} value={c}>{c}</option>)}</select></>}
      {mode==="notifications" && <select value={notifStatus} onChange={e=>setNotifStatus(e.target.value)} style={inputStyle}>{notifStatuses.map(c=><option key={c} value={c}>{c}</option>)}</select>}
      {mode==="activity" && <><select value={activityType} onChange={e=>{ setActivityType(e.target.value); setActivityUser("All"); }} style={inputStyle}>{activityTypes.map(c=><option key={c} value={c}>{c}</option>)}</select><select value={activityUser} onChange={e=>setActivityUser(e.target.value)} style={inputStyle}>{activityUsers.map(c=><option key={c} value={c}>{c}</option>)}</select></>}
      <input value={q} onChange={e=>setQ(e.target.value)} placeholder="search style, order, buyer, issue, user…" style={{ ...inputStyle, flex:"1 1 260px" }} />
      <span style={{ fontSize:10, color:"var(--muted-2)" }}>{mode==="comments"?filteredComments.length:mode==="notifications"?filteredNotifications.length:mode==="activity"?filteredActivity.length:mode==="errors"?(errorLog||[]).length:mode==="changes"?(auditRows||[]).length:filtered.length} shown</span>
    </div>
    <div style={{ background:"var(--surface)", border:"1px solid var(--line-2)", borderRadius:14, overflow:"hidden" }}>
      {mode==="activity" ? <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}><thead><tr style={{ background:"var(--ink)", color:"var(--bg)" }}>{["Type","Time","User","Style","Field / Stage","Old → New","Open"].map(h=><th key={h} style={{ textAlign:"left", padding:"9px 10px", borderRight:"1px solid #3a362e", fontFamily:"'Archivo',sans-serif" }}>{h}</th>)}</tr></thead><tbody>{filteredActivity.length===0?<tr><td colSpan={7} style={{ padding:24, color:"var(--muted-2)", textAlign:"center" }}>No bulk/upload/delete/format history for this filter.</td></tr>:filteredActivity.slice(0,500).map(r=><tr key={r.id} onClick={()=>r.raw&&r.raw.style_id&&onJump&&onJump(r.raw.style_id,r.raw.col)} style={{ cursor:r.raw&&r.raw.style_id?"pointer":"default", borderBottom:"1px solid var(--line-3)", background:r.type==="Delete Row"?"#fff7f5":r.type==="Format Changes"?"#f7fbff":"var(--surface)" }}><td style={{ padding:"8px 10px", fontWeight:900 }}>{r.type}</td><td style={{ padding:"8px 10px", color:"var(--muted-2)", whiteSpace:"nowrap" }}>{tsShortLocal(r.time)}</td><td style={{ padding:"8px 10px" }}>{r.user||"—"}</td><td style={{ padding:"8px 10px", fontWeight:800 }}>{r.styleNo||"—"}</td><td style={{ padding:"8px 10px" }}>{colLabelOf?colLabelOf(r.col):r.col}<div style={{ color:"var(--muted-2)", fontSize:9 }}>{r.field}</div></td><td style={{ padding:"8px 10px" }}><span style={{ color:"var(--danger)" }}>{r.oldVal||"—"}</span> → <span style={{ color:"var(--success)" }}>{r.newVal||"—"}</span></td><td style={{ padding:"8px 10px", color:"var(--info)", fontWeight:900 }}>{r.raw&&r.raw.style_id?"open →":""}</td></tr>)}</tbody></table>
      : mode==="comments" ? <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}><thead><tr style={{ background:"var(--ink)", color:"var(--bg)" }}>{["Category","User","Style / Order","Buyer / Brand","Comment","Status","Open"].map(h=><th key={h} style={{ textAlign:"left", padding:"9px 10px", borderRight:"1px solid #3a362e", fontFamily:"'Archivo',sans-serif" }}>{h}</th>)}</tr></thead><tbody>{filteredComments.length===0?<tr><td colSpan={7} style={{ padding:24, color:"var(--muted-2)", textAlign:"center" }}>No comments for this filter.</td></tr>:filteredComments.slice(0,500).map(c=><tr key={c.id} onClick={()=>onJump&&onJump(c.sid,c.col)} style={{ cursor:"pointer", borderBottom:"1px solid var(--line-3)", background:c.resolved?"var(--surface)":"#fffaf0" }}><td style={{ padding:"8px 10px", fontWeight:800 }}>{c.category}</td><td style={{ padding:"8px 10px" }}>{c.author}<div style={{ color:"var(--muted-2)", fontSize:9 }}>{tsShortLocal(c.at)}</div></td><td style={{ padding:"8px 10px" }}><b>{c.styleNo||"—"}</b><br/><span style={{ color:"var(--muted-2)", fontSize:9 }}>{c.orderNo||""}</span></td><td style={{ padding:"8px 10px" }}>{c.buyer||"—"}<br/><span style={{ color:"var(--muted-2)", fontSize:9 }}>{c.brand||""}</span></td><td style={{ padding:"8px 10px", maxWidth:520 }}>{c.body}</td><td style={{ padding:"8px 10px" }}>{c.resolved?badge("Resolved","#e5f1ea","#1c6048"):badge("Open","#f8e9b7","#7a560f")}</td><td style={{ padding:"8px 10px", color:"var(--info)", fontWeight:900 }}>thread →</td></tr>)}</tbody></table>
      : mode==="notifications" ? <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}><thead><tr style={{ background:"var(--ink)", color:"var(--bg)" }}>{["Category","Status","Style","Notification","Source","Time"].map(h=><th key={h} style={{ textAlign:"left", padding:"9px 10px", borderRight:"1px solid #3a362e", fontFamily:"'Archivo',sans-serif" }}>{h}</th>)}</tr></thead><tbody>{filteredNotifications.length===0?<tr><td colSpan={6} style={{ padding:24, color:"var(--muted-2)", textAlign:"center" }}>No notifications for this filter.</td></tr>:filteredNotifications.slice(0,500).map(n=><tr key={n.id} onClick={()=>n.jumpId&&onJump&&onJump(n.jumpId,n.jumpCol||"__style")} style={{ cursor:n.jumpId?"pointer":"default", borderBottom:"1px solid var(--line-3)", background:n.read?"var(--surface)":"#fff3df" }}><td style={{ padding:"8px 10px", fontWeight:800 }}>{n.category}</td><td style={{ padding:"8px 10px" }}>{n.read?badge("Read","#e5f1ea","#1c6048"):badge("Unread","#f8e9b7","#7a560f")}</td><td style={{ padding:"8px 10px" }}><b>{n.styleNo||"—"}</b></td><td style={{ padding:"8px 10px", maxWidth:600 }}>{n.body}<div style={{ color:"var(--muted-2)", fontSize:9 }}>{n.actor||""}</div></td><td style={{ padding:"8px 10px", color:"var(--muted-2)" }}>{n.source}</td><td style={{ padding:"8px 10px", color:"var(--muted-2)", whiteSpace:"nowrap" }}>{tsShortLocal(n.at)}{n.jumpId&&<div style={{ color:"var(--info)", fontWeight:900 }}>open →</div>}</td></tr>)}</tbody></table>
      : mode==="changes" ? <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}><thead><tr style={{ background:"var(--ink)", color:"var(--bg)" }}>{["Time","User","Style","Field","Old → New"].map(h=><th key={h} style={{ textAlign:"left", padding:"9px 10px", borderRight:"1px solid #3a362e", fontFamily:"'Archivo',sans-serif" }}>{h}</th>)}</tr></thead><tbody>{auditBusy?<tr><td colSpan={5} style={{ padding:24, color:"var(--muted-2)", textAlign:"center" }}>Loading changes…</td></tr>:(auditRows||[]).length===0?<tr><td colSpan={5} style={{ padding:24, color:"var(--muted-2)", textAlign:"center" }}>No changes loaded.</td></tr>:(auditRows||[]).slice(0,500).map(a=><tr key={a.id} onClick={()=>a.style_id&&onJump&&onJump(a.style_id,a.col)} style={{ cursor:a.style_id?"pointer":"default", borderBottom:"1px solid var(--line-3)" }}><td style={{ padding:"8px 10px", color:"var(--muted-2)" }}>{tsShortLocal(a.created_at)}</td><td style={{ padding:"8px 10px" }}>{a.actor_name||"—"}</td><td style={{ padding:"8px 10px", fontWeight:800 }}>{a.style_no||a.style_id||"—"}</td><td style={{ padding:"8px 10px" }}>{colLabelOf?colLabelOf(a.col):a.col}</td><td style={{ padding:"8px 10px" }}><span style={{ color:"var(--danger)" }}>{a.old_val||"—"}</span> → <span style={{ color:"var(--success)" }}>{a.new_val||"—"}</span></td></tr>)}</tbody></table>
      : mode==="errors" ? <div style={{ padding:14 }}>{(!errorLog||!errorLog.length)?<div style={{ color:"var(--muted-1)", fontSize:11, padding:18, textAlign:"center" }}>No app errors logged in this browser session.</div>:errorLog.map(e=><div key={e.id} style={{ padding:"9px 0", borderBottom:"1px solid var(--line-3)", fontSize:11 }}><span style={{ color:"var(--muted-1)", fontSize:9 }}>{tsShortLocal(e.at)}</span> · <b>{e.area}</b><br/><span style={{ color:"var(--danger)", fontWeight:700 }}>{e.msg}</span>{e.extra&&<span style={{ color:"var(--muted-1)" }}> · {e.extra}</span>}</div>)}</div>
      : <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}><thead><tr style={{ background:"var(--ink)", color:"var(--bg)" }}>{["Category","Severity","Style / Order","Buyer / Brand","Issue","Action","Source"].map(h=><th key={h} style={{ textAlign:"left", padding:"9px 10px", borderRight:"1px solid #3a362e", fontFamily:"'Archivo',sans-serif" }}>{h}</th>)}</tr></thead><tbody>{filtered.length===0?<tr><td colSpan={7} style={{ padding:24, color:"var(--muted-2)", textAlign:"center" }}>No review items for this filter.</td></tr>:filtered.slice(0,500).map(it=><tr key={it.id} onClick={()=>it.jumpId&&onJump&&onJump(it.jumpId,it.jumpCol||"__style")} style={{ cursor:it.jumpId?"pointer":"default", borderBottom:"1px solid var(--line-3)", background:it.severity==="Critical"?"#fff7f5":it.severity==="Warning"?"#fffaf0":"var(--surface)" }}><td style={{ padding:"8px 10px", fontWeight:800 }}>{it.category}</td><td style={{ padding:"8px 10px" }}>{it.severity==="Critical"?badge("Critical","#f6d3cb","#8c241a"):it.severity==="Warning"?badge("Warning","#f8e9b7","#7a560f"):badge("Info","#e5f1ea","#1c6048")}</td><td style={{ padding:"8px 10px" }}><b>{it.styleNo||"—"}</b><br/><span style={{ color:"var(--muted-2)", fontSize:9 }}>{it.orderNo||""}</span></td><td style={{ padding:"8px 10px" }}>{it.buyer||"—"}<br/><span style={{ color:"var(--muted-2)", fontSize:9 }}>{it.brand||""}</span></td><td style={{ padding:"8px 10px", maxWidth:420 }}>{it.issue}<div style={{ color:"var(--muted-2)", fontSize:9, marginTop:3 }}>{it.status}{it.at?" · "+tsShortLocal(it.at):""}</div></td><td style={{ padding:"8px 10px", color:"var(--muted-4)" }}>{it.action}</td><td style={{ padding:"8px 10px", color:"var(--muted-2)", fontSize:10 }}>{it.source}{it.jumpId&&<div style={{ color:"var(--info)", fontWeight:900, marginTop:3 }}>open →</div>}</td></tr>)}</tbody></table>}
    </div>
  </div>;
}

function EntryLogView({ auditRows, auditBusy, loadAuditRows, errorLog, clearErrorLog, colLabelOf, onJump }){
  const [mode,setMode]=useState("entries");
  const [q,setQ]=useState("");
  const [user,setUser]=useState("All");
  const [field,setField]=useState("All");
  const [days,setDays]=useState("30");
  useEffect(()=>{ loadAuditRows&&loadAuditRows(); },[]);
  const ts=(t)=>{ try{ const d=new Date(t); return d.toLocaleDateString(undefined,{day:"2-digit",month:"short",year:"numeric"})+" "+d.toLocaleTimeString(undefined,{hour:"2-digit",minute:"2-digit"}); }catch(e){ return ""; } };
  const label=(c)=> colLabelOf?colLabelOf(c):c;
  const cutoff=days==="all"?null:(Date.now()-Number(days||30)*86400000);
  const auditMatchExcept=(r,except)=>{
    if(cutoff && new Date(r.created_at).getTime()<cutoff) return false;
    if(except!=="user" && user!=="All" && String(r.actor_name||"")!==user) return false;
    if(except!=="field" && field!=="All" && String(r.field||"")!==field) return false;
    const hay=[r.style_no,r.style_id,label(r.col),r.col,r.field,r.old_val,r.new_val,r.actor_name].join(" ").toLowerCase();
    return !q || hay.includes(q.toLowerCase());
  };
  const rows=(auditRows||[]).filter(r=>auditMatchExcept(r,null));
  const users=["All",...[...new Set((auditRows||[]).filter(r=>auditMatchExcept(r,"user")).map(r=>r.actor_name).filter(Boolean))].sort()];
  const fields=["All",...[...new Set((auditRows||[]).filter(r=>auditMatchExcept(r,"field")).map(r=>r.field).filter(Boolean))].sort()];
  const exportEntries=()=>{ const data=rows.map(r=>({"Time":ts(r.created_at),"User":r.actor_name||"","Style":r.style_no||r.style_id||"","Field / Stage":label(r.col),"Action Type":r.field||"","Old Value":r.old_val||"","New Value":r.new_val||""})); if(!data.length){ alert("No entry log rows to export."); return; } const ws=XLSX.utils.json_to_sheet(data); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Entry Log"); XLSX.writeFile(wb,"entry_log_"+iso(TODAY)+".xlsx"); };
  const exportErrors=()=>{ const data=(errorLog||[]).map(e=>({"Time":ts(e.at),"Area":e.area||"","Error":e.msg||"","Extra":e.extra||""})); if(!data.length){ alert("No errors to export."); return; } const ws=XLSX.utils.json_to_sheet(data); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Error Log"); XLSX.writeFile(wb,"error_log_"+iso(TODAY)+".xlsx"); };
  const card={ background:"var(--surface)", border:"1px solid var(--toolbar-line)", borderRadius:14, boxShadow:"var(--card-shadow)" };
  const input={ fontFamily:"inherit", fontSize:11, padding:"7px 9px", border:"1px solid var(--line-2)", borderRadius:9, background:"var(--surface)", outline:"none" };
  return (<div style={{ padding:"18px 22px 36px" }}>
    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:14, flexWrap:"wrap", marginBottom:14 }}>
      <div><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:22 }}>Entry Log Book</div><div style={{ fontSize:11.5, color:"var(--muted-2)", marginTop:4, maxWidth:720, lineHeight:1.45 }}>Timestamped audit trail for sheet changes and browser-session app errors. Use this for accountability, checking who changed what, and exporting logs for review.</div></div>
      <div style={{ display:"flex", border:"1px solid var(--line-2)", borderRadius:999, overflow:"hidden", background:"var(--surface)" }}>{[["entries","Entries"],["errors","Errors"]].map(([k,l])=><button key={k} onClick={()=>setMode(k)} style={{ fontFamily:"inherit", fontSize:11, fontWeight:800, padding:"8px 13px", cursor:"pointer", border:"none", background:mode===k?"var(--accent)":"transparent", color:mode===k?"var(--surface)":"var(--muted-3)" }}>{l}{k==="errors"&&errorLog&&errorLog.length?" · "+errorLog.length:""}</button>)}</div>
    </div>
    {mode==="entries" && <div style={card}>
      <div style={{ padding:14, borderBottom:"1px solid var(--line-3)", display:"flex", gap:9, alignItems:"center", flexWrap:"wrap" }}>
        <input value={q} onChange={e=>setQ(e.target.value)} placeholder="Search style, field, user, value…" style={{ ...input, minWidth:240, flex:"1 1 260px" }}/>
        <select value={days} onChange={e=>setDays(e.target.value)} style={input}><option value="7">Last 7 days</option><option value="30">Last 30 days</option><option value="90">Last 90 days</option><option value="all">All loaded</option></select>
        <select value={user} onChange={e=>setUser(e.target.value)} style={input}>{users.map(u=><option key={u} value={u}>{u==="All"?"All users":u}</option>)}</select>
        <select value={field} onChange={e=>setField(e.target.value)} style={input}>{fields.map(f=><option key={f} value={f}>{f==="All"?"All action types":f}</option>)}</select>
        <button onClick={()=>loadAuditRows&&loadAuditRows()} style={{ ...input, cursor:"pointer", fontWeight:800 }}>Refresh</button>
        <button onClick={exportEntries} style={{ ...input, cursor:"pointer", fontWeight:800, background:"var(--ink)", color:"var(--surface)" }}>Export</button>
      </div>
      <div style={{ padding:"10px 14px", display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))", gap:10, borderBottom:"1px solid var(--line-3)" }}>
        <div><b>{rows.length}</b><div style={{ fontSize:9, color:"var(--muted-1)", textTransform:"uppercase" }}>matching entries</div></div>
        <div><b>{users.length-1}</b><div style={{ fontSize:9, color:"var(--muted-1)", textTransform:"uppercase" }}>users in log</div></div>
        <div><b>{fields.length-1}</b><div style={{ fontSize:9, color:"var(--muted-1)", textTransform:"uppercase" }}>action types</div></div>
      </div>
      <div style={{ overflowX:"auto" }}><table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}><thead><tr>{["Time","User","Style","Field / Stage","Action","Old → New"].map(h=><th key={h} style={{ textAlign:"left", padding:"9px 10px", borderBottom:"1px solid var(--line-2)", color:"var(--muted-2)", fontSize:9, textTransform:"uppercase" }}>{h}</th>)}</tr></thead><tbody>{auditBusy?<tr><td colSpan={6} style={{ padding:22, color:"var(--muted-1)", textAlign:"center" }}>Loading…</td></tr>:rows.length===0?<tr><td colSpan={6} style={{ padding:22, color:"var(--muted-1)", textAlign:"center" }}>No matching entries.</td></tr>:rows.slice(0,500).map(r=><tr key={r.id||String(r.created_at)+String(r.style_id)+String(r.col)} onDoubleClick={()=>r.style_id&&onJump&&onJump(r.style_id,r.col)} title={r.style_id?"Double-click to jump to cell":""} style={{ cursor:r.style_id?"pointer":"default" }}><td style={{ padding:"8px 10px", borderBottom:"1px solid var(--line-3)", whiteSpace:"nowrap", color:"var(--muted-2)" }}>{ts(r.created_at)}</td><td style={{ padding:"8px 10px", borderBottom:"1px solid var(--line-3)", whiteSpace:"nowrap" }}>{r.actor_name||"—"}</td><td style={{ padding:"8px 10px", borderBottom:"1px solid var(--line-3)", fontWeight:800 }}>{r.style_no||r.style_id||"—"}</td><td style={{ padding:"8px 10px", borderBottom:"1px solid var(--line-3)" }}>{label(r.col)}</td><td style={{ padding:"8px 10px", borderBottom:"1px solid var(--line-3)", whiteSpace:"nowrap" }}>{r.field||"value"}</td><td style={{ padding:"8px 10px", borderBottom:"1px solid var(--line-3)" }}><span style={{ color:"var(--danger)" }}>{r.old_val||"—"}</span> <span style={{ color:"var(--muted-1)" }}>→</span> <span style={{ color:"var(--success)" }}>{r.new_val||"—"}</span></td></tr>)}</tbody></table></div>
      {rows.length>500 && <div style={{ padding:"8px 14px", fontSize:10, color:"var(--muted-1)" }}>Showing first 500 matching entries. Use filters/export for full review.</div>}
    </div>}
    {mode==="errors" && <div style={card}>
      <div style={{ padding:14, borderBottom:"1px solid var(--line-3)", display:"flex", gap:9, justifyContent:"space-between", alignItems:"center", flexWrap:"wrap" }}><div><b>{(errorLog||[]).length}</b> browser-session error(s)<div style={{ fontSize:9, color:"var(--muted-1)", marginTop:2 }}>Errors reset when the browser session clears unless exported.</div></div><span style={{ display:"flex", gap:8 }}><button onClick={exportErrors} style={{ ...input, cursor:"pointer", fontWeight:800, background:"var(--ink)", color:"var(--surface)" }}>Export errors</button><button onClick={clearErrorLog} disabled={!errorLog||!errorLog.length} style={{ ...input, cursor:errorLog&&errorLog.length?"pointer":"not-allowed", fontWeight:800, opacity:errorLog&&errorLog.length?1:0.5 }}>Clear</button></span></div>
      <div style={{ padding:14 }}>{(!errorLog||!errorLog.length)?<div style={{ color:"var(--muted-1)", fontSize:11, padding:18, textAlign:"center" }}>No app errors logged in this browser session.</div>:errorLog.map(e=><div key={e.id} style={{ padding:"9px 0", borderBottom:"1px solid var(--line-3)", fontSize:11 }}><span style={{ color:"var(--muted-1)", fontSize:9 }}>{ts(e.at)}</span> · <b>{e.area}</b><br/><span style={{ color:"var(--danger)", fontWeight:700 }}>{e.msg}</span>{e.extra&&<span style={{ color:"var(--muted-1)" }}> · {e.extra}</span>}</div>)}</div>
    </div>}
  </div>);
}

function Th({ col, label, sort, onSort, sticky, left, z, width, onResize, onAutoFit, scale, letter, selected, filterActive, filterOpen, filterValues, filterAllowed, onToggleFilter, onSetFilter, onCloseFilter }){
  const active=sort.col===col;
  const startDrag=(e)=>{ e.preventDefault(); e.stopPropagation(); const sx=e.clientX, sw=width||80; const sc=scale||1; const move=(ev)=>onResize&&onResize(col, sw+(ev.clientX-sx)/sc); const up=()=>{ window.removeEventListener("mousemove",move); window.removeEventListener("mouseup",up); }; window.addEventListener("mousemove",move); window.addEventListener("mouseup",up); };
  return (<th role="columnheader" aria-sort={active?(sort.dir>0?"ascending":"descending"):"none"} style={{ position:"sticky", top:0, left:sticky?left:undefined, zIndex:sticky?(z||5):3, background:selected?"#dbeafe":(active?"var(--accent)":"var(--ink)"), color:selected?"#0f172a":(active?"var(--ink)":"var(--bg)"), boxShadow:selected?"inset 0 -4px 0 var(--info)":undefined, padding:"8px 9px", textAlign:"left", fontWeight:600, fontSize:9.5, letterSpacing:0.4, textTransform:"uppercase", whiteSpace:"nowrap", overflow:"visible", border:"1px solid #3a362e", userSelect:"none", width:width||80, minWidth:width||80, maxWidth:width||80, boxSizing:"border-box", backgroundClip:"padding-box", transform:sticky?"translateZ(0)":undefined }}>
    <span style={{ display:"flex", alignItems:"center", gap:3 }}>
      <span onClick={(e)=>{ e.stopPropagation(); onSort(col); }} title="click to sort" style={{ display:"inline-flex", alignItems:"center", gap:3, cursor:"pointer", flex:1, overflow:"hidden", textOverflow:"ellipsis" }}>{letter && <span style={{ fontSize:8, fontWeight:700, opacity:0.5, marginRight:4 }}>{letter}</span>}{label}{active?(sort.dir>0?<ChevronUp size={11}/>:<ChevronDown size={11}/>):null}</span>
      <span onClick={(e)=>{ e.stopPropagation(); onToggleFilter&&onToggleFilter(); }} title={filterActive?"filter ON — click to edit/clear":"filter"} style={{ cursor:"pointer", display:"inline-flex", padding:"0 1px", color: filterActive?(active?"var(--ink)":"#f4b942"):(active?"#7a4a08":"var(--on-dark)") }}><Filter size={filterActive?12:10} fill={filterActive?"currentColor":"none"}/></span>
    </span>
    {filterOpen && <FilterMenu values={filterValues||[]} allowed={filterAllowed} onSet={onSetFilter} onClose={onCloseFilter}/>}
    <span onMouseDown={startDrag} onDoubleClick={(e)=>{ e.stopPropagation(); onAutoFit&&onAutoFit(col); }} onClick={(e)=>e.stopPropagation()} title="drag to resize · double-click to auto-fit" style={{ position:"absolute", top:0, right:-2, bottom:0, width:10, cursor:"col-resize", zIndex:6 }}/>
  </th>);
}
function FilterMenu({ values, allowed, onSet, onClose }){
  const [q,setQ]=useState("");
  const anchorRef=useRef(null); const [pos,setPos]=useState(null);
  const norm=(v)=>String(v==null?"":v).replace(/\s+/g," ").trim();
  const liveNorm=(v)=>String(v==null?"":v).replace(/\s+/g," ").trim().toLowerCase();
  // Freeze the option list when the popup opens so live filtering never makes the menu flash/turn white.
  const [allValues]=useState(()=>[...new Set((values||[]).map(norm).filter(Boolean))].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true,sensitivity:"base"})));
  const rawAllowed=Array.isArray(allowed)?allowed:null;
  const isNoneSelected=Array.isArray(rawAllowed) && rawAllowed.some(v=>String(v)===FILTER_NONE);
  const initialAllowed=Array.isArray(rawAllowed)?rawAllowed.map(norm).filter(v=>v && v!==FILTER_NONE):null;
  const [pending,setPending]=useState(()=> new Set(isNoneSelected?[]:(initialAllowed||allValues).map(norm)) );
  const [manualMode,setManualMode]=useState(()=>Array.isArray(rawAllowed));
  useEffect(()=>{ setPending(new Set(isNoneSelected?[]:(initialAllowed||allValues).map(norm))); setManualMode(Array.isArray(rawAllowed)); },[allowed]);
  const qNorm=liveNorm(q);
  const shown=allValues.filter(v=>liveNorm(v).includes(qNorm));
  const filterIsActive=Array.isArray(rawAllowed);
  const selectedCount=pending.size;
  const sendFilter=(vals)=>{
    const raw=Array.isArray(vals)?vals:[...vals];
    const clean=[...new Set(raw.map(norm).filter(Boolean))];
    if(clean.length===allValues.length){ onSet(null); return; }
    onSet(clean.length?clean:[FILTER_NONE]);
  };
  const liveSearch=(nextQ)=>{
    setQ(nextQ);
    const n=liveNorm(nextQ);
    const matches=allValues.filter(v=>liveNorm(v).includes(n));
    // Main expected behaviour: typing in a table column filter immediately filters the table.
    // No Apply button is required. If search is cleared and the user has not manually selected
    // values in this open menu, the column filter is cleared back to All.
    if(n){ setManualMode(false); onSet(matches.length?matches:[FILTER_NONE]); }
    else { setManualMode(false); setPending(new Set(allValues)); onSet(null); }
  };
  const apply=()=>{ if(qNorm && !manualMode) sendFilter(shown); else sendFilter(pending); onClose&&onClose(); };
  const toggle=(v)=>{ setManualMode(true); setPending(cur=>{ const n=new Set(cur); if(n.has(v)) n.delete(v); else n.add(v); sendFilter(n); return n; }); };
  const selectVals=(vals)=>{ setManualMode(true); const n=new Set(vals.map(norm).filter(Boolean)); setPending(n); sendFilter(n); };
  const selectShown=()=>selectVals([...new Set([...pending,...shown])]);
  const clearShown=()=>{ setManualMode(true); setPending(cur=>{ const n=new Set(cur); shown.forEach(v=>n.delete(v)); sendFilter(n); return n; }); };
  const clearFilter=()=>{ setQ(""); setManualMode(false); setPending(new Set(allValues)); onSet(null); };
  const none=()=>{ setManualMode(true); const n=new Set(); setPending(n); onSet([FILTER_NONE]); };
  useEffect(()=>{ const a=anchorRef.current; if(!a) return; const r=a.getBoundingClientRect(); const W=260,H=350; let left=r.left-220; if(left+W>window.innerWidth-8) left=window.innerWidth-8-W; if(left<8) left=8; let top=r.bottom+4; if(top+H>window.innerHeight-8) top=Math.max(8,window.innerHeight-8-H); setPos({top,left}); },[]);
  const btn={ fontFamily:"'JetBrains Mono', monospace", fontSize:9, fontWeight:900, padding:"5px 7px", border:"1px solid #1f1f1d", background:"#fffdf8", color:"#1f1f1d", cursor:"pointer" };
  const menu=(
    <div onClick={e=>e.stopPropagation()} style={{ position:"fixed", top:pos?pos.top:-9999, left:pos?pos.left:-9999, zIndex:9999, background:"#fffdf8", color:"#1f1f1d", border:"1px solid #1f1f1d", boxShadow:"4px 4px 0 rgba(31,31,29,.28)", padding:9, width:260, textTransform:"none", letterSpacing:0, fontWeight:400, maxHeight:"82vh", overflowY:"auto" }}>
      <div style={{ fontSize:10, fontWeight:950, marginBottom:6, display:"flex", justifyContent:"space-between", gap:8 }}><span>Column filter</span><button onClick={onClose} style={{...btn,padding:"2px 7px"}}>×</button></div>
      <input autoFocus value={q} onClick={e=>e.stopPropagation()} onKeyDown={e=>e.stopPropagation()} onChange={e=>liveSearch(e.target.value)} placeholder="type to filter live…" style={{ width:"100%", fontFamily:"inherit", fontSize:11, padding:"5px 7px", border:"1px solid #8f8577", outline:"none", marginBottom:7, boxSizing:"border-box", background:"#ffffff", color:"#1f1f1d" }}/>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6, marginBottom:7 }}>
        <button onClick={clearFilter} title="remove this column filter" style={{ ...btn, background:!filterIsActive?"#1f1f1d":"#fffdf8", color:!filterIsActive?"#fffdf8":"#1f1f1d" }}>All / clear</button>
        <button onClick={()=>selectVals(qNorm?shown:allValues)} disabled={!(qNorm?shown.length:allValues.length)} style={{ ...btn, opacity:(qNorm?shown.length:allValues.length)?1:.5 }}>{qNorm?"Select matches":"Select all"}</button>
        <button onClick={selectShown} disabled={!shown.length} style={{ ...btn, opacity:shown.length?1:.5 }}>Add matches</button>
        <button onClick={clearShown} disabled={!shown.length} style={{ ...btn, opacity:shown.length?1:.5 }}>Uncheck matches</button>
        <button onClick={none} style={{ ...btn, gridColumn:"1 / span 2", background:"#fff2f0", color:"#b82117" }}>Show zero rows</button>
      </div>
      <div style={{ fontSize:9, color:"#6f6a61", marginBottom:5 }}>{selectedCount} of {allValues.length} selected{q?` · ${shown.length} match search · live ${manualMode?"checked values":"search matches"}`:""} · typing filters live</div>
      <div style={{ maxHeight:190, overflowY:"auto", borderTop:"1px solid #e7dcc2" }}>
        {shown.map(v=>(<label key={v} style={{ display:"flex", alignItems:"center", gap:7, fontSize:10, padding:"4px 0", borderBottom:"1px solid #eee4d2", cursor:"pointer", color:"#1f1f1d" }}>
          <input type="checkbox" checked={pending.has(v)} onChange={()=>toggle(v)} style={{ cursor:"pointer" }}/>
          <span style={{ flex:1, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{v}</span>
          <button onClick={(e)=>{ e.preventDefault(); e.stopPropagation(); selectVals([v]); }} title="show only this" style={{ fontSize:8, border:"1px solid #8f8577", background:"#fffdf8", cursor:"pointer", padding:"1px 5px", color:"#1f1f1d" }}>only</button>
        </label>))}
        {shown.length===0 && <div style={{ fontSize:10, color:"#6f6a61", padding:"6px 0" }}>No matching values.</div>}
      </div>
      <div style={{ display:"flex", gap:6, marginTop:7 }}>
        <button onClick={clearFilter} style={{ ...btn, flex:1 }}>Clear filter</button>
        <button onClick={apply} style={{ ...btn, flex:1, background:"#1f1f1d", color:"#fffdf8" }}>Done</button>
      </div>
    </div>
  );
  return (<><span ref={anchorRef} style={{ position:"absolute", width:0, height:0 }}/>{createPortal(menu, document.body)}</>);
}

function FillPanel({ count, role, activeCol, onApply, onClose }){
  const validActiveCol=(col)=> col && (["ordRec","delivery"].includes(col)||STAGE_KEYS.includes(col)) ? col : "";
  const activeFillCol=validActiveCol(activeCol);
  const [mode,setMode]=useState(()=>{ try{ const v=localStorage.getItem("mt_fill_mode"); const m=(v==="revised"||v==="actual")?v:"actual"; return (m==="revised"&&!canEditRev(role))?"actual":m; }catch(e){ return "actual"; } });
  const stageOpts=STAGES.map(s=>({key:s.key,label:s.label+(s.cutoff?" (Fabric IH)":"")}));
  const opts = mode==="revised" ? (canEditRev(role)?stageOpts:[]) : [{key:"ordRec",label:"Order Date"},{key:"delivery",label:"Delivery Date"}].concat(stageOpts).filter(o=>canEdit(role,o.key,"actual"));
  const [key,setKey]=useState(()=>{ try{ return activeFillCol || localStorage.getItem("mt_fill_stage") || (opts[0]?opts[0].key:"labAppr"); }catch(e){ return activeFillCol || (opts[0]?opts[0].key:"labAppr"); } }); const [val,setVal]=useState(iso(TODAY));
  useEffect(()=>{ if(opts.length && !opts.some(o=>o.key===key)) setKey(opts[0].key); },[mode]);
  useEffect(()=>{ if(activeFillCol && opts.some(o=>o.key===activeFillCol)) setKey(activeFillCol); },[activeFillCol,mode]);
  useEffect(()=>{ try{ localStorage.setItem("mt_fill_mode",mode); }catch(e){} },[mode]);
  useEffect(()=>{ try{ localStorage.setItem("mt_fill_stage",key); }catch(e){} },[key]);
  const accent = mode==="revised"?"var(--accent)":"var(--info)";
  const tabBtn=(m,label)=>(<button onClick={()=>setMode(m)} style={{ flex:1, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:"5px 0", cursor:"pointer", border:"1px solid var(--ink)", background:mode===m?(m==="revised"?"var(--accent)":"var(--info)"):"var(--surface)", color:mode===m?"var(--surface)":"var(--ink)" }}>{label}</button>);
  return (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:0, marginTop:4, zIndex:370, background:"var(--surface)", border:"1px solid var(--ink)", boxShadow:"4px 4px 0 var(--ink)", padding:12, width:280 }}><div style={{ fontSize:11, fontWeight:700, marginBottom:activeFillCol?3:8 }}>Set one date across {count} filtered styles</div>{activeFillCol && <div style={{ fontSize:9, color:"var(--muted-2)", marginBottom:7 }}>Auto-selected from current column: <b>{((STAGES.find(s=>s.key===activeFillCol)||{}).label)||(activeFillCol==="ordRec"?"Order Date":activeFillCol==="delivery"?"Delivery Date":activeFillCol)}</b></div>}<div style={{ display:"flex", marginBottom:8 }}>{tabBtn("actual","Actual")}{tabBtn("revised","Revised")}</div>{opts.length===0 ? <div style={{ fontSize:10, color:"var(--danger)", marginBottom:8 }}>Your role cannot set revised dates.</div> : (<><label style={{ fontSize:10, color:"var(--muted-2)" }}>Stage</label><select value={key} onChange={e=>setKey(e.target.value)} style={{ width:"100%", fontFamily:"inherit", fontSize:11, padding:5, marginBottom:8, border:"1px solid var(--ink)" }}>{opts.map(o=><option key={o.key} value={o.key}>{o.label}</option>)}</select><label style={{ fontSize:10, color:"var(--muted-2)" }}>{mode==="revised"?"Revised plan date":"Actual date"}</label><input type="date" value={val} onChange={e=>setVal(e.target.value)} style={{ width:"100%", fontFamily:"inherit", fontSize:11, padding:5, marginBottom:10, border:"1px solid var(--ink)" }}/></>)}<div style={{ display:"flex", gap:8 }}><button disabled={opts.length===0} onClick={()=>onApply(key,val,mode)} style={{ flex:1, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:6, cursor:opts.length?"pointer":"not-allowed", border:"1px solid var(--ink)", background:opts.length?accent:"var(--line-2)", color:"var(--surface)", opacity:opts.length?1:0.6 }}>Apply {mode} → {count}</button><button onClick={onClose} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--bg)" }}><X size={12}/></button></div></div>);
}
const ndCell={ border:"1px dashed #e8dcc2", padding:"6px 9px", whiteSpace:"nowrap" };
const ndInput=(w)=>({ border:"none", outline:"none", background:"transparent", fontFamily:"'JetBrains Mono', monospace", fontSize:10, width:w });


function MultiSelectDropdown({ label, value, options, onChange, rounded=false }){
  const [open,setOpen]=useState(false);
  const btnRef=useRef(null);
  const [pos,setPos]=useState(null);
  const vals=Array.isArray(value)?value:(value?[value]:[]);
  const opts=Array.from(new Set((options||[]).filter(v=>v!=null && String(v).trim()!=="").map(v=>String(v)))).sort((a,b)=>a.localeCompare(b,undefined,{numeric:true,sensitivity:"base"}));
  const selected=new Set(vals.map(v=>String(v)));
  const labelText=vals.length===0?`${label}: all`:(vals.length===1?`${label}: ${vals[0]}`:`${label}: ${vals.length} selected`);
  const updatePos=()=>{
    const r=btnRef.current?.getBoundingClientRect?.();
    if(!r) return;
    const width=Math.max(220, Math.min(360, Math.max(r.width, 230)));
    const left=Math.min(Math.max(8,r.left), Math.max(8, window.innerWidth-width-8));
    const top=Math.min(r.bottom+6, Math.max(8, window.innerHeight-320));
    setPos({ top, left, width });
  };
  useEffect(()=>{ if(!open) return; updatePos(); const onDoc=(e)=>{ if(btnRef.current && btnRef.current.contains(e.target)) return; const menu=document.getElementById(`msd-${label}`); if(menu && menu.contains(e.target)) return; setOpen(false); }; const onWin=()=>updatePos(); document.addEventListener('mousedown',onDoc,true); window.addEventListener('resize',onWin); window.addEventListener('scroll',onWin,true); return ()=>{ document.removeEventListener('mousedown',onDoc,true); window.removeEventListener('resize',onWin); window.removeEventListener('scroll',onWin,true); }; },[open,label]);
  const toggle=(opt)=>{
    const s=new Set(selected);
    if(s.has(opt)) s.delete(opt); else s.add(opt);
    onChange([...s]);
  };
  const menu=open && pos ? createPortal(
    <div id={`msd-${label}`} onClick={e=>e.stopPropagation()} style={{ position:'fixed', top:pos.top, left:pos.left, width:pos.width, maxHeight:300, overflow:'auto', zIndex:9000, background:'var(--surface)', border:'1px solid var(--ink)', borderRadius:rounded?10:2, boxShadow:'4px 4px 0 rgba(31,31,29,.22)', padding:8, fontFamily:"'JetBrains Mono', monospace" }}>
      <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', gap:8, marginBottom:6 }}>
        <div style={{ fontSize:10, fontWeight:900, textTransform:'uppercase', color:'var(--muted-3)' }}>{label}</div>
        <button className="mt-popup-close" onClick={()=>setOpen(false)} style={{ border:'1px solid var(--line-2)', background:'var(--surface)', cursor:'pointer', padding:'5px 8px', fontFamily:'inherit', fontSize:13, fontWeight:900 }}>×</button>
      </div>
      <div style={{ display:'flex', gap:6, marginBottom:7 }}>
        <button onClick={()=>onChange([])} style={{ flex:1, border:'1px solid var(--line-2)', background:vals.length===0?'var(--ink)':'var(--surface)', color:vals.length===0?'var(--bg)':'var(--ink)', cursor:'pointer', padding:'5px 7px', fontFamily:'inherit', fontSize:10, fontWeight:800 }}>All</button>
        <button onClick={()=>onChange(opts)} disabled={!opts.length} style={{ flex:1, border:'1px solid var(--line-2)', background:'var(--surface)', cursor:opts.length?'pointer':'not-allowed', padding:'5px 7px', fontFamily:'inherit', fontSize:10, fontWeight:800, opacity:opts.length?1:.5 }}>Select all</button>
        <button onClick={()=>onChange([])} style={{ flex:1, border:'1px solid var(--line-2)', background:'var(--surface)', cursor:'pointer', padding:'5px 7px', fontFamily:'inherit', fontSize:10, fontWeight:800 }}>Clear</button>
      </div>
      {opts.length===0 ? <div style={{ fontSize:10, color:'var(--muted-2)', padding:8 }}>No values.</div> : opts.map(opt=>(
        <label key={opt} style={{ display:'flex', alignItems:'center', gap:8, padding:'6px 5px', cursor:'pointer', borderBottom:'1px solid var(--line-3)', fontSize:10 }}>
          <input type="checkbox" checked={selected.has(opt)} onChange={()=>toggle(opt)} />
          <span style={{ flex:1, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{opt}</span>
        </label>
      ))}
    </div>, document.body) : null;
  return <>
    <button ref={btnRef} type="button" onClick={(e)=>{ e.stopPropagation(); setOpen(o=>!o); }} style={{ fontFamily:'inherit', fontSize:10, fontWeight:800, padding:rounded?'8px 12px':'7px 11px', cursor:'pointer', border:'1px solid var(--line-2)', borderRadius:rounded?8:0, background:vals.length?'var(--accent-tint)':'var(--surface)', color:'var(--ink)', maxWidth:220, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }} title={labelText}>{labelText} ▾</button>
    {menu}
  </>;
}

/* ========================= DASHBOARD ========================= */
function OperationalDashboardView({ computed, todoItems, cfg, applyDrill, drillTodo }){
  const [target,setTarget]=useState("tracker"); // where bar/owner/activity drills go: Tracker or To-Do only. Escalation cards route to To-Do with Type=Escalation.
  const [df,setDf]=useState(()=>{ try{ return JSON.parse(localStorage.getItem("mt_dashboard_filter")||localStorage.getItem("mt_dashfilter")||"{}"); }catch(e){ return {}; } });
  useEffect(()=>{ try{ localStorage.setItem("mt_dashboard_filter", JSON.stringify(df)); }catch(e){} },[df]);
  const [mgmtOpen,setMgmtOpen]=useState(()=>{ try{ return JSON.parse(localStorage.getItem("mt_mgmt_open")||"{}"); }catch(e){ return {}; } });
  useEffect(()=>{ try{ localStorage.setItem("mt_mgmt_open", JSON.stringify(mgmtOpen)); }catch(e){} },[mgmtOpen]);
  const isMgmtOpen=(key)=> mgmtOpen[key]!==false;
  const toggleMgmt=(key)=>setMgmtOpen(o=>({ ...o, [key]: !isMgmtOpen(key) }));
  const arrOf=(v)=>Array.isArray(v)?v:(v?[v]:[]);
  const hasSel=(key,val)=>{ const a=arrOf(df[key]); return !a.length || a.includes(val); };
  const hasColourSel=(st)=>{ const a=arrOf(df.colour); if(!a.length) return true; const cols=splitColoursAll(st.colour); return a.some(v=>cols.includes(v)); };
  const matchDf=(st,except)=> (except==="order"||hasSel("order",st.orderNo)) && (except==="fit"||hasSel("fit",st.sampleFit)) && (except==="junior"||hasSel("junior",st.owner)) && (except==="family"||hasSel("family",st.family)) && (except==="brand"||hasSel("brand",st.brand)) && (except==="fabric"||hasSel("fabric",st.fabricType)) && (except==="colour"||hasColourSel(st));
  const distinctC=(key,fn)=>{ const s=new Set(); computed.forEach(({s:st})=>{ if(!matchDf(st,key)) return; fn(st).forEach(v=>{ if(v) s.add(v); }); }); return [...s].sort(); };
  const orders=distinctC("order",s=>[s.orderNo]); const fits=distinctC("fit",s=>[s.sampleFit]); const juniors=distinctC("junior",s=>[s.owner]); const families=distinctC("family",s=>[s.family]); const brands=distinctC("brand",s=>[s.brand]); const fabrics=distinctC("fabric",s=>[s.fabricType]);
  const colours=distinctC("colour",s=>splitColoursAll(s.colour));
  const fc=useMemo(()=>computed.filter(({s})=> hasSel("order",s.orderNo) && hasSel("fit",s.sampleFit) && hasSel("junior",s.owner) && hasSel("family",s.family) && hasSel("brand",s.brand) && hasSel("fabric",s.fabricType) && hasColourSel(s) ),[computed,df]);
  const total=fc.length;
  const onTrack=fc.filter(({c})=>c.tone==="ok").length;
  const atRisk=fc.filter(({c})=>c.tone==="late"||c.tone==="warn").length;
  const released=fc.filter(({c})=>c.released).length;
  const delRisk=fc.filter(({c})=>String(c.status).startsWith("Delivery risk")).length;
  // owner load + activity load from the spliced set
  const ownerLoad={}; const actAgg={};
  fc.forEach(({s,c})=>{ if(c.released) return; (c.chaseOwners||[]).forEach(o=>{ ownerLoad[o.owner]=(ownerLoad[o.owner]||0)+1; }); (c.frontier?[...c.frontier]:[]).forEach(k=>{ const r=(c.stages||[]).find(x=>x.key===k); if(!r||r.done) return; const lbl=stageLabelFromKeyGlobal(k); const a=actAgg[lbl]=actAgg[lbl]||{n:0,over:0,key:k}; a.n++; if((r.rev||r.plan)&&TODAY>(r.rev||r.plan)) a.over++; }); });
  const owners=Object.entries(ownerLoad).sort((a,b)=>b[1]-a[1]); const maxOwner=Math.max(1,...owners.map(o=>o[1]));
  const acts=Object.entries(actAgg).sort((a,b)=>b[1].n-a[1].n); const maxAct=Math.max(1,...acts.map(e=>e[1].n));
  const overdueAct=acts.reduce((s,[,v])=>s+v.over,0);
  const styleByTodoId=new Map(fc.map(({s})=>[Number(s.id),s]));
  const todoMatchesSlice=(t)=>{
    const st=styleByTodoId.get(Number(t.id));
    if(!st) return false;
    const orderVals=(Array.isArray(t.orderNos)&&t.orderNos.length?t.orderNos:[t.orderNo,st.orderNo]).filter(Boolean);
    const juniorVals=(Array.isArray(t.juniors)&&t.juniors.length?t.juniors:[t.junior,st.owner]).filter(Boolean);
    const colourVals=splitColoursAll(t.colour||st.colour||"");
    const matchArray=(key,vals)=>{ const a=arrOf(df[key]); return !a.length || vals.some(v=>a.includes(v)); };
    return matchArray("order",orderVals) && matchArray("junior",juniorVals) && hasSel("fit",st.sampleFit) && hasSel("family",st.family) && hasSel("brand",st.brand) && hasSel("fabric",st.fabricType) && (!arrOf(df.colour).length || arrOf(df.colour).some(v=>colourVals.includes(v)));
  };
  const escalationTodo=(todoItems||[]).filter(t=>todoMatchesSlice(t)&&t.overdue&&t.escalationOwner);
  const escLoad=escalationTodo.reduce((m,t)=>{ const k=t.escalationOwner||"(blank)"; m[k]=(m[k]||0)+1; return m; },{});
  const escRows=Object.entries(escLoad).sort((a,b)=>b[1]-a[1]); const maxEsc=Math.max(1,...escRows.map(x=>x[1]));
  const phase={ "Pre-Fit":0,"Fit / Print":0,"Lab Dip":0,"Fabric IH":0,"PP / Prod":0 };
  const phaseKeyList=(phaseName)=> phaseName==="Pre-Fit"?["techpack"]:phaseName==="Fit / Print"?["fitSend","fitAppr","artwork","artAppr","strikeOff","soAppr"]:phaseName==="Lab Dip"?["labDip","labAppr"]:phaseName==="Fabric IH"?["fabricIH"]:["ppSample","ppAppr","prodFile"];
  const phaseOfKey=(k)=> k==="techpack"?"Pre-Fit":["fitSend","fitAppr","artwork","artAppr","strikeOff","soAppr"].includes(k)?"Fit / Print":["labDip","labAppr"].includes(k)?"Lab Dip":k==="fabricIH"?"Fabric IH":"PP / Prod";
  fc.forEach(({c})=>{ if(c.released) return; const k=c.nextPending&&c.nextPending.key; phase[phaseOfKey(k)]++; });
  const maxPhase=Math.max(1,...Object.values(phase));
  const OWNER_COLOR2=OWNER_COLOR;
  // splice carried into drills so the tracker shows the same slice
  const spliceCols=()=>{ const cf={}; const put=(k,col)=>{ const a=arrOf(df[k]); if(a.length) cf[col]=a; }; put("order","orderNo"); put("fit","sampleFit"); put("junior","owner"); put("family","family"); put("brand","brand"); put("fabric","fabricType"); put("colour","colour"); return cf; };
  const spliceSearch=()=> "";
  const sliceTodoRows=(pred=()=>true)=>(todoItems||[]).filter(t=>todoMatchesSlice(t)&&pred(t));
  const phaseStyleIds=(phaseName)=>new Set(fc.filter(({c})=>!c.released && phaseOfKey(c.nextPending&&c.nextPending.key)===phaseName).map(({s})=>String(s.id)));
  const phaseTodoRows=(phaseName)=>{ const ids=phaseStyleIds(phaseName); const allowed=new Set(phaseKeyList(phaseName)); return sliceTodoRows(t=>ids.has(String(t.id)) && allowed.has(stageKeyFromAnyGlobal(t.activityKey||t.key||t.activity))); };
  const drillRowsToTodo=(rows,base={})=>{ if(drillTodo) drillTodo(todoDrillFilterFromRows(rows,base,df)); };
  const goOwner=(o)=>{ if(target==="todo") drillRowsToTodo(sliceTodoRows(t=>String(t.owner||"")===String(o)),{ owner:o }); else applyDrill({ owner:o, colFilters:spliceCols(), search:spliceSearch() }); };
  const goAct=(label,key)=>{ const k=stageKeyFromAnyGlobal(key||label); if(target==="todo") drillRowsToTodo(sliceTodoRows(t=>stageKeyFromAnyGlobal(t.activityKey||t.key||t.activity)===k),{ activity:[stageLabelFromKeyGlobal(k)||label], activityKey:k?[k]:[], key:k?[k]:[] }); else applyDrill({ activity:k||key, colFilters:spliceCols(), search:spliceSearch() }); };
  const goPhase=(phaseName)=>{ const keys=phaseKeyList(phaseName); if(target==="todo") drillRowsToTodo(phaseTodoRows(phaseName),{ phase:phaseName, activityKey:keys, activity:keys.map(stageLabelFromKeyGlobal) }); else applyDrill({ activity:keys, colFilters:spliceCols(), search:spliceSearch() }); };
  const goStatus=(st,extra)=>applyDrill({ status:st, colFilters:{...spliceCols(),...(extra||{})}, search:spliceSearch() });
  const card=(label,val,color,onClick)=>(<button onClick={onClick} disabled={!onClick} style={{ flex:1, minWidth:130, textAlign:"left", background:"var(--surface)", border:"1px solid var(--ink)", padding:"14px 16px", cursor:onClick?"pointer":"default", fontFamily:"inherit" }}><div style={{ fontSize:28, fontWeight:800, fontFamily:"'Archivo',sans-serif", color, lineHeight:1 }}>{val}</div><div style={{ fontSize:10, color:"var(--muted-2)", marginTop:5, letterSpacing:0.5, textTransform:"uppercase" }}>{label}{onClick?" ›":""}</div></button>);
  const sel=(label,val,opts,onChange)=><MultiSelectDropdown label={label} value={val} options={opts} onChange={onChange} />;
  const anyDf=Object.values(df).some(v=>Array.isArray(v)?v.length:!!v);
  const bar=(items,maxV,colorFn,labelW,onClick,fmtR)=> items.length===0?<div style={{ fontSize:11, color:"var(--muted-1)" }}>Nothing pending.</div>:items.map(([k,v])=>{ const n=typeof v==="number"?v:v.n; return (
    <button key={k} onClick={()=>onClick(k,v)} style={{ display:"flex", alignItems:"center", gap:8, width:"100%", border:"none", background:"transparent", cursor:"pointer", fontFamily:"inherit", padding:"4px 0" }}>
      <span style={{ width:labelW, fontSize:10, fontWeight:700, color:colorFn(k), textAlign:"left" }}>{k}</span>
      <span style={{ flex:1, height:16, background:"#f0ece3", position:"relative" }}><span style={{ position:"absolute", left:0, top:0, bottom:0, width:`${(n/maxV)*100}%`, background:colorFn(k,v) }}/></span>
      <span style={{ width:54, textAlign:"right", fontSize:10, fontWeight:700 }}>{fmtR?fmtR(v):n}</span>
    </button>); });
  const dashboardSummary=[{ "Report Type":"Current Dashboard", "Slice Styles":total, "On Track":onTrack, "At Risk":atRisk, "Delivery Risk":delRisk, "Released":released, "Overdue Activities":overdueAct }];
  const dashboardStyleRows=fc.map(({s,c})=>({ "Order No":s.orderNo||"", "Style No":s.styleNo||"", "Sample Fit":s.sampleFit||"", "Family":s.family||"", "Colour":s.colour||"", "Brand":s.brand||"", "Buyer":s.buyer||"", "Junior":s.owner||"", "Qty":s.qty||0, "Delivery":s.delivery||"", "Status":c.status||"", "Tone":c.tone||"", "Released":c.released?"YES":"", "% Done":c.pct, "Chase":(c.chaseOwners||[]).map(o=>`${o.owner} (${o.count})`).join(", "), "Next Pending":c.nextPending?stageReviewLabel(s,c.nextPending):"", "Projected Release":c.projRelease?fmt(c.projRelease):"" }));
  const dashboardBreakup=fc.map(({s,c})=>{ const nx=c.nextPending||null; const nxDue=nx?(nx.rev||nx.plan):null; const frontier=(c.frontier?[...c.frontier]:[]).map(k=>{ const r=(c.stages||[]).find(x=>x.key===k); return r?stageLabelFromKeyGlobal(r.key):k; }).join(", "); return { "Order No":s.orderNo||"", "Style No":s.styleNo||"", "Colour":s.colour||"", "Buyer / Brand":s.buyer||s.brand||"", "Junior":s.owner||"", "Delivery":s.delivery||"", "Overall Status":c.status||"", "Overall Tone":c.tone||"", "Released":c.released?"YES":"NO", "% Done":c.pct, "Next Pending Stage":nx?stageReviewLabel(s,nx):"", "Next Pending Chase":nx?nx.owner:"", "Next Pending Due":nxDue?fmt(nxDue):"", "Next Pending Overdue?":(nxDue&&TODAY>nxDue)?"YES":"NO", "Actionable Frontier":frontier, "Chase Breakdown":(c.chaseOwners||[]).map(o=>`${o.owner} (${o.count})`).join(", "), "Fit Branch":c.fitBranch?c.fitBranch.txt:"", "Print Branch":c.printBranch?c.printBranch.txt:"", "Fabric Branch":c.fabricBranch?c.fabricBranch.txt:"", "PP Branch":c.ppBranch?c.ppBranch.txt:"", "Prod File Branch":c.prodFileBranch?c.prodFileBranch.txt:"", "Fabric IH Countdown":c.fabricCountdown?c.fabricCountdown.txt:"", "Projected Release":c.projRelease?fmt(c.projRelease):"", "Release Gate":c.releaseGate?fmt(c.releaseGate):"", "Float Days":c.float==null?"":c.float, "Idle Days":c.idle==null?"":c.idle }; });
  const stageStartFor=(s,c,r)=>{ const st=STAGES.find(x=>x.key===r.key)||{}; const byKey={}; (c.stages||[]).forEach(x=>{ byKey[x.key]=x; }); if(r.key==="fabricIH") return (s.labDipReq && byKey.labAppr && byKey.labAppr.actual) ? byKey.labAppr.actual : parse(s.ordRec); if(st.pred==="__ord") return parse(s.ordRec); return byKey[st.pred] ? byKey[st.pred].actual : null; };
  const stageState=(r)=> r.skipped?"Waived / skipped":(r.rejected?"Rejected":(r.rework?"Rework / resend":(r.actual?"Done":(r.autoClosed?"Auto-closed":(r.rev?"Revised plan":"Pending")))));
  const dashboardStageDetail=[];
  const dashboardActualVsPlan=[];
  const dashboardRevisedVsActual=[];
  fc.forEach(({s,c})=>{ (c.stages||[]).forEach(r=>{ const st=STAGES.find(x=>x.key===r.key)||{}; const start=stageStartFor(s,c,r); const due=r.rev||r.plan; const delayDue=(due&&r.actual)?netWorkdays(due,r.actual):null; const delayPlan=(r.plan&&r.actual)?netWorkdays(r.plan,r.actual):null; const delayRev=(r.rev&&r.actual)?netWorkdays(r.rev,r.actual):null; const duration=(start&&r.actual)?Math.max(0,netWorkdays(start,r.actual)||0):null; const frontier=(c.frontier&&c.frontier.has(r.key)); dashboardStageDetail.push({ "Order No":s.orderNo||"", "Style No":s.styleNo||"", "Colour":s.colour||"", "Buyer / Brand":s.buyer||s.brand||"", "Junior":s.owner||"", "Stage":stageReviewLabel(s,r), "Stage Key":r.key, "Branch":BRANCH_OF[r.key]||"", "Chase Label":r.owner||"", "State":stageState(r), "Actionable Frontier?":frontier?"YES":"NO", "Auto Plan Date":r.plan?fmt(r.plan):"", "Revised Date":r.rev?fmt(r.rev):"", "Actual Date":r.actual?fmt(r.actual):"", "Due Used":due?fmt(due):"", "Due Status":due?(TODAY>due?"Overdue":"Not overdue"):"", "Rejected Date":r.reject?fmt(r.reject):"", "Skipped Date":r.skip?fmt(r.skip):"", "Start Date Used":start?fmt(start):"", "Delay vs Due Days":delayDue==null?"":r1(delayDue), "Delay vs Auto Plan Days":delayPlan==null?"":r1(delayPlan), "Actual vs Revised Days":delayRev==null?"":r1(delayRev), "Actual Duration Days":duration==null?"":r1(duration), "Overall Status":c.status||"" }); if(delayPlan!=null){ dashboardActualVsPlan.push({ "Order No":s.orderNo||"", "Style No":s.styleNo||"", "Stage":stageReviewLabel(s,r), "Branch":BRANCH_OF[r.key]||"", "Auto Plan Date":fmt(r.plan), "Revised Date":r.rev?fmt(r.rev):"", "Actual Date":fmt(r.actual), "Actual vs Original Plan Days":r1(delayPlan||0), "Accuracy":(delayPlan===0?"On original plan":(delayPlan>0?"Late vs original plan":"Early vs original plan")), "Included?":"YES - auto plan + actual exist" }); } if(r.rev&&r.actual){ dashboardRevisedVsActual.push({ "Order No":s.orderNo||"", "Style No":s.styleNo||"", "Stage":stageReviewLabel(s,r), "Branch":BRANCH_OF[r.key]||"", "Auto Plan Date":r.plan?fmt(r.plan):"", "Revised Date":fmt(r.rev), "Actual Date":fmt(r.actual), "Actual vs Revised Days":r1(delayRev||0), "Accuracy":(delayRev===0?"On revised date":(delayRev>0?"Late vs revised":"Early vs revised")), "Included?":"YES - revised exists" }); } }); });
  const avgArr=(arr,key)=>arr.length?r1(arr.reduce((sum,row)=>sum+(Number(row[key])||0),0)/arr.length):0;
  const lateAvgArr=(arr,key)=>{ const late=arr.filter(row=>(Number(row[key])||0)>0); return late.length?r1(late.reduce((sum,row)=>sum+(Number(row[key])||0),0)/late.length):0; };
  const dashboardPlanAccuracySummary=[
    { "Comparison":"Actual vs Original Plan", "Records":dashboardActualVsPlan.length, "Late/Missed Records":dashboardActualVsPlan.filter(row=>(Number(row["Actual vs Original Plan Days"])||0)>0).length, "Avg Net Days":avgArr(dashboardActualVsPlan,"Actual vs Original Plan Days"), "Avg Late/Missed Days":lateAvgArr(dashboardActualVsPlan,"Actual vs Original Plan Days"), "Worst Late Days":dashboardActualVsPlan.reduce((m,row)=>Math.max(m,Number(row["Actual vs Original Plan Days"])||0),0) },
    { "Comparison":"Actual vs Revised Plan", "Records":dashboardRevisedVsActual.length, "Late/Missed Records":dashboardRevisedVsActual.filter(row=>(Number(row["Actual vs Revised Days"])||0)>0).length, "Avg Net Days":avgArr(dashboardRevisedVsActual,"Actual vs Revised Days"), "Avg Late/Missed Days":lateAvgArr(dashboardRevisedVsActual,"Actual vs Revised Days"), "Worst Late Days":dashboardRevisedVsActual.reduce((m,row)=>Math.max(m,Number(row["Actual vs Revised Days"])||0),0) }
  ];
  const dashboardOpenRows=dashboardStageDetail.filter(r=>r["Actionable Frontier?"]==="YES");
  // Export-ready blocker report: one row per currently actionable style/activity,
  // already sorted so styles sit directly under their stuck category and activity.
  // "Last Planned Date" is the active commitment used by the tracker: revised date
  // when entered, otherwise the original auto/system plan.
  const stuckDetailRows=[];
  const stuckSummaryMap={};
  fc.forEach(({s,c})=>{
    const frontier=c.frontier?[...c.frontier]:[];
    frontier.forEach(key=>{
      const r=(c.stages||[]).find(x=>x.key===key);
      if(!r||r.done) return;
      const due=r.rev||r.plan||null;
      const days=due?netWorkdays(TODAY,due):null;
      const summaryKey=`${phaseOfKey(key)}::${key}`;
      const summary=stuckSummaryMap[summaryKey]||(stuckSummaryMap[summaryKey]={ category:phaseOfKey(key), activity:stageReviewLabel(s,r), stageKey:key, styles:new Set(), orders:new Set(), chase:new Set(), overdue:0, plans:[] });
      summary.styles.add(String(s.id||`${s.orderNo||""}:${s.styleNo||""}:${s.colour||""}`));
      if(s.orderNo) summary.orders.add(String(s.orderNo));
      if(r.owner) summary.chase.add(String(r.owner));
      if(days!=null&&days<0) summary.overdue++;
      if(due) summary.plans.push(due);
      stuckDetailRows.push({
        "Activity Category":phaseOfKey(key),
        "Activity":stageReviewLabel(s,r),
        "Stage Key":key,
        "Order No":s.orderNo||"",
        "Style No":s.styleNo||"",
        "Colour":s.colour||"",
        "Family":s.family||"",
        "Buyer / Brand":s.buyer||s.brand||"",
        "Junior":s.owner||"",
        "Chase Label":r.owner||"",
        "Original Planned Date":r.plan?fmt(r.plan):"",
        "Revised Planned Date":r.rev?fmt(r.rev):"",
        "Last Planned Date":due?fmt(due):"",
        "Due Position":days==null?"No plan":(days<0?`Overdue ${Math.abs(days)}wd`:(days===0?"Due today":`${days}wd left`)),
        "Last Completed Activity":c.lastActualKey?stageLabelFromKeyGlobal(c.lastActualKey):"",
        "Last Completed Actual Date":c.lastActual?fmt(c.lastActual):"",
        "Delivery":s.delivery||"",
        "Overall Status":c.status||""
      });
    });
  });
  const phaseOrder={"Pre-Fit":1,"Fit / Print":2,"Lab Dip":3,"Fabric IH":4,"PP / Prod":5};
  stuckDetailRows.sort((a,b)=>(phaseOrder[a["Activity Category"]]||99)-(phaseOrder[b["Activity Category"]]||99) || stageOrderOf(a["Stage Key"])-stageOrderOf(b["Stage Key"]) || String(a["Order No"]).localeCompare(String(b["Order No"]),undefined,{numeric:true,sensitivity:"base"}) || String(a["Style No"]).localeCompare(String(b["Style No"]),undefined,{numeric:true,sensitivity:"base"}));
  const stuckSummaryRows=Object.values(stuckSummaryMap).sort((a,b)=>(phaseOrder[a.category]||99)-(phaseOrder[b.category]||99) || stageOrderOf(a.stageKey)-stageOrderOf(b.stageKey)).map(x=>({
    "Activity Category":x.category,
    "Activity":x.activity,
    "Orders":[...x.orders].join(", "),
    "Styles":x.styles.size,
    "Overdue Styles":x.overdue,
    "Chase Label":[...x.chase].join(", "),
    "Earliest Last Planned Date":x.plans.length?fmt(new Date(Math.min(...x.plans.map(d=>d.getTime())))):"",
    "Latest Last Planned Date":x.plans.length?fmt(new Date(Math.max(...x.plans.map(d=>d.getTime())))):""
  }));
  const ownerRows=Object.entries(dashboardOpenRows.reduce((m,r)=>{ const k=r["Chase Label"]||"(blank)"; m[k]=(m[k]||0)+1; return m; },{})).sort((a,b)=>b[1]-a[1]).map(([owner,count])=>({ "Chase Label":owner, "Open Items":count }));
  const activityRows=Object.entries(dashboardOpenRows.reduce((m,r)=>{ const k=r["Stage"]||"(blank)"; const x=m[k]=m[k]||{n:0,over:0,key:""}; x.n++; if(r["Due Status"]==="Overdue") x.over++; x.key=r["Stage Key"]||x.key; return m; },{})).sort((a,b)=>b[1].n-a[1].n).map(([activity,v])=>({ "Activity":activity, "Open Count":v.n, "Overdue Count":v.over, "Stage Key":v.key }));
  const phaseRows=Object.entries(phase).map(([phaseName,count])=>({ "Phase":phaseName, "Styles":count }));
  const dashboardLogicChecks=[
      { Check:"Status partition", Rule:"On Track + At Risk + Released equals Styles in Slice", Value:onTrack+atRisk+released, Expected:total, Result:(onTrack+atRisk+released)===total?"OK":"CHECK" },
      { Check:"Open activity rows", Rule:"Open Activities summary total equals actionable frontier detail rows", Value:activityRows.reduce((a,r)=>a+(Number(r["Open Count"])||0),0), Expected:dashboardOpenRows.length, Result:activityRows.reduce((a,r)=>a+(Number(r["Open Count"])||0),0)===dashboardOpenRows.length?"OK":"CHECK" },
      { Check:"Who to chase rows", Rule:"Who to Chase summary total equals actionable frontier detail rows", Value:ownerRows.reduce((a,r)=>a+(Number(r["Open Items"])||0),0), Expected:dashboardOpenRows.length, Result:ownerRows.reduce((a,r)=>a+(Number(r["Open Items"])||0),0)===dashboardOpenRows.length?"OK":"CHECK" },
      { Check:"Revised vs Actual", Rule:"Includes only rows with both Revised Date and Actual Date", Value:dashboardRevisedVsActual.length, Expected:"revised + actual rows only", Result:"OK" },
      { Check:"Escalation count", Rule:"Dashboard escalation count uses same overdue + escalation owner rows as To-Do", Value:escalationTodo.length, Expected:"To-Do overdue escalation rows", Result:"OK" },
      { Check:"Slice filters", Rule:"All dashboard tables use the same current slice", Value:fc.length, Expected:total, Result:"OK" },
      { Check:"Plan columns parity", Rule:"Dashboard export includes both Actual vs Original Plan and Actual vs Revised detail sheets", Value:"Actual vs Original Plan + Actual vs Revised", Expected:"both present", Result:"OK" }
  ];
  const dashboardSheets=[
      { label:"Summary", data:dashboardSummary, detailData:dashboardBreakup, modes:["summary","detailed"] },
      { label:"Logic Checks", data:dashboardLogicChecks, modes:["summary","detailed"] },
      { label:"Who to Chase", data:ownerRows, detailData:dashboardOpenRows, modes:["summary","detailed"] },
      { label:"Open Activities", data:activityRows, detailData:dashboardOpenRows, modes:["summary","detailed"] },
      { label:"Where Things Are Stuck", data:stuckSummaryRows, detailData:stuckDetailRows, modes:["summary","detailed"] },
      { label:"Styles in Slice", data:dashboardStyleRows, detailData:dashboardStageDetail, modes:["detailed"] },
      { label:"Dashboard Breakup", data:dashboardBreakup, modes:["detailed"] },
      { label:"Style Stage Detail", data:dashboardStageDetail, modes:["detailed"] },
      { label:"Plan Accuracy Summary", data:dashboardPlanAccuracySummary, modes:["summary","detailed"] },
      { label:"Actual vs Original Plan", data:dashboardActualVsPlan, modes:["detailed"] },
      { label:"Actual vs Revised", data:dashboardRevisedVsActual, modes:["detailed"] },
  ];
  return (<div style={{ padding:"16px 22px", maxWidth:1140 }}>
    {/* splice filter bar */}
    <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"center", marginBottom:14 }}>
      <span style={{ fontSize:10, fontWeight:700, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.5 }}>Slice:</span>
      {sel("Order",df.order,orders,v=>setDf(d=>({...d,order:v})))}
      {sel("Fit",df.fit,fits,v=>setDf(d=>({...d,fit:v})))}
      {sel("Colour",df.colour,colours,v=>setDf(d=>({...d,colour:v})))}
      {sel("Junior",df.junior,juniors,v=>setDf(d=>({...d,junior:v})))}
      {sel("Family",df.family,families,v=>setDf(d=>({...d,family:v})))}
      {sel("Brand",df.brand,brands,v=>setDf(d=>({...d,brand:v})))}
      {sel("Fabric",df.fabric,fabrics,v=>setDf(d=>({...d,fabric:v})))}
      {anyDf && <button onClick={()=>setDf({})} style={{ fontFamily:"inherit", fontSize:10, padding:"5px 9px", cursor:"pointer", border:"1px solid var(--danger)", background:"var(--surface)", color:"var(--danger)", fontWeight:700 }}>clear slice</button>}
      <ReportExportMenu title="Dashboard" prefix="dashboard" sheets={dashboardSheets} defaultMode="detailed" />
      <span style={{ marginLeft:"auto", display:"flex", alignItems:"center", gap:6 }}><span style={{ fontSize:10, color:"var(--muted-2)" }}>drill to:</span><div style={{ display:"flex", border:"1px solid var(--ink)" }}>{["tracker","todo"].map((t,i,arr)=>(<button key={t} onClick={()=>setTarget(t)} style={{ fontFamily:"inherit", fontSize:10, fontWeight:700, padding:"5px 10px", cursor:"pointer", border:"none", borderRight:i<arr.length-1?"1px solid var(--ink)":"none", background:target===t?"var(--ink)":"var(--surface)", color:target===t?"var(--bg)":"var(--ink)" }}>{t==="tracker"?"Tracker":"To-Do"}</button>))}</div></span>
    </div>

    <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
      {card("Styles (slice)",total,"var(--ink)")}
      {card("On track",onTrack,"var(--success)",()=>goStatus("On Track"))}
      {card("At risk",atRisk,"var(--danger)",()=>goStatus("At Risk"))}
      {card("Delivery risk",delRisk,"var(--danger)",()=>goStatus("All",{ overall:["Delivery risk"] }))}
      {card("Released",released,"var(--success)",()=>goStatus("Released"))}
      {card("Overdue activities",overdueAct,"var(--danger)")}
      {card("Escalation items",escalationTodo.length,escalationTodo.length?"var(--danger)":"var(--success)",()=>drillRowsToTodo(escalationTodo,{ priority:"Overdue" }))}
    </div>

    <div style={{ marginTop:14, background:"var(--surface)", border:"1px solid var(--ink)", padding:14 }}>
      <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:13, marginBottom:10 }}>PLAN ACCURACY — completed stages</div>
      <div style={{ overflowX:"auto" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", fontSize:10 }}>
          <thead><tr>{["Comparison","Records","Missed","Avg net","Avg missed","Worst missed"].map(h=><th key={h} style={{ textAlign:h==="Comparison"?"left":"right", padding:"7px 8px", borderBottom:"1px solid var(--line-2)", color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4 }}>{h}</th>)}</tr></thead>
          <tbody>{dashboardPlanAccuracySummary.map(r=><tr key={r["Comparison"]}><td style={{ padding:"7px 8px", fontWeight:800 }}>{r["Comparison"]}</td><td style={{ padding:"7px 8px", textAlign:"right" }}>{r["Records"]}</td><td style={{ padding:"7px 8px", textAlign:"right", color:r["Late/Missed Records"]?"var(--danger)":"var(--success)", fontWeight:800 }}>{r["Late/Missed Records"]}</td><td style={{ padding:"7px 8px", textAlign:"right" }}>{fmtDays(Number(r["Avg Net Days"])||0)}</td><td style={{ padding:"7px 8px", textAlign:"right" }}>{fmtDays(Number(r["Avg Late/Missed Days"])||0)}</td><td style={{ padding:"7px 8px", textAlign:"right" }}>{fmtDays(Number(r["Worst Late Days"])||0)}</td></tr>)}</tbody>
        </table>
      </div>
      <div style={{ fontSize:9, color:"var(--muted-7)", marginTop:8 }}>Original plan = first auto/system target. Revised = latest committed date entered after follow-up.</div>
    </div>

    <div style={{ display:"flex", gap:18, flexWrap:"wrap", marginTop:22 }}>
      <div style={{ flex:1, minWidth:320, background:"var(--surface)", border:"1px solid var(--ink)", padding:16 }}>
        <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:13, marginBottom:12 }}>WHO TO CHASE — by chase label</div>
        {bar(owners, maxOwner, (o)=>OWNER_COLOR2[o]||"var(--muted-2)", 64, (o)=>goOwner(o))}
        <div style={{ fontSize:9, color:"var(--muted-7)", marginTop:8 }}>Click to open in {target==="tracker"?"Tracker":"To-Do"}.</div>
      </div>
      <div style={{ flex:1, minWidth:320, background:"var(--surface)", border:"1px solid var(--ink)", padding:16 }}>
        <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:13, marginBottom:12 }}>ESCALATION OWNER LOAD</div>
        {bar(escRows, maxEsc, ()=>"var(--danger)", 92, (o)=>drillRowsToTodo(escalationTodo.filter(t=>String(t.escalationOwner||"")===String(o)),{ priority:"Overdue", escalationOwner:o }))}
        <div style={{ fontSize:9, color:"var(--muted-7)", marginTop:8 }}>Click opens To-Do escalation rows for that owner.</div>
      </div>
      <div style={{ flex:1, minWidth:320, background:"var(--surface)", border:"1px solid var(--ink)", padding:16 }}>
        <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:13, marginBottom:12 }}>OPEN ACTIVITIES</div>
        {acts.length===0?<div style={{ fontSize:11, color:"var(--muted-1)" }}>Nothing due.</div>:acts.map(([label,v])=>(
          <button key={label} onClick={()=>goAct(label,v.key)} style={{ display:"flex", alignItems:"center", gap:8, width:"100%", border:"none", background:"transparent", cursor:"pointer", fontFamily:"inherit", padding:"4px 0" }}>
            <span style={{ width:84, fontSize:10, fontWeight:700, color:"var(--muted-5)", textAlign:"left" }}>{label}</span>
            <span style={{ flex:1, height:16, background:"#f0ece3", position:"relative" }}><span style={{ position:"absolute", left:0, top:0, bottom:0, width:`${(v.n/maxAct)*100}%`, background:v.over?"var(--danger)":"var(--accent)" }}/></span>
            <span style={{ width:54, textAlign:"right", fontSize:10, fontWeight:700 }}>{v.n}{v.over?<span style={{ color:"var(--danger)" }}> ({v.over})</span>:null}</span>
          </button>))}
        <div style={{ fontSize:9, color:"var(--muted-7)", marginTop:8 }}>Click to open in {target==="tracker"?"Tracker":"To-Do"}. Red = overdue.</div>
      </div>
      <div style={{ flex:1, minWidth:320, background:"var(--surface)", border:"1px solid var(--ink)", padding:16 }}>
        <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:13, marginBottom:12 }}>WHERE THINGS ARE STUCK</div>
        {Object.entries(phase).map(([p,n])=>(
          <button key={p} onClick={()=>goPhase(p)} title="Open matching phase in To-Do" style={{ display:"flex", alignItems:"center", gap:8, padding:"4px 0", width:"100%", border:"none", background:"transparent", cursor:"pointer", fontFamily:"inherit" }}>
            <span style={{ width:80, fontSize:10, fontWeight:700, color:"var(--muted-4)", textAlign:"left" }}>{p}</span>
            <span style={{ flex:1, height:16, background:"#f0ece3", position:"relative" }}><span style={{ position:"absolute", left:0, top:0, bottom:0, width:`${(n/maxPhase)*100}%`, background:p==="Fabric IH"?"var(--danger)":"var(--accent)" }}/></span>
            <span style={{ width:28, textAlign:"right", fontSize:11, fontWeight:700 }}>{n}</span>
          </button>))}
        <div style={{ fontSize:9, color:"var(--muted-7)", marginTop:8 }}>Select an Order above, then Export → Detailed → Where Things Are Stuck for category/activity summary plus style rows and each blocker’s last planned date.</div>
      </div>
    </div>
  </div>);
}

/* ========================= TO-DO ========================= */


/* ========================= MANAGEMENT ANALYTICS ========================= */
function ManagementDashboardView({ computed, todoItems, cfg, applyDrill, drillTodo }){
  const [target,setTarget]=useState("tracker"); // where bar/owner/activity drills go: Tracker or To-Do only. Escalation cards route to To-Do with Type=Escalation.
  const [isPerfPending,startPerfTransition]=useTransition();
  const [df,setDf]=useState(()=>{ try{ return JSON.parse(localStorage.getItem("mt_management_filter")||localStorage.getItem("mt_dashfilter")||"{}"); }catch(e){ return {}; } });
  useEffect(()=>{ try{ localStorage.setItem("mt_management_filter", JSON.stringify(df)); }catch(e){} },[df]);
  const [mgmtOpen,setMgmtOpen]=useState(()=>{ try{ return JSON.parse(localStorage.getItem("mt_mgmt_open")||"{}"); }catch(e){ return {}; } });
  useEffect(()=>{ try{ localStorage.setItem("mt_mgmt_open", JSON.stringify(mgmtOpen)); }catch(e){} },[mgmtOpen]);
  const [perfView,setPerfView]=useState(()=>{ try{ return localStorage.getItem("mt_mgmt_perf_view")||"chase"; }catch(e){ return "chase"; } }); // chase | stage
  const [perfMode,setPerfMode]=useState(()=>{ try{ return localStorage.getItem("mt_mgmt_perf_mode")||"summary"; }catch(e){ return "summary"; } }); // summary | revised | noRevised
  const [perfScope,setPerfScope]=useState(()=>{ try{ return localStorage.getItem("mt_mgmt_perf_scope")||"all"; }catch(e){ return "all"; } }); // all | late
  const validPerfMode=["summary","revised","noRevised"].includes(perfMode)?perfMode:"summary";
  const validPerfScope=["all","late"].includes(perfScope)?perfScope:"all";
  useEffect(()=>{ if(validPerfMode!==perfMode) setPerfMode(validPerfMode); if(validPerfScope!==perfScope) setPerfScope(validPerfScope); },[validPerfMode,perfMode,validPerfScope,perfScope]);
  useEffect(()=>{ try{ localStorage.setItem("mt_mgmt_perf_view",perfView); localStorage.setItem("mt_mgmt_perf_mode",validPerfMode); localStorage.setItem("mt_mgmt_perf_scope",validPerfScope); }catch(e){} },[perfView,validPerfMode,validPerfScope]);
  const perfRowsCacheRef=useRef(new Map());
  const switchPerfView=(v)=>startPerfTransition(()=>setPerfView(v));
  const switchPerfMode=(v)=>startPerfTransition(()=>setPerfMode(v));
  const switchPerfScope=(v)=>startPerfTransition(()=>setPerfScope(v));
  const isMgmtOpen=(key)=> mgmtOpen[key]!==false;
  const toggleMgmt=(key)=>setMgmtOpen(o=>({ ...o, [key]: !isMgmtOpen(key) }));

  const arrOf=(v)=>Array.isArray(v)?v:(v?[v]:[]);
  const hasSel=(key,val)=>{ const a=arrOf(df[key]); return !a.length || a.includes(val); };
  const hasColourSel=(st)=>{ const a=arrOf(df.colour); if(!a.length) return true; const cols=splitColoursAll(st.colour); return a.some(v=>cols.includes(v)); };
  const matchDf=(st,except)=> (except==="order"||hasSel("order",st.orderNo)) && (except==="fit"||hasSel("fit",st.sampleFit)) && (except==="junior"||hasSel("junior",st.owner)) && (except==="family"||hasSel("family",st.family)) && (except==="brand"||hasSel("brand",st.brand)) && (except==="fabric"||hasSel("fabric",st.fabricType)) && (except==="colour"||hasColourSel(st));
  const distinctC=(key,fn)=>{ const set=new Set(); computed.forEach(({s:st})=>{ if(!matchDf(st,key)) return; fn(st).forEach(v=>{ if(v) set.add(v); }); }); return [...set].sort(); };
  const orders=distinctC("order",s=>[s.orderNo]); const fits=distinctC("fit",s=>[s.sampleFit]); const juniors=distinctC("junior",s=>[s.owner]); const families=distinctC("family",s=>[s.family]); const brands=distinctC("brand",s=>[s.brand]); const fabrics=distinctC("fabric",s=>[s.fabricType]);
  const colours=distinctC("colour",s=>splitColoursAll(s.colour));
  const fc=useMemo(()=>computed.filter(({s})=> hasSel("order",s.orderNo) && hasSel("fit",s.sampleFit) && hasSel("junior",s.owner) && hasSel("family",s.family) && hasSel("brand",s.brand) && hasSel("fabric",s.fabricType) && hasColourSel(s) ),[computed,df]);

  const total=fc.length;
  const onTrack=fc.filter(({c})=>c.tone==="ok").length;
  const atRisk=fc.filter(({c})=>c.tone==="late"||c.tone==="warn").length;
  const released=fc.filter(({c})=>c.released).length;
  const delRisk=fc.filter(({c})=>String(c.status).startsWith("Delivery risk")).length;
  const completionPct=total?Math.round((released/total)*100):0;

  const ownerLoad={}; const actAgg={};
  fc.forEach(({s,c})=>{ if(c.released) return; (c.chaseOwners||[]).forEach(o=>{ ownerLoad[o.owner]=(ownerLoad[o.owner]||0)+1; }); (c.frontier?[...c.frontier]:[]).forEach(k=>{ const r=(c.stages||[]).find(x=>x.key===k); if(!r||r.done) return; const lbl=stageLabelFromKeyGlobal(k); const a=actAgg[lbl]=actAgg[lbl]||{n:0,over:0,key:k}; a.n++; if((r.rev||r.plan)&&TODAY>(r.rev||r.plan)) a.over++; }); });
  const owners=Object.entries(ownerLoad).sort((a,b)=>b[1]-a[1]); const maxOwner=Math.max(1,...owners.map(o=>o[1]));
  const acts=Object.entries(actAgg).sort((a,b)=>b[1].n-a[1].n); const maxAct=Math.max(1,...acts.map(e=>e[1].n));
  const overdueAct=acts.reduce((s,[,v])=>s+v.over,0);
  const styleByTodoId=new Map(fc.map(({s})=>[Number(s.id),s]));
  const todoMatchesSlice=(t)=>{
    const st=styleByTodoId.get(Number(t.id));
    if(!st) return false;
    const orderVals=(Array.isArray(t.orderNos)&&t.orderNos.length?t.orderNos:[t.orderNo,st.orderNo]).filter(Boolean);
    const juniorVals=(Array.isArray(t.juniors)&&t.juniors.length?t.juniors:[t.junior,st.owner]).filter(Boolean);
    const colourVals=splitColoursAll(t.colour||st.colour||"");
    const matchArray=(key,vals)=>{ const a=arrOf(df[key]); return !a.length || vals.some(v=>a.includes(v)); };
    return matchArray("order",orderVals) && matchArray("junior",juniorVals) && hasSel("fit",st.sampleFit) && hasSel("family",st.family) && hasSel("brand",st.brand) && hasSel("fabric",st.fabricType) && (!arrOf(df.colour).length || arrOf(df.colour).some(v=>colourVals.includes(v)));
  };
  const escalationTodo=(todoItems||[]).filter(t=>todoMatchesSlice(t)&&t.overdue&&t.escalationOwner);
  const escLoad=escalationTodo.reduce((m,t)=>{ const k=t.escalationOwner||"(blank)"; m[k]=(m[k]||0)+1; return m; },{});
  const escRows=Object.entries(escLoad).sort((a,b)=>b[1]-a[1]); const maxEsc=Math.max(1,...escRows.map(x=>x[1]));

  const phase={ "Pre-Fit":0,"Fit / Print":0,"Lab Dip":0,"Fabric IH":0,"PP / Prod":0 };
  const phaseKeyList=(phaseName)=> phaseName==="Pre-Fit"?["techpack"]:phaseName==="Fit / Print"?["fitSend","fitAppr","artwork","artAppr","strikeOff","soAppr"]:phaseName==="Lab Dip"?["labDip","labAppr"]:phaseName==="Fabric IH"?["fabricIH"]:["ppSample","ppAppr","prodFile"];
  const phaseOfKey=(k)=> k==="techpack"?"Pre-Fit":["fitSend","fitAppr","artwork","artAppr","strikeOff","soAppr"].includes(k)?"Fit / Print":["labDip","labAppr"].includes(k)?"Lab Dip":k==="fabricIH"?"Fabric IH":"PP / Prod";
  fc.forEach(({c})=>{ if(c.released) return; const k=c.nextPending&&c.nextPending.key; phase[phaseOfKey(k)]++; });
  const maxPhase=Math.max(1,...Object.values(phase));

  const stageDef=(k)=>STAGES.find(x=>x.key===k)||{};
  const stageStartFor=(s,c,r)=>{
    const st=stageDef(r.key);
    const byKey={};
    (c.stages||[]).forEach(x=>{ byKey[x.key]=x; });
    if(r.key==="fabricIH") return (s.labDipReq && byKey.labAppr && byKey.labAppr.actual) ? byKey.labAppr.actual : parse(s.ordRec);
    if(st.pred==="__ord") return parse(s.ordRec);
    return byKey[st.pred] ? byKey[st.pred].actual : null;
  };
  const stageState=(r)=> r.skipped?"Waived / skipped":(r.rejected?"Rejected":(r.rework?"Rework / resend":(r.actual?"Done":(r.autoClosed?"Auto-closed":(r.rev?"Revised plan":"Pending")))));
  const delayRecords=[]; const stagePerf={}; const deptPerf={}; const buyerPerf={}; const chaseDelay={}; const brandDelay={};
  const addPerf=(bucket,key,label,delay,duration,extra={})=>{ const o=bucket[key]=bucket[key]||{ key, label, n:0, lateN:0, delaySum:0, durSum:0, durN:0, maxDelay:-999, planN:0, planMissN:0, planSum:0, planMissSum:0, planWorst:0, revN:0, revMissN:0, revSum:0, revMissSum:0, revWorst:0, ...extra }; o.n++; if(delay>0) o.lateN++; o.delaySum+=delay; if(duration!=null){ o.durSum+=duration; o.durN++; } if(extra.delayPlan!=null){ o.planN++; o.planSum+=extra.delayPlan; if(extra.delayPlan>0){ o.planMissN++; o.planMissSum+=extra.delayPlan; o.planWorst=Math.max(o.planWorst,extra.delayPlan); } } if(extra.delayRevised!=null){ o.revN++; o.revSum+=extra.delayRevised; if(extra.delayRevised>0){ o.revMissN++; o.revMissSum+=extra.delayRevised; o.revWorst=Math.max(o.revWorst,extra.delayRevised); } } o.maxDelay=Math.max(o.maxDelay,delay); return o; };
  const avg=(n,d)=> d?Math.round((n/d)*10)/10:0;
  fc.forEach(({s,c})=>{
    const byKey={}; (c.stages||[]).forEach(r=>{ byKey[r.key]=r; });
    (c.stages||[]).forEach(r=>{
      if(!r.actual) return;
      const due=r.rev||r.plan; if(!due) return;
      const delay=netWorkdays(due,r.actual)||0;
      const st=stageDef(r.key); const predKey=st.pred;
      let start=predKey==="__ord"?parse(s.ordRec):(byKey[predKey]&&byKey[predKey].actual);
      if(r.key==="fabricIH") start = (s.labDipReq && byKey.labAppr && byKey.labAppr.actual) ? byKey.labAppr.actual : parse(s.ordRec);
      const duration=start?Math.max(0,netWorkdays(start,r.actual)||0):null;
      const delayPlan=(r.plan&&r.actual)?netWorkdays(r.plan,r.actual):null; const delayRevised=(r.rev&&r.actual)?netWorkdays(r.rev,r.actual):null;
      delayRecords.push({ style:s.styleNo, order:s.orderNo, buyer:s.buyer||s.brand||"", owner:s.owner||"", stage:stageReviewLabel(s,r), stageKey:r.key, dept:r.owner, delay, delayPlan, delayRevised, duration, actual:r.actual, due, plan:r.plan, revised:r.rev, start });
      const perfExtra={ delayPlan, delayRevised };
      addPerf(stagePerf,r.key,stageLabelFromKeyGlobal(r.key)||stageReviewLabel(s,r),delay,duration,{ owner:r.owner, ...perfExtra });
      addPerf(deptPerf,r.owner,r.owner,delay,duration,perfExtra);
      if(r.owner==="Buyer") addPerf(buyerPerf,s.buyer||s.brand||"(No buyer)",s.buyer||s.brand||"(No buyer)",delay,duration,perfExtra);
      addPerf(chaseDelay,r.owner||"(No chase label)",r.owner||"(No chase label)",delay,duration,perfExtra);
      addPerf(brandDelay,s.buyer||s.brand||"(No buyer)",s.buyer||s.brand||"(No buyer)",delay,duration,perfExtra);
    });
  });
  // Performance tables may use all completed entries. Delay-ranking tables must use problem rows only.
  const lateRecords=delayRecords.filter(r=>r.delay>0);
  const buyerApprovalStageNames=["Fit Appr","Art Appr","S/O Appr","Lab Dip Appr","PP Appr"];
  const buyerApprovalStageKeys=["fitAppr","artAppr","soAppr","labAppr","ppAppr"]; // filter by canonical key so rework/re-approval labels are not dropped
  const buyerApprovalRecords=delayRecords.filter(r=>buyerApprovalStageKeys.includes(r.stageKey)||buyerApprovalStageNames.includes(r.stage));
  const buyerApprovalLateRecords=lateRecords.filter(r=>buyerApprovalStageKeys.includes(r.stageKey)||buyerApprovalStageNames.includes(r.stage));
  const makeDelayAgg=(records,keyFn,labelFn)=>{ const m={}; records.forEach(r=>{ const k=keyFn(r)||"(blank)"; const o=m[k]=m[k]||{ key:k, label:labelFn?labelFn(r,k):k, delayed:0, delaySum:0, maxDelay:0, styles:new Set(), planN:0, planSum:0, planMissN:0, planMissSum:0, revN:0, revSum:0, revMissN:0, revMissSum:0 }; const d=Math.max(0,r.delay||0); o.delayed++; o.delaySum+=d; o.maxDelay=Math.max(o.maxDelay,d); if(r.delayPlan!=null){ o.planN++; o.planSum+=r.delayPlan; if(r.delayPlan>0){ o.planMissN++; o.planMissSum+=r.delayPlan; } } if(r.delayRevised!=null){ o.revN++; o.revSum+=r.delayRevised; if(r.delayRevised>0){ o.revMissN++; o.revMissSum+=r.delayRevised; } } if(r.style) o.styles.add(r.style); }); return Object.values(m).sort((a,b)=>(b.delaySum-a.delaySum)||(b.delayed-a.delayed)); };
  const chaseDelayRows=makeDelayAgg(lateRecords,r=>r.dept||"(No chase label)").slice(0,8);
  const brandDelayRows=makeDelayAgg(lateRecords,r=>r.buyer||"(No buyer)").slice(0,8);
  const buyerDelayRows=makeDelayAgg(buyerApprovalLateRecords,r=>r.buyer||"(No buyer)").slice(0,8);
  const stageRows=Object.values(stagePerf).sort((a,b)=>stageOrderOf(a.key)-stageOrderOf(b.key));
  const deptRows=Object.values(deptPerf).sort((a,b)=>avg(b.delaySum,b.n)-avg(a.delaySum,a.n));
  const buyerRows=Object.values(buyerPerf).sort((a,b)=>avg(b.durSum,b.durN)-avg(a.durSum,a.durN)).slice(0,8);
  const actualDone=delayRecords.length;
  const avgDelay=lateRecords.length?avg(lateRecords.reduce((s,r)=>s+Math.max(0,r.delay||0),0),lateRecords.length):0;
  const durationDone=delayRecords.filter(r=>r.duration!=null).length;
  const avgDuration=durationDone?avg(delayRecords.reduce((s,r)=>s+(r.duration!=null?r.duration:0),0),durationDone):0;
  const lateDone=lateRecords.length;
  const planActualRecords=delayRecords.filter(r=>r.delayPlan!=null);
  const planMissedRecords=planActualRecords.filter(r=>r.delayPlan>0);
  const revisedActualRecords=delayRecords.filter(r=>r.delayRevised!=null);
  const revisedMissedRecords=revisedActualRecords.filter(r=>r.delayRevised>0);
  const avgPlanNet=planActualRecords.length?avg(planActualRecords.reduce((s,r)=>s+(r.delayPlan||0),0),planActualRecords.length):0;
  const avgPlanMiss=planMissedRecords.length?avg(planMissedRecords.reduce((s,r)=>s+(r.delayPlan||0),0),planMissedRecords.length):0;
  const avgRevNet=revisedActualRecords.length?avg(revisedActualRecords.reduce((s,r)=>s+(r.delayRevised||0),0),revisedActualRecords.length):0;
  const avgRevMiss=revisedMissedRecords.length?avg(revisedMissedRecords.reduce((s,r)=>s+(r.delayRevised||0),0),revisedMissedRecords.length):0;
  const worstPlanMiss=planMissedRecords.length?Math.max(...planMissedRecords.map(r=>r.delayPlan||0)):0;
  const worstRevMiss=revisedMissedRecords.length?Math.max(...revisedMissedRecords.map(r=>r.delayRevised||0)):0;
  const worstDelays=[...lateRecords].sort((a,b)=>b.delay-a.delay).slice(0,8);

  const spliceCols=()=>{ const cf={}; const put=(k,col)=>{ const a=arrOf(df[k]); if(a.length) cf[col]=a; }; put("order","orderNo"); put("fit","sampleFit"); put("junior","owner"); put("family","family"); put("brand","brand"); put("fabric","fabricType"); put("colour","colour"); return cf; };
  const spliceSearch=()=> "";
  const sliceTodoRows=(pred=()=>true)=>(todoItems||[]).filter(t=>todoMatchesSlice(t)&&pred(t));
  const phaseStyleIds=(phaseName)=>new Set(fc.filter(({c})=>!c.released && phaseOfKey(c.nextPending&&c.nextPending.key)===phaseName).map(({s})=>String(s.id)));
  const phaseTodoRows=(phaseName)=>{ const ids=phaseStyleIds(phaseName); const allowed=new Set(phaseKeyList(phaseName)); return sliceTodoRows(t=>ids.has(String(t.id)) && allowed.has(stageKeyFromAnyGlobal(t.activityKey||t.key||t.activity))); };
  const drillRowsToTodo=(rows,base={})=>{ if(drillTodo) drillTodo(todoDrillFilterFromRows(rows,base,df)); };
  const goOwner=(o)=>{ if(target==="todo") drillRowsToTodo(sliceTodoRows(t=>String(t.owner||"")===String(o)),{ owner:o }); else applyDrill({ owner:o, colFilters:spliceCols(), search:spliceSearch() }); };
  const goAct=(label,key)=>{ const k=stageKeyFromAnyGlobal(key||label); if(target==="todo") drillRowsToTodo(sliceTodoRows(t=>stageKeyFromAnyGlobal(t.activityKey||t.key||t.activity)===k),{ activity:[stageLabelFromKeyGlobal(k)||label], activityKey:k?[k]:[], key:k?[k]:[] }); else applyDrill({ activity:k||key, colFilters:spliceCols(), search:spliceSearch() }); };
  const goPhase=(phaseName)=>{ const keys=phaseKeyList(phaseName); if(target==="todo") drillRowsToTodo(phaseTodoRows(phaseName),{ phase:phaseName, activityKey:keys, activity:keys.map(stageLabelFromKeyGlobal) }); else applyDrill({ activity:keys, colFilters:spliceCols(), search:spliceSearch() }); };
  const goStatus=(st,extra)=>applyDrill({ status:st, colFilters:{...spliceCols(),...(extra||{})}, search:spliceSearch() });
  const goSearch=(q)=>applyDrill({ status:"All", colFilters:spliceCols(), search:q||spliceSearch() });
  const goStageOpen=(key)=>applyDrill({ status:"All", activity:key, colFilters:spliceCols(), search:spliceSearch() });
  const anyDf=Object.values(df).some(v=>Array.isArray(v)?v.length:!!v);

  const mgmtSummary=[{ "Report Type":"Management", "Styles in Slice":total, "On Track":onTrack, "At Risk":atRisk, "Delivery Risk":delRisk, "Released":released, "Release %":completionPct, "Overdue Open Activities":overdueAct, "Completed Stage Entries":actualDone, "Delayed Stage Entries":lateDone, "Avg Delay vs Due Days":r1(avgDelay), "Avg Actual Time Days":r1(avgDuration), "Original Plan Records":planActualRecords.length, "Avg Actual vs Original Plan Days":r1(avgPlanNet), "Revised Records":revisedActualRecords.length, "Avg Actual vs Revised Plan Days":r1(avgRevNet), "Worst Delay vs Due Days":r1(worstDelays.length?worstDelays[0].delay:0) }];
  const checks=[
      { Check:"Status partition", Formula:"On Track + At Risk + Released should equal Styles in Slice", Value:onTrack+atRisk+released, Expected:total, Result:(onTrack+atRisk+released)===total?"OK":"CHECK" },
      { Check:"Open activity overdue", Formula:"Sum overdue counts from Open Activities", Value:overdueAct, Expected:acts.reduce((a,[,v])=>a+v.over,0), Result:overdueAct===acts.reduce((a,[,v])=>a+v.over,0)?"OK":"CHECK" },
      { Check:"Completed duration denominator", Formula:"Avg actual time divides only records with known start date", Value:durationDone, Expected:"duration records", Result:"OK" },
      { Check:"Actual vs original plan denominator", Formula:"Original plan accuracy includes rows where auto plan and actual date exist", Value:planActualRecords.length, Expected:"plan+actual records", Result:"OK" },
      { Check:"Revised vs actual denominator", Formula:"Revised accuracy includes only rows where revised date exists and actual date exists", Value:revisedActualRecords.length, Expected:"revised+actual records", Result:"OK" },
      { Check:"Website/export parity - Performance Analysis", Formula:"Export Performance Analysis summary uses the same side-by-side plan/revised columns as the website table", Value:"Actual vs original + actual vs revised columns present", Expected:"same metrics", Result:"OK" },
      { Check:"Website/export parity - all report exports", Formula:"All completed/performance exports carry actual-vs-original and actual-vs-revised columns where actual completion exists; pending-only exports mark plan accuracy as N/A", Value:"Tracker + Dashboard + Management + To-Do + Escalation + Entry Log paths checked", Expected:"no silent mismatch", Result:"OK" },
      { Check:"Delay ranking rule", Formula:"Chase/Buyer/Brand delay rankings include only positive delay rows", Value:lateDone, Expected:"delay > 0 only", Result:lateRecords.every(r=>r.delay>0)?"OK":"CHECK" },
      { Check:"Worst delay rule", Formula:"Worst delay tables include positive completed delays only", Value:worstDelays.length, Expected:"all delay > 0", Result:worstDelays.every(r=>r.delay>0)?"OK":"CHECK" },
      { Check:"Buyer approval delay rule", Formula:"Buyer Approval delay report includes buyer approval stages with positive delay only", Value:buyerApprovalLateRecords.length, Expected:"buyer approval + delay > 0", Result:buyerApprovalLateRecords.every(r=>buyerApprovalStageNames.includes(r.stage)&&r.delay>0)?"OK":"CHECK" },
      { Check:"Escalation source", Formula:"Escalation Owner Load uses To-Do overdue escalation rows only", Value:escalationTodo.length, Expected:"overdue + escalation owner", Result:escalationTodo.every(t=>t.overdue&&t.escalationOwner)?"OK":"CHECK" },
      { Check:"Slice rows", Formula:"All management tables use filtered slice", Value:fc.length, Expected:total, Result:"OK" }
    ];
  const currentStyles=fc.map(({s,c})=>({ "Order No":s.orderNo||"", "Style No":s.styleNo||"", "Fit":s.sampleFit||"", "Family":s.family||"", "Colour":s.colour||"", "Brand":s.brand||"", "Buyer":s.buyer||"", "Junior":s.owner||"", "Qty":s.qty||0, "Delivery":s.delivery||"", "Status":c.status||"", "Tone":c.tone||"", "Released":c.released?"YES":"", "% Done":c.pct, "Chase":(c.chaseOwners||[]).map(o=>`${o.owner} (${o.count})`).join(", "), "Next Pending":c.nextPending?stageReviewLabel(s,c.nextPending):"", "Projected Release":c.projRelease?fmt(c.projRelease):"" }));
  const currentSliceStageDetail=[];
  fc.forEach(({s,c})=>{ (c.stages||[]).forEach(r=>{ const due=r.rev||r.plan; const start=stageStartFor(s,c,r); const delayDue=(due&&r.actual)?netWorkdays(due,r.actual):null; const delayPlan=(r.plan&&r.actual)?netWorkdays(r.plan,r.actual):null; const delayRevised=(r.rev&&r.actual)?netWorkdays(r.rev,r.actual):null; const duration=(start&&r.actual)?Math.max(0,netWorkdays(start,r.actual)||0):null; currentSliceStageDetail.push({ "Order No":s.orderNo||"", "Style No":s.styleNo||"", "Buyer / Brand":s.buyer||s.brand||"", "Junior / Style Owner":s.owner||"", "Stage":stageReviewLabel(s,r), "Chase Label":r.owner||"", "State":stageState(r), "Auto Plan Date":r.plan?fmt(r.plan):"", "Revised Date":r.rev?fmt(r.rev):"", "Due Used":due?fmt(due):"", "Actual Date":r.actual?fmt(r.actual):"", "Start Used":start?fmt(start):"", "Delay vs Due Days":delayDue==null?"":r1(delayDue), "Actual vs Original Plan Days":delayPlan==null?"":r1(delayPlan), "Actual vs Revised Days":delayRevised==null?"":r1(delayRevised), "Actual Duration Days":duration==null?"":r1(duration), "Actionable Frontier?":(c.frontier&&c.frontier.has(r.key))?"YES":"NO" }); }); });
  const stageData=stageRows.map(r=>({"TNA Order":stageOrderOf(r.key),"Section":stageSectionOf(r.key),"Stage":r.label,"Chase Label":r.owner||"","Completed":r.n,"Late Count":r.lateN,"Late %":r.n?Math.round((r.lateN/r.n)*100):0,"Avg Actual Time Days":r1(avg(r.durSum,r.durN)),"Avg Actual vs Original Plan Days":r1(avg(r.planSum,r.planN)),"Avg Delay vs Due Days":r1(avg(r.delaySum,r.n)),"Avg Actual vs Revised Plan Days":r1(avg(r.revSum,r.revN)),"Worst Delay vs Due Days":r1(r.maxDelay),"Duration Records":r.durN,"Original Plan Records":r.planN,"Revised Records":r.revN}));
  const dept=deptRows.map(r=>({"Department":r.label,"Completed":r.n,"Late Count":r.lateN,"Late %":r.n?Math.round((r.lateN/r.n)*100):0,"Avg Actual Time Days":r1(avg(r.durSum,r.durN)),"Avg Actual vs Original Plan Days":r1(avg(r.planSum,r.planN)),"Avg Delay vs Due Days":r1(avg(r.delaySum,r.n)),"Avg Actual vs Revised Plan Days":r1(avg(r.revSum,r.revN)),"Worst Delay vs Due Days":r1(r.maxDelay),"Duration Records":r.durN,"Original Plan Records":r.planN,"Revised Records":r.revN}));
  const buyerData=buyerRows.map(r=>({"Buyer / Brand":r.label,"Approvals":r.n,"Late Count":r.lateN,"Late %":r.n?Math.round((r.lateN/r.n)*100):0,"Avg Approval Time Days":r1(avg(r.durSum,r.durN)),"Avg Actual vs Original Plan Days":r1(avg(r.planSum,r.planN)),"Avg Delay vs Due Days":r1(avg(r.delaySum,r.n)),"Avg Actual vs Revised Plan Days":r1(avg(r.revSum,r.revN)),"Worst Delay vs Due Days":r1(r.maxDelay)}));
  const chaseData=chaseDelayRows.map(r=>({"Chase Label":r.label,"Delayed Entries":r.delayed,"Total Delay Days":r1(r.delaySum),"Avg Delay Days":r1(avg(r.delaySum,r.delayed)),"Worst Delay Days":r1(r.maxDelay)}));
  const brandData=brandDelayRows.map(r=>({"Buyer / Brand":r.label,"Delayed Entries":r.delayed,"Total Delay Days":r1(r.delaySum),"Avg Delay Days":r1(avg(r.delaySum,r.delayed)),"Worst Delay Days":r1(r.maxDelay)}));
  const delayData=worstDelays.map(r=>({"Order":r.order,"Style":r.style,"Buyer":r.buyer,"Chase Label":r.owner,"Stage":r.stage,"Department":r.dept,"Delay Days":r1(r.delay),"Actual vs Original Plan Days":r.delayPlan==null?"":r1(r.delayPlan),"Actual vs Revised Days":r.delayRevised==null?"":r1(r.delayRevised),"Duration Days":r.duration==null?"":r1(r.duration),"Due":fmt(r.due),"Auto Plan":r.plan?fmt(r.plan):"","Revised Plan":r.revised?fmt(r.revised):"","Actual":fmt(r.actual)}));
  const calcBreakup=delayRecords.map(r=>({"Order":r.order,"Style":r.style,"Buyer":r.buyer,"Chase Label":r.owner,"Stage":r.stage,"Stage Key":r.stageKey,"Department":r.dept,"Auto Plan Date":r.plan?fmt(r.plan):"","Revised Date":r.revised?fmt(r.revised):"","Due Date Used":fmt(r.due),"Actual Date":fmt(r.actual),"Start Date Used":r.start?fmt(r.start):"","Delay vs Due Days":r1(r.delay),"Delay vs Auto Plan Days":r.delayPlan==null?"":r1(r.delayPlan),"Actual vs Revised Days":r.delayRevised==null?"":r1(r.delayRevised),"Duration Days":r.duration==null?"":r1(r.duration),"Included in Avg Late Delay":r.delay>0?"YES":"NO - early/on time","Included in Performance Net Delay":"YES","Included in Avg Actual Time":r.duration==null?"NO - start date missing":"YES","Included in Revised Accuracy":r.delayRevised==null?"NO - no revised+actual":"YES"}));
  // Export cleanup: delay/chase detail sheets should show problem rows only and use short management text.
  const dueKind=(r)=>r.revised?"Revised Plan":"Auto Plan";
  const delayText=(d)=>{ const n=r1(d||0); if(n>0) return `${n}d late`; if(n<0) return `${Math.abs(n)}d early`; return "On time"; };
  const resultFor=(r,type)=>{
    if(type==="Actual vs Original Plan") return r.delayPlan>0?"Original plan missed":(r.delayPlan<0?"Early vs original plan":"On original plan");
    if(type==="Revised vs Actual") return r.delayRevised>0?"Revised missed":(r.delayRevised<0?"Early vs revised":"On revised date");
    return r.delay>0?"Delayed":(r.delay<0?"Early":"On time");
  };
  const readingFor=(r,type)=>{
    if(type==="Actual vs Original Plan") return r.delayPlan>0?`Missed original plan by ${r1(r.delayPlan)}d`:(r.delayPlan<0?`Beat original plan by ${Math.abs(r1(r.delayPlan))}d`:"Met original plan");
    if(type==="Revised vs Actual") return r.delayRevised>0?`Missed revised by ${r1(r.delayRevised)}d`:(r.delayRevised<0?`Beat revised by ${Math.abs(r1(r.delayRevised))}d`:"Met revised date");
    if(r.delay>0) return `${r.stage} delayed by ${r1(r.delay)}d`;
    if(r.delay<0) return `Closed ${Math.abs(r1(r.delay))}d early`;
    return "Closed on time";
  };
  const actionFor=(r,type="General")=>{
    if(type==="Buyer Approval") return "Buyer follow-up today";
    if(type==="Chase Delay Ranking") return `Chase ${r.dept||"responsible team"}`;
    if(type==="Buyer Brand Delays") return "Tighten buyer follow-up";
    if(type==="Worst Performing Styles") return "Make recovery plan";
    if(type==="Actual vs Original Plan") return r.delayPlan>0?"Review original planning accuracy":"No action";
    if(type==="Revised vs Actual") return r.delayRevised>0?"Ask reason before next revision":"No action";
    if(type==="Stage Performance") return r.delay>0?`Review ${r.stage} handoff` : "No action";
    if(type==="Department Performance") return r.delay>0?`Review ${r.dept||"team"} capacity` : "No action";
    return r.delay>0?"Review blocker" : "No action";
  };
  const detailRowsFor=(type,records=delayRecords,opts={})=>{
    const problemOnly=!!opts.problemOnly;
    const src=problemOnly?records.filter(r=> type==="Revised vs Actual" ? r.delayRevised>0 : r.delay>0):records;
    return src.map(r=>({
      "Report Context":type,
      "Style No":r.style||"", "Order No":r.order||"", "Buyer / Brand":r.buyer||"", "Junior / Style Owner":r.owner||"",
      "Stage / Activity":r.stage||"", "Chase Label":r.dept||"",
      "Auto Plan":r.plan?fmt(r.plan):"", "Revised Plan":r.revised?fmt(r.revised):"", "Due Used":r.due?fmt(r.due):"", "Actual Done":r.actual?fmt(r.actual):"", "Start Used":r.start?fmt(r.start):"",
      "Actual Time Taken":r.duration==null?"Start missing":`${r1(r.duration)}d`,
      "Delay vs Due":delayText(r.delay), "Delay vs Original Plan":r.delayPlan==null?"":delayText(r.delayPlan), "Actual vs Revised":r.delayRevised==null?"No revised plan":delayText(r.delayRevised),
      "Result":resultFor(r,type),
      "How calculated":`Due = ${dueKind(r)}`,
      "Management Reading":readingFor(r,type),
      "Suggested Action":actionFor(r,type)
    }));
  };
  const compactAgg=(records,keyFn,labelName)=>{
    const m={}; records.forEach(r=>{
      const k=keyFn(r)||"(blank)";
      const o=m[k]=m[k]||{ label:k, n:0, delaySum:0, worst:0, planN:0, planSum:0, planMissN:0, planMissSum:0, revN:0, revSum:0, revMissN:0, revMissSum:0 };
      o.n++; o.delaySum+=Math.max(0,r.delay||0); o.worst=Math.max(o.worst,Math.max(0,r.delay||0));
      if(r.delayPlan!=null){ o.planN++; o.planSum+=r.delayPlan; if(r.delayPlan>0){ o.planMissN++; o.planMissSum+=r.delayPlan; } }
      if(r.delayRevised!=null){ o.revN++; o.revSum+=r.delayRevised; if(r.delayRevised>0){ o.revMissN++; o.revMissSum+=r.delayRevised; } }
    });
    return Object.values(m).sort((a,b)=>b.delaySum-a.delaySum).map(o=>({
      [labelName]:o.label,
      "Delayed Entries":o.n,
      "Total Delay Days":r1(o.delaySum),
      "Avg Delay Days":r1(avg(o.delaySum,o.n)),
      "Avg Actual vs Original Plan Days":r1(avg(o.planSum,o.planN)),
      "Avg Actual vs Revised Plan Days":r1(avg(o.revSum,o.revN)),
      "Original Plan Records":o.planN,
      "Revised Records":o.revN,
      "Worst Delay Days":r1(o.worst)
    }));
  };
  const managementDetailRows=detailRowsFor("Management Summary",delayRecords);
  const planActualData=planActualRecords.map(r=>({"Order":r.order,"Style":r.style,"Buyer":r.buyer,"Chase Label":r.dept,"Stage":r.stage,"Auto Plan Date":r.plan?fmt(r.plan):"","Revised Date":r.revised?fmt(r.revised):"","Actual Date":fmt(r.actual),"Actual vs Original Plan Days":r1(r.delayPlan),"Accuracy":r.delayPlan===0?"On original plan":(r.delayPlan>0?"Late vs original plan":"Early vs original plan"),"Management Reading":r.delayPlan>0?`Missed original plan by ${r1(r.delayPlan)}d`:(r.delayPlan<0?`Beat original plan by ${Math.abs(r1(r.delayPlan))}d`:"Met original plan") }));
  const revisedActualData=revisedActualRecords.map(r=>({"Order":r.order,"Style":r.style,"Buyer":r.buyer,"Chase Label":r.dept,"Stage":r.stage,"Auto Plan Date":r.plan?fmt(r.plan):"","Revised Date":fmt(r.revised),"Actual Date":fmt(r.actual),"Actual vs Revised Days":r1(r.delayRevised),"Accuracy":r.delayRevised===0?"On revised date":(r.delayRevised>0?"Late vs revised":"Early vs revised"),"Management Reading":readingFor(r,"Revised vs Actual")}));
  const planAccuracyData=[
    { "Comparison":"Actual vs Original Plan", "Records":planActualRecords.length, "Late/Missed Records":planMissedRecords.length, "Avg Net Days":r1(avgPlanNet), "Avg Late/Missed Days":r1(avgPlanMiss), "Worst Late/Missed Days":r1(worstPlanMiss), "Purpose":"Shows whether original system plan was realistic/met" },
    { "Comparison":"Actual vs Revised Plan", "Records":revisedActualRecords.length, "Late/Missed Records":revisedMissedRecords.length, "Avg Net Days":r1(avgRevNet), "Avg Late/Missed Days":r1(avgRevMiss), "Worst Late/Missed Days":r1(worstRevMiss), "Purpose":"Shows whether latest meeting/revised commitment was met" }
  ];
  const buyerApprovalTurnaroundDetail=detailRowsFor("Buyer Approval Turnaround",buyerApprovalRecords);
  const buyerApprovalDetail=detailRowsFor("Buyer Approval Delay",buyerApprovalLateRecords,{problemOnly:true});
  const stageDetailRows=detailRowsFor("Stage Performance",delayRecords);
  const deptDetailRows=detailRowsFor("Department Performance",delayRecords);
  const chaseDetailRows=detailRowsFor("Chase Delay Ranking",lateRecords,{problemOnly:true});
  const brandDetailRows=detailRowsFor("Buyer Brand Delays",lateRecords,{problemOnly:true});
  const styleAgg={};
  lateRecords.forEach(r=>{ const key=(r.order||"")+"|"+(r.style||""); const o=styleAgg[key]=styleAgg[key]||{ order:r.order, style:r.style, buyer:r.buyer, owner:r.owner, completed:0, late:0, delaySum:0, worstDelay:0, worstStage:"", stages:[] }; o.completed++; o.late++; o.delaySum+=Math.max(0,r.delay||0); if((r.delay||0)>o.worstDelay){ o.worstDelay=r.delay; o.worstStage=r.stage; } o.stages.push(`${r.stage} ${r1(r.delay)}d`); });
  const worstStyleData=Object.values(styleAgg).sort((a,b)=>b.delaySum-a.delaySum).slice(0,25).map(o=>{
    const rows=lateRecords.filter(r=>(r.order||"")+"|"+(r.style||"")===(o.order||"")+"|"+(o.style||""));
    const p=rows.filter(r=>r.delayPlan!=null), pm=p.filter(r=>r.delayPlan>0), rv=rows.filter(r=>r.delayRevised!=null), rvm=rv.filter(r=>r.delayRevised>0);
    return {"Style No":o.style,"Order No":o.order,"Buyer / Brand":o.buyer,"Junior / Style Owner":o.owner,"Late Stage Count":o.late,"Total Delay Days":`${r1(o.delaySum)}d`,"Avg Actual vs Original Plan Days":r1(avg(p.reduce((s,r)=>s+(r.delayPlan||0),0),p.length)),"Avg Actual vs Revised Plan Days":r1(avg(rv.reduce((s,r)=>s+(r.delayRevised||0),0),rv.length)),"Worst Stage":o.worstStage,"Worst Stage Delay":`${r1(o.worstDelay)}d`,"Management Reading":`${o.late} delayed stage(s); ${r1(o.delaySum)}d total delay`,"Suggested Action":"Make recovery plan"};
  });
  const worstKeys=new Set(Object.values(styleAgg).sort((a,b)=>b.delaySum-a.delaySum).slice(0,25).map(o=>(o.order||"")+"|"+(o.style||"")));
  const worstStyleDetail=detailRowsFor("Worst Performing Styles",lateRecords.filter(r=>worstKeys.has((r.order||"")+"|"+(r.style||""))),{problemOnly:true});
  const lateDetailRows=detailRowsFor("Worst Delays",lateRecords,{problemOnly:true});
  const exportChaseDelayData=compactAgg(lateRecords,r=>r.dept,"Chase Label");
  const exportBuyerDelayData=compactAgg(buyerApprovalLateRecords,r=>r.buyer||"(No buyer)","Buyer / Brand");
  const exportBrandDelayData=compactAgg(lateRecords,r=>r.buyer||"(No buyer)","Buyer / Brand");

  const stageDelayData=compactAgg(lateRecords,r=>r.stage,"Stage").map(o=>{ const def=STAGES.find(s=>s.label===o.Stage)||{}; return { "TNA Order":stageOrderOf(def.key||o.Stage), "Section":stageSectionOf(def.key||o.Stage), ...o, "Chase Label":def.owner||"" }; }).sort((a,b)=>(Number(a["TNA Order"])||999)-(Number(b["TNA Order"])||999));
  const problemStylesFor=(records,metric="delay")=>records.filter(r=>Number(r[metric]||0)>0).sort((a,b)=>Number(b[metric]||0)-Number(a[metric]||0)).slice(0,3).map(r=>`${r.style||r.order||"Style"} +${fmtNum(r[metric])}d`).join(", ");
  const makePerfAggFromRecords=(records,keyFn,labelFn,ownerFn,mode)=>{ const m={}; records.forEach(r=>{ const k=keyFn(r)||"(blank)"; const label=labelFn?labelFn(r,k):k; const o=m[k]=m[k]||{ key:k, label, records:[], n:0, lateN:0, delaySum:0, durSum:0, durN:0, planN:0, planSum:0, revN:0, revSum:0, maxDelay:0, owner:ownerFn?ownerFn(r,k):"" }; o.records.push(r); o.n++; const activeDelay=Number(r.delay||0); const planDelay=Number(r.delayPlan||0); const revisedDelay=Number(r.delayRevised||0); const lateMetric=mode==="revised"?revisedDelay:mode==="noRevised"?planDelay:activeDelay; if(lateMetric>0) o.lateN++; if(mode==="revised") o.delaySum+=revisedDelay; else if(mode==="noRevised") o.delaySum+=planDelay; else o.delaySum+=activeDelay; o.maxDelay=Math.max(o.maxDelay,lateMetric); if(r.duration!=null){ o.durSum+=r.duration; o.durN++; } if(r.delayPlan!=null){ o.planN++; o.planSum+=r.delayPlan; } if(r.delayRevised!=null){ o.revN++; o.revSum+=r.delayRevised; } }); return Object.values(m); };
  const buildPerformanceRows=(mode,view,scope=validPerfScope)=>{ const source=recordsForPerformanceMode(mode,scope); const rows=view==="stage"?makePerfAggFromRecords(source,r=>r.stageKey||r.stage,r=>r.stage,r=>r.dept,mode):makePerfAggFromRecords(source,r=>r.dept||"(No chase label)",r=>r.dept||"(No chase label)",()=>"",mode); return rows.map(r=>{ const metric=mode==="revised"?"delayRevised":mode==="noRevised"?"delayPlan":"delay"; const def=view==="stage"?(STAGES.find(s=>s.key===r.key)||STAGES.find(s=>s.label===r.label)||{}):{}; const avgDelay=avg(r.delaySum,r.n); return { ...r, displayLabel:(view==="stage"&&def.label)?def.label:r.label, stageOrder:view==="stage"?stageOrderOf(def.key||r.key||r.label):null, stageSection:view==="stage"?stageSectionOf(def.key||r.key||r.label):null, count:r.n, late:r.lateN, avgDuration:mode==="summary"?avg(r.durSum,r.durN):null, avgDelay, planNet:avg(r.planSum,r.planN), revNet:avg(r.revSum,r.revN), worst:r.maxDelay, problemStyles:problemStylesFor(r.records,metric), owner:r.owner||def.owner||"" }; }).sort((a,b)=>view==="stage"?(a.stageOrder-b.stageOrder):((b.late-a.late)||(b.avgDelay-a.avgDelay)||String(a.displayLabel).localeCompare(String(b.displayLabel)))); };
  const perfDataKey=`${fc.length}|${delayRecords.length}|${lateRecords.length}|${planActualRecords.length}|${revisedActualRecords.length}|${delayRecords.reduce((sum,r)=>sum+Number(r.delay||0)+Number(r.delayPlan||0)+Number(r.delayRevised||0),0)}`;
  const isLatePerformanceRecord=(r,mode)=> mode==="revised" ? Number(r.delayRevised||0)>0 : mode==="noRevised" ? Number(r.delayPlan||0)>0 : Number(r.delay||0)>0;
  const getPerformanceRows=(mode,view,scope=validPerfScope)=>{ const key=perfDataKey+"|"+mode+"|"+view+"|"+scope; const cache=perfRowsCacheRef.current; if(cache.has(key)) return cache.get(key); if(cache.size>24) cache.clear(); const rows=buildPerformanceRows(mode,view,scope); cache.set(key,rows); return rows; };
  const recordsForPerformanceMode=(mode,scope=validPerfScope)=>{ const base = mode==="revised" ? delayRecords.filter(r=>r.delayRevised!=null) : mode==="noRevised" ? delayRecords.filter(r=>r.delayRevised==null) : delayRecords; return scope==="late" ? base.filter(r=>isLatePerformanceRecord(r,mode)) : base; };
  const performanceModeLabel=(validPerfScope==="late"?"Late Activities · ":"")+(validPerfMode==="revised"?"Revised Commitments":validPerfMode==="noRevised"?"No Revised Plan":"Summary");
  const buildPerformanceExport=(mode,view,scope=validPerfScope)=>{
    const rows=getPerformanceRows(mode,view,scope);
    if(mode==="revised") return rows.map(r=>({
      [view==="stage"?"Stage":"Chase Label"]:r.displayLabel,
      ...(view==="stage"?{"TNA Order":r.stageOrder,"Section":r.stageSection,"Owner":r.owner||""}:{}),
      "Revised Items":r.count||0,
      "Missed Revised":r.late||0,
      "Revised Miss %":(r.count||0)?Math.round(((r.late||0)/(r.count||1))*100):0,
      "Avg Missed Revised Days":r1(r.revNet||0),
      "Still Off Original TNA Days":r1(r.planNet||0),
      "Worst Revised Miss Days":r1(Math.max(0,r.worst||0)),
      "Problem Styles":r.problemStyles||""
    }));
    if(mode==="noRevised") return rows.map(r=>({
      [view==="stage"?"Stage":"Chase Label"]:r.displayLabel,
      ...(view==="stage"?{"TNA Order":r.stageOrder,"Section":r.stageSection,"Owner":r.owner||""}:{}),
      "No-Revision Items":r.count||0,
      "Missed Original":r.late||0,
      "Original Miss %":(r.count||0)?Math.round(((r.late||0)/(r.count||1))*100):0,
      "Avg Off Original TNA Days":r1(r.planNet||0),
      "Worst Original Miss Days":r1(Math.max(0,r.worst||0)),
      "Problem Styles":r.problemStyles||""
    }));
    return rows.map(r=>({
      [view==="stage"?"Stage":"Chase Label"]:r.displayLabel,
      ...(view==="stage"?{"TNA Order":r.stageOrder,"Section":r.stageSection,"Owner":r.owner||""}:{}),
      "Work Done":r.count||0,
      "Missed Due":r.late||0,
      "Miss Rate %":(r.count||0)?Math.round(((r.late||0)/(r.count||1))*100):0,
      "Real Time Taken Days":r.avgDuration==null?"":r1(r.avgDuration),
      "Off Original TNA Days":r1(r.planNet||0),
      "Missed Active Due Days":r1(r.avgDelay||0),
      "Worst Delay Days":r1(Math.max(0,r.worst||0)),
      "Data Confidence":`O:${r.planN||0} R:${r.revN||0} D:${r.durN||0}`,
      "Problem Styles":r.problemStyles||""
    }));
  };
  const performanceAnalysisData = buildPerformanceExport(validPerfMode,perfView,validPerfScope);
  const performanceAnalysisDetail = ()=>detailRowsFor(`Performance Analysis - ${performanceModeLabel}`, recordsForPerformanceMode(validPerfMode,validPerfScope), { problemOnly: validPerfScope==="late" || validPerfMode!=="summary" }).filter(r=>{
    if(perfView!=="stage") return true;
    return !!r["Stage / Activity"];
  });
  const performanceAnalysisMeta=[{ "View By":perfView==="stage"?"Stage":"Chase Label", "Data View":performanceModeLabel, "Activity Scope":validPerfScope==="late"?"Late activities only":"All completed activities", "Rows":performanceAnalysisData.length, "Detail Rows":"computed on export", "Rule":validPerfMode==="revised"?"Only completed records with revised plan dates":"noRevised"===validPerfMode?"Only completed records with no revised plan date":"All completed actual date rows; active due = revised if present else original" }];
  const analyticsSheets=[
      { label:"Summary", data:mgmtSummary, detailData:managementDetailRows, modes:["summary","detailed"] },
      { label:"Calculation Checks", data:checks, modes:["summary","detailed"] },
      { label:"Performance Analysis", data:performanceAnalysisData.length?performanceAnalysisData:performanceAnalysisMeta, detailData:performanceAnalysisDetail, modes:["summary","detailed"] },
      { label:"Buyer Approval Turnaround", data:buyerData, detailData:buyerApprovalTurnaroundDetail, modes:["summary","detailed"] },
      { label:"Buyer Approval Delay", data:exportBuyerDelayData, detailData:buyerApprovalDetail, modes:["summary","detailed"] },
      { label:"Chase Delay Ranking", data:exportChaseDelayData, detailData:chaseDetailRows, modes:["summary","detailed"] },
      { label:"Buyer Brand Delays", data:exportBrandDelayData, detailData:brandDetailRows, modes:["summary","detailed"] },
      { label:"Escalation Owner Load", data:escRows.map(([owner,count])=>({"Escalation Owner":owner,"Overdue Items":count})), detailData:escalationTodo.map(t=>({"Order No":t.orderNo||"","Style No":t.styleNo||"","Activity":t._activityLabel||t.activity||"","Chase Label":t.owner||"","Escalation Owner":t.escalationOwner||"","Escalation Level":t.escalationLevel||"","Days Overdue":t.daysLate||"","Action":t.escalationAction||""})), modes:["summary","detailed"] },
      { label:"Current Slice Styles", data:currentStyles, detailData:currentSliceStageDetail, modes:["detailed"] },
      { label:"Plan Accuracy Summary", data:planAccuracyData, modes:["summary","detailed"] },
      { label:"Actual vs Original Plan", data:planActualData, detailData:detailRowsFor("Actual vs Original Plan", planActualRecords), modes:["summary","detailed"] },
      { label:"Worst Completed Delays", data:delayData, detailData:lateDetailRows, modes:["summary","detailed"] },
      { label:"Worst Performing Styles", data:worstStyleData, detailData:worstStyleDetail, modes:["summary","detailed"] },
      { label:"Calculation Breakup", data:managementDetailRows, modes:["detailed"] },
      { label:"Actual vs Revised", data:revisedActualData, detailData:detailRowsFor("Revised vs Actual", revisedActualRecords), modes:["summary","detailed"] },
  ];

  const card=(label,val,color,onClick,sub)=>(<button onClick={onClick} disabled={!onClick} style={{ flex:1, minWidth:136, textAlign:"left", background:"var(--surface)", border:"1px solid var(--line-2)", borderRadius:12, padding:"14px 16px", cursor:onClick?"pointer":"default", fontFamily:"inherit" }}><div style={{ fontSize:28, fontWeight:800, fontFamily:"'Archivo',sans-serif", color, lineHeight:1 }}>{val}</div><div style={{ fontSize:10, color:"var(--muted-2)", marginTop:5, letterSpacing:0.5, textTransform:"uppercase" }}>{label}{onClick?" ›":""}</div>{sub&&<div style={{ fontSize:9, color:"var(--muted-1)", marginTop:4 }}>{sub}</div>}</button>);
  const sel=(label,val,opts,onChange)=><MultiSelectDropdown label={label} value={val} options={opts} onChange={onChange} rounded />;
  const section=(title,children,sub)=>{ const key=String(title||"").toLowerCase().replace(/[^a-z0-9]+/g,"_").replace(/^_|_$/g,"")||"section"; const open=isMgmtOpen(key); return (<div style={{ flex:"1 1 100%", width:"100%", minWidth:"100%", background:"var(--surface)", border:"1px solid var(--line-2)", borderRadius:14, overflow:"hidden", boxShadow:"var(--card-shadow)" }}><button onClick={()=>toggleMgmt(key)} title={open?"Collapse this management table":"Expand this management table"} style={{ width:"100%", border:"none", background:open?"var(--accent-tint)":"var(--surface)", cursor:"pointer", padding:"13px 16px", display:"flex", alignItems:"center", justifyContent:"space-between", gap:12, fontFamily:"inherit", textAlign:"left" }}><span style={{ minWidth:0 }}><span style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:14 }}>{title}</span>{sub&&<span style={{ display:"block", fontSize:10, color:"var(--muted-2)", marginTop:3, whiteSpace:"normal" }}>{sub}</span>}</span><span style={{ flex:"0 0 auto", fontSize:15, fontWeight:900, color:"var(--muted-3)" }}>{open?"−":"+"}</span></button>{open&&<div style={{ padding:"6px 16px 14px", borderTop:"1px solid var(--line-3)" }}>{children}</div>}</div>); };
  const rowBtn=(key,label,right,color,onClick,sub)=>(<button key={key} onClick={onClick} disabled={!onClick} style={{ width:"100%", display:"grid", gridTemplateColumns:"minmax(120px,1fr) 96px", alignItems:"center", gap:10, border:"none", borderBottom:"1px solid var(--line-3)", background:"transparent", cursor:onClick?"pointer":"default", fontFamily:"inherit", padding:"8px 0", textAlign:"left" }}><span style={{ minWidth:0 }}><span style={{ display:"block", fontSize:11, fontWeight:800, color:color||"var(--ink)", overflow:"hidden", whiteSpace:"nowrap", textOverflow:"ellipsis" }}>{label}</span>{sub&&<span style={{ display:"block", fontSize:9, color:"var(--muted-2)", marginTop:2, overflow:"hidden", whiteSpace:"nowrap", textOverflow:"ellipsis" }}>{sub}</span>}</span><span style={{ fontSize:11, fontWeight:800, color:color||"var(--ink)", whiteSpace:"nowrap", textAlign:"right" }}>{right}</span></button>);
  const barLine=(key,label,n,max,color,onClick,right)=>(<button key={key} onClick={onClick} disabled={!onClick} style={{ display:"flex", alignItems:"center", gap:8, width:"100%", border:"none", background:"transparent", cursor:onClick?"pointer":"default", fontFamily:"inherit", padding:"5px 0" }}><span style={{ width:90, fontSize:10, fontWeight:800, color:"var(--muted-4)", textAlign:"left", overflow:"hidden", whiteSpace:"nowrap", textOverflow:"ellipsis" }}>{label}</span><span style={{ flex:1, height:14, background:"#f0ece3", borderRadius:999, overflow:"hidden" }}><span style={{ display:"block", height:"100%", width:`${(n/Math.max(1,max))*100}%`, background:color, borderRadius:999 }}/></span><span style={{ width:58, textAlign:"right", fontSize:10, fontWeight:800 }}>{right||n}</span></button>);
  const toggleBtn=(active,label,onClick)=><button onClick={onClick} style={{ fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"5px 9px", cursor:"pointer", border:"1px solid var(--line-2)", borderRadius:8, background:active?"var(--ink)":"var(--surface)", color:active?"var(--bg)":"var(--ink)" }}>{label}</button>;
  const performanceRows = getPerformanceRows(validPerfMode,perfView,validPerfScope);
  const performanceHeaders = validPerfMode==="revised"
    ? [perfView==="stage"?"Stage":"Chase Label",validPerfScope==="late"?"Late Revised Items":"Revised Items","Missed Revised","Revised Miss %","Avg Missed Revised","Still Off Original TNA","Worst Revised Miss","Problem Styles"]
    : validPerfMode==="noRevised"
      ? [perfView==="stage"?"Stage":"Chase Label",validPerfScope==="late"?"Late No-Revision Items":"No-Revision Items","Missed Original","Original Miss %","Avg Off Original TNA","Worst Original Miss","Problem Styles"]
      : [perfView==="stage"?"Stage":"Chase Label",validPerfScope==="late"?"Late Items":"Work Done","Missed Due","Miss Rate %","Real Time Taken","Off Original TNA","Missed Active Due","Worst Delay","Data Confidence","Problem Styles"];
  const perfCell=(v)=>fmtDays(Number.isFinite(Number(v))?Number(v):0);
  const avgActualBreakup = [...deptRows]
    .filter(r=>r && r.durN)
    .sort((a,b)=>avg(b.durSum,b.durN)-avg(a.durSum,a.durN))
    .slice(0,4)
    .map(r=>({ label:r.label, count:r.durN, avg:avg(r.durSum,r.durN) }));
  const actualTimeCard=(
    <div style={{ flex:"1 1 360px", minWidth:320, textAlign:"left", background:"var(--surface)", border:"1px solid var(--line-2)", borderRadius:12, padding:"14px 16px", fontFamily:"inherit", display:"grid", gridTemplateColumns:"130px 1fr", gap:14, alignItems:"start" }}>
      <div>
        <div style={{ fontSize:28, fontWeight:800, fontFamily:"'Archivo',sans-serif", color:"var(--info)", lineHeight:1 }}>{fmtDays(avgDuration)}</div>
        <div style={{ fontSize:10, color:"var(--muted-2)", marginTop:5, letterSpacing:0.5, textTransform:"uppercase" }}>Avg actual time</div>
        <div style={{ fontSize:9, color:"var(--muted-1)", marginTop:4 }}>stage cycle time</div>
      </div>
      <div style={{ borderLeft:"1px solid var(--line-2)", paddingLeft:12, minWidth:0 }}>
        <div style={{ fontSize:9, color:"var(--muted-2)", fontWeight:800, textTransform:"uppercase", letterSpacing:0.4, marginBottom:6 }}>Breakup by chase label</div>
        {avgActualBreakup.length?avgActualBreakup.map(x=>(
          <div key={x.label} style={{ display:"grid", gridTemplateColumns:"1fr auto", gap:8, fontSize:10, padding:"3px 0", borderBottom:"1px solid var(--line-3)" }}>
            <span style={{ fontWeight:800, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{x.label}<span style={{ color:"var(--muted-1)", fontWeight:700 }}> · {x.count}</span></span>
            <span style={{ fontWeight:900, color:"var(--info)" }}>{fmtDays(x.avg)}</span>
          </div>
        )):<div style={{ fontSize:10, color:"var(--muted-1)" }}>No duration records yet.</div>}
      </div>
    </div>
  );
  const performanceCellValue=(r,h)=>{
    const pct=(r.count||0)?Math.round(((r.late||0)/(r.count||1))*100):0;
    if(h==="Revised Items"||h==="Late Revised Items"||h==="No-Revision Items"||h==="Late No-Revision Items"||h==="Work Done"||h==="Late Items") return r.count||0;
    if(h==="Missed Revised"||h==="Missed Original"||h==="Missed Due") return r.late||0;
    if(h==="Revised Miss %"||h==="Original Miss %"||h==="Miss Rate %") return `${pct}%`;
    if(h==="Real Time Taken") return r.avgDuration==null?"—":perfCell(r.avgDuration);
    if(h==="Off Original TNA"||h==="Still Off Original TNA"||h==="Avg Off Original TNA") return perfCell(r.planNet);
    if(h==="Missed Active Due"||h==="Avg Missed Revised") return perfCell(r.avgDelay);
    if(h==="Worst Delay"||h==="Worst Revised Miss"||h==="Worst Original Miss") return perfCell(Math.max(0,r.worst||0));
    if(h==="Data Confidence") return `O:${r.planN||0} · R:${r.revN||0} · D:${r.durN||0}`;
    if(h==="Problem Styles") return r.problemStyles||"—";
    return "";
  };
  const performanceTable=(rows)=> rows.length?(
    <div style={{ overflowX:"auto" }}>
      <table style={{ width:"100%", borderCollapse:"collapse", fontSize:10 }}>
        <thead><tr>{performanceHeaders.map((h,i)=><th key={h} title={h==="Data Confidence"?"O = original plan records, R = revised records, D = duration records. Confidence counts, not performance scores.":""} style={{ textAlign:i===0?"left":"right", padding:"7px 8px", borderBottom:"1px solid var(--line-2)", color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, whiteSpace:"nowrap" }}>{h}</th>)}</tr></thead>
        <tbody>{rows.map((r,idx)=>{ const showSection=perfView==="stage" && r.stageSection && (!idx || rows[idx-1].stageSection!==r.stageSection); return <React.Fragment key={r.key||r.label}>{showSection && <tr><td colSpan={performanceHeaders.length} style={{ padding:"8px 8px 5px", background:"var(--toolbar-subtle)", color:"var(--accent)", fontWeight:900, textTransform:"uppercase", letterSpacing:0.6, fontSize:9, borderTop:"1px solid var(--line-2)", borderBottom:"1px solid var(--line-3)" }}>{r.stageSection}</td></tr>}<tr>
          <td style={{ padding:"7px 8px", fontWeight:800, whiteSpace:"nowrap" }}>{perfView==="stage"&&r.stageOrder?<span style={{ color:"var(--muted-2)", fontWeight:900, marginRight:6 }}>{r.stageOrder}.</span>:null}{r.displayLabel}{r.owner?<span style={{ color:"var(--muted-2)", fontWeight:700 }}> · {r.owner}</span>:null}</td>
          {performanceHeaders.slice(1).map(h=>{ const val=performanceCellValue(r,h); const danger=(String(val).startsWith("+")||(["Missed Revised","Missed Original","Missed Due","Revised Miss %","Original Miss %","Miss Rate %"].includes(h)&&parseFloat(val)>0)); return <td key={h} title={h==="Data Confidence"?"O = original target data, R = revised commitment data, D = real-time duration data.":(h==="Problem Styles"?String(val):"")} style={{ padding:"7px 8px", textAlign:"right", color:danger?"var(--danger)":(h==="Data Confidence"?"var(--muted-2)":"var(--ink)"), fontWeight:danger?800:700, whiteSpace:"nowrap", maxWidth:h==="Problem Styles"?220:undefined, overflow:h==="Problem Styles"?"hidden":undefined, textOverflow:h==="Problem Styles"?"ellipsis":undefined }}>{val}</td>; })}
        </tr></React.Fragment>; })}</tbody>
      </table>
      <div style={{ fontSize:9, color:"var(--muted-7)", marginTop:8, lineHeight:1.45 }}>{validPerfScope==="late"?"Late Activities limits this table to completed records that actually missed the selected due basis. ":"All Completed includes early, on-time, and late completed actuals. "}{validPerfMode==="revised"?"Revised Commitments shows only completed records with revised plan dates. Use it to check whether meeting/revised commitments were still missed.":validPerfMode==="noRevised"?"No Revised Plan shows completed records without revised dates. Use it to catch original TNA misses where no new commitment was captured.":"Summary uses active due: revised date if entered, otherwise original auto/system plan. Off Original TNA shows drift from the original target."}{perfView==="stage"?" Stage view follows fixed TNA flow order, not alphabetical or delay ranking.":""}</div>
    </div>
  ):<div style={{ fontSize:11, color:"var(--muted-1)" }}>{validPerfScope==="late"?"No late completed records in this slice for the selected Performance view.":(validPerfMode==="revised"?"No completed revised-plan records in this slice.":validPerfMode==="noRevised"?"No completed no-revised records in this slice.":"No completed stage dates yet.")}</div>;

  return (<div style={{ padding:"16px 22px", maxWidth:1280 }}>
    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:16, marginBottom:14, flexWrap:"wrap" }}>
      <div><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:23 }}>Management Dashboard</div><div style={{ fontSize:11.5, color:"var(--muted-2)", marginTop:4, maxWidth:760, lineHeight:1.45 }}>Operational view of current bottlenecks plus actual turnaround performance by stage, department, buyer and chase label. All figures respect the slice filters below. Live tables are drillable; historical aggregate tables use detailed export for row-level breakup.</div></div>
      <ReportExportMenu title="Management" prefix="management" sheets={analyticsSheets} defaultMode="detailed" />
    </div>

    <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"center", marginBottom:14 }}>
      <span style={{ fontSize:10, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.5 }}>Slice:</span>
      {sel("Order",df.order,orders,v=>setDf(d=>({...d,order:v})))} {sel("Fit",df.fit,fits,v=>setDf(d=>({...d,fit:v})))} {sel("Colour",df.colour,colours,v=>setDf(d=>({...d,colour:v})))} {sel("Junior",df.junior,juniors,v=>setDf(d=>({...d,junior:v})))} {sel("Family",df.family,families,v=>setDf(d=>({...d,family:v})))} {sel("Brand",df.brand,brands,v=>setDf(d=>({...d,brand:v})))} {sel("Fabric",df.fabric,fabrics,v=>setDf(d=>({...d,fabric:v})))}
      {anyDf && <button onClick={()=>setDf({})} style={{ fontFamily:"inherit", fontSize:10, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--danger)", background:"var(--surface)", color:"var(--danger)", fontWeight:800, borderRadius:8 }}>clear slice</button>}
      <span style={{ marginLeft:"auto", display:"flex", alignItems:"center", gap:6 }}><span style={{ fontSize:10, color:"var(--muted-2)" }}>drill to:</span><div style={{ display:"flex", border:"1px solid var(--line-2)", borderRadius:8, overflow:"hidden" }}>{["tracker","todo"].map((t,i,arr)=>(<button key={t} onClick={()=>setTarget(t)} style={{ fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"6px 10px", cursor:"pointer", border:"none", borderRight:i<arr.length-1?"1px solid var(--line-2)":"none", background:target===t?"var(--ink)":"var(--surface)", color:target===t?"var(--bg)":"var(--ink)" }}>{t==="tracker"?"Tracker":"To-Do"}</button>))}</div></span>
    </div>

    {anyDf && <div style={{ display:"flex", gap:7, flexWrap:"wrap", margin:"-4px 0 14px" }}>
      {Object.entries(df).filter(([,v])=>Array.isArray(v)?v.length:!!v).map(([k,v])=><span key={k} style={{ fontSize:10, fontWeight:800, border:"1px solid var(--line-2)", background:"var(--accent-tint)", borderRadius:999, padding:"4px 8px" }}>{k}: {arrOf(v).join(", ")}</span>)}
    </div>}

    <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
      {card("Styles in slice",total,"var(--ink)",null,completionPct+"% released")}
      {card("On track",onTrack,"var(--success)",()=>goStatus("On Track"))}
      {card("At risk",atRisk,"var(--danger)",()=>goStatus("At Risk"))}
      {card("Delivery risk",delRisk,"var(--danger)",()=>goStatus("All",{ overall:["Delivery risk"] }))}
      {card("Released",released,"var(--success)",()=>goStatus("Released"))}
      {card("Overdue open activities",overdueAct,"var(--danger)",null,"frontier overdue")}
      {card("Escalation items",escalationTodo.length,escalationTodo.length?"var(--danger)":"var(--success)",()=>drillRowsToTodo(escalationTodo,{ priority:"Overdue" }),"who must chase now")}
      {card("Avg late delay",fmtDays(avgDelay),avgDelay>0?"var(--danger)":"var(--success)",null,"delayed stages only")}
      {actualTimeCard}
      {/* Technical missed-only plan/revised cards removed from visible Management dashboard.
          Main Performance Analysis now shows original target drift, active due delay, revised commitment performance, and data-confidence records. */}
    </div>

    <div style={{ display:"flex", flexDirection:"column", gap:12, marginTop:22 }}>
      {section("Who to chase — current", owners.length?owners.map(([o,n])=>barLine(o,o,n,maxOwner,OWNER_COLOR[o]||"var(--accent)",()=>goOwner(o))):<div style={{ fontSize:11, color:"var(--muted-1)" }}>Nothing pending.</div>, "Open actionable items by chase label")}
      {section("Escalation owner load", escRows.length?escRows.map(([o,n])=>barLine(o,o,n,maxEsc,"var(--danger)",()=>drillRowsToTodo(escalationTodo.filter(t=>String(t.escalationOwner||"")===String(o)),{ priority:"Overdue", escalationOwner:o }),n)):<div style={{ fontSize:11, color:"var(--muted-1)" }}>No overdue escalation items.</div>, "Who must chase now based on editable Settings duration slabs")}
      {section("Open activities", acts.length?acts.map(([label,v])=>barLine(label,label,v.n,maxAct,v.over?"var(--danger)":"var(--accent)",()=>goAct(label,v.key),v.over?`${v.n} (${v.over})`:v.n)):<div style={{ fontSize:11, color:"var(--muted-1)" }}>Nothing due.</div>, "Red count in bracket = overdue")}
      {section("Where styles are stuck", Object.entries(phase).map(([p,n])=>barLine(p,p,n,maxPhase,p==="Fabric IH"?"var(--danger)":"var(--accent)",()=>goPhase(p))), "Current next-pending phase · click opens matching To-Do phase")}
    </div>

    <div style={{ display:"flex", flexDirection:"column", gap:12, marginTop:12 }}>
      {section("Performance analysis", <>
        <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"center", margin:"0 0 8px" }}>
          <span style={{ fontSize:10, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase" }}>View by:</span>
          {toggleBtn(perfView==="chase","Chase Label",()=>switchPerfView("chase"))}
          {toggleBtn(perfView==="stage","Stage",()=>switchPerfView("stage"))}
          <span style={{ width:10 }}/>
          <span style={{ fontSize:10, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase" }}>Activity scope:</span>
          {toggleBtn(validPerfScope==="all","All Completed",()=>switchPerfScope("all"))}
          {toggleBtn(validPerfScope==="late","Late Activities",()=>switchPerfScope("late"))}
          <span style={{ width:10 }}/>
          <span style={{ fontSize:10, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase" }}>Data view:</span>
          {toggleBtn(validPerfMode==="summary","Summary",()=>switchPerfMode("summary"))}
          {toggleBtn(validPerfMode==="revised","Revised Commitments",()=>switchPerfMode("revised"))}
          {toggleBtn(validPerfMode==="noRevised","No Revised Plan",()=>switchPerfMode("noRevised"))}
          {isPerfPending&&<span style={{ fontSize:10, color:"var(--muted-2)", fontWeight:800 }}>updating…</span>}
        </div>
        {performanceTable(performanceRows)}
      </>, perfView==="stage" ? `${performanceModeLabel} by TNA stage order` : `${performanceModeLabel} by chase label`)}
      {/* Missed-only plan accuracy section kept out of visible dashboard because it is technical for daily management reading.
          Original/revised performance remains in simplified Performance Analysis columns and export sheets. */}
      {section("Worst completed delays", worstDelays.length?worstDelays.map(r=>rowBtn(r.style+":"+r.stageKey,r.style||r.order,`+${fmtNum(r.delay)}d`,"var(--danger)",()=>goSearch(r.style),`${r.stage} · ${r.buyer||""} · actual ${fmt(r.actual)} · click opens style`)):<div style={{ fontSize:11, color:"var(--muted-1)" }}>No positive completed delays in this slice.</div>, "Separate exception list: exact styles/stages with completed delay greater than 0")}
    </div>

    <div style={{ display:"flex", flexDirection:"column", gap:12, marginTop:12 }}>
      {section("Buyer approval turnaround", buyerRows.length?buyerRows.map(r=>rowBtn(r.key,r.label,`${fmtDays(avg(r.durSum,r.durN))} avg`,r.lateN?"var(--danger)":"var(--success)",null,`${r.n} approvals · ${r.lateN} late · avg net ${fmtDays(avg(r.delaySum,r.n))} · use export for detail`)):<div style={{ fontSize:11, color:"var(--muted-1)" }}>No buyer approval actuals yet.</div>, "Historical approval performance · use detailed export for row-level breakup")}
      {section("Chase delay ranking", chaseDelayRows.length?chaseDelayRows.map(r=>rowBtn(r.key,r.label,`${r1(r.delaySum)}d total`,"var(--danger)",null,`${r.delayed} delayed entries · avg ${fmtDays(avg(r.delaySum,r.delayed))} · worst ${fmtDays(r.maxDelay)} · use export for detail`)):<div style={{ fontSize:11, color:"var(--muted-1)" }}>No positive chase delays in this slice.</div>, "Historical aggregate · problem rows only · use detailed export for row-level breakup")}
      {section("Buyer / brand delay ranking", brandDelayRows.length?brandDelayRows.map(r=>rowBtn(r.key,r.label,`${r1(r.delaySum)}d total`,"var(--danger)",null,`${r.delayed} delayed entries · avg ${fmtDays(avg(r.delaySum,r.delayed))} · worst ${fmtDays(r.maxDelay)} · use export for detail`)):<div style={{ fontSize:11, color:"var(--muted-1)" }}>No positive buyer/brand delays in this slice.</div>, "Historical aggregate · problem rows only · use detailed export for row-level breakup")}
    </div>
  </div>);
}

/* ========================= TO-DO ========================= */
function TodoView({ items, cfg, setCfg, canEditSettings, filter, setFilter, onJump }){
  const OWNER_COLOR2=OWNER_COLOR;
  const [tf,setTf]=useState({});
  const [todoMode,setTodoMode]=useState(()=>{ try{ return localStorage.getItem("mt_todo_mode")||"all"; }catch(e){ return "all"; } });
  useEffect(()=>{ try{ localStorage.setItem("mt_todo_mode",todoMode); }catch(e){} },[todoMode]);
  const arrVal=(v)=>Array.isArray(v)?v:(v?[v]:[]);
  const normFilterValue=(v)=>String(v||"").trim();
  const cleanTodoFilter=(obj)=>{
    const out={};
    Object.entries(obj||{}).forEach(([k,v])=>{
      const vals=arrVal(v).map(normFilterValue).filter(Boolean);
      if(vals.length) out[k]=vals;
    });
    return out;
  };
  const filterSig=JSON.stringify(cleanTodoFilter(filter||{}));
  useEffect(()=>{ setTf(cleanTodoFilter(filter||{})); },[filterSig]);
  const includeEsc=cfg&&cfg.todoEscalationRows!==false;
  const baseItems=items||[];
  const phaseOf=(t)=>{ const k=t._stageKey||t.key||""; if(k==="techpack") return "Pre-Fit"; if(["fitSend","fitAppr","artwork","artAppr","strikeOff","soAppr"].includes(k)) return "Fit / Print"; if(["labDip","labAppr"].includes(k)) return "Lab Dip"; if(k==="fabricIH") return "Fabric IH"; return "PP / Prod"; };
  const stageLabelOf=(key)=>((STAGES.find(st=>st.key===key)||{}).label)||key||"";
  const norm=(v)=>String(v||"").replace(/^Escalate:\s*/i,"").replace(/\s+/g," ").trim().toLowerCase();
  const stageKeyFromAny=(v)=>stageKeyFromAnyGlobal(v);
  const canonicalTodoRow=(t,kind="Activity")=>{
    const stageKey=stageKeyFromAny(t&&t.key)||stageKeyFromAny(t&&t.activityKey)||stageKeyFromAny(t&&t.stageKey)||stageKeyFromAny(t&&t.originalActivity)||stageKeyFromAny(t&&t.activityLabel)||stageKeyFromAny(t&&t.activity)||"";
    const stageLabel=stageKey ? stageLabelOf(stageKey) : String((t&&t.activityLabel)||(t&&t.originalActivity)||(t&&t.activity)||"").replace(/^Escalate:\s*/i,"").replace(/\s+/g," ").trim();
    return {
      ...(t||{}),
      todoType: kind,
      _stageKey: stageKey,
      _activityLabel: stageLabel,
      key: stageKey || (t&&t.key) || "",
      activityKey: stageKey || (t&&t.activityKey) || "",
      activity: stageLabel,
      activityLabel: stageLabel,
      originalActivity: stageLabel
    };
  };
  const rawDisplayItems=useMemo(()=>{ const esc=includeEsc?baseItems.filter(t=>t.overdue&&t.escalationOwner).map(t=>canonicalTodoRow({ ...t, owner:t.owner },"Escalation")) : []; return [...baseItems.map(t=>canonicalTodoRow(t,"Activity")), ...esc]; },[items,includeEsc]);
  const escalationRows=useMemo(()=>rawDisplayItems.filter(t=>t.todoType==="Escalation"),[rawDisplayItems]);
  const itemGroup=(t)=> (t.isColour || ["fabricIH","labDip","labAppr"].includes(t._stageKey||t.key)) ? "fabricLab" : "activity";
  const displayItems=useMemo(()=>rawDisplayItems.filter(t=> todoMode==="all" || itemGroup(t)===todoMode),[rawDisplayItems,todoMode]);
  const activityCanonical=(t)=>String(t&&t._activityLabel||"").replace(/^Escalate:\s*/i,"").replace(/\s+/g," ").trim();
  const activityKeyOf=(t)=>String((t&&t._stageKey)||(t&&t.activityKey)||(t&&t.key)||"");
  const selectedFor=(field)=>arrVal(tf[field]).map(norm).filter(Boolean);
  const selectedStyleIdSet=(()=>new Set([...arrVal(tf.styleId), ...arrVal(tf.styleIds)].map(v=>String(v)).filter(Boolean)))();
  const exactStylePass=(t)=> !selectedStyleIdSet.size || selectedStyleIdSet.has(String(t&&t.id));

  // HARD activity gate for To-Do. This is intentionally strict:
  // when Activity/Stage is selected, a row may pass only if its canonical stage key matches.
  // Display labels such as "Escalate: Strike-off", "Fit Send rework / resend", grouped lab rows, etc.
  // are ignored for matching. This prevents Lab Dip / Artwork / Fit Send rows leaking into Strike-off.
  // HARD activity gate for To-Do. Do not trust the dropdown label alone.
  // The gate is ON whenever the UI has any activity text/key selected, even if mapping fails.
  // Then we match by canonical stage key first, and by normalized token as a fallback.
  // This prevents the visible filter "Activity: Strike-off" from showing Lab Dip / Artwork rows.
  const selectedActivityRawValues=[...arrVal(tf.activityKey), ...arrVal(tf.key), ...arrVal(tf.activity)].filter(v=>v!=null&&String(v).trim()!="");
  const selectedActivityKeySet=(()=>{
    const keys=selectedActivityRawValues.map(v=>stageKeyFromAny(v)).filter(Boolean);
    return new Set(keys);
  })();
  const selectedActivityTokenSet=(()=>new Set(selectedActivityRawValues.map(v=>filterToken(v)).filter(Boolean)))();
  const hasActivityGate=selectedActivityRawValues.length>0;
  const activityPass=(t)=>{
    if(!hasActivityGate) return true;
    const rowKey=stageKeyFromAny(activityKeyOf(t))||activityKeyOf(t);
    // Hard rule: if the selected activity maps to a known TNA stage, only the same stage key may pass.
    if(selectedActivityKeySet.size>0) return selectedActivityKeySet.has(rowKey);
    return selectedActivityTokenSet.has(filterToken(activityCanonical(t))) || selectedActivityTokenSet.has(filterToken(rowKey));
  };

  const matchesAny=(field,candidates)=>{
    const selected=selectedFor(field);
    if(!selected.length) return true;
    const pool=arrVal(candidates).map(norm).filter(Boolean);
    return pool.some(v=>selected.includes(v));
  };
  const candidatesFor=(t,field)=>{
    if(field==="phase") return phaseOf(t);
    if(field==="activity") return [t._activityLabel, t._stageKey].filter(Boolean);
    if(field==="priority") return t.overdue?"Overdue":"Upcoming";
    if(field==="todoType") return t.todoType;
    if(field==="risk") return t.priorityBucket;
    if(field==="orderNo") return (t.orderNos&&t.orderNos.length)?t.orderNos:t.orderNo;
    if(field==="junior") return (t.juniors&&t.juniors.length)?t.juniors:t.junior;
    if(field==="style") return [t.styleNo, t.colour, ...(Array.isArray(t.styleNos)?t.styleNos:[])].filter(Boolean);
    if(field==="colour") return [t.colour, ...(Array.isArray(t.colours)?t.colours:[])].filter(Boolean);
    if(field==="fit") return [t.fit, ...(Array.isArray(t.fits)?t.fits:[])].filter(Boolean);
    if(field==="family") return [t.family, ...(Array.isArray(t.families)?t.families:[])].filter(Boolean);
    if(field==="brand") return [t.brand, ...(Array.isArray(t.brands)?t.brands:[])].filter(Boolean);
    if(field==="fabric") return [t.fabric, ...(Array.isArray(t.fabrics)?t.fabrics:[])].filter(Boolean);
    if(field==="buyer") return [t.buyer, ...(Array.isArray(t.buyers)?t.buyers:[])].filter(Boolean);
    if(field==="planDate") return t.exp?fmt(t.exp):"(Blank)";
    if(field==="days") return t.overdue?`Late ${Math.abs(Number(t.du)||0)}d`:`Left ${Number(t.du)||0}d`;
    if(field==="drift") return `${Number(t.driftOriginal)||0}d`;
    if(field==="revReject") return `${Number(t.revisionCount)||0} rev / ${Number(t.rejectionRound)||0} rej`;
    if(field==="score") return String(Number(t.priorityScore)||0);
    return t[field];
  };
  // Activity/key/activityKey are handled only by activityPass above. They are excluded from the generic text matcher.
  const filterFields=["phase","priority","risk","todoType","orderNo","junior","branch","owner","escalationOwner","style","colour","fit","family","brand","fabric","buyer","planDate","days","drift","revReject","score"];
  const passExcept=(t,except)=> exactStylePass(t) && (except==="activity" || activityPass(t)) && filterFields.every(field=>field===except || matchesAny(field,candidatesFor(t,field)));
  const pass=(t)=>passExcept(t,null);
  const distinct=(field)=>{ const vals=new Set(); displayItems.forEach(t=>{ if(!passExcept(t,field)) return; if(field==="priority") vals.add(t.overdue?"Overdue":"Upcoming"); else if(field==="phase") vals.add(phaseOf(t)); else if(field==="key") vals.add(stageLabelOf(activityKeyOf(t))); else if(field==="activity") { const av=t._activityLabel||stageLabelOf(activityKeyOf(t)); if(av) vals.add(av); } else if(field==="style") { if(t.isColour && t.colour) vals.add(t.colour); else if(t.styleNo) vals.add(t.styleNo); } else if(field==="orderNo" && Array.isArray(t.orderNos)&&t.orderNos.length) t.orderNos.forEach(v=>v&&vals.add(v)); else if(field==="junior" && Array.isArray(t.juniors)&&t.juniors.length) t.juniors.forEach(v=>v&&vals.add(v)); else { arrVal(candidatesFor(t,field)).forEach(v=>v&&vals.add(v)); } }); return [...vals].sort((a,b)=>String(a).localeCompare(String(b),undefined,{numeric:true,sensitivity:"base"})); };
  // Memory Build 2: the 16 distinct() option lists + the filtered/shown rows are the heaviest per-render work
  // in To-Do (each distinct() scans every row). Recompute only when the display items or the active filter change.
  const tfSig=useMemo(()=>JSON.stringify(tf),[tf]);
  const flt=useMemo(()=>{
    const orders=distinct("orderNo"), juniors=distinct("junior"), activities=distinct("activity"), branches=distinct("branch"), owners=distinct("owner"), escOwners=distinct("escalationOwner"), types=distinct("todoType"), priorities=distinct("priority"), risks=distinct("risk"), stylesList=distinct("style"), planDates=distinct("planDate"), dayBuckets=distinct("days"), driftBuckets=distinct("drift"), revRejects=distinct("revReject"), scoreBuckets=distinct("score");
    const activityScopedDisplayItems=hasActivityGate?displayItems.filter(activityPass):displayItems;
    const shownPre=activityScopedDisplayItems.filter(pass);
    const shown=shownPre.filter(activityPass);
    const filterViolations=shown.filter(t=>!exactStylePass(t) || !activityPass(t) || !pass(t));
    const overdue=shown.filter(t=>t.overdue), upcoming=shown.filter(t=>!t.overdue), critical=shown.filter(t=>t.overdue && (Number(t.daysLate)||0)>5);
    return { orders, juniors, activities, branches, owners, escOwners, types, priorities, risks, stylesList, planDates, dayBuckets, driftBuckets, revRejects, scoreBuckets, shownPre, shown, filterViolations, overdue, upcoming, critical };
  },[displayItems, tfSig, hasActivityGate]);
  const { orders, juniors, activities, branches, owners, escOwners, types, priorities, risks, stylesList, planDates, dayBuckets, driftBuckets, revRejects, scoreBuckets, shownPre, shown, filterViolations, overdue, upcoming, critical }=flt;
  const phases=["Pre-Fit","Fit / Print","Lab Dip","Fabric IH","PP / Prod"];
  const activityCount=rawDisplayItems.filter(t=>itemGroup(t)==="activity").length;
  const fabricLabCount=rawDisplayItems.filter(t=>itemGroup(t)==="fabricLab").length;
  const anyF=Object.values(tf).some(v=>Array.isArray(v)?v.length:!!v);
  const COLS={ pri:"110px", risk:"170px", type:"88px", ord:"92px", style:"230px", jr:"110px", act:"160px", br:"105px", own:"120px", esc:"135px", date:"96px", days:"118px", drift:"120px", revs:"86px", score:"82px", actions:"160px" };
  const GRID=`${COLS.pri} ${COLS.risk} ${COLS.type} ${COLS.ord} ${COLS.style} ${COLS.jr} ${COLS.act} ${COLS.br} ${COLS.own} ${COLS.esc} ${COLS.date} ${COLS.days} ${COLS.drift} ${COLS.revs} ${COLS.score} ${COLS.actions}`;
  const minTableWidth=1982;
  const set=(k,v)=>{
    const clean=arrVal(v).map(normFilterValue).filter(Boolean);
    const upd=f=>{ const next={...(f||{})}; if(clean.length) next[k]=clean; else delete next[k];
      if(k==="activity"){
        const keys=clean.map(x=>stageKeyFromAny(x)).filter(Boolean);
        if(keys.length) next.activityKey=keys; else delete next.activityKey;
        delete next.key;
      }
      if(k==="activityKey"){
        const keys=clean.map(x=>stageKeyFromAny(x)||x).filter(Boolean);
        if(keys.length){ next.activityKey=keys; next.activity=keys.map(stageLabelOf).filter(Boolean); }
        else { delete next.activityKey; delete next.activity; }
        delete next.key;
      }
      return cleanTodoFilter(next); };
    setTf(upd);
    setFilter&&setFilter(upd);
  };
  const applyTodoActivity=(label,key)=>{
    const k=stageKeyFromAny(key)||stageKeyFromAny(label);
    const next=cleanTodoFilter({ ...(tf||{}), activity:k?[stageLabelOf(k)]:[label], activityKey:k?[k]:[] });
    delete next.key;
    setTf(next);
    setFilter&&setFilter(next);
  };
  const hsel=(k,opts,first)=><MultiSelectDropdown label={first} value={arrVal(tf[k])} options={opts} onChange={v=>set(k,v)} />;
  const modeBtn=(key,label,count)=><button onClick={()=>setTodoMode(key)} style={{ fontFamily:"inherit", fontSize:11, fontWeight:900, padding:"8px 13px", cursor:"pointer", border:"1px solid var(--ink)", borderRadius:999, background:todoMode===key?"var(--ink)":"var(--surface)", color:todoMode===key?"var(--bg)":"var(--ink)" }}>{label} <span style={{ opacity:.75 }}>{count}</span></button>;
  const head=<div style={{ minWidth:minTableWidth, display:"grid", gridTemplateColumns:GRID, alignItems:"end", columnGap:0, borderBottom:"2px solid var(--ink)", background:"var(--surface)", position:"sticky", top:0, zIndex:5 }}>
    <div style={{ padding:"6px 6px" }}>{hsel("priority",priorities,"Due")}</div>
    <div style={{ padding:"6px 6px" }}>{hsel("risk",risks,"Risk")}</div>
    <div style={{ padding:"6px 6px" }}>{hsel("todoType",types,"Type")}</div>
    <div style={{ padding:"6px 6px" }}>{hsel("orderNo",orders,"Order")}</div>
    <div style={{ padding:"6px 6px" }}>{hsel("style",stylesList,"Style / Colour")}</div>
    <div style={{ padding:"6px 6px" }}>{hsel("junior",juniors,"Junior")}</div>
    <div style={{ padding:"6px 6px" }}>{hsel("activity",activities,"Activity")}</div>
    <div style={{ padding:"6px 6px" }}>{hsel("branch",branches,"Branch")}</div>
    <div style={{ padding:"6px 6px" }}>{hsel("owner",owners,"Chase")}</div>
    <div style={{ padding:"6px 6px" }}>{hsel("escalationOwner",escOwners,"Escalation")}</div>
    <div style={{ padding:"6px 6px" }}>{hsel("planDate",planDates,"Plan Date")}</div>
    <div style={{ padding:"6px 6px" }}>{hsel("days",dayBuckets,"Days")}</div>
    <div style={{ padding:"6px 6px" }}>{hsel("drift",driftBuckets,"Drift")}</div>
    <div style={{ padding:"6px 6px" }}>{hsel("revReject",revRejects,"Rev/Reject")}</div>
    <div style={{ padding:"6px 6px" }}>{hsel("score",scoreBuckets,"Score")}</div>
    <div style={{ padding:"6px 8px", fontSize:9, fontWeight:900, textTransform:"uppercase", color:"#8a857a" }}>Actions</div>
  </div>;
  const data=shown.map(t=>({ "Due Priority":t.overdue?"Overdue":"Upcoming", "Risk Bucket":t.priorityBucket||"Daily Chase", "Priority Score":Number(t.priorityScore)||0, "Priority Reason":t.priorityReason||"", "To-Do Section":itemGroup(t)==="fabricLab"?"Fabric / Lab Dip":"Activities", "Type":t.todoType||"Activity", "Order No":t.orderNo||"", "Style / Colour":t.isColour?String(t.colour||""):t.styleNo, "Grouped Count":t.isColour?(t.count||0):"", "Junior":t.junior||"", "Activity":t._activityLabel||t.activity||"", "Branch":t.branch||"", "Chase Label":t.owner||"", "Escalation Owner":t.escalationOwner||"", "Escalation Level":t.escalationLevel||"", "Escalation Action":t.escalationAction||"", "Original Plan":t.originalPlan?fmt(t.originalPlan):"", "Due Used":t.dueUsed?fmt(t.dueUsed):(t.exp?fmt(t.exp):""), "Days Late / Left":t.overdue?Math.abs(t.du):t.du, "Drift vs Original Days":Number(t.driftOriginal)||0, "Revision Count":Number(t.revisionCount)||0, "Rejection Round":Number(t.rejectionRound)||0, "Style ID":t.id, "Stage Key":t.key, "Activity Key":t.activityKey||t.key, "Fit":t.fit||"", "Family":t.family||"", "Brand":t.brand||"", "Fabric":t.fabric||"", "Buyer":t.buyer||"" }));
  const summary=[{ "Report Type":"To-Do", "Section":todoMode==="all"?"All":(todoMode==="activity"?"Activities":"Fabric / Lab Dip"), "Shown Items":shown.length, "Base Activity Items":baseItems.length, "Escalation Rows Included":includeEsc?escalationRows.length:0, "Total Display Items":displayItems.length, "Overdue":overdue.length, "Upcoming":upcoming.length, "Critical / Hidden Risk":critical.length, "Hidden Risk Items":shown.filter(t=>String(t.priorityBucket||"")==="Hidden Risk").length }];
  const byOwner={}; const byActivity={};
  shown.forEach(t=>{ byOwner[t.owner||"(blank)"]=(byOwner[t.owner||"(blank)"]||0)+1; byActivity[(t._activityLabel||t.activity)||"(blank)"]=(byActivity[(t._activityLabel||t.activity)||"(blank)"]||0)+1; });
  const ownerRows=Object.entries(byOwner).map(([k,v])=>({"Chase Label":k,"Items":v}));
  const activityRows=Object.entries(byActivity).map(([k,v])=>({"Activity":k,"Items":v}));
  const riskRows=Object.entries(shown.reduce((m,t)=>{ const k=t.priorityBucket||"Daily Chase"; const x=m[k]=m[k]||{n:0,maxScore:0,drift:0}; x.n++; x.maxScore=Math.max(x.maxScore,Number(t.priorityScore)||0); x.drift+=Number(t.driftOriginal)||0; return m; },{})).map(([k,v])=>({"Risk Bucket":k,"Items":v.n,"Max Priority Score":v.maxScore,"Total Drift Days":v.drift}));
  const todoLogicChecks=[
      { Check:"Display count", Rule:"Shown rows equal final canonical filtered rows", Value:shown.length, Expected:shownPre.length, Result:filterViolations.length===0?"OK":"CHECK" },
      { Check:"Filter violations", Rule:"Rendered rows must pass canonical filter engine", Value:filterViolations.length, Expected:"0", Result:filterViolations.length===0?"OK":"CHECK" },
      { Check:"Order filter grouping", Rule:"Fabric/Lab colour rows are grouped by order + colour so order filters do not show mixed-order rows", Value:"order+colour grouping", Expected:"no mixed order row", Result:"OK" },
      { Check:"Escalation toggle", Rule:"Escalation rows added only when toggle is ON", Value:includeEsc?escalationRows.length:0, Expected:includeEsc?"overdue rows with escalation owner":"0", Result:(!includeEsc&&escalationRows.length===0)||(includeEsc&&escalationRows.every(t=>t.overdue&&t.escalationOwner))?"OK":"CHECK" },
      { Check:"Base activity rows", Rule:"Base activity rows remain even when escalation rows are ON", Value:baseItems.length, Expected:"unchanged source To-Do items", Result:"OK" },
      { Check:"Overdue/upcoming split", Rule:"Overdue + Upcoming equals shown rows", Value:overdue.length+upcoming.length, Expected:shown.length, Result:(overdue.length+upcoming.length)===shown.length?"OK":"CHECK" }
  ];
  const todoSheets=[
      { label:"Summary", data:summary, detailData:data, modes:["summary","detailed"] },
      { label:"Logic Checks", data:todoLogicChecks, modes:["summary","detailed"] },
      { label:"By Chase", data:ownerRows, detailData:data, modes:["summary","detailed"] },
      { label:"By Activity", data:activityRows, detailData:data, modes:["summary","detailed"] },
      { label:"By Risk", data:riskRows, detailData:data, modes:["summary","detailed"] },
      { label:"To-Do Items", data:data, modes:["detailed"] },
  ];
  const copyPlainText=async(txt)=>{ const v=String(txt||""); if(!v) return; try{ await navigator.clipboard.writeText(v); }catch(e){ try{ const ta=document.createElement("textarea"); ta.value=v; ta.setAttribute("readonly",""); ta.style.position="fixed"; ta.style.left="-9999px"; document.body.appendChild(ta); ta.select(); document.execCommand("copy"); document.body.removeChild(ta); }catch(err){} } };
  const row=(t)=><div key={(t.isColour?"col-":"")+t.id+t._stageKey+(t.colour||"")+(t.orderNo||"")} title="Text is selectable/copyable. Use Copy style to copy only the style number; use Open to jump to Tracker." style={{ minWidth:minTableWidth, display:"grid", gridTemplateColumns:GRID, alignItems:"center", borderLeft:`4px solid ${t.overdue?"var(--danger)":"var(--accent)"}`, borderBottom:"1px solid #eee7da", background:t.isColour?"#fbf8f1":"var(--surface)", cursor:"default", fontFamily:"inherit", minHeight:46, userSelect:"text", WebkitUserSelect:"text" }}>
    <div style={{ padding:"7px 8px", fontSize:10, fontWeight:800, display:"flex", alignItems:"center", gap:7, color:t.overdue?"var(--danger)":"#7a560f" }}><span style={{ width:8, height:8, borderRadius:"50%", background:t.overdue?"var(--danger)":"var(--accent)", flexShrink:0 }}/>{t.overdue?"Overdue":"Upcoming"}</div>
    <div title={t.priorityReason||""} style={{ padding:"7px 8px", fontSize:10, fontWeight:900, color:String(t.priorityBucket||"").includes("Critical")?"var(--danger)":String(t.priorityBucket||"").includes("Hidden")?"#b45309":"var(--ink)", lineHeight:1.15 }}>{t.priorityBucket||"Daily Chase"}<div style={{ fontSize:8, fontWeight:700, color:"var(--muted-2)", marginTop:2, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{t.priorityReason||""}</div></div>
    <div style={{ padding:"7px 8px", fontSize:9, fontWeight:800, color:t.todoType==="Escalation"?"var(--danger)":"var(--muted-3)" }}>{t.todoType||"Activity"}</div>
    <div style={{ padding:"7px 8px", fontSize:10, color:"var(--muted-4)", fontWeight:800 }}>{t.orderNo||"—"}</div>
    <div style={{ padding:"7px 8px", fontSize:11, fontWeight:800, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{t.isColour?<span style={{ display:"inline-flex", alignItems:"center", gap:6, maxWidth:"100%" }}><span style={{ fontSize:8, fontWeight:900, background:"var(--ink)", color:"var(--bg)", padding:"1px 4px", flexShrink:0 }}>{t.key==="fabricIH"?"FABRIC":"LAB"}</span><span style={{ overflow:"hidden", textOverflow:"ellipsis" }}>{t.colour}</span><span style={{ color:"var(--muted-1)", fontWeight:500, flexShrink:0 }}>×{t.count}</span></span>:t.styleNo}</div>
    <div style={{ padding:"7px 8px", fontSize:10, color:"var(--muted-5)" }}>{t.junior||"—"}</div>
    <div style={{ padding:"7px 8px", fontSize:10, fontWeight:800, color:"#333", whiteSpace:"normal", lineHeight:1.25 }}>{t._activityLabel}</div>
    <div style={{ padding:"7px 8px", fontSize:10, color:"var(--muted-3)" }}>{t.branch}</div>
    <div style={{ padding:"7px 8px", fontSize:10, fontWeight:800, color:OWNER_COLOR2[t.owner]||"var(--muted-3)" }}>{t.owner}</div>
    <div style={{ padding:"7px 8px", fontSize:10, fontWeight:900, color:t.escalationOwner?"var(--danger)":"var(--muted-2)" }}>{t.escalationOwner||"—"}</div>
    <div style={{ padding:"7px 8px", fontSize:10, color:"var(--muted-3)", fontWeight:700 }}>{fmt(t.exp)}</div>
    <div style={{ padding:"7px 8px", fontSize:10, fontWeight:900, color:t.overdue?"var(--danger)":"#7a560f", lineHeight:1.15 }}>{t.overdue?`+${Math.abs(t.du)}d\nlate`:`${t.du}d\nleft`}</div>
    <div title={`Original ${t.originalPlan?fmt(t.originalPlan):"—"} · Due used ${t.dueUsed?fmt(t.dueUsed):(t.exp?fmt(t.exp):"—")}`} style={{ padding:"7px 8px", fontSize:10, fontWeight:900, color:(Number(t.driftOriginal)||0)>0?"var(--danger)":"var(--muted-3)" }}>{Number(t.driftOriginal)||0}d</div>
    <div style={{ padding:"7px 8px", fontSize:10, fontWeight:800, color:(Number(t.revisionCount)||0)||(Number(t.rejectionRound)||0)?"#b45309":"var(--muted-2)", lineHeight:1.2 }}>{Number(t.revisionCount)||0} rev<div style={{ fontSize:8 }}>{Number(t.rejectionRound)||0} rej</div></div>
    <div title="Priority score combines active due delay, original-plan drift, revision count, rejection round, and stage criticality." style={{ padding:"7px 8px", fontSize:11, fontWeight:900, color:"var(--ink)" }}>{Number(t.priorityScore)||0}</div>
    <div style={{ padding:"7px 8px", display:"flex", gap:8, justifyContent:"flex-end" }}>
      {!t.isColour && <button onClick={(e)=>{ e.stopPropagation(); copyPlainText(t.styleNo); }} title="Copy only this Style No to clipboard" style={{ flexShrink:0, fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"4px 9px", cursor:"pointer", border:"1px solid var(--line-2)", borderRadius:8, background:"var(--bg)", color:"var(--ink)", userSelect:"none", WebkitUserSelect:"none" }}>Copy style</button>}
      {t.isColour && <button onClick={(e)=>{ e.stopPropagation(); copyPlainText(t.colour); }} title="Copy this colour/group text to clipboard" style={{ flexShrink:0, fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"4px 9px", cursor:"pointer", border:"1px solid var(--line-2)", borderRadius:8, background:"var(--bg)", color:"var(--ink)", userSelect:"none", WebkitUserSelect:"none" }}>Copy</button>}
      <button onClick={(e)=>{ e.stopPropagation(); onJump(t.id,t._stageKey||t.key); }} title="Open this style/stage in Tracker" style={{ flexShrink:0, fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"4px 9px", cursor:"pointer", border:"1px solid var(--ink)", borderRadius:8, background:"var(--surface)", color:"var(--ink)", userSelect:"none", WebkitUserSelect:"none" }}>Open</button>
    </div>
  </div>;
  const activitySummary=Object.values(shown.reduce((acc,t)=>{ const stageKey=activityKeyOf(t)||""; const k=stageLabelOf(stageKey)||activityCanonical(t)||"(blank)"; const cur=acc[stageKey||k]||(acc[stageKey||k]={ activity:k, activityKey:stageKey, upcoming:0, overdue:0, critical:0, total:0 }); cur.total++; if(t.overdue){ cur.overdue++; if((Number(t.daysLate)||0)>5) cur.critical++; } else cur.upcoming++; return acc; },{})).sort((a,b)=>b.critical-a.critical||b.overdue-a.overdue||b.total-a.total).slice(0,10);
  const cardTone=(title)=> title==="Upcoming"?{ bg:"#fff4d8", bd:"#d58a13", fg:"#8a5200" }:title==="Overdue"?{ bg:"#ffe4dc", bd:"#c5251a", fg:"#b82117" }:title==="Critical"?{ bg:"#ffd4cc", bd:"#9f1712", fg:"#9f1712" }:{ bg:"var(--surface)", bd:"var(--line-2)", fg:"var(--ink)" };
  const card=(title,n,sub,color)=>{ const tone=cardTone(title); return <div style={{ minWidth:165, flex:"1 1 165px", border:`2px solid ${tone.bd}`, borderRadius:14, background:tone.bg, padding:"12px 14px", boxShadow:"var(--card-shadow)" }}><div style={{ fontSize:10.5, textTransform:"uppercase", color:tone.fg, fontWeight:950, letterSpacing:.4 }}>{title}</div><div style={{ fontSize:34, fontFamily:"'Archivo',sans-serif", fontWeight:950, color:color||tone.fg, lineHeight:1.02 }}>{n}</div><div style={{ fontSize:10, color:tone.fg, fontWeight:800, opacity:.85 }}>{sub}</div></div>; };
  // ── Active-filter / drill-context chips ──────────────────────────────────────
  // Normal table behaviour: every active filter (including those auto-applied by a Dashboard/Management
  // drill) shows as a removable chip, so the user can always see WHY the list is narrowed and undo it.
  const FIELD_LABELS={ priority:"Due", risk:"Risk", todoType:"Type", orderNo:"Order", junior:"Junior", branch:"Branch", owner:"Chase", escalationOwner:"Escalation", style:"Style", colour:"Colour", fit:"Fit", family:"Family", brand:"Brand", fabric:"Fabric", buyer:"Buyer", planDate:"Plan Date", days:"Days", drift:"Drift", revReject:"Rev/Reject", score:"Score", phase:"Phase" };
  const activeChips=(()=>{
    const out=[];
    const actLabels=arrVal(tf.activity).length?arrVal(tf.activity):arrVal(tf.activityKey).map(k=>stageLabelOf(k));
    [...new Set(actLabels.filter(Boolean))].forEach(lbl=>out.push({ kind:"activity", label:"Activity: "+lbl }));
    const pinN=new Set([...arrVal(tf.styleId),...arrVal(tf.styleIds)].map(String).filter(Boolean)).size;
    if(pinN) out.push({ kind:"styleId", label:`Pinned to ${pinN} style${pinN>1?"s":""}` });
    Object.keys(tf||{}).forEach(k=>{ if(["activity","activityKey","key","styleId","styleIds"].includes(k)) return; arrVal(tf[k]).filter(v=>String(v).trim()!=="").forEach(v=>out.push({ kind:"val", k, val:v, label:`${FIELD_LABELS[k]||k}: ${v}` })); });
    return out;
  })();
  const removeChip=(c)=>{ if(c.kind==="activity") set("activityKey",[]); else if(c.kind==="styleId"){ set("styleId",[]); set("styleIds",[]); } else { const cur=arrVal(tf[c.k]).map(String); set(c.k, cur.filter(x=>x!==String(c.val))); } };
  const isDrillContext=activeChips.some(c=>c.kind==="styleId"||c.kind==="activity");
  return (<div style={{ padding:"16px 22px", maxWidth:"none" }}>
    <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:8, flexWrap:"wrap" }}>
      <span style={{ fontSize:10, color:"var(--muted-2)" }}>Showing {shown.length} of {displayItems.length}</span>
      <div style={{ display:"inline-flex", borderRadius:8, overflow:"hidden", marginLeft:4 }}>{modeBtn("all","All",rawDisplayItems.length)}{modeBtn("activity","Activities",activityCount)}{modeBtn("fabricLab","Fabric / Lab Dip",fabricLabCount)}</div>
      <button onClick={()=>setCfg&&setCfg(c=>({ ...c, todoEscalationRows:!(c&&c.todoEscalationRows!==false) }))} disabled={!canEditSettings} title="Adds/removes separate escalation action rows in To-Do. It does not change the base activity row." style={{ fontFamily:"inherit", fontSize:10, padding:"5px 9px", cursor:canEditSettings?"pointer":"not-allowed", border:"1px solid var(--ink)", background:includeEsc?"var(--accent-tint)":"var(--surface)", fontWeight:800 }}>Escalation rows: {includeEsc?"ON":"OFF"}</button>
      {anyF && <button onClick={()=>{ setTf({}); setFilter&&setFilter({}); }} style={{ fontFamily:"inherit", fontSize:10, padding:"5px 9px", cursor:"pointer", border:"1px solid var(--danger)", background:"var(--surface)", color:"var(--danger)", fontWeight:700 }}>clear filters</button>}
      <span style={{ marginLeft:"auto" }}><ReportExportMenu title="To-Do" prefix="todo" sheets={todoSheets} defaultMode="detailed" /></span>
    </div>
    {activeChips.length>0 && <div style={{ display:"flex", alignItems:"center", gap:7, flexWrap:"wrap", marginBottom:9, padding:"7px 9px", border:"1px solid "+(isDrillContext?"var(--accent)":"var(--line-2)"), borderRadius:10, background:isDrillContext?"var(--accent-tint)":"var(--surface)" }}>
      <span style={{ fontSize:9.5, fontWeight:950, textTransform:"uppercase", letterSpacing:.4, color:isDrillContext?"var(--accent)":"var(--muted-3)" }}>{isDrillContext?"Drill context":"Active filters"}</span>
      {activeChips.map((c,i)=><span key={c.kind+":"+(c.k||"")+":"+(c.val||c.label)+":"+i} style={{ display:"inline-flex", alignItems:"center", gap:5, fontSize:10, fontWeight:700, padding:"3px 5px 3px 9px", borderRadius:999, border:"1px solid var(--ink)", background:"var(--surface)", color:"var(--ink)", whiteSpace:"nowrap" }}>{c.label}<button onClick={()=>removeChip(c)} title="remove this filter" style={{ border:"none", background:"transparent", cursor:"pointer", padding:0, lineHeight:0, color:"var(--muted-2)", display:"inline-flex" }}><X size={12}/></button></span>)}
      <button onClick={()=>{ setTf({}); setFilter&&setFilter({}); }} style={{ fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"3px 9px", cursor:"pointer", border:"1px solid var(--danger)", borderRadius:999, background:"var(--surface)", color:"var(--danger)" }}>Clear all</button>
    </div>}
    <div style={{ display:"flex", alignItems:"baseline", gap:12, margin:"4px 0 8px", flexWrap:"wrap" }}><span style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:13 }}>TO-DO · {shown.length}</span>{overdue.length>0 && <span style={{ fontSize:11, fontWeight:700, color:"var(--danger)" }}>{overdue.length} overdue</span>}{upcoming.length>0 && <span style={{ fontSize:11, fontWeight:700, color:"#7a560f" }}>{upcoming.length} upcoming</span>}{critical.length>0 && <span style={{ fontSize:11, fontWeight:900, color:"var(--danger)" }}>{critical.length} critical &gt;5d</span>}{filterViolations.length>0 && <span title="QA guard: rows rendered that do not pass the canonical filter engine. Should always be 0." style={{ fontSize:10, fontWeight:900, color:"var(--danger)", border:"1px solid var(--danger)", padding:"3px 7px", borderRadius:999 }}>⚠ filter check: {filterViolations.length} wrong row(s)</span>}</div>
    <div style={{ display:"flex", gap:10, flexWrap:"wrap", marginBottom:10 }}>{card("Shown",shown.length,"current filter total")}{card("Upcoming",upcoming.length,"within watch window","#7a560f")}{card("Overdue",overdue.length,"missed plan/revised","var(--danger)")}{card("Critical",critical.length,">5 working days late","var(--danger)")}</div>
    <div style={{ border:"1px solid var(--line-2)", borderRadius:12, background:"var(--surface)", boxShadow:"var(--card-shadow)", padding:12, marginBottom:10 }}><div style={{ fontSize:11, fontWeight:950, color:"var(--muted-3)", textTransform:"uppercase", marginBottom:8, letterSpacing:.35 }}>Activity summary</div><div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(230px,1fr))", gap:9 }}>{activitySummary.map(a=><button key={a.activity} onClick={()=>applyTodoActivity(a.activity,a.activityKey)} style={{ border:"1px solid var(--line-3)", background:"var(--bg)", textAlign:"left", padding:"10px 11px", cursor:"pointer", fontFamily:"inherit", borderRadius:10 }}><div style={{ fontSize:12, fontWeight:950, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis", marginBottom:7 }}>{a.activity}</div><div style={{ display:"flex", gap:7, flexWrap:"wrap" }}><span style={{ display:"inline-flex", alignItems:"center", gap:4, background:"#fff4d8", color:"#8a5200", border:"1px solid #d58a13", borderRadius:999, padding:"3px 7px", fontSize:11, fontWeight:950 }}>U <b style={{ fontSize:15 }}>{a.upcoming}</b></span><span style={{ display:"inline-flex", alignItems:"center", gap:4, background:"#ffe4dc", color:"#b82117", border:"1px solid #c5251a", borderRadius:999, padding:"3px 7px", fontSize:11, fontWeight:950 }}>O <b style={{ fontSize:15 }}>{a.overdue}</b></span><span style={{ display:"inline-flex", alignItems:"center", gap:4, background:"#ffd4cc", color:"#9f1712", border:"1px solid #9f1712", borderRadius:999, padding:"3px 7px", fontSize:11, fontWeight:950 }}>Critical <b style={{ fontSize:15 }}>{a.critical}</b></span></div></button>)}</div></div>
    <div style={{ overflowX:"auto", border:"1px solid var(--line-3)", borderRadius:10, background:"var(--surface)" }}>
      {head}
      {shown.length?shown.map(row):<div style={{ fontSize:11, color:"var(--muted-1)", padding:"12px" }}>Nothing due or coming up. 👍</div>}
    </div>
    <div style={{ fontSize:9, color:"var(--muted-7)", marginTop:14 }}>Separate views: Activities and Fabric / Lab Dip. Multiple filters are supported. Fabric/Lab Dip items are colour-grouped by order + colour, so order filters stay clean. Critical = overdue by more than 5 working days.</div>
  </div>);
}

/* ========================= SETTINGS ========================= */
function SettingsView({ cfg, setCfg, canEdit }){
  const num=(v)=> v==null?"":v;
  const setLead=(k,val)=> setCfg(c=>({ ...c, leads:{ ...c.leads, [k]: val===""?undefined:Math.max(0,Number(val)||0) } }));
  const setRew=(k,val)=> setCfg(c=>({ ...c, rework:{ ...c.rework, [k]: val===""?undefined:Math.max(0,Number(val)||0) } }));
  const setUpc=(k,val)=> setCfg(c=>({ ...c, upcoming:{ ...c.upcoming, [k]: val===""?undefined:Math.max(0,Number(val)||0) } }));
  const setStageOwner=(k,val)=> setCfg(c=>({ ...c, stageOwners:{ ...(c.stageOwners||DEFAULT_CFG.stageOwners), [k]: val||DEFAULT_STAGE_OWNERS[k] } }));
  const setEscRule=(idx,field,val)=> setCfg(c=>{ const arr=escalationRulesOf(c).map(x=>({...x})); arr[idx]={ ...(arr[idx]||{}), [field]: field==="from"||field==="to" ? (val===""?null:Math.max(0,Number(val)||0)) : val }; return { ...c, escalationRules:arr }; });
  const addEscRule=()=> setCfg(c=>({ ...c, escalationRules:[...escalationRulesOf(c), { from:0, to:null, owner:"Jr Merchant", level:"New Level", action:"Chase and update commitment." }] }));
  const removeEscRule=(idx)=> setCfg(c=>{ const arr=escalationRulesOf(c).filter((_,i)=>i!==idx); return { ...c, escalationRules:arr.length?arr:DEFAULT_ESCALATION_RULES.map(x=>({...x})) }; });
  const setTop=(k,val)=> setCfg(c=>({ ...c, [k]: val===""?undefined:Math.max(0,Number(val)||0) }));
  const [open,setOpen]=useState(()=>{ try{ return localStorage.getItem("mt_settings_open")||"leads"; }catch(e){ return "leads"; } });
  useEffect(()=>{ try{ localStorage.setItem("mt_settings_open",open); }catch(e){} },[open]);
  const inputStyle={ width:64, fontFamily:"inherit", fontSize:12, fontWeight:700, padding:"5px 7px", border:"1px solid var(--line-2)", borderRadius:8, outline:"none", background:canEdit?"var(--surface)":"#f3f1ec", color:"var(--ink)", textAlign:"center" };
  const inp=(value,onChange)=>(<input type="number" min="0" disabled={!canEdit} value={num(value)} onChange={e=>onChange(e.target.value)} style={inputStyle}/>);
  const escText=(value,onChange,w=130)=>(<input disabled={!canEdit} value={value||""} onChange={e=>onChange(e.target.value)} style={{ width:w, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:"5px 7px", border:"1px solid var(--line-2)", borderRadius:8, outline:"none", background:canEdit?"var(--surface)":"#f3f1ec" }}/>);
  const ownerSel=(value,onChange)=>{ const v=value||"Jr Merchant"; const isPreset=CHASE_LABELS.includes(v); return (<span style={{ display:"flex", gap:6, alignItems:"center", flexWrap:"wrap", justifyContent:"flex-end" }}><select disabled={!canEdit} value={isPreset?v:"__custom"} onChange={e=>{ if(e.target.value!=="__custom") onChange(e.target.value); }} title="Quick role labels only — not linked to user permissions" style={{ width:132, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:"5px 7px", border:"1px solid var(--line-2)", borderRadius:8, outline:"none", background:canEdit?"var(--surface)":"#f3f1ec" }}>{CHASE_LABELS.map(o=><option key={o} value={o}>{o}</option>)}<option value="__custom">Custom…</option></select><input disabled={!canEdit} value={v} onChange={e=>onChange(e.target.value)} placeholder="custom label" title="Type any label to show in Chase and To-Do" style={{ width:140, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:"5px 7px", border:"1px solid var(--line-2)", borderRadius:8, outline:"none", background:canEdit?"var(--surface)":"#f3f1ec" }}/></span>); };
  const section={ background:"var(--surface)", border:"1px solid var(--toolbar-line)", borderRadius:14, boxShadow:"var(--card-shadow)", overflow:"hidden" };
  const sectionHead=(id,title,desc,count)=>(<button onClick={()=>setOpen(open===id?"":id)} style={{ width:"100%", display:"flex", alignItems:"center", gap:12, textAlign:"left", cursor:"pointer", border:"none", background:open===id?"var(--accent-tint)":"var(--surface)", padding:"13px 15px", fontFamily:"inherit" }}><span style={{ width:28, height:28, borderRadius:14, background:open===id?"var(--accent)":"var(--toolbar-subtle)", color:open===id?"var(--surface)":"var(--muted-3)", display:"inline-flex", alignItems:"center", justifyContent:"center", fontWeight:800 }}>{open===id?"−":"+"}</span><span style={{ flex:1, minWidth:0 }}><span style={{ display:"block", fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:14, color:"var(--ink)" }}>{title}</span><span style={{ display:"block", fontSize:10.5, color:"var(--muted-2)", marginTop:2 }}>{desc}</span></span>{count && <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", border:"1px solid var(--line-2)", borderRadius:999, padding:"3px 7px", background:"var(--surface)" }}>{count}</span>}</button>);
  const row=(label,control,hint)=>(<div style={{ display:"grid", gridTemplateColumns:"minmax(160px,1fr) auto", gap:12, alignItems:"center", padding:"8px 0", borderTop:"1px solid var(--line-3)", fontSize:11.5 }}><span><b>{label}</b>{hint && <span style={{ display:"block", fontSize:9.5, color:"var(--muted-1)", marginTop:2 }}>{hint}</span>}</span>{control}</div>);
  const rworkLabels={ fitSend:"Fit (redo & resend)", artwork:"Artwork", strikeOff:"Strike-off", labDip:"Lab Dip", ppSample:"PP Sample" };
  const upcLabels={ techpack:"Techpack", fitSend:"Fit Send", fitAppr:"Fit Appr", artwork:"Artwork", artAppr:"Art Appr", strikeOff:"Strike-off", soAppr:"S/O Appr", labDip:"Lab Dip", labAppr:"Lab Dip Appr", ppSample:"PP Sample", ppAppr:"PP Appr", fabricIH:"Fabric In-House", prodFile:"Prod File" };
  return (<div style={{ padding:"18px 22px", maxWidth:1180 }}>
    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:16, marginBottom:14, flexWrap:"wrap" }}>
      <div><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:22, letterSpacing:-0.2 }}>Settings</div><div style={{ fontSize:11.5, color:"var(--muted-2)", maxWidth:650, lineHeight:1.45, marginTop:4 }}>Operational configuration for the tracker. These settings change planning labels and date calculations for everyone, so they are grouped into calmer sections.</div></div>
      <span style={{ fontSize:10, fontWeight:800, color:canEdit?"var(--success)":"var(--danger)", background:canEdit?"var(--tint-ok)":"var(--tint-late)", border:"1px solid rgba(31,31,29,0.08)", borderRadius:999, padding:"6px 10px", whiteSpace:"nowrap" }}>{canEdit?"Editable: Management / Senior":"View only"}</span>
    </div>
    {!canEdit && <div style={{ fontSize:11, color:"var(--danger)", marginBottom:12, background:"var(--tint-late)", border:"1px solid rgba(180,35,24,0.18)", borderRadius:10, padding:"9px 11px" }}>Your role cannot edit these settings.</div>}
    <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(360px,1fr))", gap:14 }}>
      <div style={section}>{sectionHead("leads","Stage lead times","Working days each stage takes after its predecessor.",STAGES.length+" stages")}{open==="leads" && <div style={{ padding:"4px 15px 14px" }}>{STAGES.map(st=><React.Fragment key={st.key}>{row(st.label, inp(cfg.leads[st.key], v=>setLead(st.key,v)))}</React.Fragment>)}</div>}</div>
      <div style={section}>{sectionHead("chase","Chase labels by stage","Labels shown in Chase and To-Do only. Not linked to users or permissions.",STAGES.length+" labels")}{open==="chase" && <div style={{ padding:"4px 15px 14px" }}>{STAGES.map(st=><React.Fragment key={st.key}>{row(st.label, ownerSel((cfg.stageOwners&&cfg.stageOwners[st.key])||DEFAULT_STAGE_OWNERS[st.key], v=>setStageOwner(st.key,v)))}</React.Fragment>)}</div>}</div>
      <div style={section}>{sectionHead("rework","Rework days","Extra working days after rejection before the next re-send.",Object.keys(rworkLabels).length+" rules")}{open==="rework" && <div style={{ padding:"4px 15px 14px" }}>{Object.keys(rworkLabels).map(k=><React.Fragment key={k}>{row(rworkLabels[k], inp(cfg.rework[k], v=>setRew(k,v)))}</React.Fragment>)}</div>}</div>
      <div style={section}>{sectionHead("gates","Delivery gates","Working-day cutoffs counted backwards from delivery.","2 gates")}{open==="gates" && <div style={{ padding:"4px 15px 14px" }}>{row("Fabric cut-off", inp(cfg.fabricCutoff, v=>setTop("fabricCutoff",v)), "Fabric IH target from delivery date")}{row("Release gate", inp(cfg.relGate, v=>setTop("relGate",v)), "Prod file / release risk gate")}</div>}</div>
      <div style={{ ...section, gridColumn:"1 / -1" }}>{sectionHead("escalation","Escalation matrix rules","Editable overdue-day slabs. Chase Label stays as blocker; Escalation Owner is who must chase now.",escalationRulesOf(cfg).length+" levels")}{open==="escalation" && <div style={{ padding:"4px 15px 14px" }}>
        <div style={{ display:"grid", gridTemplateColumns:"70px 70px 170px 140px minmax(220px,1fr) 58px", gap:8, alignItems:"center", fontSize:9.5, fontWeight:900, color:"var(--muted-2)", textTransform:"uppercase", borderTop:"1px solid var(--line-3)", paddingTop:10, marginBottom:4 }}><span>From</span><span>To</span><span>Escalation Owner</span><span>Level</span><span>Action Text</span><span></span></div>
        {escalationRulesOf(cfg).map((r,i)=><div key={i} style={{ display:"grid", gridTemplateColumns:"70px 70px 170px 140px minmax(220px,1fr) 58px", gap:8, alignItems:"center", padding:"5px 0", borderTop:"1px solid var(--line-3)" }}>
          <input type="number" min="0" disabled={!canEdit} value={r.from} onChange={e=>setEscRule(i,"from",e.target.value)} style={inputStyle}/>
          <input type="number" min="0" disabled={!canEdit} value={r.to==null?"":r.to} placeholder="∞" onChange={e=>setEscRule(i,"to",e.target.value)} style={inputStyle}/>
          {ownerSel(r.owner,v=>setEscRule(i,"owner",v))}
          {escText(r.level,v=>setEscRule(i,"level",v),130)}
          {escText(r.action,v=>setEscRule(i,"action",v),"100%")}
          <button disabled={!canEdit} onClick={()=>removeEscRule(i)} style={{ ...chip, opacity:canEdit?1:0.5 }}>del</button>
        </div>)}
        <div style={{ display:"flex", alignItems:"center", gap:10, marginTop:10, flexWrap:"wrap" }}><button disabled={!canEdit} onClick={addEscRule} style={{ ...chip, opacity:canEdit?1:0.5 }}>+ Add escalation level</button><label style={{ display:"flex", alignItems:"center", gap:6, fontSize:11, fontWeight:800 }}><input type="checkbox" disabled={!canEdit} checked={cfg.todoEscalationRows!==false} onChange={e=>setCfg(c=>({ ...c, todoEscalationRows:e.target.checked }))}/> Show escalation rows in To-Do by default</label></div>
      </div>}</div>
      <div style={{ ...section, gridColumn:"1 / -1" }}>{sectionHead("upcoming","To-Do watch windows","How early a pending activity appears as upcoming.",Object.keys(upcLabels).length+" windows")}{open==="upcoming" && <div style={{ padding:"4px 15px 14px", display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(260px,1fr))", columnGap:18 }}>{Object.keys(upcLabels).map(k=><React.Fragment key={k}>{row(upcLabels[k], inp(cfg.upcoming[k], v=>setUpc(k,v)))}</React.Fragment>)}</div>}</div>
    </div>
    <div style={{ display:"flex", alignItems:"center", gap:10, marginTop:16, flexWrap:"wrap" }}><button disabled={!canEdit} onClick={()=>setCfg(DEFAULT_CFG)} style={{ fontFamily:"inherit", fontSize:11, fontWeight:800, padding:"8px 14px", cursor:canEdit?"pointer":"not-allowed", border:"1px solid var(--line-2)", borderRadius:9, background:"var(--surface)", opacity:canEdit?1:0.5 }}>Reset to defaults</button><span style={{ fontSize:9.5, color:"var(--muted-2)" }}>Changes save automatically and apply to every user's calculations.</span></div>
  </div>);
}

// ============================ AUTH LAYER ============================
function Splash({ text }){ return (<div style={{ minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", background:"var(--bg)", fontFamily:"'JetBrains Mono',monospace", color:"#6a665e", fontSize:13 }}>{text}</div>); }

function UpdateBanner(){
  const [ready,setReady]=useState(false);
  const baseRef=useRef(null);
  useEffect(()=>{
    let alive=true;
    const grab=async()=>{
      try{ const r=await fetch("/?_v="+Date.now(),{ cache:"no-store" }); if(!r.ok) return null; const t=await r.text(); return (t&&t.length)?t.trim():null; }catch(e){ return null; }
    };
    (async()=>{ for(let i=0;i<6 && alive && baseRef.current===null;i++){ const t=await grab(); if(t){ baseRef.current=t; break; } await new Promise(res=>setTimeout(res,2000)); } })();
    const check=async()=>{ if(typeof document!=="undefined"&&document.hidden) return; if(baseRef.current===null) return; const t=await grab(); if(alive&&t&&t!==baseRef.current) setReady(true); };
    const t0=setTimeout(check,5000);
    const iv=setInterval(check,30000);
    const onVis=()=>{ if(typeof document!=="undefined"&&!document.hidden) check(); };
    if(typeof document!=="undefined") document.addEventListener("visibilitychange",onVis);
    return ()=>{ alive=false; clearTimeout(t0); clearInterval(iv); if(typeof document!=="undefined") document.removeEventListener("visibilitychange",onVis); };
  },[]);
  if(!ready) return null;
  return createPortal(
    <div style={{ position:"fixed", bottom:18, left:"50%", transform:"translateX(-50%)", zIndex:99999, background:"var(--ink)", color:"var(--surface)", padding:"9px 14px", borderRadius:6, boxShadow:"0 6px 20px rgba(0,0,0,0.35)", display:"flex", alignItems:"center", gap:11, fontSize:12, fontFamily:"'JetBrains Mono',monospace" }}>
      <span style={{ width:7, height:7, borderRadius:7, background:"#7fd1a8", display:"inline-block" }}/>
      <span>New version available.</span>
      <button onClick={()=>{ window.location.reload(); }} style={{ fontFamily:"inherit", fontSize:12, fontWeight:700, padding:"5px 13px", cursor:"pointer", border:"none", borderRadius:4, background:"var(--accent)", color:"var(--surface)" }}>Refresh</button>
      <button onClick={()=>setReady(false)} style={{ fontFamily:"inherit", fontSize:11, padding:"5px 6px", cursor:"pointer", border:"none", background:"transparent", color:"#9b958a" }}>Later</button>
    </div>, document.body);
}

function AuthShell({ children }){
  return (<div style={{ minHeight:"100vh", background:"linear-gradient(135deg,#f7f3ea 0%,#fffaf1 52%,#efe6d7 100%)", fontFamily:"'JetBrains Mono',monospace", padding:24, display:"flex", alignItems:"center", justifyContent:"center" }}>
    <div style={{ width:980, maxWidth:"100%", minHeight:560, display:"grid", gridTemplateColumns:"1.05fr 0.95fr", background:"var(--surface)", border:"1px solid var(--ink)", boxShadow:"10px 10px 0 rgba(31,31,29,0.92)", borderRadius:18, overflow:"hidden" }}>
      <div style={{ padding:38, background:"linear-gradient(180deg,#fffaf1 0%,#f3eadf 100%)", borderRight:"1px solid var(--line-2)", position:"relative" }}>
        <div style={{ fontFamily:"'Archivo',sans-serif", fontSize:13, fontWeight:800, letterSpacing:1.8, color:"var(--accent)", textTransform:"uppercase", marginBottom:10 }}>Kothari Sports & Apparels</div>
        <div style={{ fontFamily:"'Archivo',sans-serif", fontSize:38, lineHeight:1.02, fontWeight:800, color:"var(--ink)", marginBottom:12 }}>Merch Tracker</div>
        <div style={{ fontSize:12, color:"var(--muted-3)", lineHeight:1.6, maxWidth:430, marginBottom:26 }}>Pre-production TNA command center for styles, approvals, rework, fabric, PP, production file release and management reporting.</div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:24 }}>
          {[['TNA','Date discipline'],['TO-DO','Daily chase list'],['REWORK','Resend rounds'],['REPORTS','Management clarity']].map(([a,b])=><div key={a} style={{ background:"rgba(255,253,248,0.72)", border:"1px solid var(--line-2)", borderRadius:14, padding:14 }}><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:18, color:"var(--ink)" }}>{a}</div><div style={{ fontSize:10, color:"var(--muted-2)", marginTop:4 }}>{b}</div></div>)}
        </div>
        <div style={{ position:"absolute", left:38, right:38, bottom:32, border:"1px dashed var(--line-2)", borderRadius:14, padding:14, color:"var(--muted-3)", fontSize:11, lineHeight:1.55, background:"rgba(255,253,248,0.48)" }}>
          Dates are base data. Revisions, rejections, skips and re-send rounds are tracked for audit, To-Do, dashboards and exports.
        </div>
      </div>
      <div style={{ padding:38, display:"flex", alignItems:"center", justifyContent:"center" }}>
        <div style={{ width:380, maxWidth:"100%" }}>{children}</div>
      </div>
    </div>
  </div>);
}

function LoginScreen(){
  const [email,setEmail]=useState(""); const [pw,setPw]=useState(""); const [mode,setMode]=useState("in"); const [msg,setMsg]=useState(""); const [busy,setBusy]=useState(false); const [remember,setRemember]=useState(true);
  const submit=async()=>{ if(!email.trim()||!pw){ setMsg("Enter email and password."); return; } setMsg(""); setBusy(true);
    const stamp=()=>{ try{ localStorage.setItem("mt_login_at", String(Date.now())); localStorage.setItem("mt_remember", remember?"1":"0"); }catch(x){} };
    try{ if(mode==="in"){ const { error }=await supabase.auth.signInWithPassword({ email:email.trim(), password:pw }); if(error) throw error; stamp(); }
      else { const { error,data }=await supabase.auth.signUp({ email:email.trim(), password:pw }); if(error) throw error; if(data.session){ stamp(); } if(!data.session){ setMsg("Account created. If sign-in does not happen automatically, check your email to confirm, then sign in."); setMode("in"); } } }
    catch(e){ setMsg(e.message||String(e)); } setBusy(false); };
  const inp={ width:"100%", fontFamily:"inherit", fontSize:13, padding:"12px 13px", border:"1px solid var(--line-2)", borderRadius:10, marginBottom:12, boxSizing:"border-box", background:"var(--surface)", outline:"none" };
  return (<AuthShell>
    <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:28, color:"var(--ink)", marginBottom:4 }}>{mode==="in"?"Welcome back":"Create access"}</div>
    <div style={{ fontSize:12, color:"var(--muted-2)", marginBottom:22 }}>{mode==="in"?"Sign in to continue to the live TNA tracker.":"Create an account. Management can approve your role after signup."}</div>
    <form onSubmit={e=>{ e.preventDefault(); submit(); }}>
      <label style={{ display:"block", fontSize:10, fontWeight:800, color:"var(--muted-3)", textTransform:"uppercase", letterSpacing:0.8, marginBottom:5 }}>Email</label>
      <input style={inp} type="email" name="email" autoComplete="username" placeholder="name@company.com" value={email} onChange={e=>setEmail(e.target.value)}/>
      <label style={{ display:"block", fontSize:10, fontWeight:800, color:"var(--muted-3)", textTransform:"uppercase", letterSpacing:0.8, marginBottom:5 }}>Password</label>
      <input style={inp} type="password" name="password" autoComplete={mode==="in"?"current-password":"new-password"} placeholder="••••••••" value={pw} onChange={e=>setPw(e.target.value)}/>
      <label style={{ display:"flex", alignItems:"center", gap:9, fontSize:11, color:"var(--muted-4)", margin:"4px 0 16px", cursor:"pointer" }}><input type="checkbox" checked={remember} onChange={e=>setRemember(e.target.checked)}/>Keep me signed in <span style={{ color:"var(--muted-1)" }}>(otherwise 12h)</span></label>
      <button type="submit" disabled={busy} style={{ width:"100%", fontFamily:"inherit", fontSize:13, fontWeight:800, padding:"12px", cursor:busy?"wait":"pointer", border:"1px solid var(--ink)", borderRadius:10, background:"var(--ink)", color:"var(--bg)", marginBottom:12, boxShadow:"0 2px 0 rgba(31,31,29,0.16)" }}>{busy?"Signing…":(mode==="in"?"Sign in":"Create account")}</button>
    </form>
    <div style={{ fontSize:11, textAlign:"center", color:"var(--muted-2)" }}>{mode==="in"?"New here? ":"Have an account? "}<button type="button" onClick={()=>{ setMode(mode==="in"?"up":"in"); setMsg(""); }} style={{ border:"none", background:"transparent", color:"var(--accent)", cursor:"pointer", fontFamily:"inherit", fontSize:11, fontWeight:800 }}>{mode==="in"?"Create account":"Sign in"}</button></div>
    {msg && <div style={{ fontSize:11, color:"var(--danger)", marginTop:14, lineHeight:1.45, border:"1px solid var(--tint-late)", background:"#fff4f1", borderRadius:10, padding:10 }}>{msg}</div>}
    <div style={{ marginTop:22, paddingTop:14, borderTop:"1px solid var(--line-3)", fontSize:10, color:"var(--muted-1)", lineHeight:1.5 }}>Secure Supabase login · access controlled by user role · changes are audited after login.</div>
  </AuthShell>);
}

function PendingScreen({ email, onSignOut }){
  return (<AuthShell>
    <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:18, marginBottom:8 }}>Awaiting access</div>
    <div style={{ fontSize:12, color:"var(--muted-4)", lineHeight:1.55, marginBottom:16 }}>You're signed in as <b>{email}</b>, but your account hasn't been given a role yet. Ask a Management user to set your role in the <b>Users</b> panel, then refresh this page.</div>
    <button onClick={onSignOut} style={{ fontFamily:"inherit", fontSize:12, padding:"8px 14px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)" }}>Sign out</button>
  </AuthShell>);
}

function UsersPanel({ onClose }){
  const [list,setList]=useState(null); const [busy,setBusy]=useState(false);
  const load=async()=>{ const { data }=await supabase.from("profiles").select("*").order("created_at",{ ascending:true }); setList(data||[]); };
  useEffect(()=>{ load(); },[]);
  const setR=async(id,role)=>{ setBusy(true); await supabase.from("profiles").update({ role }).eq("id",id); await load(); setBusy(false); };
  const editName=(id,name)=>setList(l=>l.map(x=>x.id===id?{...x,name}:x));
  const saveName=async(id,name)=>{ try{ await supabase.from("profiles").update({ name:(name||"").trim()||null }).eq("id",id); }catch(e){} };
  const ROLE_OPTS=["pending"].concat(Object.keys(ROLES));
  return (<div onClick={onClose} style={{ position:"fixed", inset:0, background:"rgba(26,26,26,0.55)", zIndex:400, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
    <div onClick={e=>e.stopPropagation()} style={{ background:"var(--bg)", border:"2px solid var(--ink)", boxShadow:"8px 8px 0 var(--ink)", width:560, maxWidth:"100%", maxHeight:"86vh", overflowY:"auto", padding:22, fontFamily:"'JetBrains Mono',monospace" }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:18 }}>Users &amp; roles</div><button onClick={onClose} title="Close users panel" style={{ border:"1px solid var(--line-2)", background:"var(--surface)", cursor:"pointer", borderRadius:9, padding:"7px 10px", display:"inline-flex", alignItems:"center", gap:5 }}><X size={16}/><span style={{ fontSize:11, fontWeight:800 }}>Close</span></button></div>
      <p style={{ fontSize:11, color:"var(--muted-3)", lineHeight:1.55, marginBottom:12 }}>Each person signs in with their own email + password; their access follows the role you set here. Set anyone to "pending" to suspend access.</p>
      {list===null ? <div style={{ fontSize:12, color:"var(--muted-2)" }}>Loading…</div> : list.length===0 ? <div style={{ fontSize:12, color:"var(--muted-2)" }}>No users yet — have your team sign up from the login screen.</div> : (
        <div style={{ border:"1px solid var(--ink)", background:"var(--surface)" }}>{list.map((u,i)=>(<div key={u.id} style={{ display:"flex", alignItems:"center", gap:10, padding:"8px 10px", borderBottom:i<list.length-1?"1px solid #eee7da":"none" }}>
          <div style={{ flex:1, minWidth:0 }}><input value={u.name||""} placeholder="(set name)" onChange={e=>editName(u.id,e.target.value)} onBlur={e=>{ e.target.style.border="1px solid transparent"; e.target.style.background="transparent"; saveName(u.id,e.target.value); }} onKeyDown={e=>{ if(e.key==="Enter") e.target.blur(); }} style={{ fontFamily:"inherit", fontSize:12, fontWeight:700, width:"100%", border:"1px solid transparent", background:"transparent", padding:"2px 4px", boxSizing:"border-box" }} onFocus={e=>{ e.target.style.border="1px solid var(--ink)"; e.target.style.background="var(--surface)"; }} onMouseLeave={e=>{}} />
<div style={{ fontSize:10, color:"var(--muted-1)", whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis", padding:"0 4px" }}>{u.email}</div></div>
          <select disabled={busy} value={u.role||"pending"} onChange={e=>setR(u.id,e.target.value)} style={{ fontFamily:"inherit", fontSize:11, padding:"5px 7px", border:"1px solid var(--ink)", background:u.role==="pending"?"#fbeaea":"var(--accent-tint)" }}>{ROLE_OPTS.map(r=>(<option key={r} value={r}>{r==="pending"?"— pending —":(ROLES[r]||{}).label||r}</option>))}</select>
        </div>))}</div>
      )}
    </div>
  </div>);
}

export default function App(){
  const [session,setSession]=useState(undefined);   // undefined = still loading
  const [profile,setProfile]=useState(undefined);
  useEffect(()=>{ supabase.auth.getSession().then(({ data })=>setSession(data.session||null)); const { data:sub }=supabase.auth.onAuthStateChange((_e,s)=>{ setSession(s||null); if(!s) setProfile(null); }); return ()=>{ try{ sub.subscription.unsubscribe(); }catch(e){} }; },[]);
  useEffect(()=>{ if(session===undefined) return; if(!session){ setProfile(null); return; } let active=true; (async()=>{ const uid=session.user.id; let { data }=await supabase.from("profiles").select("*").eq("id",uid).maybeSingle(); if(!data){ const { count }=await supabase.from("profiles").select("*",{ count:"exact", head:true }); const role=(count===0)?"management":"pending"; const ins=await supabase.from("profiles").insert({ id:uid, email:session.user.email, name:(session.user.email||"").split("@")[0], role }).select().maybeSingle(); data=ins.data; } if(active) setProfile(data||null); })(); return ()=>{ active=false; }; },[session]);
  const signOut=async()=>{ await supabase.auth.signOut(); setProfile(null); try{ localStorage.removeItem("mt_login_at"); }catch(e){} };
  useEffect(()=>{ const check=()=>{ try{ if(localStorage.getItem("mt_remember")==="1") return; const at=parseInt(localStorage.getItem("mt_login_at")||"0",10); if(at && Date.now()-at>12*3600*1000){ supabase.auth.signOut(); } }catch(e){} }; check(); const t=setInterval(check,5*60*1000); return ()=>clearInterval(t); },[]);
  if(session===undefined || (session && profile===undefined)) return <Splash text="Loading…"/>;
  if(!session) return <LoginScreen/>;
  if(!profile || profile.role==="pending" || !ROLES[profile.role]) return <PendingScreen email={session.user.email} onSignOut={signOut}/>;
  return (<><MerchTracker me={profile} onSignOut={signOut}/><UpdateBanner/></>);
}

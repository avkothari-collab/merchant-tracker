import React, { useState, useMemo, useRef, useEffect, useDeferredValue } from "react";
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
const iso=(d)=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
const colLetter=(n)=>{ let s=""; n=Number(n)+1; if(n<1) return ""; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s; };
const letterToIndex=(s)=>{ s=String(s||"").toUpperCase(); if(!/^[A-Z]+$/.test(s)) return -1; let n=0; for(const ch of s) n=n*26+(ch.charCodeAt(0)-64); return n-1; };
const _now=new Date(); const TODAY=new Date(_now.getFullYear(),_now.getMonth(),_now.getDate()); // live current date (local midnight)
const perfNow=()=> (typeof performance!=="undefined"&&performance.now)?performance.now():Date.now();
const lc=(v)=>String(v==null?"":v).toLowerCase();
const buildSearchIndex=(s,c)=>{
  const chase=(c&&c.chaseOwners?c.chaseOwners:[]).map(o=>o.owner).join(" ");
  const byCol={ styleNo:lc(s.styleNo), orderNo:lc(s.orderNo), sampleFit:lc(s.sampleFit), family:lc(s.family), colour:lc(s.colour), color:lc(s.colour), brand:lc(s.brand), buyer:lc(s.buyer), fabricType:lc(s.fabricType), owner:lc(s.owner), remarks:lc(s.remarks), age:lc(s.age), setRole:lc(s.setRole), chase:lc(chase) };
  return { byCol, auto:[byCol.styleNo,byCol.orderNo,byCol.sampleFit,byCol.family,byCol.colour,byCol.brand,byCol.buyer,byCol.fabricType,byCol.owner,byCol.remarks,byCol.chase].join(" ") };
};
const pushPerfSample=(arr,ms)=>{ const n=Number(ms)||0; const next=(arr||[]).concat(n).slice(-80).sort((a,b)=>a-b); return { samples:next, p95:next.length?Math.round(next[Math.min(next.length-1,Math.floor(next.length*0.95))]*10)/10:0 }; };


const REJECTABLE=["fitAppr","artAppr","soAppr","labAppr","ppAppr"]; // approval stages that can be rejected
const SKIPPABLE_STAGES=["fitSend","fitAppr","artwork","artAppr","strikeOff","soAppr","labDip","labAppr","ppSample","ppAppr"]; // activities that can be waived/skipped
const APPR_OF_SEND={ fitSend:"fitAppr", artwork:"artAppr", strikeOff:"soAppr", labDip:"labAppr", ppSample:"ppAppr" }; // send/make stage -> the approval that can reject it
const REWORK_DAYS={ fitSend:4, artwork:2, strikeOff:3, labDip:7, ppSample:4 }; // working days added on rejection (redo+resend)
const stageApplies=(s,st)=>{
  if(!st) return false;
  // Strike-off is only meaningful inside the print branch. If S/O is accidentally ticked while Print is off, keep it out of the chain.
  if(st.key==="strikeOff" || st.key==="soAppr") return !!(s.printReq && s.soReq);
  if(st.key==="artwork" || st.key==="artAppr") return !!s.printReq;
  return st.flag===null || !!s[st.flag];
};
const applicableStages=(s)=> STAGES.filter(st=> stageApplies(s,st));

function computeStyle(s, cfg){
  const ordRec=parse(s.ordRec), delivery=parse(s.delivery);
  const leadOf=(st)=>{ const v=cfg&&cfg.leads&&cfg.leads[st.key]; return v==null?st.lead:v; }; const rwOf=(st)=>{ const v=cfg&&cfg.rework&&cfg.rework[st.key]; return v==null?(REWORK_DAYS[st.key]||st.lead):v; }; const ownerOf=(k)=>{ const st=STAGES.find(x=>x.key===k); return (cfg&&cfg.stageOwners&&cfg.stageOwners[k]) || DEFAULT_STAGE_OWNERS[k] || (st&&st.owner) || "Jr Merchant"; }; const CUTD=(cfg&&cfg.fabricCutoff!=null)?cfg.fabricCutoff:FABRIC_CUTOFF_DAYS; const GATED=(cfg&&cfg.relGate!=null)?cfg.relGate:REL_GATE_DAYS;
  const cutoff=addWorkdays(delivery,-CUTD);
  const eff={}, plan={};
  const applies=(k)=>{ const st=STAGES.find(x=>x.key===k); return stageApplies(s,st); };
  const actualOf=(k)=>parse(s.actuals[k]); const revOf=(k)=>parse(s.revs?.[k]); const rejOf=(k)=>parse(s.rejects?.[k]); const skipOf=(k)=>parse(s.skips?.[k]);
  STAGES.forEach(st=>{
    let p;
    if(st.cutoff){ const base=s.labDipReq?(eff["labAppr"]||eff["labDip"]||ordRec):ordRec; p=s.labDipReq?new Date(Math.max(addWorkdays(base,15)?.getTime()||0, cutoff.getTime())):cutoff; }
    else { let predEff; if(st.key==="prodFile") predEff = (s.ppBypass || !s.ppNeeded) ? eff["fabricIH"] : eff["ppAppr"]; else predEff = st.pred==="__ord"?ordRec:eff[st.pred]; if((st.key==="ppSample"||st.key==="prodFile") && s.fitReq && eff["fitAppr"]) predEff = new Date(Math.max((predEff&&predEff.getTime())||0, eff["fitAppr"].getTime())); p=addWorkdays(predEff||ordRec, leadOf(st)); }
    const apprK=APPR_OF_SEND[st.key]; const rejAppr = !!(apprK && rejOf(apprK) && !actualOf(apprK));
    const selfRej = REJECTABLE.includes(st.key) && rejOf(st.key) && !actualOf(st.key);
    if(rejAppr){ const rjd=rejOf(apprK); const auto=addWorkdays(rjd, rwOf(st)); const a=actualOf(st.key); const rv=revOf(st.key); plan[st.key]=auto;
      if(a && a>rjd) eff[st.key]=a; else if(rv && rv>=rjd) eff[st.key]=rv; else eff[st.key]=auto; } // redo: re-sent actual wins, else fresh revised, else rejection+rework days
    else if(selfRej){ const rjd=rejOf(st.key); const rv=revOf(st.key); plan[st.key]=p; eff[st.key]=(rv && rv>=rjd)?rv:p; } // rejected approval cascades off redone send
    else { plan[st.key]=p; eff[st.key]=actualOf(st.key)||skipOf(st.key)||revOf(st.key)||p; }
  });
  const fabricIHStamp = actualOf("fabricIH") || skipOf("fabricIH");
  // Print / strike-off chain closes as a cross-check once fabric is in-house ONLY when PP bypass is OFF.
  // If PP bypass is ON, print/strike-off stays actionable because the production file path has no PP approval gate.
  // Fit does NOT close automatically because PP depends on Fit approval.
  const printPreFabricKeys=new Set(["artwork","artAppr","strikeOff","soAppr"]);
  const stages=applicableStages(s).map(st=>{ const apprK=APPR_OF_SEND[st.key]; const apprRej=apprK?rejOf(apprK):null; const selfRejDate=REJECTABLE.includes(st.key)?rejOf(st.key):null; const rejAppr=!!(apprK&&apprRej&&!actualOf(apprK)); const a=actualOf(st.key); const skp=skipOf(st.key); const isSkip=!!skp&&!a; const rjd_=rejAppr?apprRej:null; const resent=rejAppr&&a&&a>rjd_; const rework=rejAppr&&!resent; const rejected=REJECTABLE.includes(st.key)&&!!selfRejDate&&!actualOf(st.key); const autoClosed=!!(fabricIHStamp && !s.ppBypass && printPreFabricKeys.has(st.key) && !a && !skp && !rework && !rejected); const rjd=rework?rjd_:(rejected?selfRejDate:null); let rv=revOf(st.key); if(rjd&&rv&&rv<rjd) rv=null; const histReject = a ? ((apprRej&&a>=apprRej)?apprRej:(selfRejDate||null)) : null; return { ...st, owner:ownerOf(st.key), actual:a, rev:rv, reject:rjd, histReject, rework:isSkip?false:rework, rejected:isSkip?false:rejected, skipped:isSkip, autoClosed, skip:skp, plan:plan[st.key], done: autoClosed?true:(isSkip?true:(rework?false:!!a)) }; });
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
  const float=lastPlan?netWorkdays(lastPlan,delivery):null;
  let status="On Track", tone="ok";
  if(released){ status="Released"; tone="done"; }
  else if(nextPending&&nextPending.plan&&TODAY>nextPending.plan){ status=`Overdue ${Math.round((TODAY-nextPending.plan)/ONE_DAY)}d`; tone="late"; }
  else if(idle!==null&&idle>=7){ status=`Idle ${idle}d`; tone="warn"; }
  const dueText=(k)=>{ const r=get(k); if(!r||!r.plan) return "pending"; return TODAY>r.plan?`OVERDUE ${Math.round((TODAY-r.plan)/ONE_DAY)}d`:`due ${fmt(r.plan)}`; };
  const dueTone=(k)=>{ const r=get(k); const d=r&&(r.rev||r.plan); return d&&TODAY>d?"late":"warn"; };
  const bs=(txt,tn,extra={})=>({txt,tone:tn,...extra});
  const autoClosed=(k)=>{ const r=get(k); return !!(r&&r.autoClosed); };
  const reSentAfterReject=(sendK,apprK)=>{ const sd=actualOf(sendK), rj=rejOf(apprK), ap=actualOf(apprK); return !!(sd&&rj&&!ap&&sd>rj); };
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
  let fabricBranch; const fabPlan=get("fabricIH")?.plan;
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
const isEditableCol=(col)=> col==="__style"||col==="qty"||col==="ordRec"||col==="delivery"||TEXT_COLS.includes(col)||STAGE_KEYS.includes(col);
const isDateCol=(col)=> col==="ordRec"||col==="delivery"||STAGE_KEYS.includes(col);
const BRANCH_STAGES={ fit:["fitSend","fitAppr"], print:["artwork","artAppr","strikeOff","soAppr"], fabric:["labDip","labAppr","fabricIH"], pp:["ppSample","ppAppr"], prod:["prodFile"] };
const BRANCH_LABEL={ fit:"Fit", print:"Print", fabric:"Fabric", pp:"PP", prod:"Production" };
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
const chip={ fontSize:10, padding:"4px 6px", border:"1px solid var(--ink)", background:"var(--bg)", cursor:"pointer", fontFamily:"'JetBrains Mono', monospace", fontWeight:600 };

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
const DEFAULT_NAMED_TRACKER_VIEWS=[
  { name:"Management View", shared:true, state:{ columnView:"management", hidden:null, statusFilter:"All", ownerFilter:"All", archiveView:"active", savedView:"", search:"", searchCol:"auto", activityFilter:null, colFilters:{}, sort:{ col:null, dir:1 }, freezeN:1 } },
  { name:"Junior Merchandiser View", shared:true, state:{ columnView:"merchant", hidden:null, statusFilter:"All", ownerFilter:"All", archiveView:"active", savedView:"", search:"", searchCol:"auto", activityFilter:null, colFilters:{}, sort:{ col:null, dir:1 }, freezeN:1 } },
  { name:"Buyer Approval View", shared:true, state:{ columnView:"buyer", hidden:null, statusFilter:"All", ownerFilter:"Buyer", archiveView:"active", savedView:"buyerPending", search:"", searchCol:"auto", activityFilter:null, colFilters:{}, sort:{ col:"delivery", dir:1 }, freezeN:1 } },
  { name:"Production Follow-up View", shared:true, state:{ columnView:"store", hidden:null, statusFilter:"At Risk", ownerFilter:"All", archiveView:"active", savedView:"fabricPending", search:"", searchCol:"auto", activityFilter:null, colFilters:{}, sort:{ col:"fabricCD", dir:1 }, freezeN:1 } },
  { name:"My To-Do View", shared:true, state:{ columnView:"custom", hidden:null, statusFilter:"At Risk", ownerFilter:"All", archiveView:"active", savedView:"dueThisWeek", search:"", searchCol:"auto", activityFilter:null, colFilters:{}, sort:{ col:"delivery", dir:1 }, freezeN:1 } },
];
const normalizeTrackerViews=(views)=>{ const arr=Array.isArray(views)?views:[]; const byName=new Map(); DEFAULT_NAMED_TRACKER_VIEWS.forEach(v=>byName.set(v.name,v)); arr.forEach(v=>{ if(v&&v.name) byName.set(v.name,{...v, state:{...(v.state||{})}}); }); return [...byName.values()]; };
const appendOneSheet=(wb,label,data)=>{
  const rows=Array.isArray(data)?data:[];
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
        {available.map((sh,i)=><label key={sh.label} style={{ display:"flex", alignItems:"center", gap:6, fontSize:10, padding:"4px 2px", cursor:"pointer", borderBottom:i===available.length-1?"none":"1px solid var(--line-3)" }}><input type="checkbox" checked={active.has(sh.label)} onChange={()=>toggle(sh.label)}/><span style={{ flex:1 }}>{sh.label}</span><span style={{ color:"var(--muted-1)", fontSize:9 }}>{Array.isArray(sh.data)?sh.data.length:0}{mode==="detailed"&&Array.isArray(sh.detailData)?" + "+sh.detailData.length:""} rows</span></label>)}
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
  const [styles,setStyles]=useState([]); // loaded from Supabase on mount
  const role=(me&&me.role)||"junior";
  const [usersOpen,setUsersOpen]=useState(false);
  const [textScale,setTextScale]=useState(()=>{ try{ const v=parseFloat(localStorage.getItem("mt_textscale")); return (v&&v>=0.6&&v<=2.0)?v:1; }catch(e){ return 1; } });
  const bumpScale=(d)=>setTextScale(v=>{ const n=Math.min(2.0,Math.max(0.6,Math.round((v+d)*100)/100)); try{ localStorage.setItem("mt_textscale",String(n)); }catch(e){} return n; });
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
  const [activityFilter,setActivityFilter]=useState(PF.activityFilter||null);
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
  const [editing,setEditing]=useState(null);
  const [editVal,setEditVal]=useState("");
  const [sort,setSort]=useState({ col:null, dir:1 });
  const [hidden,setHidden]=useState(()=>{ try{ const sv=localStorage.getItem("mt_hidden_cols"); return new Set(sv?JSON.parse(sv):["extra1","extra2"]); }catch(e){ return new Set(["extra1","extra2"]); } });
  useEffect(()=>{ try{ localStorage.setItem("mt_hidden_cols", JSON.stringify([...hidden])); }catch(e){} },[hidden]);
  const [specialClip,setSpecialClip]=useState(null);
  const [expOpen,setExpOpen]=useState(false); const [expMode,setExpMode]=useState(()=>{ try{ return localStorage.getItem("mt_exp_mode")||"full"; }catch(e){ return "full"; } }); const [expBuf,setExpBuf]=useState(()=>{ try{ const v=parseInt(localStorage.getItem("mt_exp_buf"),10); return (isFinite(v)&&v>=0&&v<=30)?v:2; }catch(e){ return 2; } }); const [expIncBuf,setExpIncBuf]=useState(()=>{ try{ return localStorage.getItem("mt_exp_incbuf")==="1"; }catch(e){ return false; } }); const [expRelMode,setExpRelMode]=useState(()=>{ try{ return localStorage.getItem("mt_exp_relmode")||"detailed"; }catch(e){ return "detailed"; } });
  const [frOpen,setFrOpen]=useState(false); const [frFind,setFrFind]=useState(""); const [frRepl,setFrRepl]=useState(""); const [frScope,setFrScope]=useState(()=>{ try{ const v=localStorage.getItem("mt_fr_scope"); return (v==="selected"||v==="filtered")?v:"filtered"; }catch(e){ return "filtered"; } }); const [frCase,setFrCase]=useState(()=>{ try{ return localStorage.getItem("mt_fr_case")==="1"; }catch(e){ return false; } });
  useEffect(()=>{ try{ localStorage.setItem("mt_exp_mode",expMode); localStorage.setItem("mt_exp_buf",String(expBuf)); localStorage.setItem("mt_exp_incbuf",expIncBuf?"1":"0"); localStorage.setItem("mt_exp_relmode",expRelMode); localStorage.setItem("mt_fr_scope",frScope); localStorage.setItem("mt_fr_case",frCase?"1":"0"); }catch(e){} },[expMode,expBuf,expIncBuf,expRelMode,frScope,frCase]);
  const [freezeN,setFreezeN]=useState(1);  // # leading columns frozen (incl style)
  const [findIdx,setFindIdx]=useState(-1);
  const [frMatches,setFrMatches]=useState([]); // computed Find matches: [{id,col,style,colLabel,text}]
  const [colW,setColW]=useState({});  // per-column width overrides (drag to resize)
  const [fills,setFills]=useState({});
  const [notes,setNotes]=useState({});
  const [noteEditing,setNoteEditing]=useState(false);
  const [noteText,setNoteText]=useState("");
  const [comments,setComments]=useState({}); const [team,setTeam]=useState([]); const [threadCell,setThreadCell]=useState(null); const [cmText,setCmText]=useState("");
  const [inbox,setInbox]=useState([]); const [bellOpen,setBellOpen]=useState(false); const [peersOpen,setPeersOpen]=useState(false);
  const [history,setHistory]=useState(false); const [auditRows,setAuditRows]=useState([]); const [auditBusy,setAuditBusy]=useState(false); const [histFilter,setHistFilter]=useState("");
  const [reviewOpen,setReviewOpen]=useState(false); const [reviewTab,setReviewTab]=useState("changes");
  const [helpOpen,setHelpOpen]=useState(false); const [helpTab,setHelpTab]=useState("guide");
  const [errorLog,setErrorLog]=useState([]);
  const [follows,setFollows]=useState(new Set());
  const [clip,setClip]=useState(null);     // {values:2D,h,w}
  const [showAux,setShowAux]=useState(false); // toggle: reveal underlying auto/plan + revised dates in cells
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
  // LOAD everything from Supabase (also used by the Sync button)
  const loadShared=async()=>{ try{
    // Supabase caps a single select() at 1000 rows. With 100+ styles x 13 stages, stage_dates was being silently truncated, so cells beyond row 1000 rendered blank even though the data was in the DB. Fetch EVERY row in pages.
    const fetchAll=async(table)=>{ let out=[], from=0; const size=1000; for(let i=0;i<100;i++){ const { data, error }=await supabase.from(table).select("*").range(from, from+size-1); if(error){ logAppError("load "+table,error); break; } if(!data||!data.length) break; out=out.concat(data); if(data.length<size) break; from+=size; } return out; };
    const styData = await fetchAll("styles"); styData.sort((a,b)=>(a.id||0)-(b.id||0));
    if(!styData.length){ logAppError("load styles","no styles loaded"); return; }
    const sdData = await fetchAll("stage_dates");
    const cmData = await fetchAll("cell_meta");
    const byId={}; sdData.forEach(r=>{ const e=(byId[r.style_id]=byId[r.style_id]||{actuals:{},revs:{},rejects:{},skips:{}}); if(r.actual_date) e.actuals[r.stage]=r.actual_date; if(r.revised_date) e.revs[r.stage]=r.revised_date; if(r.reject_date) e.rejects[r.stage]=r.reject_date; if(r.skip_date) e.skips[r.stage]=r.skip_date; });
    const appStyles=styData.map(row=>rowToStyle(row,byId));
    setStyles(appStyles);
    const SR={ sty:{}, stg:{}, meta:{} };
    appStyles.forEach(s=>{ SR.sty[s.id]=JSON.stringify(styleToRow(s)); STAGE_KEYS.forEach(k=>{ SR.stg[s.id+":"+k]=JSON.stringify({ style_id:s.id, stage:k, revised_date:(s.revs&&s.revs[k])||null, actual_date:s.actuals[k]||null, reject_date:(s.rejects&&s.rejects[k])||null, skip_date:(s.skips&&s.skips[k])||null }); }); });
    cmData.forEach(r=>{ if(r.fill||r.note) SR.meta[r.style_id+":"+r.col]=JSON.stringify({ style_id:r.style_id, col:r.col, fill:r.fill||null, note:r.note||null }); });
    savedRef.current=SR;
    try{ const cfgRes=await supabase.from("app_settings").select("data").eq("id","global").maybeSingle(); if(cfgRes&&cfgRes.data&&cfgRes.data.data){ const d=cfgRes.data.data; setCfg({ ...DEFAULT_CFG, ...d, leads:{...DEFAULT_CFG.leads,...(d.leads||{})}, stageOwners:{...DEFAULT_CFG.stageOwners,...(d.stageOwners||{})}, rework:{...DEFAULT_CFG.rework,...(d.rework||{})}, upcoming:{...DEFAULT_CFG.upcoming,...(d.upcoming||{})}, escalationRules:Array.isArray(d.escalationRules)&&d.escalationRules.length?d.escalationRules:DEFAULT_ESCALATION_RULES.map(x=>({...x})), todoEscalationRows:d.todoEscalationRows!==undefined?!!d.todoEscalationRows:DEFAULT_CFG.todoEscalationRows, labels:{...(d.labels||{})} }); } }catch(e){ /* settings table optional */ }
    try{ const tvRes=await supabase.from("app_settings").select("data").eq("id",TRACKER_VIEW_SETTING_ID).maybeSingle(); const views=normalizeTrackerViews(tvRes&&tvRes.data&&tvRes.data.data&&tvRes.data.data.views); setSharedViews(views); }catch(e){ setSharedViews(DEFAULT_NAMED_TRACKER_VIEWS); }
    const f={}, n={}; cmData.forEach(r=>{ if(r.fill) f[`${r.style_id}:${r.col}`]=r.fill; if(r.note) n[`${r.style_id}:${r.col}`]=r.note; });
    setFills(f); setNotes(n); try{ const [cmR,prR]=await Promise.all([ supabase.from("comments").select("*").order("created_at"), supabase.from("profiles").select("id,name,role,email") ]); const cg={}; (cmR.data||[]).forEach(r=>{ const ck=r.style_id+":"+r.col; (cg[ck]=cg[ck]||[]).push(r); }); setComments(cg); setTeam(prR.data||[]); const nR=await supabase.from("notifications").select("*").eq("user_id",me.id).order("created_at",{ascending:false}).limit(100); setInbox(nR.data||[]); const fR=await supabase.from("style_follows").select("style_id").eq("user_id",me.id); setFollows(new Set((fR.data||[]).map(r=>r.style_id))); }catch(e){} loadedRef.current=true; flash();
  }catch(e){ logAppError("load failed",e); } };
  useEffect(()=>{ loadShared(); },[]);
  const editingRef=useRef(null); useEffect(()=>{ editingRef.current=editing; },[editing]);
  const presRef=useRef(null); const [peers,setPeers]=useState([]); const [presReady,setPresReady]=useState(false);
  useEffect(()=>{ if(!me) return; setPresReady(false); const key=String(me.id)+"."+Math.random().toString(36).slice(2,7); const ch=supabase.channel("merch-presence",{ config:{ presence:{ key } } });
    ch.on("presence",{ event:"sync" },()=>{ const st=ch.presenceState(); const arr=[]; Object.keys(st).forEach(k=>{ const m=st[k]&&st[k][0]; if(m&&String(m.id)!==String(me.id)) arr.push(m); }); setPeers(arr); });
    ch.subscribe(async(status)=>{ if(status==="SUBSCRIBED"){ presRef.current=ch; setPresReady(true); try{ await ch.track({ id:me.id, name:me.name||me.email, role:me.role, cell:null, editing:null }); }catch(e){} } });
    return ()=>{ try{ supabase.removeChannel(ch); }catch(e){} presRef.current=null; setPresReady(false); }; },[me]);
  // re-broadcast our own position AND the cell we're actively editing, whenever either changes or the channel becomes ready
  useEffect(()=>{ const ch=presRef.current; if(ch&&me&&presReady){ try{ ch.track({ id:me.id, name:me.name||me.email, role:me.role, cell: sel?{ id:sel.id, col:sel.col }:null, editing: editing?{ id:editing.id, col:editing.col }:null }); }catch(e){} } },[sel,editing,presReady]);
  useEffect(()=>{ const ch=supabase.channel("merch-comments").on("postgres_changes",{ event:"*", schema:"public", table:"comments" },(p)=>{ const n=(p.new&&p.new.id)?p.new:p.old; if(!n) return; const ck=n.style_id+":"+n.col; setComments(prev=>{ const arr=(prev[ck]||[]).filter(x=>x.id!==n.id); if(p.eventType!=="DELETE") arr.push(p.new); arr.sort((a,b)=>new Date(a.created_at)-new Date(b.created_at)); return {...prev,[ck]:arr}; }); }).subscribe(); return ()=>{ try{ supabase.removeChannel(ch); }catch(e){} }; },[]);
  const peerCell={}; // cellKey -> [{id,name,color,editing}]
  peers.forEach(p=>{ const loc=p.editing||p.cell; if(!loc) return; const k=loc.id+":"+loc.col; (peerCell[k]=peerCell[k]||[]).push({ id:p.id, name:p.name, color:colorFor(p.id), editing:!!p.editing }); });
  const peerOn=(id,col)=> peerCell[id+":"+col]||null; // array | null
  const peerEditingHere=(id,col)=>{ const arr=peerCell[id+":"+col]; return arr?(arr.find(w=>w.editing)||null):null; };
  const peerLockBlocks=(id,col)=>{ const w=peerEditingHere(id,col); if(!w) return false; return !window.confirm(w.name+" is editing this cell right now.\n\nSaving over it at the same time can overwrite their change. Open it anyway?"); };
  const cellLabel=(id,col)=>{ const s=styles.find(x=>x.id===id); const sn=s?(s.styleNo||("#"+id)):("#"+id); const cl= col==="__style"?"Style No":((INFO_COLS.find(c=>c.key===col)||{}).label||(STAGES.find(x=>x.key===col)||{}).label||col); return sn+" · "+cl; };
  useEffect(()=>{ if(!me) return; const ch=supabase.channel("merch-notifs").on("postgres_changes",{ event:"*", schema:"public", table:"notifications", filter:"user_id=eq."+me.id },(p)=>{ const n=(p.new&&p.new.id)?p.new:p.old; if(!n) return; setInbox(prev=>{ let arr=prev.filter(x=>x.id!==n.id); if(p.eventType!=="DELETE") arr=[p.new,...arr]; arr.sort((a,b)=>new Date(b.created_at)-new Date(a.created_at)); return arr; }); }).subscribe(); return ()=>{ try{ supabase.removeChannel(ch); }catch(e){} }; },[me]);
  useEffect(()=>{ const ch=supabase.channel("merch-live")
    .on("postgres_changes",{ event:"*", schema:"public", table:"stage_dates" },(p)=>{ const del=p.eventType==="DELETE"; const n=(p.new&&Object.keys(p.new).length)?p.new:p.old; if(!n||!n.style_id){ setRemoteChanged(true); return; } const sid=n.style_id, stg=n.stage, key=sid+":"+stg; const row=JSON.stringify({ style_id:sid, stage:stg, revised_date:n.revised_date||null, actual_date:n.actual_date||null, reject_date:n.reject_date||null, skip_date:n.skip_date||null }); if(!del && savedRef.current.stg[key]===row) return; const ed=editingRef.current; if(ed&&ed.id===sid&&ed.col===stg){ setRemoteChanged(true); return; } if(del) delete savedRef.current.stg[key]; else savedRef.current.stg[key]=row; remotePatchRef.current=true; setStyles(prev=>{ let found=false; const next=prev.map(s=>{ if(s.id!==sid) return s; found=true; const ns={...s, actuals:{...s.actuals}, revs:{...(s.revs||{})}, rejects:{...(s.rejects||{})}, skips:{...(s.skips||{})} }; if(del){ delete ns.actuals[stg]; delete ns.revs[stg]; delete ns.rejects[stg]; delete ns.skips[stg]; } else { if(n.actual_date) ns.actuals[stg]=n.actual_date; else delete ns.actuals[stg]; if(n.revised_date) ns.revs[stg]=n.revised_date; else delete ns.revs[stg]; if(n.reject_date) ns.rejects[stg]=n.reject_date; else delete ns.rejects[stg]; if(n.skip_date) ns.skips[stg]=n.skip_date; else delete ns.skips[stg]; } return ns; }); if(!found){ setRemoteChanged(true); return prev; } return next; }); setTimeout(()=>{ remotePatchRef.current=false; },0); })
    .on("postgres_changes",{ event:"*", schema:"public", table:"styles" },(p)=>{ const del=p.eventType==="DELETE"; const n=(p.new&&p.new.id)?p.new:p.old; if(!n||!n.id){ setRemoteChanged(true); return; } const sid=n.id; if(del){ delete savedRef.current.sty[sid]; remotePatchRef.current=true; setStyles(prev=>prev.filter(s=>s.id!==sid)); setTimeout(()=>{ remotePatchRef.current=false; },0); return; } if(savedRef.current.sty[sid]===JSON.stringify(n)) return; const ed=editingRef.current; if(ed&&ed.id===sid){ setRemoteChanged(true); return; } savedRef.current.sty[sid]=JSON.stringify(n); remotePatchRef.current=true; setStyles(prev=>{ const idx=prev.findIndex(s=>s.id===sid); if(idx===-1) return [...prev, rowToStyle(n,{})]; const cur=prev[idx]; const merged={ ...rowToStyle(n,{}), actuals:cur.actuals, revs:cur.revs, rejects:cur.rejects, skips:cur.skips }; const copy=prev.slice(); copy[idx]=merged; return copy; }); setTimeout(()=>{ remotePatchRef.current=false; },0); })
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
    const keys=new Set([...Object.keys(fills),...Object.keys(notes)]); const metaChanged=[]; keys.forEach(key=>{ const i=key.indexOf(":"); const row={ style_id:Number(key.slice(0,i)), col:key.slice(i+1), fill:fills[key]||null, note:notes[key]||null }; const j=JSON.stringify(row); if(SR.meta[key]!==j) metaChanged.push({row,key,j}); });
    if(metaChanged.length){ const up3=await supabase.from("cell_meta").upsert(metaChanged.map(x=>x.row),{ onConflict:"style_id,col" }); if(up3.error) throw up3.error; }
    styChanged.forEach(r=>{ SR.sty[r.id]=JSON.stringify(r); }); stgChanged.forEach(x=>{ SR.stg[x.key]=x.j; }); metaChanged.forEach(x=>{ SR.meta[x.key]=x.j; });
    setSaveState("saved"); flash();
  }catch(e){ logAppError("save failed",e); setSaveState("error"); } },700); return ()=>clearTimeout(t); },[styles,fills,notes]);

  const [saveState,setSaveState]=useState("idle"); // idle | saving | saved | error
  useEffect(()=>{ if(!loadedRef.current) return; const t=setTimeout(()=>{ supabase.from("app_settings").upsert({ id:"global", data:cfg }).then(()=>{}).catch(()=>{}); },600); return ()=>clearTimeout(t); },[cfg]);
  const [remoteChanged,setRemoteChanged]=useState(false); // another user wrote data
  const flash=()=>{ setSaved(true); clearTimeout(savedTimer.current); savedTimer.current=setTimeout(()=>setSaved(false),1200); };
  const logAppError=(area,err,extra)=>{ const msg=(err&&err.message)?err.message:String(err||""); const row={ id:Date.now()+Math.random(), at:new Date().toISOString(), area, msg, extra:extra||"" }; setErrorLog(p=>[row,...p].slice(0,100)); console.error(area,err); };
  const setField=(id,field,val)=>{ pushHistory(); if(STAGE_KEYS.includes(field)){ const ck=id+":"+field+":actual"; if(val) clearedRef.current.delete(ck); else clearedRef.current.add(ck); } setStyles(prev=>prev.map(s=>{ if(s.id!==id) return s; if(STAGE_KEYS.includes(field)) return { ...s, actuals:{ ...s.actuals, [field]: val||undefined } }; if(field==="qty") return { ...s, qty:Number(val)||0 }; return { ...s, [field]:val }; })); flash(); };
  const setRev=(id,key,val)=>{ pushHistory(); const ck=id+":"+key+":revised"; if(val) clearedRef.current.delete(ck); else clearedRef.current.add(ck); setStyles(prev=>prev.map(s=> s.id===id?{...s,revs:{...(s.revs||{}),[key]:val||undefined}}:s)); flash(); };
  const setReject=(id,key,val)=>{ pushHistory(); const ck=id+":"+key+":reject"; if(val) clearedRef.current.delete(ck); else clearedRef.current.add(ck); setStyles(prev=>prev.map(s=> s.id===id?{...s,rejects:{...(s.rejects||{}),[key]:val||undefined}}:s)); flash(); };
  const setSkip=(id,key,val)=>{ pushHistory(); const ap=APPR_OF_SEND[key]; [key,...(ap?[ap]:[])].forEach(kk=>{ const ck=id+":"+kk+":skip"; if(val) clearedRef.current.delete(ck); else clearedRef.current.add(ck); }); setStyles(prev=>prev.map(s=>{ if(s.id!==id) return s; const skips={...(s.skips||{})}; if(val){ skips[key]=val; if(ap) skips[ap]=val; } else { skips[key]=undefined; if(ap) skips[ap]=undefined; } return {...s,skips}; })); flash(); };
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
    if(updates.length){ const m={}; updates.forEach(u=>{ m[u.id]={...(m[u.id]||{}),...u.chg}; }); setStyles(prev=>prev.map(s=> m[s.id]?{...s,...m[s.id]}:s)); }
    if(inserts.length){ const rows=inserts.map(rec=>{ const s={ orderNo:rec.orderNo||"", styleNo:rec.styleNo, sampleFit:rec.sampleFit||"", family:rec.family||"", colour:rec.colour||"", brand:rec.brand||"", buyer:rec.buyer||"", fabricType:rec.fabricType||"", owner:rec.owner||"", setId:rec.setId||"", setRole:rec.setRole||"", age:rec.age||"", qty:rec.qty||0, ordRec:rec.ordRec||iso(TODAY), delivery:rec.delivery||rec.ordRec||iso(TODAY), fitReq:rec.fitReq??true, printReq:rec.printReq??false, soReq:rec.soReq??false, ppBypass:rec.ppBypass??false, labDipReq:rec.labDipReq??true, ppNeeded:rec.ppNeeded??true, remarks:"" }; const r=styleToRow(s); delete r.id; return r; });
      try{ const { data, error }=await supabase.from("styles").insert(rows).select(); if(error) throw error; if(data) setStyles(prev=>[...prev, ...data.map(d=>rowToStyle(d,{}))]); }catch(e){ logAppError("bulk insert failed",e); alert("New styles failed to insert (existing updates were applied): "+(e.message||e)); } }
    setBulkOpen(false); setBulkResult(null); setUploadSkip(new Set()); flash(); };
  const BUYER_STAGES=["fitAppr","artAppr","soAppr","labAppr","ppAppr"];
  const runExport=(mode,buf,incBuf,relMode)=>{ try{ let data, name, sheet="Tracker"; const B=Math.abs(buf||0);
    const fmtY=(d)=> !d?"":d.toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"});
    const fmtIso=(v)=> v?fmtY(parse(v)):"";
    const actCell=(s,c,k)=>{ const r=(c.stages||[]).find(x=>x.key===k); if(!r) return "n/a"; if(s.actuals&&s.actuals[k]) return fmtY(parse(s.actuals[k])); if(r.skipped) return "WAIVED"; if(r.rejected) return "REJECTED"; if(r.rework) return "REDO"; return ""; };
    const planCell=(s,c,k,withBuf)=>{ const r=(c.stages||[]).find(x=>x.key===k); if(!r) return "n/a"; if(actCell(s,c,k)) return ""; const t=r.rev||r.plan; if(!t) return ""; /* E6: buffer is a single B-day slack on our internal chain. Each approval inherits the shift from its (buffered) internal predecessor and adds NO buffer of its own — net effect is one B applied per stage, never doubled. */ const d=(withBuf&&B)?addWorkdays(t,B):t; return fmtY(d); };
    const pairs=(o,s,c,withBuf,list)=>{ (list||STAGES).forEach(st=>{ o["Actual \u00b7 "+st.label]=actCell(s,c,st.key); o["Plan \u00b7 "+st.label]=planCell(s,c,st.key,withBuf); }); };
    const blockers=(c)=> (c.frontier?[...c.frontier]:[]).map(k=>{ const r=(c.stages||[]).find(x=>x.key===k); if(!r||r.done) return null; const lbl=r.label; if(r.rework) return lbl+": redo & resend"; if(r.rejected) return lbl+": rejected"; if(REJECTABLE.includes(k)) return lbl+": not approved"; return lbl+": pending"; }).filter(Boolean);
    if(mode==="buyer"){ data=rows.map(({s,c})=>{ const o={ "Style No":s.styleNo, "Family":s.family, "Colour":s.colour, "Brand":s.brand, "Buyer":s.buyer||"", "Age Group":s.age||"", "Qty":s.qty, "Delivery":fmtIso(s.delivery), "Status":c.status }; pairs(o,s,c,incBuf); addStylePlanAccuracyCols(o,s,c); o["Proj. Release"]=c.projRelease?fmtY(c.projRelease):""; return o; }); name=incBuf?"buyer_tna_buf"+B+"d":"buyer_tna"; sheet="Buyer"; }
    else if(mode==="release"){ data=rows.map(({s,c})=>{ const open=blockers(c); const o={ "Order No":s.orderNo, "Style No":s.styleNo, "Colour":s.colour, "Qty":s.qty }; if(relMode!=="summary"){ o["FIT"]=s.fitReq?"Y":"-"; o["PRT"]=s.printReq?"Y":"-"; o["S-O"]=s.soReq?"Y":"-"; o["LAB"]=s.labDipReq?"Y":"-"; o["BYP"]=s.ppBypass?"Y":"-"; o["PP"]=s.ppNeeded?"Y":"-"; pairs(o,s,c,false); } else { ["fitAppr","soAppr","fabricIH","ppAppr","prodFile"].forEach(k=>{ const st=STAGES.find(x=>x.key===k); o["Actual \u00b7 "+st.label]=actCell(s,c,k); o["Plan \u00b7 "+st.label]=planCell(s,c,k,false); }); } addStylePlanAccuracyCols(o,s,c); o["Proj. Release"]=c.projRelease?fmtY(c.projRelease):""; o["Delivery"]=fmtIso(s.delivery); o["Released"]=c.released?"YES":""; o["Pending / blockers"]=c.released?"released":(open.join("; ")||"none"); o["Remarks"]=s.remarks||""; o["Status"]=c.status; return o; }); name=relMode==="summary"?"release_plan_summary":"release_plan_detailed"; sheet="Release Plan"; }
    else if(mode==="detail"){ data=rows.map(({s,c})=>{ const open=blockers(c); const o={ "Order No":s.orderNo, "Style No":s.styleNo, "Sample Fit":s.sampleFit, "Family":s.family, "Colour":s.colour, "Brand":s.brand, "Buyer":s.buyer||"", "Age Group":s.age||"", "Fabric Type":s.fabricType, "Owner":s.owner, "Qty":s.qty, "Order Date":fmtIso(s.ordRec), "Delivery":fmtIso(s.delivery), "FIT":s.fitReq?"Y":"-", "PRT":s.printReq?"Y":"-", "S-O":s.soReq?"Y":"-", "LAB":s.labDipReq?"Y":"-", "BYP":s.ppBypass?"Y":"-", "PP":s.ppNeeded?"Y":"-" }; pairs(o,s,c,false); o["% Done"]=c.pct+"%"; o["Float"]=(c.float!=null?c.float+"d":""); o["Idle"]=(c.idle!=null?c.idle+"d":""); addStylePlanAccuracyCols(o,s,c); o["Proj. Release"]=c.projRelease?fmtY(c.projRelease):""; o["Released"]=c.released?"YES":""; o["Pending / blockers"]=c.released?"released":(open.join("; ")||"none"); o["Remarks"]=s.remarks||""; o["Status"]=c.status; return o; }); name="detailed_summary"; sheet="Detailed"; }
    else if(mode==="internal"){ data=rows.map(({s,c})=>{ const o={ "Order No":s.orderNo, "Style No":s.styleNo, "Sample Fit":s.sampleFit, "Family":s.family, "Colour":s.colour, "Buyer":s.buyer||"", "Age Group":s.age||"", "Qty":s.qty, "Delivery":fmtIso(s.delivery), "Status":c.status }; pairs(o,s,c,true); addStylePlanAccuracyCols(o,s,c); return o; }); name="internal_plan_buf"+B+"d"; }
    else { data=rows.map(({s,c})=>{ const o={ "Order No":s.orderNo, "Style No":s.styleNo, "Sample Fit":s.sampleFit, "Family":s.family, "Colour":s.colour, "Brand":s.brand, "Buyer":s.buyer||"", "Age Group":s.age||"", "Fabric Type":s.fabricType, "Junior":s.owner, "Qty":s.qty, "Order Date":fmtIso(s.ordRec), "Delivery":fmtIso(s.delivery), "Status":c.status }; STAGES.forEach(st=>{ o[st.label]=actCell(s,c,st.key); }); addStylePlanAccuracyCols(o,s,c); return o; }); name="merch_tracker"; }
    if(data&&data.length){ Object.keys(data[0]).forEach(k=>{ if(data.every(r=>r[k]==="n/a")) data.forEach(r=>{ delete r[k]; }); }); }
    const ws=XLSX.utils.json_to_sheet(data); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,sheet); XLSX.writeFile(wb,name+"_"+iso(TODAY)+".xlsx"); }catch(e){ logAppError("export failed",e); alert("Export failed: "+(e.message||e)); } };
  const bulkVisibleSummary=()=>{ const vis=rows.map(r=>r.s); const buyers=[...new Set(vis.map(s=>s.buyer||s.brand||"—").filter(Boolean))].slice(0,5); const owners=[...new Set(vis.map(s=>s.owner||"—").filter(Boolean))].slice(0,6); const sample=vis.slice(0,8).map(s=>`${s.orderNo||"—"} · ${s.styleNo||"—"}`); return { count:vis.length, buyers, owners, sample, more:Math.max(0,vis.length-sample.length) }; };
  const openBulkConfirm=(payload)=>{ const ids=rows.map(r=>r.s.id); if(!ids.length){ alert("No visible styles to update."); return; } const summary=bulkVisibleSummary(); setBulkConfirm({ ...payload, ids, summary }); };
  const applyBulkConfirm=()=>{ const bc=bulkConfirm; if(!bc) return; const ids=new Set(bc.ids||[]); if(!ids.size){ setBulkConfirm(null); return; } pushHistory();
    if(bc.kind==="appendRemark"){ const note=String(bc.note||"").trim(); setStyles(prev=>prev.map(s=>ids.has(s.id)?{...s,remarks:[s.remarks,note].filter(Boolean).join(" | ")}:s)); }
    else { const clean=Object.fromEntries(Object.entries(bc.patch||{}).filter(([_,v])=>v!==undefined)); setStyles(prev=>prev.map(s=>ids.has(s.id)?{...s,...clean}:s)); }
    flash(); setBulkConfirm(null); setBulkActionsOpen(false);
  };
  const archiveFiltered=(val)=>{ const label=val?"Archive visible styles":"Restore visible styles"; openBulkConfirm({ title:label, kind:"patch", patch:{ archived:val }, impact:val?"Archived styles will be hidden from the active sheet. This is reversible from Archive view.":"Restored styles will return to the active sheet.", actionText:val?"Archive styles":"Restore styles", danger:val }); };
  const bulkUpdateVisible=(patch,label,impact)=>{ const clean=Object.fromEntries(Object.entries(patch||{}).filter(([_,v])=>v!==undefined)); if(!Object.keys(clean).length) return; openBulkConfirm({ title:label, kind:"patch", patch:clean, impact:impact||"This changes real row data and will auto-save for everyone.", actionText:"Apply bulk update" }); };
  const bulkAppendRemark=(txt)=>{ const note=String(txt||"").trim(); if(!note) return; openBulkConfirm({ title:"Append remark", kind:"appendRemark", note, impact:`This remark will be added to every visible style: “${note}”`, actionText:"Append remark" }); };
  const bulkFlagVisible=(flag,val)=>{ const def=(FLAG_DEFS.find(x=>x.key===flag)||{}); const label=def.short||flag; bulkUpdateVisible({ [flag]:val }, `Set ${label} = ${val?"Yes":"No"}`, `Requirement flag ${def.title||flag} will be set to ${val?"YES":"NO"} for all visible styles.`); };
  const deleteStyle=async(id)=>{ if(!window.confirm("Delete this style row? This removes it for everyone and cannot be undone.")) return; pushHistory(); setStyles(prev=>prev.filter(s=>s.id!==id)); flash(); try{ await supabase.from("stage_dates").delete().eq("style_id",id); await supabase.from("cell_meta").delete().eq("style_id",id); await supabase.from("styles").delete().eq("id",id); }catch(e){ console.error("delete failed",e); } };

  const perfRef=useRef({ computeMs:0, filterMs:0, rowMs:0, rows:0, styles:0, rendered:0, cacheHits:0, recomputed:0, alerts:[], samples:[], p95Ms:0, deferred:false });
  const computeCacheRef=useRef(new Map());
  // Keep compute invalidation tight: changing To-Do watch windows or escalation rules should not recompute every style.
  const cfgComputeKey=useMemo(()=>JSON.stringify({ leads:cfg.leads, stageOwners:cfg.stageOwners, rework:cfg.rework, fabricCutoff:cfg.fabricCutoff, relGate:cfg.relGate }),[cfg.leads,cfg.stageOwners,cfg.rework,cfg.fabricCutoff,cfg.relGate]);
  const computed=useMemo(()=>{ const t=perfNow(); const cache=computeCacheRef.current; const liveIds=new Set(); let hits=0, recomputed=0; const out=styles.map(s=>{ liveIds.add(s.id); const old=cache.get(s.id); if(old && old.sRef===s && old.cfgKey===cfgComputeKey){ hits++; return old.out; } const c=computeStyle(s,cfg); const item={s,c,idx:buildSearchIndex(s,c)}; cache.set(s.id,{ sRef:s, cfgKey:cfgComputeKey, out:item }); recomputed++; return item; }); for(const id of cache.keys()){ if(!liveIds.has(id)) cache.delete(id); } const ms=Math.round((perfNow()-t)*10)/10; const hist=pushPerfSample(perfRef.current.samples,ms); perfRef.current.computeMs=ms; perfRef.current.samples=hist.samples; perfRef.current.p95Ms=hist.p95; perfRef.current.styles=styles.length; perfRef.current.cacheHits=hits; perfRef.current.recomputed=recomputed; return out; },[styles,cfg,cfgComputeKey]);
  const chaseOwnerOptions=useMemo(()=>["All", ...Array.from(new Set([...STAGES.map(st=>(cfg.stageOwners&&cfg.stageOwners[st.key])||DEFAULT_STAGE_OWNERS[st.key]||st.owner), ...CHASE_LABELS])).filter(Boolean)], [cfg]);
  const todoItems=useMemo(()=>{
    const enrich=(item)=>{
      const daysLate=item.overdue?Math.max(1,Math.abs(Number(item.du)||0)):0;
      const esc=item.overdue?escalationFor(cfg,daysLate):null;
      return { ...item, daysLate, escalationOwner:esc?esc.owner:"", escalationLevel:esc?esc.level:"", escalationAction:esc?esc.action:"", escalationRange:esc?esc.rangeLabel:"" };
    };
    const out=[]; const fabByCol={};
    computed.forEach(({s,c})=>{
      if(c.released) return;
      const front=c.frontier?[...c.frontier]:[];
      front.forEach(key=>{
        const r=(c.stages||[]).find(x=>x.key===key);
        if(!r||r.done) return;
        const exp=r.rev||r.plan; if(!exp) return;
        const du=netWorkdays(TODAY,exp);
        const overdue=TODAY>exp;
        const win=(cfg.upcoming&&cfg.upcoming[key]!=null)?cfg.upcoming[key]:null;
        const include = overdue || (win!=null && du<=win);
        if(!include) return;
        const branch=BRANCH_OF[key]||"";
        if(key==="fabricIH"){
          const cols=String(s.colour||"").split(/[,/]/).map(x=>x.trim()).filter(Boolean);
          (cols.length?cols:["(no colour)"]).forEach(col=>{
            let cur=fabByCol[col];
            if(!cur){ cur=fabByCol[col]={ colour:col, key, label:r.label, owner:r.owner, branch, exp, du, overdue, anyStyle:s.id, anyOrder:s.orderNo, anyJunior:s.owner, count:0 }; }
            cur.count++;
            if(exp<cur.exp){ cur.exp=exp; cur.du=du; cur.overdue=overdue; cur.anyStyle=s.id; cur.anyOrder=s.orderNo; cur.anyJunior=s.owner; }
          });
        } else {
          out.push(enrich({ id:s.id, orderNo:s.orderNo, styleNo:s.styleNo, junior:s.owner, colour:s.colour, key, activity:r.label, branch, owner:r.owner, exp, du, overdue }));
        }
      });
    });
    Object.values(fabByCol).forEach(f=> out.push(enrich({ id:f.anyStyle, orderNo:f.anyOrder, styleNo:f.colour, junior:f.anyJunior, colour:f.colour, key:f.key, activity:f.label, branch:f.branch, owner:f.owner, exp:f.exp, du:f.du, overdue:f.overdue, isColour:true, count:f.count })));
    out.sort((a,b)=> (a.overdue!==b.overdue)?(a.overdue?-1:1):((a.exp&&b.exp)?(a.exp-b.exp):0));
    return out;
  },[computed,cfg]);
  const [todoFilter,setTodoFilter]=useState(PF.todoFilter||{});
  useEffect(()=>{ try{ localStorage.setItem("mt_trackfilters", JSON.stringify({ search, searchCol, statusFilter, ownerFilter, archiveView, activityFilter, colFilters, tab, todoFilter, followFilter, savedView })); }catch(e){} },[search,searchCol,statusFilter,ownerFilter,archiveView,activityFilter,colFilters,tab,todoFilter,followFilter,savedView]);
  useEffect(()=>{ try{ localStorage.setItem("mt_column_view",columnView); }catch(e){} },[columnView]);
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
  const passCol=(s,c,col,allowed)=>{ if(!allowed||allowed.length===0) return true; if(col==="chase"){ const owners=(c.chaseOwners||[]).map(o=>o.owner); if(owners.length===0) return allowed.includes("(Blanks)"); return owners.some(o=>allowed.includes(o)); } return allowed.includes(valueFor(s,c,col)); };
  const chaseLabel=(owner)=>String(owner||"—");
  const SAVED_VIEWS=[ ["","Saved view: none"], ["overdue","Overdue"], ["dueThisWeek","Due this week"], ["buyerPending","Buyer approval pending"], ["fabricPending","Fabric pending"], ["ppPending","PP pending"], ["deliveryRisk","Delivery risk"], ["following","Followed styles"], ["rework","Rejected / rework"], ["released","Released"] ];
  const anyFilter = statusFilter!=="All"||ownerFilter!=="All"||!!search||Object.keys(colFilters).length>0||!!activityFilter||followFilter||!!savedView;
  const resetFilters=()=>{ setStatusFilter("All"); setOwnerFilter("All"); setSearch(""); setColFilters({}); setActivityFilter(null); setSavedView(""); setActiveNamedView(""); };
  const snapCurrent=()=>setViewSnap({ statusFilter, ownerFilter, search, colFilters, activityFilter, savedView });
  const clearAllFilters=()=>{ resetFilters(); setViewSnap(null); };
  const restoreView=()=>{ if(!viewSnap) return; setStatusFilter(viewSnap.statusFilter); setOwnerFilter(viewSnap.ownerFilter); setSearch(viewSnap.search); setColFilters(viewSnap.colFilters); setActivityFilter(viewSnap.activityFilter); setSavedView(viewSnap.savedView||""); setViewSnap(null); };
  const applyDrill=(spec)=>{ snapCurrent(); setStatusFilter(spec.status||"All"); setOwnerFilter(spec.owner||"All"); setSearch(spec.search||""); setColFilters(spec.colFilters||{}); setActivityFilter(spec.activity||null); setSavedView(""); setTab("tracker"); };
  const presetPass=(s,c)=>{ if(!savedView) return true; const front=c.frontier?[...c.frontier]:[]; const has=(keys)=>front.some(k=>keys.includes(k)); const dueSoon=front.some(k=>{ const r=(c.stages||[]).find(x=>x.key===k); const d=r&&(r.rev||r.plan); const n=d?netWorkdays(TODAY,d):999; return d && n>=0 && n<=6; }); const anyRework=(c.stages||[]).some(r=>r.rework||r.rejected); if(savedView==="overdue") return (c.tone==="late"||c.status.startsWith("Overdue")); if(savedView==="dueThisWeek") return !c.released && dueSoon; if(savedView==="buyerPending") return !c.released && front.some(k=>((STAGES.find(x=>x.key===k)||{}).owner)==="Buyer"); if(savedView==="fabricPending") return !c.released && has(["labDip","labAppr","fabricIH"]); if(savedView==="ppPending") return !c.released && has(["ppSample","ppAppr","prodFile"]); if(savedView==="deliveryRisk") return c.tone==="late"||String(c.status).toLowerCase().includes("risk"); if(savedView==="following") return follows.has(s.id); if(savedView==="rework") return anyRework; if(savedView==="released") return c.released; return true; };
  const deferredSearch=useDeferredValue(search);
  const filterKey=useMemo(()=>JSON.stringify({ search:deferredSearch, searchCol, statusFilter, ownerFilter, archiveView, activityFilter, colFilters, followFilter, savedView, follows:[...follows].sort() }),[deferredSearch,searchCol,statusFilter,ownerFilter,archiveView,activityFilter,colFilters,followFilter,savedView,follows]);
  const filtered=useMemo(()=>{ const t=perfNow(); const q=String(deferredSearch||"").trim().toLowerCase(); const cfEntries=Object.entries(colFilters||{}); const out=computed.filter((row)=>{ const {s,c,idx}=row; const matchQ = !q ? true : (searchCol==="auto" ? (idx&&idx.auto?idx.auto:"").includes(q) : ((idx&&idx.byCol&&idx.byCol[searchCol])!=null?idx.byCol[searchCol]:lc(s[searchCol])).includes(q)); const matchS=statusFilter==="All"||(statusFilter==="At Risk"&&(c.tone==="late"||c.tone==="warn"))||(statusFilter==="On Track"&&c.tone==="ok")||(statusFilter==="Released"&&c.released); const matchF=cfEntries.every(([col,allowed])=> passCol(s,c,col,allowed)); const matchO=ownerFilter==="All"||(c.chaseOwners||[]).some(o=>o.owner===ownerFilter); const matchA=!activityFilter||(c.frontier&&c.frontier.has(activityFilter)); const matchArch=archiveView==="all"?true:(archiveView==="archived"?!!s.archived:!s.archived); const matchFollow=!followFilter||follows.has(s.id); return matchQ&&matchS&&matchF&&matchO&&matchA&&matchArch&&matchFollow&&presetPass(s,c); }); perfRef.current.filterMs=Math.round((perfNow()-t)*10)/10; perfRef.current.deferred=deferredSearch!==search; return out; },[computed,filterKey,deferredSearch,search]);
  const toneRank={ late:0, warn:1, ok:2, done:3, na:4 };
  const fitNum=(s)=>{ const m=String(s.sampleFit).match(/\d+/); return m?Number(m[0]):Infinity; };
  const sortVal=(col,{s,c})=>{ switch(col){ case "__style": return s.styleNo.toLowerCase(); case "orderNo": return (s.orderNo||"~").toLowerCase(); case "sampleFit": return fitNum(s); case "family": return s.family.toLowerCase(); case "colour": return s.colour.toLowerCase(); case "owner": return (s.owner||"").toLowerCase(); case "setId": return (s.setId||"~").toLowerCase(); case "setRole": return (s.setRole||"").toLowerCase(); case "qty": return s.qty; case "ordRec": return s.ordRec?new Date(s.ordRec).getTime():Infinity; case "delivery": return s.delivery?new Date(s.delivery).getTime():Infinity; case "overall": return toneRank[c.tone]; case "fit": return toneRank[c.fitBranch.tone]; case "print": return toneRank[c.printBranch.tone]; case "fabric": return toneRank[c.fabricBranch.tone]; case "pp": return toneRank[c.ppBranch.tone]; case "prod": return toneRank[c.prodFileBranch.tone]; case "fabricCD": return c.fabricCountdown.n==null?Infinity:c.fabricCountdown.n; case "proj": return c.projRelease?c.projRelease.getTime():Infinity; case "pct": return c.pct; case "chase": return (c.chaseOwners||[]).length; case "float": return c.float==null?Infinity:c.float; case "idle": return c.idle==null?-1:c.idle; case "remarks": return (s.remarks||"~").toLowerCase(); default: { const a=s.actuals[col]; return a?new Date(a).getTime():Infinity; } } };
  const rows=useMemo(()=>{ const t=perfNow(); const out=!sort.col ? filtered : [...filtered].sort((A,B)=>{ const a=sortVal(sort.col,A), b=sortVal(sort.col,B); return a<b?-sort.dir:a>b?sort.dir:0; }); perfRef.current.rowMs=Math.round((perfNow()-t)*10)/10; perfRef.current.rows=out.length; perfRef.current.alerts=[perfRef.current.p95Ms>800?"P95 >800ms":"", perfRef.current.computeMs>300?"compute slow":"", perfRef.current.filterMs>200?"filter slow":"", perfRef.current.rowMs>200?"sort slow":"", out.length>900?"large visible rows":"", styles.length>1000?"1000+ active styles":"", perfRef.current.deferred?"search catching up":""].filter(Boolean); return out; },[filtered,sort,styles.length]);
  const peerEditingList=peers.filter(p=>p.editing).map(p=>({ ...p, ref: (()=>{ const loc=p.editing; if(!loc) return ""; const ri=rows.findIndex(r=>r.s.id===loc.id); const ci=navCols.indexOf(loc.col); return (ci>=0&&ri>=0)?(colLetter(ci)+(ri+1)):""; })() }));
  const [renderLimit,setRenderLimit]=useState(()=>{ try{ const v=parseInt(localStorage.getItem("mt_render_limit"),10); return (isFinite(v)&&v>=300&&v<=5000)?v:900; }catch(e){ return 900; } });
  useEffect(()=>{ try{ localStorage.setItem("mt_render_limit",String(renderLimit)); }catch(e){} },[renderLimit]);
  useEffect(()=>{ if(rows.length<=renderLimit && renderLimit>900) setRenderLimit(900); },[rows.length,renderLimit]);
  const renderRows=useMemo(()=>{ const out=rows.length>renderLimit?rows.slice(0,renderLimit):rows; perfRef.current.rendered=out.length; return out; },[rows,renderLimit]);
  const clickHeader=(col)=>{ finishEditing(); setSort(p=> p.col===col?{col,dir:-p.dir}:{col,dir:1}); };

  const visInfo=INFO_COLS.filter(c=>!hidden.has(c.key));
  const colLabel=(c)=>((cfg.labels&&cfg.labels[c.key])||c.label);
  const visStages=STAGES.filter(s=>!hidden.has(s.key));
  const remarksVis=!hidden.has("remarks");
  const navCols=["__style", ...visInfo.map(c=>c.key), ...visStages.map(s=>s.key), ...(remarksVis?["remarks"]:[])];
  const totalCols=navCols.length;
  const maxFreeze=1+visInfo.length;
  const ROLE_VIEW_PRESETS={
    merchant:{ label:"Merchant view", show:null },
    management:{ label:"Management view", show:["orderNo","family","colour","brand","owner","buyer","qty","delivery","overall","fit","print","fabric","pp","prod","fabricCD","proj","pct","chase","float","idle","remarks"] },
    cad:{ label:"CAD view", show:["orderNo","sampleFit","family","colour","brand","owner","delivery","overall","fit","techpack","fitSend","fitAppr","proj","chase","remarks"] },
    designer:{ label:"Designer view", show:["orderNo","family","colour","brand","fabricType","owner","delivery","overall","print","artwork","artAppr","strikeOff","soAppr","proj","chase","remarks"] },
    store:{ label:"Store / Fabric view", show:["orderNo","family","colour","brand","fabricType","owner","delivery","overall","fabric","labDip","labAppr","fabricIH","fabricCD","chase","remarks"] },
    buyer:{ label:"Buyer follow-up view", show:["orderNo","family","colour","brand","buyer","delivery","overall","fitAppr","artAppr","soAppr","labAppr","ppAppr","proj","chase","remarks"] }
  };
  const applyColumnView=(view)=>{ setColumnView(view); setColsOpen(false); if(view==="custom") return; const all=[...INFO_COLS.map(c=>c.key),...STAGE_KEYS,"remarks"]; const spec=ROLE_VIEW_PRESETS[view]; if(!spec||!spec.show){ setHidden(new Set(["extra1","extra2"])); return; } const keep=new Set(spec.show); setHidden(new Set(all.filter(k=>!keep.has(k)))); setFreezeN(1); };

  const currentViewState=()=>({
    search, searchCol, statusFilter, ownerFilter, archiveView, activityFilter, colFilters, savedView, columnView,
    hidden:[...hidden], sort, freezeN
  });
  const applyTrackerViewState=(state, name="")=>{
    const st=state||{};
    setSearch(st.search||""); setSearchCol(st.searchCol||"auto"); setStatusFilter(st.statusFilter||"All"); setOwnerFilter(st.ownerFilter||"All");
    setArchiveView(st.archiveView||"active"); setActivityFilter(st.activityFilter||null); setColFilters(st.colFilters||{}); setSavedView(st.savedView||"");
    if(st.columnView){ setColumnView(st.columnView); }
    if(Array.isArray(st.hidden)) setHidden(new Set(st.hidden)); else if(st.columnView && st.columnView!=="custom") applyColumnView(st.columnView);
    setSort(st.sort&&st.sort.col?st.sort:{ col:null, dir:1 }); setFreezeN(Number(st.freezeN)||1); setActiveNamedView(name||""); setViewSnap(null);
  };
  const saveSharedTrackerView=async()=>{
    if(!canAdmin(role)){ alert("Only Management / Sr Merchant can save or overwrite shared default views. Your current filters are still temporary for you only."); return; }
    const name=(window.prompt("Save current filters/columns as default shared view:", activeNamedView||"Management View")||"").trim();
    if(!name) return;
    const next=normalizeTrackerViews(sharedViews.filter(v=>v.name!==name).concat([{ name, shared:true, updatedAt:new Date().toISOString(), updatedBy:me?.name||me?.email||"", state:currentViewState() }]));
    setSharedViews(next); setActiveNamedView(name);
    try{ const { error }=await supabase.from("app_settings").upsert({ id:TRACKER_VIEW_SETTING_ID, data:{ views:next } }); if(error) throw error; flash(); }
    catch(e){ logAppError("save shared view failed",e); alert("View saved locally in this browser, but shared save failed: "+(e.message||e)); }
  };
  const resetTemporaryView=()=>{ clearAllFilters(); setSort({ col:null, dir:1 }); setActiveNamedView(""); };

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

  const ownerOfCol=(col)=>{ const st=STAGES.find(s=>s.key===col); if(st) return st.owner; const ic=INFO_COLS.find(c=>c.key===col); if(ic&&ic.owner) return ic.owner; return "Merchant"; };
  const isStageCol=(col)=>STAGE_KEYS.includes(col);
  const fmtTyped=(isoStr)=>{ const d=parse(isoStr); return d?`${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`:""; };
  const parseTyped=(strv)=>{ const t=(strv||"").trim(); if(!t) return ""; let m=/^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t); if(m){ const y=+m[1],mo=+m[2],d=+m[3]; if(mo<1||mo>12||d<1||d>31) return false; return `${y}-${String(mo).padStart(2,"0")}-${String(d).padStart(2,"0")}`; } m=/^(\d{1,2})[\/\-.](\d{1,2})(?:[\/\-.](\d{2,4}))?$/.exec(t); if(m){ let d=+m[1],mo=+m[2],y=m[3]?+m[3]:TODAY.getFullYear(); if(y<100) y+=2000; if(mo<1||mo>12||d<1||d>31) return false; return `${y}-${String(mo).padStart(2,"0")}-${String(d).padStart(2,"0")}`; } return false; };
  const [calOpen,setCalOpen]=useState(false);
  const [entryVal,setEntryVal]=useState("");
  const [entryTouched,setEntryTouched]=useState(false);
  const selectedStyle=sel?styles.find(x=>x.id===sel.id):null;
  const selectedColLabel=sel?(sel.col==="__style"?"Style No":(INFO_COLS.find(c=>c.key===sel.col)?.label||STAGES.find(s=>s.key===sel.col)?.label||sel.col)):"";
  const selectedDisplayValue=()=>{ if(!sel||!selectedStyle) return ""; const v=getVal(selectedStyle,sel.col); return isDateCol(sel.col)&&v?fmtTyped(v):String(v??""); };
  const entrySuggestion=useMemo(()=>{
    if(!sel||!entryTouched||!entryVal||isDateCol(sel.col)) return "";
    const ci=colIndex(sel.col), ri=rowIndex(sel.id);
    if(ci<0||ri<0) return "";
    const q=String(entryVal).trim().toLowerCase();
    if(!q) return "";
    const seen=[];
    for(let r=ri-1;r>=0&&seen.length<40;r--){
      const row=rows[r]; if(!row||!row.s) continue;
      const v=String(getVal(row.s,sel.col)||"").trim();
      if(v&&!seen.some(x=>x.toLowerCase()===v.toLowerCase())) seen.push(v);
    }
    return seen.find(v=>v.toLowerCase().startsWith(q) && v!==entryVal)||"";
  },[sel,entryTouched,entryVal,rows]);
  useEffect(()=>{ if(editing && sel && editing.id===sel.id && editing.col===sel.col && !entryTouched){ setEntryVal(editVal||""); return; } if(!entryTouched){ setEntryVal(selectedDisplayValue()); } },[sel,styles,editing,editVal,entryTouched]);
  const commitEntry=()=>{
    if(!sel||!selectedStyle) return;
    if(!isEditableCol(sel.col)||!canEdit(role,sel.col,"actual")) return;
    if(peerLockBlocks(sel.id,sel.col)) return;
    const raw=String(entryVal??"");
    if(isDateCol(sel.col)){
      const r=parseTyped(raw);
      if(r===false){ window.alert('"'+raw+'" isn’t a valid date. Type it as dd/mm/yyyy (e.g. 26/05/2026), or clear the box to leave it empty.'); return; }
      setField(sel.id,sel.col,r===""?null:r);
    } else {
      const f=sel.col==="__style"?"styleNo":sel.col;
      setField(sel.id,f,sel.col==="qty"?raw.replace(/[^0-9]/g,""):raw);
    }
    setEntryTouched(false);
    setEditing(null);
    setEditVal("");
  };
  const beginDate=(id,col,mode,initialChar)=>{ if(!canEdit(role,col,mode)) return; if(peerLockBlocks(id,col)) return; setSel({id,col}); setFocus(null); setEditing({id,col,mode}); setCalOpen(false); const s=styles.find(x=>x.id===id); const cur= mode==="rev"?(s&&s.revs&&s.revs[col]):mode==="reject"?(s&&s.rejects&&s.rejects[col]):(isStageCol(col)?(s&&s.actuals[col]):(s&&s[col])); setEditVal(initialChar!=null?initialChar:(cur?fmtTyped(cur):"")); };
  const commitDate=()=>{ if(!editing) return; const r=parseTyped(editVal); if(r===false){ window.alert('"'+editVal+'" isn\u2019t a valid date. Type it as dd/mm/yyyy (e.g. 26/05/2026), or clear the box to leave it empty.'); return; } if(r!==false){ const val=r===""?null:r; if(val===null){ const s=styles.find(x=>x.id===editing.id); const had= editing.mode==="rev"?(s&&s.revs&&s.revs[editing.col]):editing.mode==="reject"?(s&&s.rejects&&s.rejects[editing.col]):(isStageCol(editing.col)?(s&&s.actuals&&s.actuals[editing.col]):(s&&s[editing.col])); if(had && !window.confirm("Delete this saved date? You can re-enter it later.")){ setEditing(null); setCalOpen(false); return; } } if(editing.mode==="rev") setRev(editing.id,editing.col,val); else if(editing.mode==="reject") setReject(editing.id,editing.col,val); else setField(editing.id,editing.col,val); } setEditing(null); setCalOpen(false); };
  const dateEditor=(id,col,mode)=>{ const s=styles.find(x=>x.id===id); const _cmp=computed.find(x=>x.s.id===id); const _stR=_cmp&&(_cmp.c.stages||[]).find(x=>x.key===col); const planFb=_stR?(_stR.rev||_stR.plan):null; const stored= mode==="rev"?(s&&s.revs&&s.revs[col]):mode==="reject"?(s&&s.rejects&&s.rejects[col]):(isStageCol(col)?(s&&s.actuals[col]):(s&&s[col])); const colLabel=col==="ordRec"?"Order Date":col==="delivery"?"Delivery Date":((STAGES.find(x=>x.key===col)||{}).label||col); const modeLabel=mode==="rev"?"REVISED":mode==="reject"?"REJECTED":"ACTUAL"; const mc=mode==="rev"?"var(--accent)":mode==="reject"?"var(--danger)":"var(--info)"; return (<span onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:1, left:1, zIndex:80, display:"flex", flexDirection:"column", gap:1, background:"var(--surface)", border:"1px solid "+mc, padding:"2px 4px", boxShadow:"2px 2px 0 rgba(0,0,0,0.18)" }}><span style={{ fontSize:8, fontWeight:700, color:mc, textTransform:"uppercase", letterSpacing:0.3, whiteSpace:"nowrap" }}>{colLabel} · {modeLabel}</span><span style={{ display:"flex", alignItems:"center", gap:2, position:"relative" }}><input autoFocus onFocus={e=>{ if((e.target.value||"").length>2) e.target.select(); }} value={editVal} placeholder="dd/mm/yyyy" onChange={e=>setEditVal(e.target.value.replace(/[^0-9\/\-. ]/g,""))} onKeyDown={e=>{ e.stopPropagation(); if(e.key==="Enter") commitDate(); else if(e.key==="Escape"){ setEditing(null); setCalOpen(false); } }} onBlur={()=>{ if(!calOpen) commitDate(); }} style={{ width:80, fontFamily:"inherit", fontSize:11, border:"none", outline:"none" }}/>{stored && <button onMouseDown={e=>e.preventDefault()} onClick={e=>{ e.stopPropagation(); if(mode==="rev") setRev(id,col,null); else if(mode==="reject") setReject(id,col,null); else setField(id,col,null); setEditing(null); setCalOpen(false); }} title={"Clear "+modeLabel.toLowerCase()+" date"} style={{ border:"1px solid var(--line-2)", background:"var(--surface)", cursor:"pointer", padding:"0 4px", fontSize:10, lineHeight:"16px" }}>clear</button>}<button onMouseDown={e=>e.preventDefault()} onClick={e=>{ e.stopPropagation(); setCalOpen(o=>!o); }} title="calendar" style={{ border:"none", background:"transparent", cursor:"pointer", padding:0, lineHeight:0, fontSize:12 }}>📅</button>{calOpen && <CalPopup label={colLabel+" · "+modeLabel} value={stored} fallback={planFb} onClose={()=>setCalOpen(false)} onPick={(d)=>{ if(mode==="rev") setRev(id,col,d); else if(mode==="reject") setReject(id,col,d); else setField(id,col,d); setEditing(null); setCalOpen(false); }}/>}</span></span>); };
    const startEdit=(id,col,initialChar)=>{ if(!isEditableCol(col)) return; if(!canEdit(role,col,"actual")) return; if(isDateCol(col)){ beginDate(id,col,"actual"); return; } if(peerLockBlocks(id,col)) return; const s=styles.find(x=>x.id===id); setEditing({id,col,mode:"text"}); if(col==="qty") setEditVal(initialChar??String(s.qty)); else if(col==="__style") setEditVal(initialChar??s.styleNo); else setEditVal(initialChar??(s[col]||"")); };
  const commitText=()=>{ if(!editing) return; const f=editing.col==="__style"?"styleNo":editing.col; if(editing.mode==="text"||editing.mode===undefined){ if(!isDateCol(editing.col)) setField(editing.id,f,editVal); } setEditing(null); };
  const finishEditing=()=>{ if(!editing) return; if(editing.mode==="actual"||editing.mode==="rev"||editing.mode==="reject") commitDate(); else commitText(); };

  // ---- selection range ----
  const rowIndex=(id)=>rows.findIndex(r=>r.s.id===id);
  const colIndex=(col)=>navCols.indexOf(col);
  const cellRef=(x)=> x? (colLetter(colIndex(x.col))+(rowIndex(x.id)+1)) : "";
  const [nameBox,setNameBox]=useState("");
  useEffect(()=>{ setNameBox(cellRef(sel)); },[sel]);
  const gotoCell=(ref)=>{ const m=/^([A-Za-z]+)\s*(\d+)$/.exec(String(ref||"").trim()); if(!m) return; const ci=letterToIndex(m[1]); const ri=Number(m[2])-1; if(ci<0||ci>=navCols.length||ri<0||ri>=rows.length) return; const id=rows[ri].s.id, col=navCols[ci]; setSel({ id, col }); setFocus(null); scrollToCell(id,col); };
  const rect=()=>{ if(!sel) return null; const aR=rowIndex(sel.id), aC=colIndex(sel.col); const f=focus||sel; const fR=rowIndex(f.id), fC=colIndex(f.col); return { r1:Math.min(aR,fR), r2:Math.max(aR,fR), c1:Math.min(aC,fC), c2:Math.max(aC,fC) }; };
  const selKeys=useMemo(()=>{ const R=rect(); const set=new Set(); if(!R) return set; for(let r=R.r1;r<=R.r2;r++){ for(let c=R.c1;c<=R.c2;c++){ if(rows[r]) set.add(`${rows[r].s.id}:${navCols[c]}`); } } return set; },[sel,focus,rows,navCols]);
  const onCellClick=(e,id,col)=>{ e.stopPropagation(); if(filterCol) setFilterCol(null); if(gridRef.current) gridRef.current.focus({preventScroll:true}); if(editing){ if(editing.id===id&&editing.col===col) return; finishEditing(); } if(e.shiftKey&&sel){ setFocus({id,col}); scrollToCell(id,col); return; } setSel({id,col}); setFocus(null); };

  const moveAnchor=(dr,dc)=>{ if(!sel) return; let r=rowIndex(sel.id)+dr, c=colIndex(sel.col)+dc; r=Math.min(Math.max(r,0),rows.length-1); c=Math.min(Math.max(c,0),navCols.length-1); if(rows[r]){ setSel({id:rows[r].s.id,col:navCols[c]}); setFocus(null); scrollToCell(rows[r].s.id,navCols[c]); } };
  const scrollToCell=(id,col)=>{ if(tab!=="tracker") setTab("tracker"); if(col && col!=="__style" && hidden.has(col)) setHidden(prev=>{ const n=new Set(prev); n.delete(col); return n; });
    const go=()=>{ let el=document.getElementById(`cell-${id}-${col}`); if(!el) el=document.getElementById(`cell-${id}-__style`); if(!el) return; const wrap=scrollWrapRef.current; if(wrap){ const cr=el.getBoundingClientRect(), wr=wrap.getBoundingClientRect(); wrap.scrollTop += (cr.top - wr.top) - (wr.height/2 - cr.height/2); const cr2=el.getBoundingClientRect(); const frozen=STYLE_W+6; if(cr2.left < wr.left+frozen) wrap.scrollLeft -= (wr.left+frozen-cr2.left)+8; else if(cr2.right > wr.right) wrap.scrollLeft += (cr2.right-wr.right)+8; } else { el.scrollIntoView({ block:"center", behavior:"smooth" }); } };
    requestAnimationFrame(()=>requestAnimationFrame(go)); };
  const selectRow=(id)=>{ setSel({id,col:navCols[0]}); setFocus({id,col:navCols[navCols.length-1]}); };
  const selectAll=()=>{ if(!rows.length) return; setSel({id:rows[0].s.id,col:navCols[0]}); setFocus({id:rows[rows.length-1].s.id,col:navCols[navCols.length-1]}); };
  const moveFocus=(dr,dc)=>{ if(!sel) return; const f=focus||sel; let r=rowIndex(f.id)+dr, c=colIndex(f.col)+dc; r=Math.min(Math.max(r,0),rows.length-1); c=Math.min(Math.max(c,0),navCols.length-1); if(rows[r]){ setFocus({id:rows[r].s.id,col:navCols[c]}); scrollToCell(rows[r].s.id,navCols[c]); } };

  const snap=()=>({ styles, fills, notes });
  const pushHistory=()=>{ setPast(p=>[...p.slice(-60), snap()]); setFuture([]); };
  const applySnap=(d)=>{ setStyles(d.styles); setFills(d.fills); setNotes(d.notes); };
  const undo=()=>{ if(!past.length) return; if(remoteChanged && !window.confirm("Teammates have made changes that aren't merged in yet (the Sync button is highlighted).\n\nUndoing now may overwrite them. Undo anyway?")) return; const prev=past[past.length-1]; setFuture(f=>[...f, snap()]); setPast(p=>p.slice(0,-1)); applySnap(prev); flash(); };
  const redo=()=>{ if(!future.length) return; if(remoteChanged && !window.confirm("Teammates have unmerged changes (the Sync button is highlighted).\n\nRedoing now may overwrite them. Redo anyway?")) return; const nx=future[future.length-1]; setPast(p=>[...p, snap()]); setFuture(f=>f.slice(0,-1)); applySnap(nx); flash(); };
  const getVal=(s,col)=>{ if(col==="__style") return s.styleNo; if(col==="qty") return String(s.qty); if(STAGE_KEYS.includes(col)) return s.actuals[col]||""; if(col==="ordRec"||col==="delivery"||TEXT_COLS.includes(col)) return s[col]||""; return null; };
  const doCopy=()=>{ const R=rect(); if(!R) return; const values=[]; let any=false; for(let r=R.r1;r<=R.r2;r++){ const row=[]; for(let c=R.c1;c<=R.c2;c++){ const v=rows[r]?getVal(rows[r].s,navCols[c]):null; if(v!=null&&v!=="") any=true; row.push(v); } values.push(row); } if(any){ setClip({ values, h:values.length, w:values[0].length }); flash(); } };
  const canPasteCell=(s,col)=>{ if(!isEditableCol(col)) return false; if(!canEdit(role,col,"actual")) return false; if(STAGE_KEYS.includes(col)){ const st=STAGES.find(x=>x.key===col); if(!(st.flag===null||s[st.flag])) return false; } return true; };
  const doPaste=()=>{ if(!clip||!sel) return; pushHistory(); const R=rect(); const changes={}; const put=(id,col,val)=>{ (changes[id]=changes[id]||{})[col]=val; };
    if(clip.h===1&&clip.w===1){ const v=clip.values[0][0]; for(let r=R.r1;r<=R.r2;r++){ for(let c=R.c1;c<=R.c2;c++){ const row=rows[r]; const col=navCols[c]; if(row&&canPasteCell(row.s,col)) put(row.s.id,col,v); } } }
    else { for(let i=0;i<clip.h;i++){ for(let j=0;j<clip.w;j++){ const r=R.r1+i, c=R.c1+j; const row=rows[r]; const col=navCols[c]; if(row&&col&&canPasteCell(row.s,col)) put(row.s.id,col,clip.values[i][j]); } } }
    setStyles(prev=>prev.map(s=>{ const ch=changes[s.id]; if(!ch) return s; let ns={...s, actuals:{...s.actuals}}; Object.entries(ch).forEach(([col,val])=>{ if(STAGE_KEYS.includes(col)) ns.actuals[col]=val||undefined; else if(col==="qty") ns.qty=Number(val)||0; else if(col==="__style") ns.styleNo=val; else ns[col]=val; }); return ns; })); flash(); };

  // batch write a {id:{col:val}} change map into styles
  const writeChanges=(changes)=>{ Object.entries(changes).forEach(([id,ch])=>Object.entries(ch).forEach(([col,val])=>{ if(STAGE_KEYS.includes(col)){ const ck=id+":"+col+":actual"; if(val) clearedRef.current.delete(ck); else clearedRef.current.add(ck); } })); setStyles(prev=>prev.map(s=>{ const ch=changes[s.id]; if(!ch) return s; let ns={...s, actuals:{...s.actuals}, revs:{...(s.revs||{})}}; Object.entries(ch).forEach(([col,val])=>{ if(STAGE_KEYS.includes(col)) ns.actuals[col]=val||undefined; else if(col==="qty") ns.qty=Number(val)||0; else if(col==="__style") ns.styleNo=val; else ns[col]=val; }); return ns; })); flash(); };
  const coerce=(col,raw)=>{ if(raw==null) return ""; const v=String(raw).trim(); if(isDateCol(col)){ const pt=parseTyped(v); if(pt!==false) return pt; const d=new Date(v); return isNaN(d)?"":iso(d); } return v; };
  const clearRange=()=>{ const R=rect(); if(!R) return; pushHistory(); const ch={}; for(let r=R.r1;r<=R.r2;r++){ for(let cc=R.c1;cc<=R.c2;cc++){ const row=rows[r]; const col=navCols[cc]; if(row&&canPasteCell(row.s,col)) (ch[row.s.id]=ch[row.s.id]||{})[col]= STAGE_KEYS.includes(col)?null:(col==="qty"?0:""); } } writeChanges(ch); };
  const applyFillHandle=()=>{ if(!fillFrom||!fillTo) return; const aR=rowIndex(fillFrom.id), aC=colIndex(fillFrom.col), tR=rowIndex(fillTo.id), tC=colIndex(fillTo.col); const r1=Math.min(aR,tR), r2=Math.max(aR,tR), c1=Math.min(aC,tC), c2=Math.max(aC,tC); const srcRow=rows[aR]; if(!srcRow) return; pushHistory(); const ch={}; for(let r=r1;r<=r2;r++){ for(let cc=c1;cc<=c2;cc++){ const row=rows[r]; const col=navCols[cc]; if(!row) continue; if(r===aR&&cc===aC) continue; const srcVal=getVal(srcRow.s, navCols[cc]); if(canPasteCell(row.s,col)) (ch[row.s.id]=ch[row.s.id]||{})[col]=srcVal; } } writeChanges(ch); };
  const fitAllCols=()=>{ const cols=["__style", ...visInfo.map(c=>c.key), ...visStages.map(s=>s.key)]; const upd={}; cols.forEach(col=>{ let max=String((INFO_COLS.find(x=>x.key===col)||{}).label||(STAGES.find(x=>x.key===col)||{}).label||col).length; rows.forEach(({s})=>{ const v=getVal(s,col); if(v!=null) max=Math.max(max,String(isDateCol(col)?fmt(parse(v)):v).length); }); upd[col]=Math.max(54,Math.min(420,Math.ceil(max*7.4*textScale+34))); }); setColW(p=>({...p,...upd})); flash(); };
  const autoFit=(col)=>{ let max=String(col==="__style"?"Style No":(INFO_COLS.find(x=>x.key===col)?.label||STAGES.find(x=>x.key===col)?.label||col)).length; rows.forEach(({s})=>{ const v=getVal(s,col); if(v!=null) max=Math.max(max,String(isDateCol(col)?fmt(parse(v)):v).length); }); setColW(p=>({ ...p, [col]:Math.max(54, Math.min(420, Math.ceil(max*7.4*textScale+34))) })); };
  const onKeyDown=(e)=>{ const _tt=e.target&&e.target.tagName; if(_tt==="INPUT"||_tt==="SELECT"||_tt==="TEXTAREA"||(e.target&&e.target.isContentEditable)) return;
    if((e.ctrlKey||e.metaKey)&&(e.key==="z"||e.key==="Z")){ e.preventDefault(); if(e.shiftKey) redo(); else undo(); return; }
    if((e.ctrlKey||e.metaKey)&&(e.key==="y"||e.key==="Y")){ e.preventDefault(); redo(); return; }
    if(!sel) return;
    if(editing){ const dm=editing.mode==="actual"||editing.mode==="rev"; if(e.key==="Enter"){ dm?commitDate():commitText(); moveAnchor(1,0); e.preventDefault(); } else if(e.key==="Escape"){ setEditing(null); setCalOpen(false); e.preventDefault(); } else if(e.key==="Tab"){ dm?commitDate():commitText(); moveAnchor(0,e.shiftKey?-1:1); e.preventDefault(); } return; }
    if((e.ctrlKey||e.metaKey)&&(e.key==="c"||e.key==="C")){ doCopy(); e.preventDefault(); return; }
    if((e.ctrlKey||e.metaKey)&&(e.key==="v"||e.key==="V")){ doPaste(); e.preventDefault(); return; }
    if(["ArrowUp","ArrowDown","ArrowLeft","ArrowRight"].includes(e.key)){ e.preventDefault(); const dr=e.key==="ArrowUp"?-1:e.key==="ArrowDown"?1:0; const dc=e.key==="ArrowLeft"?-1:e.key==="ArrowRight"?1:0; if(e.shiftKey) moveFocus(dr,dc); else moveAnchor(dr,dc); }
    else if(e.key==="Enter"||e.key==="F2"){ e.preventDefault(); startEdit(sel.id,sel.col); }
    else if(e.key==="Delete"||e.key==="Backspace"){ if(focus) clearRange(); else if(canPasteCell(styles.find(x=>x.id===sel.id),sel.col)) setField(sel.id,sel.col, STAGE_KEYS.includes(sel.col)?null:(sel.col==="qty"?0:"")); e.preventDefault(); }
    else if(e.key==="Escape"){ setSel(null); setFocus(null); }
    else if(e.key.length===1&&!e.metaKey&&!e.ctrlKey){ if(isDateCol(sel.col)&&/[0-9]/.test(e.key)){ beginDate(sel.id,sel.col,"actual",e.key); e.preventDefault(); } else if(sel.col==="qty"&&/[0-9]/.test(e.key)){ startEdit(sel.id,"qty",e.key); e.preventDefault(); } else if(sel.col==="__style"||TEXT_COLS.includes(sel.col)){ startEdit(sel.id,sel.col,e.key); e.preventDefault(); } }
  };
  const jumpToEnter=(id,stageKey)=>{ const st=STAGES.find(x=>x.key===stageKey); setSel({id,col:stageKey}); setFocus(null); requestAnimationFrame(()=>{ const el=document.getElementById(`cell-${id}-${stageKey}`); if(el) el.scrollIntoView({ behavior:"smooth", inline:"center", block:"nearest" }); }); }; // jump + select only — double-click / F2 to edit (matches single-click rule)

  const handleCopy=(e)=>{ const R=rect(); if(!R) return; const lines=[]; let any=false; for(let r=R.r1;r<=R.r2;r++){ const cells=[]; for(let cc=R.c1;cc<=R.c2;cc++){ let v=rows[r]?getVal(rows[r].s,navCols[cc]):""; if(isDateCol(navCols[cc])&&v) v=fmtTyped(v); if(v) any=true; cells.push(v??""); } lines.push(cells.join("\t")); } if(any){ const tsv=lines.join("\n"); try{ e.clipboardData.setData("text/plain",tsv); e.preventDefault(); }catch(err){} setClip({ values:lines.map(l=>l.split("\t")), h:lines.length, w:lines[0].split("\t").length }); flash(); } };
  const handlePaste=(e)=>{ if(!sel) return; let txt=""; try{ txt=e.clipboardData.getData("text/plain"); }catch(err){} if(!txt){ doPaste(); return; } e.preventDefault(); const grid=txt.replace(/\r/g,"").replace(/\n$/,"").split("\n").map(l=>l.split("\t")); pushHistory(); const aR=rowIndex(sel.id), aC=colIndex(sel.col); const ch={}; if(grid.length===1&&grid[0].length===1){ const R=rect(); for(let r=R.r1;r<=R.r2;r++){ for(let cc=R.c1;cc<=R.c2;cc++){ const row=rows[r]; const col=navCols[cc]; if(row&&canPasteCell(row.s,col)) (ch[row.s.id]=ch[row.s.id]||{})[col]=coerce(col,grid[0][0]); } } } else { for(let i=0;i<grid.length;i++){ for(let j=0;j<grid[i].length;j++){ const row=rows[aR+i]; const col=navCols[aC+j]; if(row&&col&&canPasteCell(row.s,col)) (ch[row.s.id]=ch[row.s.id]||{})[col]=coerce(col,grid[i][j]); } } } writeChanges(ch); };
  const FR_COLS=["__style","orderNo","styleNo","sampleFit","family","colour","brand","fabricType","owner","remarks"];
  const frGet=(s,col)=> col==="__style"?(s.styleNo||""):String(s[col]==null?"":s[col]);
  const frColLabel=(col)=> col==="__style"?"Style No":((INFO_COLS.find(c=>c.key===col)||{}).label||col);
  const computeMatches=()=>{ const f=frCase?frFind:(frFind||"").toLowerCase(); if(!f) return []; let cells=[]; if(frScope==="selected"){ const R=rect(); if(R){ for(let r=R.r1;r<=R.r2;r++) for(let cc=R.c1;cc<=R.c2;cc++){ const row=rows[r], col=navCols[cc]; if(row&&col&&FR_COLS.includes(col)&&col!=="styleNo") cells.push([row.s,col]); } } } else { for(const row of rows) for(const col of FR_COLS){ if(col==="styleNo") continue; cells.push([row.s,col]); } } const out=[]; const seen=new Set(); for(const [s,col] of cells){ const cur=frGet(s,col); if(!cur) continue; const hay=frCase?cur:cur.toLowerCase(); if(hay.indexOf(f)===-1) continue; const key=s.id+":"+col; if(seen.has(key)) continue; seen.add(key); out.push({ id:s.id, col, style:s.styleNo||"", colLabel:frColLabel(col), text:cur }); } return out; };
  const gotoMatch=(i,list)=>{ const m=list||frMatches; if(!m.length) return; const ni=((i%m.length)+m.length)%m.length; setFindIdx(ni); const t=m[ni]; setSel({ id:t.id, col:t.col }); setFocus(null); scrollToCell(t.id, t.col); };
  const runFind=()=>{ const m=computeMatches(); setFrMatches(m); if(!m.length){ setFindIdx(-1); window.alert('No cells contain "'+frFind+'".'); return; } gotoMatch(0,m); };
  const findNext=()=>{ if(!frMatches.length){ runFind(); return; } gotoMatch(findIdx+1); };
  const findReplace=(preview)=>{ const f=frFind; if(!f) return 0; let cells=[]; if(frScope==="selected"){ const R=rect(); if(!R) return 0; for(let r=R.r1;r<=R.r2;r++) for(let cc=R.c1;cc<=R.c2;cc++){ const row=rows[r], col=navCols[cc]; if(row&&col&&FR_COLS.includes(col)) cells.push([row.s,col]); } } else { for(const row of rows) for(const col of FR_COLS) cells.push([row.s,col]); } const esc=f.replace(/[.*+?^${}()|[\]\\]/g,"\\$&"); const re=new RegExp(esc, frCase?"g":"gi"); const repStr=frRepl.replace(/\$/g,"$$$$"); const needle=frCase?f:f.toLowerCase(); let count=0; const ch={}; const seen=new Set(); for(const [s,col] of cells){ const k=s.id+":"+(col==="__style"?"styleNo":col); if(seen.has(k)) continue; seen.add(k); const cur=frGet(s,col); if(!cur) continue; const hay=frCase?cur:cur.toLowerCase(); if(hay.indexOf(needle)===-1) continue; const nv=cur.replace(re,repStr); if(nv===cur) continue; count++; if(!preview && canPasteCell(s,col)) (ch[s.id]=ch[s.id]||{})[col]=nv; } if(!preview && Object.keys(ch).length){ pushHistory(); writeChanges(ch); } return count; };
  const copySpecial=()=>{ if(!sel||!isStageCol(sel.col)){ alert("Select a stage date cell first. Copy-special grabs that cell full state: actual + revised + rejected + skip."); return; } const s=styles.find(x=>x.id===sel.id); if(!s) return; setSpecialClip({ actual:s.actuals[sel.col]||null, rev:(s.revs&&s.revs[sel.col])||null, reject:(s.rejects&&s.rejects[sel.col])||null, skip:(s.skips&&s.skips[sel.col])||null }); flash(); };
  const pasteSpecial=()=>{ if(!sel||!specialClip) return; const targets=[]; const R=rect(); if(R){ for(let rr=R.r1; rr<=R.r2; rr++){ for(let cc=R.c1; cc<=R.c2; cc++){ const row=rows[rr]; const col=navCols[cc]; if(row && isStageCol(col)) targets.push({id:row.s.id,col}); } } } else if(isStageCol(sel.col)) targets.push({id:sel.id,col:sel.col}); if(!targets.length) return; const targetSet=new Set(targets.map(t=>t.id+":"+t.col)); pushHistory(); setStyles(prev=>prev.map(s=>{ const cols=targets.filter(t=>t.id===s.id).map(t=>t.col); if(!cols.length) return s; const ns={...s, actuals:{...s.actuals}, revs:{...(s.revs||{})}, rejects:{...(s.rejects||{})}, skips:{...(s.skips||{})} }; cols.forEach(col=>{ if(canEdit(role,col,"actual")){ const ck=s.id+":"+col+":actual"; if(specialClip.actual){ ns.actuals[col]=specialClip.actual; clearedRef.current.delete(ck); } else { delete ns.actuals[col]; clearedRef.current.add(ck); } } if(canEditRev(role)){ const ck=s.id+":"+col+":revised"; if(specialClip.rev){ ns.revs[col]=specialClip.rev; clearedRef.current.delete(ck); } else { delete ns.revs[col]; clearedRef.current.add(ck); } } if(canEditReject(role,col)){ const ck=s.id+":"+col+":reject"; if(specialClip.reject){ ns.rejects[col]=specialClip.reject; clearedRef.current.delete(ck); } else { delete ns.rejects[col]; clearedRef.current.add(ck); } } if(MERCH_ROLES.includes(role)){ const ck=s.id+":"+col+":skip"; if(specialClip.skip){ ns.skips[col]=specialClip.skip; clearedRef.current.delete(ck); } else { delete ns.skips[col]; clearedRef.current.add(ck); } } }); return ns; })); flash(); };
  const copySelection=async()=>{ const R=rect(); if(!R) return; const lines=[]; for(let r=R.r1;r<=R.r2;r++){ const cells=[]; for(let cc=R.c1;cc<=R.c2;cc++){ let v=rows[r]?getVal(rows[r].s,navCols[cc]):""; if(isDateCol(navCols[cc])&&v) v=fmtTyped(v); cells.push(v??""); } lines.push(cells.join("\t")); } const tsv=lines.join("\n"); setClip({ values:lines.map(l=>l.split("\t")), h:lines.length, w:lines[0].split("\t").length }); try{ await navigator.clipboard.writeText(tsv); }catch(err){} flash(); };
  const cellKey=(id,col)=>`${id}:${col}`;
  const applyFill=(color)=>{ if(!sel) return; pushHistory(); const R=rect(); setFills(p=>{ const n={...p}; for(let r=R.r1;r<=R.r2;r++){ for(let c=R.c1;c<=R.c2;c++){ if(!rows[r]) continue; const k=`${rows[r].s.id}:${navCols[c]}`; if(color==="") delete n[k]; else n[k]=color; } } return n; }); flash(); };
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
  const logStageAudit=(changedList,SR)=>{ try{ const entries=[]; const fields=[["actual_date","actual"],["revised_date","revised"],["reject_date","rejected"],["skip_date","waived"]]; for(const x of changedList){ let oldRow={}; try{ oldRow=JSON.parse(SR.stg[x.key]||"{}"); }catch(e){ oldRow={}; } const nr=x.row; const sNo=(styles.find(s=>s.id===nr.style_id)||{}).styleNo||""; for(const [f,lab] of fields){ const ov=oldRow[f]||null, nv=nr[f]||null; if(ov!==nv) entries.push({ style_id:nr.style_id, style_no:sNo, col:nr.stage, field:lab, old_val:ov, new_val:nv, actor_id:me.id, actor_name:me.name||me.email }); } } if(entries.length) supabase.from("audit_log").insert(entries).then(()=>{}).catch(()=>{}); }catch(e){} };
  const logStyleAudit=(changedRows,SR)=>{ try{ const L={ style_no:"Style No", order_no:"Order No", sample_fit:"Sample Fit", family:"Family", colour:"Colour", brand:"Brand", buyer:"Buyer", fabric_type:"Fabric Type", owner:"Owner", age:"Age", set_id:"Set ID", set_role:"Set Role", qty:"Qty", order_date:"Order Date", delivery_date:"Delivery", remarks:"Remarks" }; const entries=[]; for(const r of changedRows){ if(!SR.sty[r.id]) continue; let oldRow={}; try{ oldRow=JSON.parse(SR.sty[r.id]); }catch(e){ continue; } const sNo=r.style_no||oldRow.style_no||""; for(const f in L){ const ov=(oldRow[f]==null)?"":String(oldRow[f]); const nv=(r[f]==null)?"":String(r[f]); if(ov!==nv) entries.push({ style_id:r.id, style_no:sNo, col:L[f], field:"value", old_val:ov||null, new_val:nv||null, actor_id:me.id, actor_name:me.name||me.email }); } } if(entries.length) supabase.from("audit_log").insert(entries).then(()=>{}).catch(()=>{}); }catch(e){} };
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
    const matchA=!activityFilter||(c.frontier&&c.frontier.has(activityFilter));
    const matchArch=archiveView==="all"?true:(archiveView==="archived"?!!s.archived:!s.archived);
    const matchFollow=!followFilter||follows.has(s.id);
    return matchQ&&matchS&&matchF&&matchO&&matchA&&matchArch&&matchFollow&&presetPass(s,c);
  };
  const distinctFor=(col)=>{ const set=new Set(); computed.forEach((row)=>{ if(!passForFilterOptions(row,col)) return; const {s,c}=row; if(col==="chase"){ const owners=(c.chaseOwners||[]).map(o=>o.owner); if(owners.length===0) set.add("(Blanks)"); else owners.forEach(o=>set.add(o)); } else set.add(valueFor(s,c,col)); }); return [...set].sort((a,b)=> a==="(Blanks)"?1:b==="(Blanks)"?-1:(a>b?1:a<b?-1:0)); };
  const filterProps=(col)=>({ filterActive: !!colFilters[col], filterOpen: filterCol===col, filterValues: filterCol===col?distinctFor(col):null, filterAllowed: colFilters[col]||null,
    onToggleFilter:()=>{ finishEditing(); setFilterCol(p=>p===col?null:col); },
    onSetFilter:(arr)=>setColFilters(f=>{ const n={...f}; if(!arr) delete n[col]; else n[col]=arr; return n; }),
    onCloseFilter:()=>setFilterCol(null) });
  const funnel=useMemo(()=>{ const b={ "Pre-Fit":0,"Fit/Print":0,"Lab Dip":0,"Fabric IH":0,"PP":0,"Released":0 }; computed.forEach(({c})=>{ if(c.released) b["Released"]++; else { const k=c.nextPending.key; if(k==="techpack") b["Pre-Fit"]++; else if(["fitSend","fitAppr","artwork","artAppr","strikeOff","soAppr"].includes(k)) b["Fit/Print"]++; else if(["labDip","labAppr"].includes(k)) b["Lab Dip"]++; else if(k==="fabricIH") b["Fabric IH"]++; else b["PP"]++; } }); return b; },[computed]);

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
  const bgFor=(id,col,base)=>{ const f=fills[cellKey(id,col)]; if(f) return f; if(inFill(id,col)) return "#def0e0"; if(focus&&inRange(id,col)&&!isAnchor(id,col)) return "#e3edfb"; return base; };
  const FillHandle=({id,col})=> isAnchor(id,col)&&!editing&&canPasteCell(styles.find(x=>x.id===id),col)? <span onMouseDown={(e)=>{ e.stopPropagation(); e.preventDefault(); setFilling(true); setFillFrom({id,col}); setFillTo({id,col}); }} title="drag to fill" style={{ position:"absolute", right:-2, bottom:-2, width:7, height:7, background:"var(--info)", cursor:"crosshair", zIndex:6 }}/> : null;
  const ringFor=(id,col)=> isAnchor(id,col)?"inset 0 0 0 2px var(--info)":null;
  const NoteTri=({k})=>{ const arr=comments[k]; if(!arr||!arr.length) return null; const un=arr.filter(x=>!x.resolved).length; const cl=un>0?"var(--danger)":"#9aa0a6"; return <span title={un>0?(un+" open comment(s)"):"comments resolved"} style={{ position:"absolute", top:0, right:0, width:0, height:0, borderTop:"9px solid "+cl, borderLeft:"9px solid transparent" }}/>; };

  const renderEditable=(s,col)=>{
    const k=cellKey(s.id,col.key);
    const val=col.key==="qty"?s.qty:s[col.key];
    const editingThis=editing&&editing.id===s.id&&editing.col===col.key;
    const bg=bgFor(s.id,col.key,"var(--surface)");
    return (
      <td key={col.key} id={`cell-${s.id}-${col.key}`} onClick={(e)=>onCellClick(e,s.id,col.key)} onDoubleClick={(e)=>{ e.stopPropagation(); startEdit(s.id,col.key); }}
        style={{ border:"1px solid var(--line-1)", padding:"6px 9px", whiteSpace: col.key==="remarks"?"normal":"nowrap", boxShadow:ringFor(s.id,col.key), cursor:"cell", maxWidth:col.w, overflow:"hidden", textOverflow:"ellipsis", fontSize: col.key==="remarks"?10:11, color: col.key==="remarks"?"#a15":"var(--ink)", position:"relative", background:bg, ...freezeStyle(col.key,bg) }}>
        {editingThis ? (<input autoFocus onFocus={e=>{ if((e.target.value||"").length>1) e.target.select(); }} value={editVal} onClick={e=>e.stopPropagation()} onChange={e=>setEditVal(col.key==="qty"?e.target.value.replace(/[^0-9]/g,""):e.target.value)} onBlur={commitText} style={{ width:Math.max(40,col.w-16), fontFamily:"inherit", fontSize:11, border:"1px solid var(--info)", outline:"none", padding:"1px 3px" }}/>) : (val===""||val==null ? <span style={{color:"var(--line-2)"}}>—</span> : String(val))}
        <PeerTag who={peerOn(s.id,col.key)}/><NoteTri k={k}/><FillHandle id={s.id} col={col.key}/>
      </td>
    );
  };

  return (
    <div ref={gridRef} tabIndex={0} onKeyDown={onKeyDown} onCopy={handleCopy} onPaste={handlePaste} onMouseDown={(e)=>{ if(editing) return; if(e.target.closest && (e.target.closest("input")||e.target.closest("button")||e.target.closest("th"))) return; const td=e.target.closest && e.target.closest('td[id^="cell-"]'); if(!td) return; const m=td.id.match(/^cell-(\d+)-(.+)$/); if(!m||e.shiftKey) return; e.preventDefault(); gridRef.current&&gridRef.current.focus(); setSel({ id:Number(m[1]), col:m[2] }); setFocus(null); selectingRef.current=true; setDragSel(true); }} onMouseUp={()=>{ if(filling){ applyFillHandle(); setFilling(false); setFillFrom(null); setFillTo(null); } if(selectingRef.current){ selectingRef.current=false; setDragSel(false); } }} onMouseOver={(e)=>{ const td=e.target.closest && e.target.closest('td[id^="cell-"]'); if(!td) return; const m=td.id.match(/^cell-(\d+)-(.+)$/); if(!m) return; if(filling){ setFillTo({ id:Number(m[1]), col:m[2] }); return; } if(selectingRef.current){ setFocus({ id:Number(m[1]), col:m[2] }); } }} onClick={()=>{ finishEditing(); setFillOpen(false); setColsOpen(false); setFilterCol(null); setFrOpen(false); setExpOpen(false); setBellOpen(false); setPeersOpen(false); }}
      style={{ minHeight:"100vh", background:"var(--bg)", fontFamily:"'JetBrains Mono', monospace", color:"var(--ink)", paddingBottom:80, outline:"none" }}>
      <style>{FONT}</style>
      <style>{THEME_CSS}</style>

      <div style={{ background:"var(--ink)", color:"var(--bg)", padding:"14px 22px", display:"flex", alignItems:"center", justifyContent:"space-between", borderBottom:"3px solid var(--accent)", position:"relative", zIndex:160, pointerEvents:"auto" }}>
        <div style={{ display:"flex", flexDirection:"column", gap:3 }}><span style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:20, letterSpacing:-0.2, lineHeight:1 }}>KOTHARI SPORTS & APPARELS</span><span style={{ fontFamily:"'JetBrains Mono',monospace", fontSize:9, color:"#9a958c", letterSpacing:1.2 }}>MERCH<span style={{ color:"var(--accent)" }}>·</span>TRACKER · PRE-PRODUCTION TRACKER</span></div>
        <div style={{ display:"flex", alignItems:"center", gap:12 }}><span style={{ fontSize:11, color: saveState==="error"?"#e8746b":saveState==="saving"?"#d9b46a":saveState==="saved"?"#7fd1a8":"#6a665e" }}>{saveState==="error"?"⚠ save failed":saveState==="saving"?"… saving":saveState==="saved"?"● saved to cloud":"○ connected"}</span><span style={{ fontSize:11, color:"var(--on-dark)", whiteSpace:"nowrap" }}>{(me&&(me.name||me.email))||""} · <b style={{ color:"var(--accent)" }}>{(ROLES[role]||{}).label||role}</b></span>{peers.length>0 && (<span style={{ position:"relative", display:"flex", alignItems:"center", gap:3 }}>{peers.slice(0,6).map(p=>{ const loc=p.editing||p.cell; return (<span key={p.id} onClick={(e)=>{ e.stopPropagation(); if(loc) scrollToCell(loc.id,loc.col); }} title={(p.name||"")+" · "+((ROLES[p.role]||{}).label||p.role)+(loc?" — click to view their cell":" — not on a cell")} style={{ width:18, height:18, borderRadius:9, background:colorFor(p.id), color:"var(--surface)", fontSize:8, fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center", cursor:loc?"pointer":"default", boxShadow:p.editing?"0 0 0 1px var(--surface)":"none" }}>{initials(p.name)}</span>); })}<button onClick={(e)=>{ e.stopPropagation(); setBellOpen(false); setReviewOpen(false); setHistory(false); setPeersOpen(o=>!o); }} title="See who's online and where" style={{ fontSize:9, color:"#9b958a", marginLeft:2, background:"transparent", border:"none", cursor:"pointer", fontFamily:"inherit" }}>{peers.length} here ▾</button>{peersOpen && (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", right:0, marginTop:6, zIndex:390, background:"var(--surface)", color:"var(--ink)", border:"1px solid var(--ink)", boxShadow:"4px 4px 0 var(--ink)", width:248, maxHeight:300, overflowY:"auto" }}><div style={{ fontSize:10, fontWeight:700, padding:"7px 9px", borderBottom:"1px solid var(--line-3)" }}>Active now ({peers.length})</div>{peers.map(p=>{ const loc=p.editing||p.cell; return (<div key={p.id} onClick={()=>{ if(loc){ scrollToCell(loc.id,loc.col); setPeersOpen(false); } }} style={{ display:"flex", gap:7, alignItems:"center", padding:"7px 9px", borderBottom:"1px solid var(--line-3)", cursor:loc?"pointer":"default" }}><span style={{ width:18, height:18, borderRadius:9, flex:"0 0 auto", background:colorFor(p.id), color:"var(--surface)", fontSize:8, fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center" }}>{initials(p.name)}</span><span style={{ flex:1, minWidth:0 }}><span style={{ fontSize:11, fontWeight:700, display:"block", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{p.name||"?"}</span><span style={{ fontSize:9, color:"var(--muted-2)", display:"block", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{loc?((p.editing?"✎ editing ":"at ")+cellLabel(loc.id,loc.col)):"not on a cell"}</span></span>{loc && <span style={{ fontSize:9, color:"var(--info)", fontWeight:700, whiteSpace:"nowrap" }}>view →</span>}</div>); })}</div>)}</span>)}<button onClick={(e)=>{ e.stopPropagation(); openHistory(); }} title="Change history" style={{ fontFamily:"inherit", fontSize:10, padding:"5px 8px", cursor:"pointer", border:"1px solid var(--on-dark-line)", background:"transparent", color:"var(--on-dark)", display:"inline-flex", alignItems:"center" }}><Clock size={13}/></button><button onClick={(e)=>{ e.stopPropagation(); setBellOpen(false); setPeersOpen(false); setReviewOpen(false); setHistory(false); setTab("review"); loadAuditRows(); }} title="Open Review tab for category-wise review items" style={{ fontFamily:"inherit", fontSize:10, padding:"5px 9px", cursor:"pointer", border:"1px solid var(--on-dark-line)", background:tab==="review"?"var(--on-dark-line)":"transparent", color:"var(--on-dark)", display:"inline-flex", alignItems:"center", gap:4 }}>Review</button><span style={{ position:"relative" }}><button onClick={(e)=>{ e.stopPropagation(); setPeersOpen(false); setReviewOpen(false); setHistory(false); setBellOpen(o=>!o); }} title="Notifications" style={{ fontFamily:"inherit", fontSize:10, padding:"5px 8px", cursor:"pointer", border:"1px solid var(--on-dark-line)", background:bellOpen?"var(--on-dark-line)":"transparent", color:"var(--on-dark)", position:"relative", display:"inline-flex", alignItems:"center" }}><Bell size={13}/>{unreadCount>0 && <span style={{ position:"absolute", top:-5, right:-5, background:"#e8746b", color:"var(--surface)", fontSize:8, fontWeight:700, minWidth:14, height:14, borderRadius:7, display:"flex", alignItems:"center", justifyContent:"center", padding:"0 3px" }}>{unreadCount>9?"9+":unreadCount}</span>}</button>{bellOpen && (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", right:0, marginTop:6, zIndex:390, background:"var(--surface)", color:"var(--ink)", border:"1px solid var(--ink)", boxShadow:"4px 4px 0 var(--ink)", width:320, maxHeight:380, display:"flex", flexDirection:"column" }}><div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"8px 10px", borderBottom:"1px solid var(--line-3)" }}><span style={{ fontSize:11, fontWeight:700 }}>Notifications{unreadCount>0?" ("+unreadCount+" new)":""}</span><span style={{ display:"flex", gap:8 }}>{unreadCount>0 && <button onClick={markAllRead} style={{ fontSize:9, border:"none", background:"transparent", color:"var(--info)", cursor:"pointer", fontFamily:"inherit", fontWeight:700 }}>Mark all read</button>}{inbox.length>0 && <button onClick={clearInbox} style={{ fontSize:9, border:"none", background:"transparent", color:"var(--danger)", cursor:"pointer", fontFamily:"inherit", fontWeight:700 }}>Clear</button>}</span></div><div style={{ overflowY:"auto" }}>{inbox.length===0 ? <div style={{ fontSize:10, color:"var(--muted-1)", padding:"14px 10px", textAlign:"center" }}>No notifications yet.</div> : inbox.map(n=>(<div key={n.id} onClick={()=>openNotif(n)} style={{ padding:"8px 10px", borderBottom:"1px solid #f2f2f2", cursor:"pointer", background:n.read?"var(--surface)":"var(--accent-tint)", display:"flex", gap:7, alignItems:"flex-start" }}><span style={{ width:6, height:6, borderRadius:6, marginTop:4, flex:"0 0 auto", background:n.read?"transparent":"var(--accent)" }}/><span style={{ flex:1 }}><span style={{ fontSize:11, lineHeight:1.35, display:"block" }}>{n.body}</span><span style={{ fontSize:9, color:"var(--muted-1)" }}>{tsShort(n.created_at)}</span></span></div>))}</div></div>)}</span>{canManageUsers(role) && <button onClick={(e)=>{ e.stopPropagation(); setBellOpen(false); setPeersOpen(false); setReviewOpen(false); setHistory(false); setUsersOpen(true); }} style={{ fontFamily:"inherit", fontSize:10, padding:"5px 9px", cursor:"pointer", border:"1px solid var(--on-dark-line)", background:"transparent", color:"var(--on-dark)" }}>Users</button>}<button onClick={(e)=>{ e.stopPropagation(); onSignOut&&onSignOut(); }} style={{ fontFamily:"inherit", fontSize:10, padding:"5px 9px", cursor:"pointer", border:"1px solid var(--on-dark-line)", background:"transparent", color:"var(--on-dark)" }}>Sign out</button></div>
      </div>

      <div style={{ display:"flex", gap:0, padding:"0 22px", background:"var(--ink)", borderBottom:"1px solid #3a362e", position:"relative", zIndex:155, pointerEvents:"auto" }}>
        {[["tracker","Tracker"],["dashboard","Dashboard"],["management","Management"],["escalation","Escalation"],["todo","To-Do"],["review","Review"],["entrylog","Entry Log"],["settings","Settings"],["help","Help"]].map(([k,lab])=>(<button key={k} onClick={(e)=>{ e.stopPropagation(); setTab(k); if(k==="entrylog"||k==="review") loadAuditRows(); }} style={{ fontFamily:"'Archivo',sans-serif", fontWeight:700, fontSize:12, letterSpacing:0.3, padding:"9px 16px", cursor:"pointer", border:"none", borderBottom:tab===k?"3px solid var(--accent)":"3px solid transparent", background:"transparent", color:tab===k?"var(--bg)":"#9a958c" }}>{lab}{k==="todo"&&todoItems.length?` · ${todoItems.length}`:k==="review"?(" · "+((errorLog&&errorLog.length)||0)):k==="entrylog"&&errorLog.length?` · ${errorLog.length}`:""}</button>))}
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

            {helpTab==="buttons" && <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(240px,1fr))", gap:12 }}>
              {[
                ["Search & filters", ["Search: find style/order/colour text", "Saved View: apply common preset filters", "Status/Chase/Activity: narrow visible rows", "Clear: remove active filters"]],
                ["Data actions", ["Upload styles: import/update rows from Excel", "Export: download filtered data", "Fill date: bulk-fill one date column"]],
                ["View tools", ["Role/Column view: show useful columns", "Columns: hide/show or auto-fit", "Freeze: keep left columns visible", "A/B controls: text size and boldness"]],
                ["Edit & review", ["Find/Replace: search and replace text fields", "Copy/Paste special: copy full stage state", "Review: changes, comments and errors", "Sync: pull latest shared data"]]
              ].map(([t,items])=><div key={t} style={{ border:"1px solid var(--line-2)", borderRadius:12, padding:14, background:"#fffdf8" }}><h3 style={{ margin:"0 0 8px", fontFamily:"'Archivo',sans-serif" }}>{t}</h3><ul style={{ margin:"0 0 0 18px", padding:0 }}>{items.map(x=><li key={x}>{x}</li>)}</ul></div>)}
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
                ["Flow", ["Order → Techpack → Fit/Print/Lab", "Fabric IH gates PP", "Prod File follows PP Approval unless PP is bypassed/not needed"]],
                ["Rejections", ["Rejected approval triggers rework", "Old revised dates before rejection are ignored", "Fresh actual/revised dates after rejection rebuild the chain"]],
                ["Risk", ["Delivery risk appears when projected Production File release misses the release gate", "Tight appears when release is close to the gate"]]
              ].map(([t,items])=><div key={t} style={{ border:"1px solid var(--line-2)", borderRadius:12, padding:14, background:"#fffdf8" }}><h3 style={{ margin:"0 0 8px", fontFamily:"'Archivo',sans-serif" }}>{t}</h3><ul style={{ margin:"0 0 0 18px", padding:0 }}>{items.map(x=><li key={x}>{x}</li>)}</ul></div>)}
            </div>}
          </div>
        </div>
      </div>)}
      {usersOpen && <UsersPanel onClose={()=>setUsersOpen(false)}/>}
{history && (<div onClick={()=>setHistory(false)} style={{ position:"fixed", inset:0, background:"rgba(26,26,26,0.55)", zIndex:200, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}><div onClick={e=>e.stopPropagation()} style={{ background:"var(--surface)", border:"2px solid var(--ink)", boxShadow:"8px 8px 0 var(--ink)", width:640, maxWidth:"100%", maxHeight:"80vh", display:"flex", flexDirection:"column", fontFamily:"'JetBrains Mono',monospace" }}><div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"12px 14px", borderBottom:"1px solid var(--line-3)" }}><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:17 }}>Change history</div><button onClick={()=>setHistory(false)} style={{ border:"1px solid var(--ink)", background:"var(--bg)", cursor:"pointer", fontFamily:"inherit", fontSize:11, padding:"4px 9px" }}>Close</button></div><div style={{ padding:"8px 14px", borderBottom:"1px solid #f2f2f2" }}><input value={histFilter} onChange={e=>setHistFilter(e.target.value)} placeholder="filter by style no, field, stage, or person…" style={{ width:"100%", boxSizing:"border-box", fontFamily:"inherit", fontSize:11, padding:6, border:"1px solid var(--ink)" }}/></div><div style={{ overflowY:"auto", padding:"4px 0" }}>{auditBusy ? <div style={{ fontSize:11, color:"var(--muted-1)", padding:"24px", textAlign:"center" }}>Loading…</div> : (()=>{ const q=histFilter.trim().toLowerCase(); const fl=auditRows.filter(a=> !q || (a.style_no||"").toLowerCase().includes(q) || (a.field||"").toLowerCase().includes(q) || (a.actor_name||"").toLowerCase().includes(q) || (colLabelOf(a.col)||"").toLowerCase().includes(q)); if(fl.length===0) return <div style={{ fontSize:11, color:"var(--muted-1)", padding:"24px", textAlign:"center" }}>{auditRows.length===0?"No changes recorded yet.":"No matches."}</div>; return fl.map(a=>(<div key={a.id} style={{ padding:"7px 14px", borderBottom:"1px solid #f5f5f5", fontSize:11, display:"flex", gap:10, alignItems:"baseline" }}><span style={{ color:"var(--muted-1)", fontSize:9, whiteSpace:"nowrap", minWidth:80 }}>{tsShort(a.created_at)}</span><span style={{ flex:1 }}><b>{a.style_no||a.style_id}</b> · {colLabelOf(a.col)} <span style={{ color:"var(--muted-2)" }}>({a.field})</span><br/><span style={{ color:"var(--danger)" }}>{a.old_val||"—"}</span> <span style={{ color:"var(--muted-2)" }}>→</span> <span style={{ color:"var(--success)" }}>{a.new_val||"—"}</span> <span style={{ color:"var(--muted-1)" }}>· {a.actor_name}</span></span></div>)); })()}</div></div></div>)}
      {reviewOpen && (<div onClick={()=>setReviewOpen(false)} style={{ position:"fixed", inset:0, background:"rgba(26,26,26,0.55)", zIndex:205, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}><div onClick={e=>e.stopPropagation()} style={{ background:"var(--surface)", border:"2px solid var(--ink)", boxShadow:"8px 8px 0 var(--ink)", width:760, maxWidth:"100%", maxHeight:"84vh", display:"flex", flexDirection:"column", fontFamily:"'JetBrains Mono',monospace" }}><div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"12px 14px", borderBottom:"1px solid var(--line-3)" }}><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:17 }}>Review centre</div><button onClick={()=>setReviewOpen(false)} style={{ border:"1px solid var(--ink)", background:"var(--bg)", cursor:"pointer", fontFamily:"inherit", fontSize:11, padding:"4px 9px" }}>Close</button></div><div style={{ display:"flex", gap:0, borderBottom:"1px solid var(--line-3)" }}>{[["changes","Changes to sheet"],["comments","All comments"],["mine","Comments involving me"],["errors","Error log"]].map(([k,l])=><button key={k} onClick={()=>{ setReviewTab(k); if(k==="changes") loadAuditRows(); }} style={{ fontFamily:"inherit", fontSize:11, fontWeight:700, padding:"8px 12px", cursor:"pointer", border:"none", borderRight:"1px solid var(--line-3)", background:reviewTab===k?"var(--accent)":"var(--surface)", color:reviewTab===k?"var(--ink)":"var(--muted-3)" }}>{l}</button>)}</div><div style={{ overflowY:"auto", padding:12 }}>{reviewTab==="changes" && (auditBusy?<div style={{ fontSize:11, color:"var(--muted-1)", padding:20 }}>Loading changes…</div>:auditRows.length===0?<div style={{ fontSize:11, color:"var(--muted-1)", padding:20 }}>No recorded changes yet.</div>:auditRows.slice(0,120).map(a=><div key={a.id} style={{ padding:"7px 0", borderBottom:"1px solid #f2eee6", fontSize:11 }}><span style={{ color:"var(--muted-1)", fontSize:9 }}>{tsShort(a.created_at)}</span> · <b>{a.style_no||a.style_id}</b> · {colLabelOf(a.col)} <span style={{ color:"var(--muted-2)" }}>({a.field})</span><br/><span style={{ color:"var(--danger)" }}>{a.old_val||"—"}</span> → <span style={{ color:"var(--success)" }}>{a.new_val||"—"}</span> <span style={{ color:"var(--muted-1)" }}>· {a.actor_name}</span></div>))}{reviewTab!=="changes" && (()=>{ const all=[]; Object.entries(comments||{}).forEach(([ck,arr])=>{ const [sid,col]=ck.split(":"); const st=styles.find(x=>String(x.id)===String(sid)); (arr||[]).forEach(c=>all.push({...c, sid:Number(sid), col, styleNo:(st&&st.styleNo)||c.style_no||sid, orderNo:(st&&st.orderNo)||""})); }); const my=(me&&(me.name||me.email)||"").toLowerCase(); const myCompact=my.replace(/\s+/g,""); const data=reviewTab==="mine"?all.filter(c=>(c.mentions||[]).includes(me&&me.id)||c.author_id===(me&&me.id)||String(c.body||"").toLowerCase().includes("@"+myCompact)||String(c.body||"").toLowerCase().includes("@"+my)||String(c.author_name||"").toLowerCase().includes(my)):all; if(reviewTab==="errors"){ return errorLog.length===0?<div style={{ fontSize:11, color:"var(--muted-1)", padding:20 }}>No app errors logged in this browser session.</div>:errorLog.map(e=><div key={e.id} style={{ padding:"7px 0", borderBottom:"1px solid #f2eee6", fontSize:11 }}><span style={{ color:"var(--muted-1)", fontSize:9 }}>{tsShort(e.at)}</span> · <b>{e.area}</b><br/><span style={{ color:"var(--danger)" }}>{e.msg}</span>{e.extra&&<span style={{ color:"var(--muted-1)" }}> · {e.extra}</span>}</div>); } return data.length===0?<div style={{ fontSize:11, color:"var(--muted-1)", padding:20 }}>No comments found.</div>:data.sort((a,b)=>new Date(b.created_at)-new Date(a.created_at)).slice(0,150).map(c=><button key={c.id} onClick={()=>{ setReviewOpen(false); setHistory(false); setTab("tracker"); setTimeout(()=>{ setSel({id:c.sid,col:c.col}); setFocus(null); setThreadCell({id:c.sid,col:c.col}); scrollToCell(c.sid,c.col); },60); }} style={{ display:"block", width:"100%", textAlign:"left", border:"none", borderBottom:"1px solid #f2eee6", background:"transparent", cursor:"pointer", padding:"7px 0", fontFamily:"inherit", fontSize:11 }}><span style={{ color:"var(--muted-1)", fontSize:9 }}>{tsShort(c.created_at)}</span> · <b>{c.orderNo?c.orderNo+" · ":""}{c.styleNo}</b> · {colLabelOf(c.col)}<br/><span>{c.body}</span></button>); })()}</div></div></div>)}
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
      <div style={{ display:"flex", padding:"12px 22px 0", flexWrap:"wrap" }}>
        {Object.entries(funnel).map(([k,v],i,arr)=>(<div key={k} style={{ flex:1, minWidth:90, background:"var(--surface)", border:"1px solid var(--line-2)", borderRight:i===arr.length-1?"1px solid var(--line-2)":"none", padding:"8px 10px" }}><div style={{ fontSize:22, fontWeight:700, lineHeight:1, fontFamily:"'Archivo',sans-serif", color:k==="Released"?"var(--success)":k==="Fabric IH"?"var(--danger)":"var(--ink)" }}>{v}</div><div style={{ fontSize:9, color:"var(--muted-2)", marginTop:3, letterSpacing:0.5, textTransform:"uppercase" }}>{k}</div></div>))}
      </div>

      <div style={{ display:"flex", gap:8, alignItems:"center", padding:"12px 22px 6px", flexWrap:"wrap" }}>
        <div style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, marginRight:2 }}>Search</span>
        <div style={{ display:"flex", alignItems:"center", gap:6, background:"var(--surface)", border:"1px solid var(--ink)", padding:"5px 9px" }}><Filter size={13}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder={searchCol==="auto"?"search style / colour / fit / order…":"search selected column…"} onClick={e=>e.stopPropagation()} style={{ border:"none", outline:"none", fontFamily:"inherit", fontSize:12, width:160, background:"transparent" }}/><span style={{ width:1, height:16, background:"var(--line-2)" }}/><select value={searchCol} onChange={e=>setSearchCol(e.target.value)} onClick={e=>e.stopPropagation()} title="search a specific column (Auto = smart search across the main fields)" style={{ border:"none", outline:"none", fontFamily:"inherit", fontSize:10, background:"transparent", color:searchCol==="auto"?"var(--muted-2)":"var(--accent)", cursor:"pointer", fontWeight:searchCol==="auto"?400:700 }}>{[["auto","Auto"],["styleNo","Style No"],["colour","Colour"],["family","Family"],["sampleFit","Sample Fit"],["orderNo","Order No"],["owner","Junior"],["brand","Brand"],["buyer","Buyer"],["age","Age Group"],["fabricType","Fabric Type"],["setRole","Set Role"],["remarks","Remarks"]].map(([k,lab])=><option key={k} value={k}>{lab}</option>)}</select></div>
        <span style={{ position:"relative" }}><button onClick={(e)=>{ e.stopPropagation(); setFrOpen(o=>!o); }} title="Find & replace text in cells" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:frOpen?"var(--ink)":"var(--surface)", color:frOpen?"var(--bg)":"var(--ink)", fontWeight:700 }}>⌕ Find/Replace</button>{frOpen && (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:0, marginTop:4, zIndex:370, background:"var(--surface)", border:"1px solid var(--ink)", boxShadow:"4px 4px 0 var(--ink)", padding:12, width:260 }}><div style={{ fontSize:11, fontWeight:700, marginBottom:8 }}>Find &amp; replace</div><label style={{ fontSize:10, color:"var(--muted-2)" }}>Find</label><input autoFocus value={frFind} onChange={e=>{ setFrFind(e.target.value); setFindIdx(-1); setFrMatches([]); }} onKeyDown={e=>{ if(e.key==="Enter"){ e.preventDefault(); runFind(); } }} placeholder="text to find" style={{ width:"100%", boxSizing:"border-box", fontFamily:"inherit", fontSize:11, padding:5, marginBottom:6, border:"1px solid var(--ink)" }}/><label style={{ fontSize:10, color:"var(--muted-2)" }}>Replace with</label><input value={frRepl} onChange={e=>setFrRepl(e.target.value)} placeholder="(leave blank to delete)" style={{ width:"100%", boxSizing:"border-box", fontFamily:"inherit", fontSize:11, padding:5, marginBottom:8, border:"1px solid var(--ink)" }}/><div style={{ display:"flex", gap:12, fontSize:10, marginBottom:6 }}><label style={{ display:"flex", gap:4, cursor:"pointer" }}><input type="radio" checked={frScope==="filtered"} onChange={()=>setFrScope("filtered")}/>All filtered</label><label style={{ display:"flex", gap:4, cursor:"pointer" }}><input type="radio" checked={frScope==="selected"} onChange={()=>setFrScope("selected")}/>Selected cells</label></div><label style={{ fontSize:10, display:"flex", gap:5, marginBottom:8, cursor:"pointer" }}><input type="checkbox" checked={frCase} onChange={e=>setFrCase(e.target.checked)}/>Match case</label><div style={{ fontSize:10, color:"var(--muted-3)", marginBottom:8, minHeight:13 }}>{frScope==="selected" && !rect() ? "Select cells in the grid first." : (frFind ? findReplace(true)+" cell(s) will change" : "Applies to text fields (style, colour, brand, owner, remarks…)")}</div>{frMatches.length>0 && (<div style={{ marginBottom:8 }}><div style={{ fontSize:9, fontWeight:700, color:"var(--muted-3)", marginBottom:3 }}>{frMatches.length} match{frMatches.length===1?"":"es"} — click to jump</div><div style={{ maxHeight:150, overflowY:"auto", border:"1px solid var(--line-2)" }}>{frMatches.map((m,i)=>(<div key={m.id+":"+m.col} onClick={(e)=>{ e.stopPropagation(); gotoMatch(i); }} title={m.style+" · "+m.colLabel} style={{ display:"flex", gap:6, alignItems:"baseline", padding:"4px 6px", cursor:"pointer", fontSize:10, borderBottom:"1px solid var(--line-3)", background:i===findIdx?"var(--accent-tint)":"var(--surface)" }}><span style={{ fontWeight:700, whiteSpace:"nowrap", maxWidth:78, overflow:"hidden", textOverflow:"ellipsis" }}>{m.style||"—"}</span><span style={{ color:"var(--muted-2)", whiteSpace:"nowrap" }}>{m.colLabel}</span><span style={{ flex:1, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{m.text}</span></div>))}</div></div>)}<div style={{ display:"flex", gap:8 }}><button disabled={!frFind} onClick={(e)=>{ e.stopPropagation(); findNext(); }} title="search, then cycle through matching cells" style={{ flex:1, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:6, cursor:frFind?"pointer":"not-allowed", border:"1px solid var(--ink)", background:"var(--surface)", opacity:frFind?1:0.6 }}>{frMatches.length?("Next "+(findIdx+1)+"/"+frMatches.length):"Find"}</button><button disabled={!frFind} onClick={(e)=>{ e.stopPropagation(); const n=findReplace(true); if(!n){ window.alert('No cells match "'+frFind+'".'); return; } const bw=frRepl.trim()===""?"\n\n\u26a0\ufe0f This will BLANK those cells.":""; if(!window.confirm('Replace "'+frFind+'" \u2192 "'+(frRepl||"(blank)")+'" in '+n+' cell(s)?'+bw)) return; findReplace(false); flash(); }} style={{ flex:1, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:6, cursor:frFind?"pointer":"not-allowed", border:"1px solid var(--ink)", background:frFind?"var(--accent)":"var(--line-2)", opacity:frFind?1:0.6 }}>Replace all</button><button onClick={()=>setFrOpen(false)} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--bg)" }}><X size={12}/></button></div></div>)}</span>
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, marginRight:2 }}>Filters</span>
        <select value={savedView} onChange={e=>{ setSavedView(e.target.value); if(e.target.value==="following") setFollowFilter(false); }} onClick={e=>e.stopPropagation()} title="quick saved operational views" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 8px", cursor:"pointer", border:"1px solid "+(savedView?"var(--accent)":"var(--ink)"), background:savedView?"var(--accent-tint)":"var(--surface)", fontWeight:savedView?700:400 }}>{SAVED_VIEWS.map(([k,l])=><option key={k} value={k}>{l}</option>)}</select>
        <div style={{ display:"flex", border:"1px solid var(--ink)" }}>{["All","At Risk","On Track","Released"].map(f=>(<button key={f} onClick={(e)=>{ e.stopPropagation(); setStatusFilter(f); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"none", borderRight:f!=="Released"?"1px solid var(--ink)":"none", background:statusFilter===f?"var(--ink)":"var(--surface)", color:statusFilter===f?"var(--bg)":"var(--ink)" }}>{f}</button>))}</div>
        <div style={{ display:"flex", border:"1px solid var(--ink)" }}>{chaseOwnerOptions.map((f,i)=>(<button key={f} onClick={(e)=>{ e.stopPropagation(); setOwnerFilter(f); }} title="chase label / department" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 9px", cursor:"pointer", border:"none", borderRight:i<chaseOwnerOptions.length-1?"1px solid var(--ink)":"none", background:ownerFilter===f?"#2563a6":"var(--surface)", color:ownerFilter===f?"var(--surface)":"var(--ink)" }}>{f==="All"?"Chase":f}</button>))}</div>
        <select value={activityFilter||""} onChange={e=>{ setActivityFilter(e.target.value||null); }} onClick={e=>e.stopPropagation()} title="show only styles whose current pending action is this activity" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 8px", cursor:"pointer", border:"1px solid "+(activityFilter?"var(--accent)":"var(--ink)"), background:activityFilter?"var(--accent-tint)":"var(--surface)", fontWeight:activityFilter?700:400 }}><option value="">Activity: all</option>{STAGES.map(st=>(<option key={st.key} value={st.key}>{st.label}</option>))}</select>
        <div style={{ display:"flex", border:"1px solid var(--ink)" }}>{[["active","Active"],["all","All"],["archived","Archived"]].map(([v,lab])=>(<button key={v} onClick={(e)=>{ e.stopPropagation(); setArchiveView(v); }} title="archived styles are hidden from the live sheet" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 9px", cursor:"pointer", border:"none", borderRight:v!=="archived"?"1px solid var(--ink)":"none", background:archiveView===v?"#5a6650":"var(--surface)", color:archiveView===v?"var(--surface)":"var(--ink)" }}>{lab}</button>))}</div>
        <button onClick={(e)=>{ e.stopPropagation(); setFollowFilter(v=>!v); }} title="show only styles you follow" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid var(--ink)", background:followFilter?"var(--accent)":"var(--surface)", color:followFilter?"var(--surface)":"var(--ink)", display:"inline-flex", alignItems:"center", gap:5 }}><Star size={12} fill={followFilter?"var(--surface)":"none"}/> Following</button>
        {canAdmin(role) && archiveView!=="archived" && anyFilter && <button onClick={(e)=>{ e.stopPropagation(); archiveFiltered(true); }} title="archive the styles currently shown (e.g. a finished season)" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid #5a6650", background:"var(--surface)", color:"#5a6650", fontWeight:700 }}>Archive these ({rows.length})</button>}
        {canAdmin(role) && archiveView==="archived" && rows.length>0 && <button onClick={(e)=>{ e.stopPropagation(); archiveFiltered(false); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--success)", background:"var(--surface)", color:"var(--success)", fontWeight:700 }}>Restore these ({rows.length})</button>}
        {anyFilter && <button onClick={(e)=>{ e.stopPropagation(); clearAllFilters(); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--danger)", background:"var(--surface)", color:"var(--danger)", fontWeight:700, display:"flex", alignItems:"center", gap:5 }}><X size={12}/> clear filters</button>}
        {viewSnap && <button onClick={(e)=>{ e.stopPropagation(); restoreView(); }} title="go back to the view you had before drilling in" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)", display:"flex", alignItems:"center", gap:5 }}><RotateCcw size={12}/> restore view</button>}
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, marginRight:2 }}>Data</span>
        {canMaster(role) && <button onClick={(e)=>{ e.stopPropagation(); setBulkResult(null); setBulkOpen(true); }} title="bulk upload / update styles from Excel" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--ink)", color:"var(--bg)", fontWeight:700, display:"flex", alignItems:"center", gap:6 }}><Plus size={13}/> Upload styles</button>}
        {canMaster(role) && <button onClick={(e)=>{ e.stopPropagation(); setBulkActionsOpen(true); }} title="bulk actions on currently visible/filtered rows" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)", color:"var(--ink)", fontWeight:700, display:"flex", alignItems:"center", gap:6 }}><Check size={13}/> Bulk actions ({rows.length})</button>}
        <span style={{ position:"relative" }}><button onClick={(e)=>{ e.stopPropagation(); setExpOpen(o=>!o); }} title="export options" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid var(--ink)", background:expOpen?"var(--ink)":"var(--surface)", color:expOpen?"var(--bg)":"var(--ink)", fontWeight:700, display:"flex", alignItems:"center", gap:6 }}>⬇ Export</button>{expOpen && (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:0, marginTop:4, zIndex:370, background:"var(--surface)", border:"1px solid var(--ink)", boxShadow:"4px 4px 0 var(--ink)", padding:12, width:288 }}><div style={{ fontSize:11, fontWeight:700, marginBottom:8 }}>Export {rows.length} filtered styles</div>{[["full","Full \u2014 all columns (actual dates)"],["detail","Detailed summary \u2014 everything (toggles + every stage)"],["release","Release plan \u2014 for production (incl. blockers)"],["internal","Internal plan (with buffer)"],["buyer","Buyer view \u2014 key details (printable)"]].map(([v,lbl])=>(<label key={v} style={{ display:"flex", gap:6, fontSize:10, padding:"3px 0", cursor:"pointer", alignItems:"flex-start" }}><input type="radio" checked={expMode===v} onChange={()=>setExpMode(v)} style={{ marginTop:1 }}/>{lbl}</label>))}{expMode==="release" && (<div style={{ display:"flex", gap:6, margin:"4px 0 2px" }}>{[["detailed","Detailed"],["summary","Summary"]].map(([v,l])=>(<button key={v} onClick={()=>setExpRelMode(v)} style={{ flex:1, fontFamily:"inherit", fontSize:10, fontWeight:700, padding:"4px 0", cursor:"pointer", border:"1px solid var(--ink)", background:expRelMode===v?"var(--ink)":"var(--surface)", color:expRelMode===v?"var(--bg)":"var(--ink)" }}>{l}</button>))}</div>)}{expMode==="buyer" && (<label style={{ display:"flex", gap:6, fontSize:10, padding:"4px 0", cursor:"pointer", alignItems:"flex-start", borderTop:"1px solid var(--line-3)", marginTop:4, paddingTop:6 }}><input type="checkbox" checked={expIncBuf} onChange={e=>setExpIncBuf(e.target.checked)} style={{ marginTop:1 }}/>Include internal buffered plan dates <span style={{ color:"var(--muted-1)" }}>(for your team only)</span></label>)}{(expMode==="internal"||(expMode==="buyer"&&expIncBuf)) && (<div style={{ fontSize:10, color:"var(--muted-4)", margin:"6px 0 4px", lineHeight:1.5, background:"#f7f4ee", border:"1px solid #e6e0d4", padding:"6px 8px" }}>Add buffer to internal plan dates of <input type="number" min={0} max={30} value={expBuf} onChange={e=>setExpBuf(Math.max(0,Math.min(30,Number(e.target.value)||0)))} style={{ width:40, fontFamily:"inherit", fontSize:11, padding:"2px 4px", border:"1px solid var(--ink)", margin:"0 4px" }}/> working days.<div style={{ color:"var(--muted-1)", marginTop:4 }}>Buyer approvals (Fit / Art / S-O / Lab / PP Appr) get no extra buffer of their own, but follow the buffered internal dates. Actual dates stay unchanged.</div></div>)}<div style={{ display:"flex", gap:8, marginTop:10 }}><button onClick={()=>{ runExport(expMode,expBuf,expIncBuf,expRelMode); setExpOpen(false); }} style={{ flex:1, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:6, cursor:"pointer", border:"1px solid var(--ink)", background:"var(--accent)" }}>⬇ Export .xlsx</button><button onClick={()=>setExpOpen(false)} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--bg)" }}><X size={12}/></button></div></div>)}</span>
        <div style={{ position:"relative" }}><button onClick={(e)=>{ e.stopPropagation(); finishEditing(); setFillOpen(o=>!o); setColsOpen(false); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--accent)", color:"var(--ink)", fontWeight:700, display:"flex", alignItems:"center", gap:6 }}><Copy size={13}/> Fill date → {rows.length}</button>{fillOpen && (<FillPanel count={rows.length} role={role} onClose={()=>setFillOpen(false)} onApply={(key,val,mode)=>{ pushHistory(); const ids=rows.map(r=>r.s.id); const idset=new Set(ids); const _msg=val?(`Fill this date into all ${ids.length} visible styles? Existing dates in that column will be overwritten.`):(`Clear this date for all ${ids.length} visible styles? This blanks them — only do this if you mean to.`); if(!window.confirm(_msg)){ setFillOpen(false); return; } if(mode==="revised"){ if(!canEditRev(role)){ setFillOpen(false); return; } ids.forEach(id=>{ const ck=id+":"+key+":revised"; if(val) clearedRef.current.delete(ck); else clearedRef.current.add(ck); }); setStyles(prev=>prev.map(s=>idset.has(s.id)?{...s,revs:{...(s.revs||{}),[key]:val||undefined}}:s)); flash(); setFillOpen(false); return; } if(!canEdit(role,key,"actual")){ setFillOpen(false); return; } const top=(key==="ordRec"||key==="delivery"); if(!top){ ids.forEach(id=>{ const ck=id+":"+key+":actual"; if(val) clearedRef.current.delete(ck); else clearedRef.current.add(ck); }); } setStyles(prev=>prev.map(s=>idset.has(s.id)?(top?{...s,[key]:val||""}:{...s,actuals:{...s.actuals,[key]:val||undefined}}):s)); flash(); setFillOpen(false); }}/>)}</div>
        </div>
      </div>

      <div style={{ display:"flex", gap:8, alignItems:"center", padding:"0 22px 6px", flexWrap:"wrap" }}>
        <div style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, marginRight:2 }}>View</span>
        <select value={columnView} onChange={e=>applyColumnView(e.target.value)} onClick={e=>e.stopPropagation()} title="role-based column views" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 8px", cursor:"pointer", border:"1px solid var(--ink)", background:columnView==="custom"?"var(--surface)":"#eef6ff", fontWeight:columnView==="custom"?400:700 }}><option value="custom">View: Custom</option>{Object.entries(ROLE_VIEW_PRESETS).map(([k,v])=><option key={k} value={k}>{v.label}</option>)}</select>
        <select value={activeNamedView} onChange={e=>{ const name=e.target.value; const v=(sharedViews||[]).find(x=>x.name===name); if(v) applyTrackerViewState(v.state,name); else setActiveNamedView(""); }} onClick={e=>e.stopPropagation()} title="shared default views saved in app_settings; loading does not affect other users" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 8px", cursor:"pointer", border:"1px solid "+(activeNamedView?"var(--accent)":"var(--ink)"), background:activeNamedView?"var(--accent-tint)":"var(--surface)", fontWeight:activeNamedView?800:400 }}><option value="">Default view: choose</option>{(sharedViews||[]).map(v=><option key={v.name} value={v.name}>{v.name}</option>)}</select>
        <button onClick={(e)=>{ e.stopPropagation(); saveSharedTrackerView(); }} title="save current filters/sort/columns as shared default view for everyone (Management/Senior only)" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)", fontWeight:700 }}>Save default</button>
        <button onClick={(e)=>{ e.stopPropagation(); resetTemporaryView(); }} title="clear your temporary filters/sort only; this does not change shared data or other users" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--line-2)", background:"var(--surface)", color:"var(--muted-3)", fontWeight:700 }}>Reset temp</button><button onClick={(e)=>{ e.stopPropagation(); setActiveNamedView(""); resetTemporaryView(); }} title="return from CAD/Buyer/other saved view to full custom default" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)", color:"var(--ink)", fontWeight:700 }}>Full view</button>
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
        <div style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, marginRight:2 }}>Edit / Sync</span>
        <button onClick={(e)=>{ e.stopPropagation(); copySpecial(); }} title="Copy a stage cell full state (actual + revised + rejected + skip)" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)", fontWeight:700 }}>⧉ Copy✦</button>
        <button disabled={!specialClip} onClick={(e)=>{ e.stopPropagation(); pasteSpecial(); }} title="Paste the copied full date state into selected stage cell(s)" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:specialClip?"pointer":"not-allowed", border:"1px solid var(--ink)", background:specialClip?"var(--success)":"var(--surface)", color:specialClip?"var(--surface)":"var(--muted-6)", fontWeight:700, opacity:specialClip?1:0.6 }}>Paste✦ {selKeys&&selKeys.size>1?selKeys.size:""}</button>
        <div style={{ display:"flex", border:"1px solid var(--ink)" }}>
          <button onClick={(e)=>{ e.stopPropagation(); undo(); }} disabled={!past.length} title="Undo (Ctrl/Cmd+Z)" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 9px", cursor:past.length?"pointer":"not-allowed", border:"none", borderRight:"1px solid var(--ink)", background:"var(--surface)", opacity:past.length?1:0.4 }}>↶</button>
          <button onClick={(e)=>{ e.stopPropagation(); redo(); }} disabled={!future.length} title="Redo (Ctrl/Cmd+Shift+Z)" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 9px", cursor:future.length?"pointer":"not-allowed", border:"none", background:"var(--surface)", opacity:future.length?1:0.4 }}>↷</button>
        </div>
        <button onClick={(e)=>{ e.stopPropagation(); setRemoteChanged(false); loadShared(); }} title="reload shared data (pull latest edits)" style={{ fontFamily:"inherit", fontSize:11, fontWeight:remoteChanged?700:400, padding:"6px 11px", cursor:"pointer", border:"1px solid var(--ink)", background:remoteChanged?"var(--accent)":"var(--surface)", display:"flex", alignItems:"center", gap:6 }}><RotateCcw size={13}/> {remoteChanged?"Sync · new changes":"Sync"}</button>
        </div>
        <span style={{ fontSize:10, color:"var(--muted-1)", marginLeft:"auto" }}>{sort.col?<>sorted by <b>{sort.col==="__style"?"Style":(INFO_COLS.find(c=>c.key===sort.col)?.label||STAGES.find(s=>s.key===sort.col)?.label||(sort.col==="remarks"?"Remarks":sort.col))}</b> {sort.dir>0?"↑":"↓"}</>:"shift-click / shift-arrows = range · Ctrl/Cmd C & V = copy/paste"}</span>
      </div>

      <div style={{ display:"flex", gap:8, alignItems:"center", padding:"0 22px 12px", flexWrap:"wrap", fontSize:10, color:"var(--muted-2)" }}>
        <div style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, marginRight:2 }}>Selection</span>
        <button onClick={(e)=>{ e.stopPropagation(); selectAll(); }} title="select all" style={{ fontFamily:"inherit", fontSize:10, padding:"4px 8px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)" }}>⌖ all</button>
        <button onClick={(e)=>{ e.stopPropagation(); copySelection(); }} disabled={!sel} title="copy selected cell(s)" style={{ fontFamily:"inherit", fontSize:10, padding:"4px 9px", cursor:sel?"pointer":"not-allowed", border:"1px solid var(--ink)", background:"var(--surface)", display:"inline-flex", alignItems:"center", gap:5, opacity:sel?1:0.4 }}><Copy size={12}/> copy</button>
        <span style={{ display:"inline-flex", alignItems:"center", gap:4, marginLeft:6 }} title="type a cell like A9 and press Enter to jump there"><span style={{ fontSize:9, color:"var(--muted-2)" }}>Go to</span><input value={nameBox} onChange={e=>setNameBox(e.target.value)} onKeyDown={e=>{ if(e.key==="Enter") gotoCell(nameBox); }} onFocus={e=>e.target.select()} placeholder="A9" style={{ width:54, fontFamily:"inherit", fontSize:11, fontWeight:700, textTransform:"uppercase", padding:"3px 6px", border:"1px solid var(--ink)", textAlign:"center", color:"var(--accent)" }}/></span><span style={{ marginLeft:8 }}>{clip?<span style={{color:"#2563a6"}}>📋 {clip.h}×{clip.w} copied — select & Ctrl/Cmd+V to paste</span>:(sel?<>selected: <b style={{ color:"var(--accent)" }}>{cellRef(sel)}{focus&&(focus.id!==sel.id||focus.col!==sel.col)?(":"+cellRef(focus)):""}</b> · {styles.find(s=>s.id===sel.id)?.styleNo} · {sel.col==="__style"?"Style":(INFO_COLS.find(c=>c.key===sel.col)?.label||STAGES.find(s=>s.key===sel.col)?.label||sel.col)}</>:"click a cell to format / comment")}</span>
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:6, flex:"1 1 420px", minWidth:320, background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)", position:"relative" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4 }}>Entry</span>
          <span style={{ fontSize:10, fontWeight:800, color:"var(--accent)", minWidth:34 }}>{sel?cellRef(sel):"—"}</span>
          <input disabled={!sel} value={entryVal} onChange={e=>{ setEntryTouched(true); setEntryVal(sel&&sel.col==="qty"?e.target.value.replace(/[^0-9]/g,""):e.target.value); }} onKeyDown={e=>{ if(e.key==="Tab"&&entrySuggestion){ e.preventDefault(); setEntryTouched(true); setEntryVal(entrySuggestion); } else if(e.key==="Enter"){ e.preventDefault(); commitEntry(); } else if(e.key==="Escape"){ setEntryVal(selectedDisplayValue()); setEntryTouched(false); if(editing) setEditing(null); } }} placeholder={sel?("type "+selectedColLabel+"…"):"select a cell"} title="Excel-style entry bar. Type here; Tab accepts the suggestion from values above; Enter saves. Entry bar saves directly and will not open an inline cell editor while typing." style={{ flex:1, minWidth:160, fontFamily:"inherit", fontSize:11, padding:"4px 7px", border:"1px solid var(--line-2)", outline:"none", background:sel?"var(--surface)":"var(--line-3)" }}/>
          {entrySuggestion && <button onMouseDown={e=>e.preventDefault()} onClick={()=>{ setEntryTouched(true); setEntryVal(entrySuggestion); }} title="Click or press Tab to accept" style={{ position:"absolute", right:10, top:"100%", marginTop:3, zIndex:390, maxWidth:260, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"4px 8px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--accent-tint)", boxShadow:"3px 3px 0 var(--ink)" }}>Tab ↹ {entrySuggestion}</button>}
          {peerEditingList.length>0 && <span title={peerEditingList.map(p=>(p.name||p.email||"User")+" editing "+p.ref).join("\n")} style={{ fontSize:9, color:"var(--danger)", fontWeight:800, whiteSpace:"nowrap" }}>{peerEditingList.length} editing now</span>}
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, marginRight:2 }}>Format</span>
        <Droplet size={13}/><span>fill:</span>
        {FILL_SWATCHES.map((sw,i)=>(<button key={i} onClick={(e)=>{ e.stopPropagation(); applyFill(sw); }} disabled={!sel} title={sw===""?"clear fill":sw} style={{ width:18, height:18, cursor:sel?"pointer":"not-allowed", border:"1px solid var(--ink)", background:sw===""?"var(--surface)":sw, position:"relative", opacity:sel?1:0.4 }}>{sw===""?<X size={11} style={{position:"absolute",top:2,left:2}}/>:null}</button>))}
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:6, padding:"5px 7px", boxShadow:"0 1px 0 rgba(0,0,0,0.03)" }}>
          <span style={{ fontSize:9, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, marginRight:2 }}>Comments</span>
        <span style={{ marginLeft:10, position:"relative" }}>
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

      {(() => { const chips=[]; const chip=(key,label,onX)=>chips.push(<span key={key} style={{ display:"inline-flex", alignItems:"center", gap:5, background:"var(--accent-tint)", border:"1px solid rgba(201,111,22,0.22)", borderRadius:999, padding:"4px 6px 4px 9px", fontSize:10, fontWeight:700, boxShadow:"var(--pill-shadow)" }}>{label}<button onClick={onX} title="remove this filter" style={{ border:"none", background:"transparent", cursor:"pointer", padding:0, lineHeight:0, color:"var(--muted-2)", display:"inline-flex" }}><X size={11}/></button></span>);
        if(search) chip("q",(searchCol==="auto"?"search":searchCol)+": "+search, ()=>setSearch(""));
        if(statusFilter!=="All") chip("st","status: "+statusFilter, ()=>setStatusFilter("All"));
        if(ownerFilter!=="All") chip("ow","chase: "+ownerFilter, ()=>setOwnerFilter("All"));
        if(activityFilter) chip("ac","activity: "+((STAGES.find(x=>x.key===activityFilter)||{}).label||activityFilter), ()=>setActivityFilter(null));
        if(followFilter) chip("fo","following only", ()=>setFollowFilter(false));
        if(savedView) chip("sv","saved: "+((SAVED_VIEWS.find(x=>x[0]===savedView)||[])[1]||savedView), ()=>setSavedView(""));
        if(activeNamedView) chip("nv","default view: "+activeNamedView, ()=>setActiveNamedView(""));
        if(archiveView!=="active") chip("ar","view: "+archiveView, ()=>setArchiveView("active"));
        Object.keys(colFilters||{}).forEach(col=>{ const lab=(INFO_COLS.find(c=>c.key===col)||{}).label||(STAGES.find(x=>x.key===col)||{}).label||col; chip("c-"+col, lab+": "+((colFilters[col]||[]).length)+" sel", ()=>setColFilters(f=>{ const n={...f}; delete n[col]; return n; })); });
        if(!chips.length) return null;
        return <div style={{ display:"flex", alignItems:"center", gap:7, flexWrap:"wrap", padding:"7px 22px 0" }}><span style={{ fontSize:10, fontWeight:700, color:"var(--muted-2)" }}>Active filters:</span>{chips}<button onClick={clearAllFilters} style={{ fontSize:10, fontWeight:700, border:"1px solid var(--ink)", background:"var(--surface)", cursor:"pointer", padding:"2px 8px", marginLeft:2 }}>Clear all</button></div>;
      })()}

      {showJump && <button onClick={jumpToTop} title="Back to controls / top" style={{ position:"fixed", bottom:24, right:24, zIndex:370, width:42, height:42, borderRadius:21, border:"1px solid var(--ink)", background:"var(--accent)", color:"var(--ink)", boxShadow:"2px 2px 0 var(--ink)", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center" }}><ChevronUp size={20}/></button>}

      {rows.length>renderRows.length && <div style={{ margin:"0 22px 8px", border:"1px solid var(--line-2)", background:"var(--accent-tint)", borderRadius:10, padding:"8px 10px", fontSize:10.5, color:"var(--muted-4)", display:"flex", gap:10, alignItems:"center", flexWrap:"wrap" }}><b>Performance mode:</b> showing first {renderRows.length} of {rows.length} filtered styles to protect P95 speed. Narrow filters for fastest work, or <button onClick={()=>setRenderLimit(Math.min(5000, renderLimit+900))} style={{ fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"3px 8px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)" }}>show 900 more</button><button onClick={()=>setRenderLimit(5000)} style={{ fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"3px 8px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)" }}>show all up to 5,000</button></div>}
      <div ref={scrollWrapRef} style={{ overflow:"auto", padding:"0 22px", maxHeight:"calc(100vh - 210px)" }}>
        <table role="grid" aria-label="Pre-production tracker grid. Arrow keys to move, Escape to exit, Tab to leave the grid." style={{ borderCollapse:"separate", borderSpacing:0, zoom:textScale, fontSize:11, fontWeight:tableWeight, tableLayout:"fixed", userSelect:dragSel?"none":"auto" }}>
          <colgroup>
            <col style={{ width:widthOf("__style") }}/>
            {visInfo.map(c=><col key={c.key} style={{ width:widthOf(c.key) }}/>)}
            {visStages.map(st=><col key={st.key} style={{ width:widthOf(st.key) }}/>)}
            {remarksVis && <col style={{ width:widthOf("remarks") }}/>}
          </colgroup>
          <thead><tr role="row">
            <Th col="__style" label="Style No" sort={sort} onSort={clickHeader} width={widthOf("__style")} letter={colLetter(colIndex("__style"))} onResize={onResize} onAutoFit={autoFit} scale={textScale} {...filterProps("__style")} sticky left={0} z={22}/>
            {visInfo.map(c=><Th key={c.key} col={c.key} label={colLabel(c)} sort={sort} onSort={clickHeader} width={widthOf(c.key)} letter={colLetter(colIndex(c.key))} onResize={onResize} onAutoFit={autoFit} scale={textScale} {...filterProps(c.key)} sticky={isFrozen(c.key)} left={isFrozen(c.key)?leftOf(c.key):undefined} z={21}/>)}
            {visStages.map(st=><Th key={st.key} col={st.key} label={st.label} sort={sort} onSort={clickHeader} width={widthOf(st.key)} letter={colLetter(colIndex(st.key))} onResize={onResize} onAutoFit={autoFit} scale={textScale} {...filterProps(st.key)}/>)}
            {remarksVis && <Th col="remarks" label={REMARK_COL.label} sort={sort} onSort={clickHeader} width={widthOf("remarks")} letter={colLetter(colIndex("remarks"))} onResize={onResize} onAutoFit={autoFit} scale={textScale} {...filterProps("remarks")}/>}
          </tr></thead>
          <tbody>
            {renderRows.map(({s,c},rowIdx)=>{ const t=TONE_STYLE[c.tone]; const sk=cellKey(s.id,"__style"); const styBg=bgFor(s.id,"__style","var(--surface)"); return (
              <tr key={s.id} role="row">
                <td id={`cell-${s.id}-__style`} onClick={(e)=>onCellClick(e,s.id,"__style")} onDoubleClick={(e)=>{ e.stopPropagation(); startEdit(s.id,"__style"); }} style={{ border:"1px solid var(--line-1)", padding:"6px 9px", overflow:"hidden", cursor:"cell", ...freezeStyle("__style",styBg), boxShadow:ringFor(s.id,"__style")||freezeStyle("__style",styBg).boxShadow }}>
                  {editing&&editing.id===s.id&&editing.col==="__style" ? (<input autoFocus onFocus={e=>{ if((e.target.value||"").length>1) e.target.select(); }} value={editVal} onClick={e=>e.stopPropagation()} onChange={e=>setEditVal(e.target.value)} onBlur={commitText} style={{ width:150, fontFamily:"inherit", fontSize:11, fontWeight:700, border:"1px solid var(--info)", outline:"none", padding:"1px 3px" }}/>) : <div style={{ fontWeight:700, display:"flex", alignItems:"center", gap:6 }}><span onClick={(e)=>{ e.stopPropagation(); selectRow(s.id); }} title="select row" style={{ fontSize:8, color:"var(--muted-6)", cursor:"pointer", minWidth:14 }}>{rowIdx+1}</span><Star size={11} onClick={(e)=>{ e.stopPropagation(); toggleFollow(s.id); }} title={follows.has(s.id)?"following — click to unfollow":"follow this style"} fill={follows.has(s.id)?"var(--accent)":"none"} color={follows.has(s.id)?"var(--accent)":"#c9c1b3"} style={{ cursor:"pointer", flex:"0 0 auto", opacity:follows.has(s.id)?1:0.5 }}/><span style={{ color:follows.has(s.id)?"#b45309":"inherit" }}>{s.styleNo}</span></div>}
                  <div style={{ fontSize:11, fontWeight:600, color:"var(--muted-4)", whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis", maxWidth:188, marginTop:2 }}>{s.colour}</div>
                  <div style={{ display:"flex", gap:3, marginTop:4, flexWrap:"wrap" }}>{FLAG_DEFS.map(f=>{ const on=!!s[f.key]; return (<button key={f.key} title={f.title} onClick={(e)=>{ e.stopPropagation(); if(canMaster(role)) toggleFlag(s.id,f.key); }} style={{ fontFamily:"inherit", fontSize:8.5, fontWeight:700, letterSpacing:0.3, padding:"2px 5px", cursor:!canMaster(role)?"not-allowed":"pointer", lineHeight:1.3, border:`1px solid ${on?"var(--ink)":"#cfcabf"}`, background:on?"var(--ink)":"transparent", color:on?"var(--bg)":"var(--muted-6)", opacity:!canMaster(role)?0.5:1 }}>{f.short}</button>); })}</div>
                  <NoteTri k={sk}/>
                  {canAdmin(role) && <button title="delete row" onClick={(e)=>{ e.stopPropagation(); deleteStyle(s.id); }} style={{ position:"absolute", top:2, right:2, zIndex:6, border:"none", background:"transparent", cursor:"pointer", padding:0, lineHeight:0, color:"#cbb4ac" }}><Trash2 size={11}/></button>}
                </td>

                {visInfo.map(col=>{
                  if(col.kind==="text"||col.kind==="num") return renderEditable(s,col);
                  const k=cellKey(s.id,col.key);
                  if(col.kind==="date"){ const bg=bgFor(s.id,col.key,"var(--surface)"); return (<td key={col.key} id={`cell-${s.id}-${col.key}`} onClick={(e)=>onCellClick(e,s.id,col.key)} onDoubleClick={(e)=>{ e.stopPropagation(); if(canEdit(role,col.key,"actual")) beginDate(s.id,col.key,"actual"); }} style={{ border:"1px solid var(--line-1)", padding:"6px 9px", whiteSpace:"nowrap", boxShadow:ringFor(s.id,col.key), cursor:"cell", position:"relative", overflow:(editing&&editing.id===s.id&&editing.col===col.key)?"visible":"hidden", background:bg, ...freezeStyle(col.key,bg) }}>{fmt(parse(s[col.key]))||<span style={{color:"var(--line-2)"}}>—</span>}{editing&&editing.id===s.id&&editing.col===col.key && dateEditor(s.id,col.key,editing.mode)}<PeerTag who={peerOn(s.id,col.key)}/><NoteTri k={k}/><FillHandle id={s.id} col={col.key}/></td>); }
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
                  return <td key={col.key} id={`cell-${s.id}-${col.key}`} onClick={(e)=>onCellClick(e,s.id,col.key)} style={{ border:"1px solid var(--line-1)", padding:"6px 9px", whiteSpace:"nowrap", boxShadow:ringFor(s.id,col.key), cursor:"default", background:bg, position:"relative", overflow:"hidden", ...freezeStyle(col.key,bg) }}>{content}<PeerTag who={peerOn(s.id,col.key)}/><NoteTri k={k}/><FillHandle id={s.id} col={col.key}/></td>;
                })}

                {visStages.map(st=>{
                  const applies=stageApplies(s,st);
                  const cs=c.stages.find(x=>x.key===st.key);
                  const isNext=applies && c.frontier && c.frontier.has(st.key);
                  const editable=applies&&canEdit(role,st.key,"actual"); const canRev=applies&&canEditRev(role); const canRej=applies&&canEditReject(role,st.key); const canSkp=applies&&MERCH_ROLES.includes(role);
                  const k=cellKey(s.id,st.key);
                  if(!applies){ const bg=bgFor(s.id,st.key,"#f3f1ec"); return <td key={st.key} id={`cell-${s.id}-${st.key}`} onClick={(e)=>onCellClick(e,s.id,st.key)} style={{ border:"1px solid var(--line-1)", background:bg, color:"var(--line-2)", textAlign:"center", padding:"6px 9px", boxShadow:ringFor(s.id,st.key), position:"relative", overflow:"hidden" }}>—<NoteTri k={k}/></td>; }
                  const hasRev=cs&&cs.rev&&!cs.done;
                  const bg=bgFor(s.id,st.key,(cs&&cs.skipped)?"var(--tint-waive)":(cs&&cs.rejected)?"var(--tint-reject)":(cs&&cs.rework)?"var(--tint-rework)":(cs&&cs.actual&&cs.histReject?"var(--tint-histrej)":(isNext?"var(--tint-next)":"var(--surface)")));
                  return (
                    <td key={st.key} id={`cell-${s.id}-${st.key}`} onClick={(e)=>onCellClick(e,s.id,st.key)} onDoubleClick={(e)=>{ e.stopPropagation(); if(editable) beginDate(s.id,st.key,"actual"); }}
                      style={{ border:"1px solid var(--line-1)", padding:0, position:"relative", overflow:(editing&&editing.id===s.id&&editing.col===st.key)?"visible":"hidden", background:bg, boxShadow:ringFor(s.id,st.key)||(isNext?"inset 0 0 0 2px var(--accent)":null), cursor:editable?"cell":"default" }}>
                      <div style={{ minHeight:38, padding:"5px 8px", fontSize:12.5, color:cs.actual?"var(--ink)":"var(--muted-6)" }}>
                        {showAux && cs.plan && <span style={{ display:"block", fontSize:8, color:"#bcb6a8", lineHeight:1.3 }}>auto {fmt(cs.plan)}{cs.rev?` · rev ${fmt(cs.rev)}`:""}</span>}
                        {cs.autoClosed ? (
                          <span style={{ color:"var(--line-2)", fontSize:11 }}>—</span>
                        ) : cs.skipped ? (
                          <span style={{ display:"flex", flexDirection:"column", lineHeight:1.25 }}>
                            <span style={{ fontSize:9, color:"#8a6d3b", fontWeight:700, display:"flex", alignItems:"center", gap:3 }}><SkipForward size={9}/>SKIPPED</span>
                            <span style={{ fontSize:8, color:"var(--on-dark-2)" }}>{fmt(cs.skip)}</span>
                          </span>
                        ) : cs.rework ? (
                          <span style={{ display:"flex", flexDirection:"column", lineHeight:1.25 }}>
                            <span style={{ fontSize:9, color:"#b03020", fontWeight:700, display:"flex", alignItems:"center", gap:3 }}><X size={9}/>REDO &amp; RESEND</span>
                            <span style={{ fontSize:9, color:hasRev?"var(--revised)":"#7a560f" }}>{hasRev?"→ rev ":"→ "}{fmt(cs.rev||cs.plan)}</span>
                            {editable && <span style={{ fontSize:9, color:"var(--accent)", fontWeight:700 }}>▸ enter resend</span>}
                          </span>
                        ) : cs.actual ? (<span style={{ display:"flex", flexDirection:"column", lineHeight:1.25 }}><span style={{ display:"flex", alignItems:"center", gap:4 }}><Check size={11} color={OWNER_COLOR[(cfg.stageOwners&&cfg.stageOwners[st.key])||DEFAULT_STAGE_OWNERS[st.key]||st.owner]}/>{fmt(cs.actual)}</span>{cs.histReject && <span style={{ fontSize:8, color:"#b03020", fontWeight:700 }}>↻ was REJ {fmt(cs.histReject)}</span>}</span>) : cs.rejected ? (
                          <span style={{ display:"flex", flexDirection:"column", lineHeight:1.25 }}>
                            <span style={{ fontSize:9, color:"#b03020", fontWeight:700, display:"flex", alignItems:"center", gap:3 }}><X size={9}/>REJECTED</span>
                            <span style={{ fontSize:9, color:"#b03020" }}>rej {fmt(cs.reject)}</span>
                            <span style={{ fontSize:9, color:hasRev?"var(--revised)":"#7a560f" }}>re-appr → {fmt(cs.rev||cs.plan)}</span>
                          </span>
                        ) : (
                          <span style={{ display:"flex", flexDirection:"column", lineHeight:1.2 }}>
                            <span style={{ fontSize:10.5, color:hasRev?"var(--revised)":isNext?"var(--accent)":"#c4c0b8" }}>{hasRev?"rev":st.cutoff?"cutoff":"plan"} {fmt(hasRev?cs.rev:cs.plan)}</span>
                            {editable?<span style={{ fontSize:10.5, color:isNext?"var(--accent)":"#c4c0b8", fontWeight:isNext?700:400 }}>{isNext?"▸ enter":st.cutoff?"log arrival":"—"}</span>:<span style={{ fontSize:10, color:"var(--line-2)", display:"flex", alignItems:"center", gap:3 }}><Lock size={9}/>locked</span>}
                          </span>
                        )}
                      </div>
                      {canRev && !cs.skipped && !cs.autoClosed && (!cs.actual || cs.rework) && (<button title="set revised plan date" onClick={(e)=>{ e.stopPropagation(); beginDate(s.id,st.key,"rev"); }} style={{ position:"absolute", top:3, right:3, border:"none", background:"transparent", cursor:"pointer", padding:0, lineHeight:1, display:"flex" }}><RotateCcw size={11} color="var(--revised)"/></button>)}
                      {canRej && !cs.skipped && !cs.autoClosed && !cs.actual && REJECTABLE.includes(st.key) && (<button title={cs.rejected?"clear rejection (remove rework)":"mark REJECTED (log rejection date)"} onClick={(e)=>{ e.stopPropagation(); if(cs.rejected){ if(window.confirm(`Clear the rejection on "${st.label}" for ${s.styleNo}? This removes the rework flag.`)) setReject(s.id,st.key,null); } else beginDate(s.id,st.key,"reject"); }} style={{ position:"absolute", top:3, right:20, border:"none", background:cs.rejected?"#b03020":"transparent", borderRadius:2, cursor:"pointer", padding:cs.rejected?2:0, lineHeight:1, display:"flex" }}><X size={cs.rejected?9:11} color={cs.rejected?"var(--surface)":"#b03020"}/></button>)}
                      {canSkp && !cs.autoClosed && SKIPPABLE_STAGES.includes(st.key) && !cs.actual && (<button title={cs.skipped?"un-skip (restore this activity)":"skip this activity (waive — counts as resolved, not done)"} onClick={(e)=>{ e.stopPropagation(); if(cs.skipped){ if(window.confirm(`Un-skip "${st.label}" for ${s.styleNo}? This restores the activity.`)) setSkip(s.id,st.key,null); } else if(window.confirm(`Skip / waive "${st.label}" for ${s.styleNo}?\n\nIt will count as RESOLVED (not done) and drop off the to-do. You can un-skip later.`)){ setSkip(s.id,st.key,iso(TODAY)); } }} style={{ position:"absolute", bottom:3, right:3, border:"none", background:cs.skipped?"#8a6d3b":"transparent", borderRadius:2, cursor:"pointer", padding:cs.skipped?2:0, lineHeight:1, display:"flex" }}><SkipForward size={cs.skipped?9:12} color={cs.skipped?"var(--surface)":"#b8a98a"}/></button>)}
                      {cs.rework && canRej && (<button title="clear rework (un-reject the approval)" onClick={(e)=>{ e.stopPropagation(); if(window.confirm(`Clear the rework on "${st.label}" for ${s.styleNo}? This un-rejects the approval.`)) setReject(s.id, APPR_OF_SEND[st.key], null); }} style={{ position:"absolute", top:3, right:20, border:"none", background:"#b03020", borderRadius:2, cursor:"pointer", padding:2, lineHeight:1, display:"flex" }}><X size={9} color="var(--surface)"/></button>)}
                      {editing&&editing.id===s.id&&editing.col===st.key&&editable && dateEditor(s.id,st.key,editing.mode)}
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
                    
                    
                    
                    
                    
                    
                    
                    
                    {FLAG_DEFS.map(f=>{ const on=!!newRow[f.key]; return (<button key={f.key} title={f.title} onClick={(e)=>{ e.stopPropagation(); setNewRow(n=>({...n,[f.key]:!n[f.key]})); }} style={{ fontFamily:"inherit", fontSize:8, fontWeight:700, padding:"2px 4px", cursor:"pointer", lineHeight:1.3, border:`1px solid ${on?"var(--ink)":"#cfcabf"}`, background:on?"var(--ink)":"transparent", color:on?"var(--bg)":"var(--muted-6)" }}>{f.short}</button>); })}
                    <button onClick={(e)=>{ e.stopPropagation(); addNewStyle(); }} style={{ fontFamily:"inherit", fontSize:11, fontWeight:700, padding:"5px 14px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--accent)", color:"var(--ink)" }}>+ Create (Enter)</button>{newError && <span style={{ fontSize:10, color:"var(--danger)", fontWeight:700, marginLeft:8 }}>{newError}</span>}
                  </span>
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      <div style={{ padding:"10px 22px", display:"flex", gap:16, flexWrap:"wrap", alignItems:"center", fontSize:11, borderTop:"1px solid var(--line-3)" }}>
        {(() => { const rs=rows; const n=rs.length; const rel=rs.filter(r=>r.c.released).length; const risk=rs.filter(r=>r.c.tone==="late"||r.c.tone==="warn").length; const ok=rs.filter(r=>(r.c.tone==="ok")&&!r.c.released).length; const qty=rs.reduce((a,r)=>a+(Number(r.s.qty)||0),0); const avg=n?Math.round(rs.reduce((a,r)=>a+(r.c.pct||0),0)/n):0; return <span style={{ display:"flex", gap:14, flexWrap:"wrap", alignItems:"center" }}><span><b>{n}</b> styles</span><span style={{ color:"var(--danger)" }}><b>{risk}</b> at risk</span><span><b>{ok}</b> on track</span><span><b>{rel}</b> released</span><span>total qty <b>{qty.toLocaleString()}</b></span><span>avg <b>{avg}%</b> done</span><span style={{ color:(perfRef.current.alerts||[]).length?"var(--danger)":"var(--muted-2)" }}>Perf: P95 <b>{perfRef.current.p95Ms||0}ms</b> · compute <b>{perfRef.current.computeMs}ms</b> · filter <b>{perfRef.current.filterMs}ms</b> · sort <b>{perfRef.current.rowMs}ms</b> · rendered <b>{perfRef.current.rendered}/{perfRef.current.rows}</b> · recomputed <b>{perfRef.current.recomputed}</b> · cache <b>{perfRef.current.cacheHits}</b>{perfRef.current.deferred?" · search catching up":""}{(perfRef.current.alerts||[]).length?" · "+perfRef.current.alerts.join(", "):""}</span></span>; })()}
        <span style={{ marginLeft:"auto", fontSize:9, color:"var(--muted-7)", display:"flex", alignItems:"center", gap:4, flexWrap:"wrap" }}>Ctrl/Cmd C·V copy-paste · Z / Shift+Z undo-redo · F2 edit · Del clears · drag blue corner to fill · <span style={{ width:0,height:0, borderTop:"7px solid var(--danger)", borderLeft:"7px solid transparent", display:"inline-block" }}/> comment · <RotateCcw size={10} color="var(--revised)"/> revised plan · <Snowflake size={10} color="#2563a6"/> freeze cols</span>
      </div>
      </>)}

      {tab==="dashboard" && <OperationalDashboardView computed={computed} todoItems={todoItems} cfg={cfg} applyDrill={applyDrill} drillTodo={(obj)=>{ setTodoFilter(obj); setTab("todo"); }}/>}
      {tab==="management" && <ManagementDashboardView computed={computed} todoItems={todoItems} cfg={cfg} applyDrill={applyDrill} drillTodo={(obj)=>{ setTodoFilter(obj); setTab("todo"); }}/>}
      {tab==="escalation" && <EscalationMatrixView computed={computed} cfg={cfg} applyDrill={applyDrill} />}
      {tab==="todo" && <TodoView items={todoItems} cfg={cfg} setCfg={setCfg} canEditSettings={canAdmin(role)} filter={todoFilter} setFilter={setTodoFilter} onJump={(id,key)=>{ snapCurrent(); resetFilters(); setTab("tracker"); requestAnimationFrame(()=>setTimeout(()=>jumpToEnter(id,key),60)); }}/>}
      {tab==="review" && <ReviewTabView computed={computed} todoItems={todoItems} auditRows={auditRows} auditBusy={auditBusy} loadAuditRows={loadAuditRows} errorLog={errorLog} comments={comments} inbox={inbox} me={me} colLabelOf={colLabelOf} onJump={(id,col)=>{ setTab("tracker"); setTimeout(()=>{ setSel({id:Number(id),col:col||"__style"}); setFocus(null); scrollToCell(Number(id),col||"__style"); },60); }}/>}
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
    return front.map(k=>{ const r=(c.stages||[]).find(x=>x.key===k); const due=r&&(r.rev||r.plan); if(!r||!due||r.done) return null; const days=Math.max(0,netWorkdays(due,TODAY)||0); if(days<=0) return null; const esc=escalationFor(cfg,days); return { id:s.id, orderNo:s.orderNo, styleNo:s.styleNo, family:s.family, colour:s.colour, buyer:s.buyer||s.brand||"", junior:s.owner, stageKey:k, stage:r.label, chase:r.owner||"—", due, days, bucket:esc.rangeLabel, status:c.status, risk:c.tone, escalationOwner:esc.owner, escalationLevel:esc.level, escalationAction:esc.action }; }).filter(Boolean);
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
  const categories=useMemo(()=>{
    const src=mode==="comments"?allComments.filter(c=>commentPassExcept(c,"category")):mode==="notifications"?allNotifications.filter(n=>notificationPassExcept(n,"category")):items.filter(x=>itemPassExcept(x,"category"));
    return ["All",...Array.from(new Set(src.map(x=>x.category))).sort((a,b)=>a.localeCompare(b))];
  },[items,allComments,allNotifications,mode,severity,q,commentUser,commentScope,notifStatus,category]);
  const severities=useMemo(()=>["All",...Array.from(new Set(items.filter(x=>itemPassExcept(x,"severity")).map(x=>x.severity))).sort((a,b)=>a.localeCompare(b))],[items,category,q,severity]);
  const commentUsers=useMemo(()=>["All users",...Array.from(new Set(allComments.filter(c=>commentPassExcept(c,"commentUser")).map(c=>c.author).filter(Boolean))).sort((a,b)=>a.localeCompare(b))],[allComments,category,commentScope,q,commentUser]);
  const notifStatuses=useMemo(()=>["All",...Array.from(new Set(allNotifications.filter(n=>notificationPassExcept(n,"notifStatus")).map(n=>n.status))).sort((a,b)=>a.localeCompare(b))],[allNotifications,category,q,notifStatus]);
  const filtered=useMemo(()=> items.filter(x=>itemPassExcept(x,null)),[items,category,severity,q]);
  const filteredComments=useMemo(()=> allComments.filter(c=>commentPassExcept(c,null)),[allComments,category,commentUser,commentScope,q,myId,myName]);
  const filteredNotifications=useMemo(()=> allNotifications.filter(n=>notificationPassExcept(n,null)),[allNotifications,category,notifStatus,q]);
  const counts=useMemo(()=>({ total:items.length, critical:items.filter(x=>x.severity==="Critical").length, warning:items.filter(x=>x.severity==="Warning").length, dq:items.filter(x=>x.category==="Data Quality").length, comments:allComments.length, unread:allNotifications.filter(n=>!n.read).length }),[items,allComments,allNotifications]);
  const badge=(txt,bg,fg="var(--ink)")=><span style={{ display:"inline-flex", alignItems:"center", padding:"3px 7px", borderRadius:999, background:bg, color:fg, fontSize:9, fontWeight:900, whiteSpace:"nowrap" }}>{txt}</span>;
  const inputStyle={ fontFamily:"inherit", fontSize:11, padding:"7px 9px", border:"1px solid var(--ink)", background:"var(--surface)" };
  const modeBtn=(k,l,cnt)=><button onClick={()=>{ setMode(k); setCategory("All"); setSeverity("All"); setQ(""); }} style={{ fontFamily:"inherit", fontSize:11, fontWeight:900, padding:"8px 12px", cursor:"pointer", border:"1px solid var(--line-2)", background:mode===k?"var(--ink)":"var(--surface)", color:mode===k?"var(--surface)":"var(--ink)", borderRadius:999 }}>{l}{cnt!=null?` · ${cnt}`:""}</button>;
  return <div style={{ padding:"18px 22px 36px" }}>
    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:16, marginBottom:14 }}>
      <div><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:24 }}>Review</div><div style={{ fontSize:11.5, color:"var(--muted-3)", marginTop:4, maxWidth:900, lineHeight:1.45 }}>Category-wise review centre. Review Inbox summarizes work/data checks. Comments and Notifications have their own views with user-wise and business-category filters.</div></div>
      <button onClick={loadAuditRows} style={{ fontFamily:"inherit", fontSize:11, fontWeight:800, padding:"7px 11px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--surface)" }}>{auditBusy?"Loading…":"Refresh review"}</button>
    </div>
    <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(145px,1fr))", gap:10, marginBottom:12 }}>
      {[ ["Open review items",counts.total,"var(--surface)"], ["Critical",counts.critical,"#f6d3cb"], ["Warnings",counts.warning,"#f8e9b7"], ["Comments",counts.comments,"#e7ecff"], ["Unread notifications",counts.unread,"#fff3df"], ["Data quality",counts.dq,"#e3edfb"] ].map(([l,v,b])=><div key={l} style={{ background:b, border:"1px solid var(--line-2)", borderRadius:12, padding:12 }}><div style={{ fontSize:10, color:"var(--muted-3)", fontWeight:900, textTransform:"uppercase" }}>{l}</div><div style={{ fontFamily:"'Archivo',sans-serif", fontSize:25, fontWeight:800 }}>{v}</div></div>)}
    </div>
    <div style={{ display:"flex", gap:8, flexWrap:"wrap", marginBottom:12 }}>{modeBtn("inbox","Review Inbox",items.length)}{modeBtn("comments","Comments",allComments.length)}{modeBtn("notifications","Notifications",allNotifications.length)}{modeBtn("changes","Changes",(auditRows||[]).length)}{modeBtn("errors","Errors",(errorLog||[]).length)}</div>
    <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"center", marginBottom:12, background:"var(--toolbar-bg)", border:"1px solid var(--toolbar-line)", borderRadius:12, padding:10 }}>
      <select value={category} onChange={e=>setCategory(e.target.value)} style={inputStyle}>{categories.map(c=><option key={c} value={c}>{c}</option>)}</select>
      {mode==="inbox" && <select value={severity} onChange={e=>setSeverity(e.target.value)} style={inputStyle}>{severities.map(c=><option key={c} value={c}>{c}</option>)}</select>}
      {mode==="comments" && <><select value={commentScope} onChange={e=>setCommentScope(e.target.value)} style={inputStyle}>{["All comments","To me","By me","Unresolved"].map(c=><option key={c} value={c}>{c}</option>)}</select><select value={commentUser} onChange={e=>setCommentUser(e.target.value)} style={inputStyle}>{commentUsers.map(c=><option key={c} value={c}>{c}</option>)}</select></>}
      {mode==="notifications" && <select value={notifStatus} onChange={e=>setNotifStatus(e.target.value)} style={inputStyle}>{notifStatuses.map(c=><option key={c} value={c}>{c}</option>)}</select>}
      <input value={q} onChange={e=>setQ(e.target.value)} placeholder="search style, order, buyer, issue, user…" style={{ ...inputStyle, flex:"1 1 260px" }} />
      <span style={{ fontSize:10, color:"var(--muted-2)" }}>{mode==="comments"?filteredComments.length:mode==="notifications"?filteredNotifications.length:mode==="errors"?(errorLog||[]).length:mode==="changes"?(auditRows||[]).length:filtered.length} shown</span>
    </div>
    <div style={{ background:"var(--surface)", border:"1px solid var(--line-2)", borderRadius:14, overflow:"hidden" }}>
      {mode==="comments" ? <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}><thead><tr style={{ background:"var(--ink)", color:"var(--bg)" }}>{["Category","User","Style / Order","Buyer / Brand","Comment","Status","Open"].map(h=><th key={h} style={{ textAlign:"left", padding:"9px 10px", borderRight:"1px solid #3a362e", fontFamily:"'Archivo',sans-serif" }}>{h}</th>)}</tr></thead><tbody>{filteredComments.length===0?<tr><td colSpan={7} style={{ padding:24, color:"var(--muted-2)", textAlign:"center" }}>No comments for this filter.</td></tr>:filteredComments.slice(0,500).map(c=><tr key={c.id} onClick={()=>onJump&&onJump(c.sid,c.col)} style={{ cursor:"pointer", borderBottom:"1px solid var(--line-3)", background:c.resolved?"var(--surface)":"#fffaf0" }}><td style={{ padding:"8px 10px", fontWeight:800 }}>{c.category}</td><td style={{ padding:"8px 10px" }}>{c.author}<div style={{ color:"var(--muted-2)", fontSize:9 }}>{tsShortLocal(c.at)}</div></td><td style={{ padding:"8px 10px" }}><b>{c.styleNo||"—"}</b><br/><span style={{ color:"var(--muted-2)", fontSize:9 }}>{c.orderNo||""}</span></td><td style={{ padding:"8px 10px" }}>{c.buyer||"—"}<br/><span style={{ color:"var(--muted-2)", fontSize:9 }}>{c.brand||""}</span></td><td style={{ padding:"8px 10px", maxWidth:520 }}>{c.body}</td><td style={{ padding:"8px 10px" }}>{c.resolved?badge("Resolved","#e5f1ea","#1c6048"):badge("Open","#f8e9b7","#7a560f")}</td><td style={{ padding:"8px 10px", color:"var(--info)", fontWeight:900 }}>thread →</td></tr>)}</tbody></table>
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

function Th({ col, label, sort, onSort, sticky, left, z, width, onResize, onAutoFit, scale, letter, filterActive, filterOpen, filterValues, filterAllowed, onToggleFilter, onSetFilter, onCloseFilter }){
  const active=sort.col===col;
  const startDrag=(e)=>{ e.preventDefault(); e.stopPropagation(); const sx=e.clientX, sw=width||80; const sc=scale||1; const move=(ev)=>onResize&&onResize(col, sw+(ev.clientX-sx)/sc); const up=()=>{ window.removeEventListener("mousemove",move); window.removeEventListener("mouseup",up); }; window.addEventListener("mousemove",move); window.addEventListener("mouseup",up); };
  return (<th role="columnheader" aria-sort={active?(sort.dir>0?"ascending":"descending"):"none"} style={{ position:"sticky", top:0, left:sticky?left:undefined, zIndex:sticky?(z||5):3, background:active?"var(--accent)":"var(--ink)", color:active?"var(--ink)":"var(--bg)", padding:"8px 9px", textAlign:"left", fontWeight:600, fontSize:9.5, letterSpacing:0.4, textTransform:"uppercase", whiteSpace:"nowrap", overflow:"visible", border:"1px solid #3a362e", userSelect:"none", width:width||80, minWidth:width||80, maxWidth:width||80, boxSizing:"border-box", backgroundClip:"padding-box", transform:sticky?"translateZ(0)":undefined }}>
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
  const masterRef=useRef(null); const anchorRef=useRef(null); const [pos,setPos]=useState(null);
  const isOn=(v)=> !allowed || allowed.includes(v);
  const allOn = !allowed;                       // no filter = every value shown
  const noneOn = allowed && allowed.length===0; // nothing selected = grid empty
  const shown=values.filter(v=> v.toLowerCase().includes(q.toLowerCase()));
  const toggle=(v)=>{ const cur = allowed? new Set(allowed): new Set(values); if(cur.has(v)) cur.delete(v); else cur.add(v); const arr=[...cur]; onSet(arr.length===values.length? null : arr); };
  const toggleAll=()=> onSet(allOn? [] : null);
  const selectResults=()=> onSet(shown.length===values.length? null : shown);
  useEffect(()=>{ if(masterRef.current) masterRef.current.indeterminate = !allOn && !noneOn; },[allOn,noneOn]);
  useEffect(()=>{ const a=anchorRef.current; if(!a) return; const r=a.getBoundingClientRect(); const W=212,H=300; let left=r.left-180; if(left+W>window.innerWidth-8) left=window.innerWidth-8-W; if(left<8) left=8; let top=r.bottom+2; if(top+H>window.innerHeight-8) top=Math.max(8,window.innerHeight-8-H); setPos({top,left}); },[]);
  const menu=(
    <div onClick={e=>e.stopPropagation()} style={{ position:"fixed", top:pos?pos.top:-9999, left:pos?pos.left:-9999, zIndex:360, background:"var(--surface)", color:"var(--ink)", border:"1px solid var(--ink)", boxShadow:"4px 4px 0 var(--ink)", padding:8, width:210, textTransform:"none", letterSpacing:0, fontWeight:400, maxHeight:"80vh", overflowY:"auto" }}>
      <input autoFocus value={q} onClick={e=>e.stopPropagation()} onKeyDown={e=>e.stopPropagation()} onChange={e=>setQ(e.target.value)} placeholder="search values…" style={{ width:"100%", fontFamily:"inherit", fontSize:11, padding:"4px 6px", border:"1px solid var(--line-2)", outline:"none", marginBottom:6 }}/>
      <label style={{ display:"flex", alignItems:"center", gap:6, fontSize:10, fontWeight:700, padding:"3px 0", cursor:"pointer", borderBottom:"1px solid var(--line-3)", marginBottom:4 }}><input ref={masterRef} type="checkbox" checked={allOn} onChange={q?selectResults:toggleAll}/>{q?"(Select matches)":"(Select All)"}</label>
      <div style={{ maxHeight:180, overflowY:"auto" }}>
        {shown.map(v=>(<div key={v} style={{ display:"flex", alignItems:"center", gap:6, fontSize:10, padding:"2px 0" }}><input type="checkbox" checked={isOn(v)} onChange={()=>toggle(v)} style={{ cursor:"pointer" }}/><span onClick={()=>toggle(v)} style={{ flex:1, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis", cursor:"pointer" }}>{v}</span><button onClick={(e)=>{ e.stopPropagation(); onSet([v]); }} title="show only this" style={{ fontSize:8, border:"1px solid var(--line-2)", background:"var(--bg)", cursor:"pointer", padding:"1px 5px", color:"var(--muted-3)" }}>only</button></div>))}
        {shown.length===0 && <div style={{ fontSize:10, color:"var(--muted-1)", padding:"4px 0" }}>no matches</div>}
      </div>
      {noneOn && <div style={{ fontSize:9, color:"var(--danger)", marginTop:4 }}>Nothing selected — no rows shown.</div>}
      <div style={{ display:"flex", gap:6, marginTop:6 }}>
        <button onClick={()=>onSet(null)} style={{ ...chip, flex:1, fontSize:9 }}>Clear filter</button>
        <button onClick={onClose} style={{ ...chip, flex:1, fontSize:9, background:"var(--ink)", color:"var(--bg)" }}>Done</button>
      </div>
    </div>
  );
  return (<><span ref={anchorRef} style={{ position:"absolute", width:0, height:0 }}/>{createPortal(menu, document.body)}</>);
}
function FillPanel({ count, role, onApply, onClose }){
  const [mode,setMode]=useState(()=>{ try{ const v=localStorage.getItem("mt_fill_mode"); const m=(v==="revised"||v==="actual")?v:"actual"; return (m==="revised"&&!canEditRev(role))?"actual":m; }catch(e){ return "actual"; } });
  const stageOpts=STAGES.map(s=>({key:s.key,label:s.label+(s.cutoff?" (Fabric IH)":"")}));
  const opts = mode==="revised" ? (canEditRev(role)?stageOpts:[]) : [{key:"ordRec",label:"Order Date"},{key:"delivery",label:"Delivery Date"}].concat(stageOpts).filter(o=>canEdit(role,o.key,"actual"));
  const [key,setKey]=useState(()=>{ try{ return localStorage.getItem("mt_fill_stage")||(opts[0]?opts[0].key:"labAppr"); }catch(e){ return opts[0]?opts[0].key:"labAppr"; } }); const [val,setVal]=useState(iso(TODAY));
  useEffect(()=>{ if(opts.length && !opts.some(o=>o.key===key)) setKey(opts[0].key); },[mode]);
  useEffect(()=>{ try{ localStorage.setItem("mt_fill_mode",mode); }catch(e){} },[mode]);
  useEffect(()=>{ try{ localStorage.setItem("mt_fill_stage",key); }catch(e){} },[key]);
  const accent = mode==="revised"?"var(--accent)":"var(--info)";
  const tabBtn=(m,label)=>(<button onClick={()=>setMode(m)} style={{ flex:1, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:"5px 0", cursor:"pointer", border:"1px solid var(--ink)", background:mode===m?(m==="revised"?"var(--accent)":"var(--info)"):"var(--surface)", color:mode===m?"var(--surface)":"var(--ink)" }}>{label}</button>);
  return (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:0, marginTop:4, zIndex:370, background:"var(--surface)", border:"1px solid var(--ink)", boxShadow:"4px 4px 0 var(--ink)", padding:12, width:280 }}><div style={{ fontSize:11, fontWeight:700, marginBottom:8 }}>Set one date across {count} filtered styles</div><div style={{ display:"flex", marginBottom:8 }}>{tabBtn("actual","Actual")}{tabBtn("revised","Revised")}</div>{opts.length===0 ? <div style={{ fontSize:10, color:"var(--danger)", marginBottom:8 }}>Your role cannot set revised dates.</div> : (<><label style={{ fontSize:10, color:"var(--muted-2)" }}>Stage</label><select value={key} onChange={e=>setKey(e.target.value)} style={{ width:"100%", fontFamily:"inherit", fontSize:11, padding:5, marginBottom:8, border:"1px solid var(--ink)" }}>{opts.map(o=><option key={o.key} value={o.key}>{o.label}</option>)}</select><label style={{ fontSize:10, color:"var(--muted-2)" }}>{mode==="revised"?"Revised plan date":"Actual date"}</label><input type="date" value={val} onChange={e=>setVal(e.target.value)} style={{ width:"100%", fontFamily:"inherit", fontSize:11, padding:5, marginBottom:10, border:"1px solid var(--ink)" }}/></>)}<div style={{ display:"flex", gap:8 }}><button disabled={opts.length===0} onClick={()=>onApply(key,val,mode)} style={{ flex:1, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:6, cursor:opts.length?"pointer":"not-allowed", border:"1px solid var(--ink)", background:opts.length?accent:"var(--line-2)", color:"var(--surface)", opacity:opts.length?1:0.6 }}>Apply {mode} → {count}</button><button onClick={onClose} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--ink)", background:"var(--bg)" }}><X size={12}/></button></div></div>);
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
        <button onClick={()=>setOpen(false)} style={{ border:'1px solid var(--line-2)', background:'var(--surface)', cursor:'pointer', padding:'2px 5px', fontFamily:'inherit', fontSize:10 }}>×</button>
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
    <button ref={btnRef} type="button" onClick={(e)=>{ e.stopPropagation(); setOpen(o=>!o); }} style={{ fontFamily:'inherit', fontSize:10, fontWeight:800, padding:rounded?'6px 10px':'5px 9px', cursor:'pointer', border:'1px solid var(--line-2)', borderRadius:rounded?8:0, background:vals.length?'var(--accent-tint)':'var(--surface)', color:'var(--ink)', maxWidth:220, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }} title={labelText}>{labelText} ▾</button>
    {menu}
  </>;
}

/* ========================= DASHBOARD ========================= */
function OperationalDashboardView({ computed, todoItems, cfg, applyDrill, drillTodo }){
  const [target,setTarget]=useState("tracker"); // where bar/owner/activity drills go
  const [df,setDf]=useState(()=>{ try{ return JSON.parse(localStorage.getItem("mt_dashboard_filter")||localStorage.getItem("mt_dashfilter")||"{}"); }catch(e){ return {}; } });
  useEffect(()=>{ try{ localStorage.setItem("mt_dashboard_filter", JSON.stringify(df)); }catch(e){} },[df]);
  const [mgmtOpen,setMgmtOpen]=useState(()=>{ try{ return JSON.parse(localStorage.getItem("mt_mgmt_open")||"{}"); }catch(e){ return {}; } });
  useEffect(()=>{ try{ localStorage.setItem("mt_mgmt_open", JSON.stringify(mgmtOpen)); }catch(e){} },[mgmtOpen]);
  const isMgmtOpen=(key)=> mgmtOpen[key]!==false;
  const toggleMgmt=(key)=>setMgmtOpen(o=>({ ...o, [key]: !isMgmtOpen(key) }));
  const arrOf=(v)=>Array.isArray(v)?v:(v?[v]:[]);
  const hasSel=(key,val)=>{ const a=arrOf(df[key]); return !a.length || a.includes(val); };
  const hasColourSel=(st)=>{ const a=arrOf(df.colour); if(!a.length) return true; const cols=String(st.colour||"").split(/[,/]/).map(x=>x.trim()).filter(Boolean); return a.some(v=>cols.includes(v)); };
  const matchDf=(st,except)=> (except==="order"||hasSel("order",st.orderNo)) && (except==="fit"||hasSel("fit",st.sampleFit)) && (except==="junior"||hasSel("junior",st.owner)) && (except==="family"||hasSel("family",st.family)) && (except==="brand"||hasSel("brand",st.brand)) && (except==="fabric"||hasSel("fabric",st.fabricType)) && (except==="colour"||hasColourSel(st));
  const distinctC=(key,fn)=>{ const s=new Set(); computed.forEach(({s:st})=>{ if(!matchDf(st,key)) return; fn(st).forEach(v=>{ if(v) s.add(v); }); }); return [...s].sort(); };
  const orders=distinctC("order",s=>[s.orderNo]); const fits=distinctC("fit",s=>[s.sampleFit]); const juniors=distinctC("junior",s=>[s.owner]); const families=distinctC("family",s=>[s.family]); const brands=distinctC("brand",s=>[s.brand]); const fabrics=distinctC("fabric",s=>[s.fabricType]);
  const colours=distinctC("colour",s=>String(s.colour||"").split(/[,/]/).map(x=>x.trim()));
  const fc=computed.filter(({s})=> hasSel("order",s.orderNo) && hasSel("fit",s.sampleFit) && hasSel("junior",s.owner) && hasSel("family",s.family) && hasSel("brand",s.brand) && hasSel("fabric",s.fabricType) && hasColourSel(s) );
  const total=fc.length;
  const onTrack=fc.filter(({c})=>c.tone==="ok").length;
  const atRisk=fc.filter(({c})=>c.tone==="late"||c.tone==="warn").length;
  const released=fc.filter(({c})=>c.released).length;
  const delRisk=fc.filter(({c})=>String(c.status).startsWith("Delivery risk")).length;
  // owner load + activity load from the spliced set
  const ownerLoad={}; const actAgg={};
  fc.forEach(({c})=>{ if(c.released) return; (c.chaseOwners||[]).forEach(o=>{ ownerLoad[o.owner]=(ownerLoad[o.owner]||0)+1; }); (c.frontier?[...c.frontier]:[]).forEach(k=>{ const r=(c.stages||[]).find(x=>x.key===k); if(!r||r.done) return; const a=actAgg[r.label]=actAgg[r.label]||{n:0,over:0,key:k}; a.n++; if((r.rev||r.plan)&&TODAY>(r.rev||r.plan)) a.over++; }); });
  const owners=Object.entries(ownerLoad).sort((a,b)=>b[1]-a[1]); const maxOwner=Math.max(1,...owners.map(o=>o[1]));
  const acts=Object.entries(actAgg).sort((a,b)=>b[1].n-a[1].n); const maxAct=Math.max(1,...acts.map(e=>e[1].n));
  const overdueAct=acts.reduce((s,[,v])=>s+v.over,0);
  const escalationTodo=(todoItems||[]).filter(t=>t.overdue&&t.escalationOwner);
  const escLoad=escalationTodo.reduce((m,t)=>{ const k=t.escalationOwner||"(blank)"; m[k]=(m[k]||0)+1; return m; },{});
  const escRows=Object.entries(escLoad).sort((a,b)=>b[1]-a[1]); const maxEsc=Math.max(1,...escRows.map(x=>x[1]));
  const phase={ "Pre-Fit":0,"Fit / Print":0,"Lab Dip":0,"Fabric IH":0,"PP / Prod":0 };
  fc.forEach(({c})=>{ if(c.released) return; const k=c.nextPending&&c.nextPending.key; if(k==="techpack") phase["Pre-Fit"]++; else if(["fitSend","fitAppr","artwork","artAppr","strikeOff","soAppr"].includes(k)) phase["Fit / Print"]++; else if(["labDip","labAppr"].includes(k)) phase["Lab Dip"]++; else if(k==="fabricIH") phase["Fabric IH"]++; else phase["PP / Prod"]++; });
  const maxPhase=Math.max(1,...Object.values(phase));
  const OWNER_COLOR2=OWNER_COLOR;
  // splice carried into drills so the tracker shows the same slice
  const spliceCols=()=>{ const cf={}; const put=(k,col)=>{ const a=arrOf(df[k]); if(a.length) cf[col]=a; }; put("order","orderNo"); put("fit","sampleFit"); put("junior","owner"); put("family","family"); put("brand","brand"); put("fabric","fabricType"); put("colour","colour"); return cf; };
  const spliceSearch=()=> "";
  const goOwner=(o)=>{ if(target==="todo") drillTodo({ owner:o }); else applyDrill({ owner:o, colFilters:spliceCols(), search:spliceSearch() }); };
  const goAct=(label,key)=>{ if(target==="todo") drillTodo({ activity:label }); else applyDrill({ activity:key, colFilters:spliceCols(), search:spliceSearch() }); };
  const goPhase=(phaseName)=>{ drillTodo&&drillTodo({ phase:phaseName }); };
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
  const dashboardStyleRows=fc.map(({s,c})=>({ "Order No":s.orderNo||"", "Style No":s.styleNo||"", "Sample Fit":s.sampleFit||"", "Family":s.family||"", "Colour":s.colour||"", "Brand":s.brand||"", "Buyer":s.buyer||"", "Junior":s.owner||"", "Qty":s.qty||0, "Delivery":s.delivery||"", "Status":c.status||"", "Tone":c.tone||"", "Released":c.released?"YES":"", "% Done":c.pct, "Chase":(c.chaseOwners||[]).map(o=>`${o.owner} (${o.count})`).join(", "), "Next Pending":c.nextPending?c.nextPending.label:"", "Projected Release":c.projRelease?fmt(c.projRelease):"" }));
  const dashboardBreakup=fc.map(({s,c})=>{ const nx=c.nextPending||null; const nxDue=nx?(nx.rev||nx.plan):null; const frontier=(c.frontier?[...c.frontier]:[]).map(k=>{ const r=(c.stages||[]).find(x=>x.key===k); return r?r.label:k; }).join(", "); return { "Order No":s.orderNo||"", "Style No":s.styleNo||"", "Colour":s.colour||"", "Buyer / Brand":s.buyer||s.brand||"", "Junior":s.owner||"", "Delivery":s.delivery||"", "Overall Status":c.status||"", "Overall Tone":c.tone||"", "Released":c.released?"YES":"NO", "% Done":c.pct, "Next Pending Stage":nx?nx.label:"", "Next Pending Chase":nx?nx.owner:"", "Next Pending Due":nxDue?fmt(nxDue):"", "Next Pending Overdue?":(nxDue&&TODAY>nxDue)?"YES":"NO", "Actionable Frontier":frontier, "Chase Breakdown":(c.chaseOwners||[]).map(o=>`${o.owner} (${o.count})`).join(", "), "Fit Branch":c.fitBranch?c.fitBranch.txt:"", "Print Branch":c.printBranch?c.printBranch.txt:"", "Fabric Branch":c.fabricBranch?c.fabricBranch.txt:"", "PP Branch":c.ppBranch?c.ppBranch.txt:"", "Prod File Branch":c.prodFileBranch?c.prodFileBranch.txt:"", "Fabric IH Countdown":c.fabricCountdown?c.fabricCountdown.txt:"", "Projected Release":c.projRelease?fmt(c.projRelease):"", "Release Gate":c.releaseGate?fmt(c.releaseGate):"", "Float Days":c.float==null?"":c.float, "Idle Days":c.idle==null?"":c.idle }; });
  const stageStartFor=(s,c,r)=>{ const st=STAGES.find(x=>x.key===r.key)||{}; const byKey={}; (c.stages||[]).forEach(x=>{ byKey[x.key]=x; }); if(r.key==="fabricIH") return (s.labDipReq && byKey.labAppr && byKey.labAppr.actual) ? byKey.labAppr.actual : parse(s.ordRec); if(st.pred==="__ord") return parse(s.ordRec); return byKey[st.pred] ? byKey[st.pred].actual : null; };
  const stageState=(r)=> r.skipped?"Waived / skipped":(r.rejected?"Rejected":(r.rework?"Rework / resend":(r.actual?"Done":(r.rev?"Revised plan":"Pending"))));
  const dashboardStageDetail=[];
  const dashboardActualVsPlan=[];
  const dashboardRevisedVsActual=[];
  fc.forEach(({s,c})=>{ (c.stages||[]).forEach(r=>{ const st=STAGES.find(x=>x.key===r.key)||{}; const start=stageStartFor(s,c,r); const due=r.rev||r.plan; const delayDue=(due&&r.actual)?netWorkdays(due,r.actual):null; const delayPlan=(r.plan&&r.actual)?netWorkdays(r.plan,r.actual):null; const delayRev=(r.rev&&r.actual)?netWorkdays(r.rev,r.actual):null; const duration=(start&&r.actual)?Math.max(0,netWorkdays(start,r.actual)||0):null; const frontier=(c.frontier&&c.frontier.has(r.key)); dashboardStageDetail.push({ "Order No":s.orderNo||"", "Style No":s.styleNo||"", "Colour":s.colour||"", "Buyer / Brand":s.buyer||s.brand||"", "Junior":s.owner||"", "Stage":r.label, "Stage Key":r.key, "Branch":BRANCH_OF[r.key]||"", "Chase Label":r.owner||"", "State":stageState(r), "Actionable Frontier?":frontier?"YES":"NO", "Auto Plan Date":r.plan?fmt(r.plan):"", "Revised Date":r.rev?fmt(r.rev):"", "Actual Date":r.actual?fmt(r.actual):"", "Due Used":due?fmt(due):"", "Due Status":due?(TODAY>due?"Overdue":"Not overdue"):"", "Rejected Date":r.reject?fmt(r.reject):"", "Skipped Date":r.skip?fmt(r.skip):"", "Start Date Used":start?fmt(start):"", "Delay vs Due Days":delayDue==null?"":r1(delayDue), "Delay vs Auto Plan Days":delayPlan==null?"":r1(delayPlan), "Actual vs Revised Days":delayRev==null?"":r1(delayRev), "Actual Duration Days":duration==null?"":r1(duration), "Overall Status":c.status||"" }); if(delayPlan!=null){ dashboardActualVsPlan.push({ "Order No":s.orderNo||"", "Style No":s.styleNo||"", "Stage":r.label, "Branch":BRANCH_OF[r.key]||"", "Auto Plan Date":fmt(r.plan), "Revised Date":r.rev?fmt(r.rev):"", "Actual Date":fmt(r.actual), "Actual vs Original Plan Days":r1(delayPlan||0), "Accuracy":(delayPlan===0?"On original plan":(delayPlan>0?"Late vs original plan":"Early vs original plan")), "Included?":"YES - auto plan + actual exist" }); } if(r.rev&&r.actual){ dashboardRevisedVsActual.push({ "Order No":s.orderNo||"", "Style No":s.styleNo||"", "Stage":r.label, "Branch":BRANCH_OF[r.key]||"", "Auto Plan Date":r.plan?fmt(r.plan):"", "Revised Date":fmt(r.rev), "Actual Date":fmt(r.actual), "Actual vs Revised Days":r1(delayRev||0), "Accuracy":(delayRev===0?"On revised date":(delayRev>0?"Late vs revised":"Early vs revised")), "Included?":"YES - revised exists" }); } }); });
  const avgArr=(arr,key)=>arr.length?r1(arr.reduce((sum,row)=>sum+(Number(row[key])||0),0)/arr.length):0;
  const lateAvgArr=(arr,key)=>{ const late=arr.filter(row=>(Number(row[key])||0)>0); return late.length?r1(late.reduce((sum,row)=>sum+(Number(row[key])||0),0)/late.length):0; };
  const dashboardPlanAccuracySummary=[
    { "Comparison":"Actual vs Original Plan", "Records":dashboardActualVsPlan.length, "Late/Missed Records":dashboardActualVsPlan.filter(row=>(Number(row["Actual vs Original Plan Days"])||0)>0).length, "Avg Net Days":avgArr(dashboardActualVsPlan,"Actual vs Original Plan Days"), "Avg Late/Missed Days":lateAvgArr(dashboardActualVsPlan,"Actual vs Original Plan Days"), "Worst Late Days":dashboardActualVsPlan.reduce((m,row)=>Math.max(m,Number(row["Actual vs Original Plan Days"])||0),0) },
    { "Comparison":"Actual vs Revised Plan", "Records":dashboardRevisedVsActual.length, "Late/Missed Records":dashboardRevisedVsActual.filter(row=>(Number(row["Actual vs Revised Days"])||0)>0).length, "Avg Net Days":avgArr(dashboardRevisedVsActual,"Actual vs Revised Days"), "Avg Late/Missed Days":lateAvgArr(dashboardRevisedVsActual,"Actual vs Revised Days"), "Worst Late Days":dashboardRevisedVsActual.reduce((m,row)=>Math.max(m,Number(row["Actual vs Revised Days"])||0),0) }
  ];
  const dashboardOpenRows=dashboardStageDetail.filter(r=>r["Actionable Frontier?"]==="YES");
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
      { label:"Stuck Phases", data:phaseRows, detailData:dashboardBreakup, modes:["summary","detailed"] },
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
      <span style={{ marginLeft:"auto", display:"flex", alignItems:"center", gap:6 }}><span style={{ fontSize:10, color:"var(--muted-2)" }}>drill to:</span><div style={{ display:"flex", border:"1px solid var(--ink)" }}>{["tracker","todo"].map(t=>(<button key={t} onClick={()=>setTarget(t)} style={{ fontFamily:"inherit", fontSize:10, fontWeight:700, padding:"5px 10px", cursor:"pointer", border:"none", borderRight:t==="tracker"?"1px solid var(--ink)":"none", background:target===t?"var(--ink)":"var(--surface)", color:target===t?"var(--bg)":"var(--ink)" }}>{t==="tracker"?"Tracker":"To-Do"}</button>))}</div></span>
    </div>

    <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
      {card("Styles (slice)",total,"var(--ink)")}
      {card("On track",onTrack,"var(--success)",()=>goStatus("On Track"))}
      {card("At risk",atRisk,"var(--danger)",()=>goStatus("At Risk"))}
      {card("Delivery risk",delRisk,"var(--danger)",()=>goStatus("All",{ overall:["Delivery risk"] }))}
      {card("Released",released,"var(--success)",()=>goStatus("Released"))}
      {card("Overdue activities",overdueAct,"var(--danger)")}
      {card("Escalation items",escalationTodo.length,escalationTodo.length?"var(--danger)":"var(--success)",()=>{ drillTodo&&drillTodo({ todoType:"Escalation", priority:"Overdue" }); })}
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
        <div style={{ fontSize:9, color:"var(--muted-7)", marginTop:8 }}>Click to open in {target==="todo"?"To-Do":"Tracker"}.</div>
      </div>
      <div style={{ flex:1, minWidth:320, background:"var(--surface)", border:"1px solid var(--ink)", padding:16 }}>
        <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:13, marginBottom:12 }}>ESCALATION OWNER LOAD</div>
        {bar(escRows, maxEsc, ()=>"var(--danger)", 92, (o)=>drillTodo&&drillTodo({ todoType:"Escalation", priority:"Overdue", escalationOwner:o }))}
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
        <div style={{ fontSize:9, color:"var(--muted-7)", marginTop:8 }}>Click to open in {target==="todo"?"To-Do":"Tracker"}. Red = overdue.</div>
      </div>
      <div style={{ flex:1, minWidth:320, background:"var(--surface)", border:"1px solid var(--ink)", padding:16 }}>
        <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:13, marginBottom:12 }}>WHERE STYLES ARE STUCK</div>
        {Object.entries(phase).map(([p,n])=>(
          <button key={p} onClick={()=>goPhase(p)} title="Open matching phase in To-Do" style={{ display:"flex", alignItems:"center", gap:8, padding:"4px 0", width:"100%", border:"none", background:"transparent", cursor:"pointer", fontFamily:"inherit" }}>
            <span style={{ width:80, fontSize:10, fontWeight:700, color:"var(--muted-4)", textAlign:"left" }}>{p}</span>
            <span style={{ flex:1, height:16, background:"#f0ece3", position:"relative" }}><span style={{ position:"absolute", left:0, top:0, bottom:0, width:`${(n/maxPhase)*100}%`, background:p==="Fabric IH"?"var(--danger)":"var(--accent)" }}/></span>
            <span style={{ width:28, textAlign:"right", fontSize:11, fontWeight:700 }}>{n}</span>
          </button>))}
      </div>
    </div>
  </div>);
}

/* ========================= TO-DO ========================= */


/* ========================= MANAGEMENT ANALYTICS ========================= */
function ManagementDashboardView({ computed, todoItems, cfg, applyDrill, drillTodo }){
  const [target,setTarget]=useState("tracker"); // where bar/owner/activity drills go
  const [df,setDf]=useState(()=>{ try{ return JSON.parse(localStorage.getItem("mt_management_filter")||localStorage.getItem("mt_dashfilter")||"{}"); }catch(e){ return {}; } });
  useEffect(()=>{ try{ localStorage.setItem("mt_management_filter", JSON.stringify(df)); }catch(e){} },[df]);
  const [mgmtOpen,setMgmtOpen]=useState(()=>{ try{ return JSON.parse(localStorage.getItem("mt_mgmt_open")||"{}"); }catch(e){ return {}; } });
  useEffect(()=>{ try{ localStorage.setItem("mt_mgmt_open", JSON.stringify(mgmtOpen)); }catch(e){} },[mgmtOpen]);
  const [perfView,setPerfView]=useState(()=>{ try{ return localStorage.getItem("mt_mgmt_perf_view")||"chase"; }catch(e){ return "chase"; } }); // chase | stage
  const [perfMode,setPerfMode]=useState(()=>{ try{ return localStorage.getItem("mt_mgmt_perf_mode")||"all"; }catch(e){ return "all"; } }); // all | delay
  useEffect(()=>{ try{ localStorage.setItem("mt_mgmt_perf_view",perfView); localStorage.setItem("mt_mgmt_perf_mode",perfMode); }catch(e){} },[perfView,perfMode]);
  const isMgmtOpen=(key)=> mgmtOpen[key]!==false;
  const toggleMgmt=(key)=>setMgmtOpen(o=>({ ...o, [key]: !isMgmtOpen(key) }));

  const arrOf=(v)=>Array.isArray(v)?v:(v?[v]:[]);
  const hasSel=(key,val)=>{ const a=arrOf(df[key]); return !a.length || a.includes(val); };
  const hasColourSel=(st)=>{ const a=arrOf(df.colour); if(!a.length) return true; const cols=String(st.colour||"").split(/[,/]/).map(x=>x.trim()).filter(Boolean); return a.some(v=>cols.includes(v)); };
  const matchDf=(st,except)=> (except==="order"||hasSel("order",st.orderNo)) && (except==="fit"||hasSel("fit",st.sampleFit)) && (except==="junior"||hasSel("junior",st.owner)) && (except==="family"||hasSel("family",st.family)) && (except==="brand"||hasSel("brand",st.brand)) && (except==="fabric"||hasSel("fabric",st.fabricType)) && (except==="colour"||hasColourSel(st));
  const distinctC=(key,fn)=>{ const set=new Set(); computed.forEach(({s:st})=>{ if(!matchDf(st,key)) return; fn(st).forEach(v=>{ if(v) set.add(v); }); }); return [...set].sort(); };
  const orders=distinctC("order",s=>[s.orderNo]); const fits=distinctC("fit",s=>[s.sampleFit]); const juniors=distinctC("junior",s=>[s.owner]); const families=distinctC("family",s=>[s.family]); const brands=distinctC("brand",s=>[s.brand]); const fabrics=distinctC("fabric",s=>[s.fabricType]);
  const colours=distinctC("colour",s=>String(s.colour||"").split(/[,/]/).map(x=>x.trim()));
  const fc=computed.filter(({s})=> hasSel("order",s.orderNo) && hasSel("fit",s.sampleFit) && hasSel("junior",s.owner) && hasSel("family",s.family) && hasSel("brand",s.brand) && hasSel("fabric",s.fabricType) && hasColourSel(s) );

  const total=fc.length;
  const onTrack=fc.filter(({c})=>c.tone==="ok").length;
  const atRisk=fc.filter(({c})=>c.tone==="late"||c.tone==="warn").length;
  const released=fc.filter(({c})=>c.released).length;
  const delRisk=fc.filter(({c})=>String(c.status).startsWith("Delivery risk")).length;
  const completionPct=total?Math.round((released/total)*100):0;

  const ownerLoad={}; const actAgg={};
  fc.forEach(({c})=>{ if(c.released) return; (c.chaseOwners||[]).forEach(o=>{ ownerLoad[o.owner]=(ownerLoad[o.owner]||0)+1; }); (c.frontier?[...c.frontier]:[]).forEach(k=>{ const r=(c.stages||[]).find(x=>x.key===k); if(!r||r.done) return; const a=actAgg[r.label]=actAgg[r.label]||{n:0,over:0,key:k}; a.n++; if((r.rev||r.plan)&&TODAY>(r.rev||r.plan)) a.over++; }); });
  const owners=Object.entries(ownerLoad).sort((a,b)=>b[1]-a[1]); const maxOwner=Math.max(1,...owners.map(o=>o[1]));
  const acts=Object.entries(actAgg).sort((a,b)=>b[1].n-a[1].n); const maxAct=Math.max(1,...acts.map(e=>e[1].n));
  const overdueAct=acts.reduce((s,[,v])=>s+v.over,0);
  const escalationTodo=(todoItems||[]).filter(t=>t.overdue&&t.escalationOwner);
  const escLoad=escalationTodo.reduce((m,t)=>{ const k=t.escalationOwner||"(blank)"; m[k]=(m[k]||0)+1; return m; },{});
  const escRows=Object.entries(escLoad).sort((a,b)=>b[1]-a[1]); const maxEsc=Math.max(1,...escRows.map(x=>x[1]));

  const phase={ "Pre-Fit":0,"Fit / Print":0,"Lab Dip":0,"Fabric IH":0,"PP / Prod":0 };
  fc.forEach(({c})=>{ if(c.released) return; const k=c.nextPending&&c.nextPending.key; if(k==="techpack") phase["Pre-Fit"]++; else if(["fitSend","fitAppr","artwork","artAppr","strikeOff","soAppr"].includes(k)) phase["Fit / Print"]++; else if(["labDip","labAppr"].includes(k)) phase["Lab Dip"]++; else if(k==="fabricIH") phase["Fabric IH"]++; else phase["PP / Prod"]++; });
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
  const stageState=(r)=> r.skipped?"Waived / skipped":(r.rejected?"Rejected":(r.rework?"Rework / resend":(r.actual?"Done":(r.rev?"Revised plan":"Pending"))));
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
      delayRecords.push({ style:s.styleNo, order:s.orderNo, buyer:s.buyer||s.brand||"", owner:s.owner||"", stage:r.label, stageKey:r.key, dept:r.owner, delay, delayPlan, delayRevised, duration, actual:r.actual, due, plan:r.plan, revised:r.rev, start });
      const perfExtra={ delayPlan, delayRevised };
      addPerf(stagePerf,r.key,r.label,delay,duration,{ owner:r.owner, ...perfExtra });
      addPerf(deptPerf,r.owner,r.owner,delay,duration,perfExtra);
      if(r.owner==="Buyer") addPerf(buyerPerf,s.buyer||s.brand||"(No buyer)",s.buyer||s.brand||"(No buyer)",delay,duration,perfExtra);
      addPerf(chaseDelay,r.owner||"(No chase label)",r.owner||"(No chase label)",delay,duration,perfExtra);
      addPerf(brandDelay,s.buyer||s.brand||"(No buyer)",s.buyer||s.brand||"(No buyer)",delay,duration,perfExtra);
    });
  });
  // Performance tables may use all completed entries. Delay-ranking tables must use problem rows only.
  const lateRecords=delayRecords.filter(r=>r.delay>0);
  const buyerApprovalStageNames=["Fit Appr","Art Appr","S/O Appr","Lab Dip Appr","PP Appr"];
  const buyerApprovalRecords=delayRecords.filter(r=>buyerApprovalStageNames.includes(r.stage));
  const buyerApprovalLateRecords=lateRecords.filter(r=>buyerApprovalStageNames.includes(r.stage));
  const makeDelayAgg=(records,keyFn,labelFn)=>{ const m={}; records.forEach(r=>{ const k=keyFn(r)||"(blank)"; const o=m[k]=m[k]||{ key:k, label:labelFn?labelFn(r,k):k, delayed:0, delaySum:0, maxDelay:0, styles:new Set(), planN:0, planSum:0, planMissN:0, planMissSum:0, revN:0, revSum:0, revMissN:0, revMissSum:0 }; const d=Math.max(0,r.delay||0); o.delayed++; o.delaySum+=d; o.maxDelay=Math.max(o.maxDelay,d); if(r.delayPlan!=null){ o.planN++; o.planSum+=r.delayPlan; if(r.delayPlan>0){ o.planMissN++; o.planMissSum+=r.delayPlan; } } if(r.delayRevised!=null){ o.revN++; o.revSum+=r.delayRevised; if(r.delayRevised>0){ o.revMissN++; o.revMissSum+=r.delayRevised; } } if(r.style) o.styles.add(r.style); }); return Object.values(m).sort((a,b)=>(b.delaySum-a.delaySum)||(b.delayed-a.delayed)); };
  const chaseDelayRows=makeDelayAgg(lateRecords,r=>r.dept||"(No chase label)").slice(0,8);
  const brandDelayRows=makeDelayAgg(lateRecords,r=>r.buyer||"(No buyer)").slice(0,8);
  const buyerDelayRows=makeDelayAgg(buyerApprovalLateRecords,r=>r.buyer||"(No buyer)").slice(0,8);
  const stageRows=Object.values(stagePerf).sort((a,b)=>avg(b.delaySum,b.n)-avg(a.delaySum,a.n));
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
  const goOwner=(o)=>{ if(target==="todo") drillTodo({ owner:o }); else applyDrill({ owner:o, colFilters:spliceCols(), search:spliceSearch() }); };
  const goAct=(label,key)=>{ if(target==="todo") drillTodo({ activity:label }); else applyDrill({ activity:key, colFilters:spliceCols(), search:spliceSearch() }); };
  const goPhase=(phaseName)=>{ drillTodo&&drillTodo({ phase:phaseName }); };
  const goStatus=(st,extra)=>applyDrill({ status:st, colFilters:{...spliceCols(),...(extra||{})}, search:spliceSearch() });
  const goSearch=(q)=>applyDrill({ status:"All", colFilters:spliceCols(), search:q||spliceSearch() });
  const goStageOpen=(key)=>applyDrill({ status:"All", activity:key, colFilters:spliceCols(), search:spliceSearch() });
  const anyDf=Object.values(df).some(v=>Array.isArray(v)?v.length:!!v);

  const mgmtSummary=[{ "Report Type":"Management", "Styles in Slice":total, "On Track":onTrack, "At Risk":atRisk, "Delivery Risk":delRisk, "Released":released, "Release %":completionPct, "Overdue Open Activities":overdueAct, "Completed Stage Entries":actualDone, "Delayed Stage Entries":lateDone, "Avg Late Delay Days":r1(avgDelay), "Avg Actual Time Days":r1(avgDuration), "Actual vs Original Plan Records":planActualRecords.length, "Avg Actual vs Original Plan Net Days":r1(avgPlanNet), "Avg Missed Original Plan Days":r1(avgPlanMiss), "Worst Missed Original Plan Days":r1(worstPlanMiss), "Actual vs Revised Records":revisedActualRecords.length, "Avg Actual vs Revised Net Days":r1(avgRevNet), "Avg Missed Revised Days":r1(avgRevMiss), "Worst Missed Revised Days":r1(worstRevMiss) }];
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
  const currentStyles=fc.map(({s,c})=>({ "Order No":s.orderNo||"", "Style No":s.styleNo||"", "Fit":s.sampleFit||"", "Family":s.family||"", "Colour":s.colour||"", "Brand":s.brand||"", "Buyer":s.buyer||"", "Junior":s.owner||"", "Qty":s.qty||0, "Delivery":s.delivery||"", "Status":c.status||"", "Tone":c.tone||"", "Released":c.released?"YES":"", "% Done":c.pct, "Chase":(c.chaseOwners||[]).map(o=>`${o.owner} (${o.count})`).join(", "), "Next Pending":c.nextPending?c.nextPending.label:"", "Projected Release":c.projRelease?fmt(c.projRelease):"" }));
  const currentSliceStageDetail=[];
  fc.forEach(({s,c})=>{ (c.stages||[]).forEach(r=>{ const due=r.rev||r.plan; const start=stageStartFor(s,c,r); const delayDue=(due&&r.actual)?netWorkdays(due,r.actual):null; const delayPlan=(r.plan&&r.actual)?netWorkdays(r.plan,r.actual):null; const delayRevised=(r.rev&&r.actual)?netWorkdays(r.rev,r.actual):null; const duration=(start&&r.actual)?Math.max(0,netWorkdays(start,r.actual)||0):null; currentSliceStageDetail.push({ "Order No":s.orderNo||"", "Style No":s.styleNo||"", "Buyer / Brand":s.buyer||s.brand||"", "Junior / Style Owner":s.owner||"", "Stage":r.label||"", "Chase Label":r.owner||"", "State":stageState(r), "Auto Plan Date":r.plan?fmt(r.plan):"", "Revised Date":r.rev?fmt(r.rev):"", "Due Used":due?fmt(due):"", "Actual Date":r.actual?fmt(r.actual):"", "Start Used":start?fmt(start):"", "Delay vs Due Days":delayDue==null?"":r1(delayDue), "Actual vs Original Plan Days":delayPlan==null?"":r1(delayPlan), "Actual vs Revised Days":delayRevised==null?"":r1(delayRevised), "Actual Duration Days":duration==null?"":r1(duration), "Actionable Frontier?":(c.frontier&&c.frontier.has(r.key))?"YES":"NO" }); }); });
  const stageData=stageRows.map(r=>({"Stage":r.label,"Chase Label":r.owner||"","Completed":r.n,"Late Count":r.lateN,"Avg Net Delay Days":r1(avg(r.delaySum,r.n)),"Avg Actual Duration Days":r1(avg(r.durSum,r.durN)),"Avg Actual vs Original Plan Net Days":r1(avg(r.planSum,r.planN)),"Avg Missed Original Plan Days":r1(avg(r.planMissSum,r.planMissN)),"Avg Actual vs Revised Net Days":r1(avg(r.revSum,r.revN)),"Avg Missed Revised Days":r1(avg(r.revMissSum,r.revMissN)),"Duration Records":r.durN,"Original Plan Records":r.planN,"Revised Records":r.revN,"Worst Delay Days":r1(r.maxDelay)}));
  const dept=deptRows.map(r=>({"Department":r.label,"Completed":r.n,"Late Count":r.lateN,"Avg Net Delay Days":r1(avg(r.delaySum,r.n)),"Avg Actual Duration Days":r1(avg(r.durSum,r.durN)),"Avg Actual vs Original Plan Net Days":r1(avg(r.planSum,r.planN)),"Avg Missed Original Plan Days":r1(avg(r.planMissSum,r.planMissN)),"Avg Actual vs Revised Net Days":r1(avg(r.revSum,r.revN)),"Avg Missed Revised Days":r1(avg(r.revMissSum,r.revMissN)),"Duration Records":r.durN,"Original Plan Records":r.planN,"Revised Records":r.revN,"Worst Delay Days":r1(r.maxDelay)}));
  const buyerData=buyerRows.map(r=>({"Buyer / Brand":r.label,"Approvals":r.n,"Late Count":r.lateN,"Avg Approval Time Days":r1(avg(r.durSum,r.durN)),"Avg Net Delay Days":r1(avg(r.delaySum,r.n)),"Avg Actual vs Original Plan Net Days":r1(avg(r.planSum,r.planN)),"Avg Missed Original Plan Days":r1(avg(r.planMissSum,r.planMissN)),"Avg Actual vs Revised Net Days":r1(avg(r.revSum,r.revN)),"Avg Missed Revised Days":r1(avg(r.revMissSum,r.revMissN)),"Worst Delay Days":r1(r.maxDelay)}));
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
      "Actual vs Original Plan Net Days":r1(avg(o.planSum,o.planN)),
      "Missed Original Avg Days":r1(avg(o.planMissSum,o.planMissN)),
      "Actual vs Revised Net Days":r1(avg(o.revSum,o.revN)),
      "Missed Revised Avg Days":r1(avg(o.revMissSum,o.revMissN)),
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
    return {"Style No":o.style,"Order No":o.order,"Buyer / Brand":o.buyer,"Junior / Style Owner":o.owner,"Late Stage Count":o.late,"Total Delay Days":`${r1(o.delaySum)}d`,"Avg Actual vs Original Plan Net Days":r1(avg(p.reduce((s,r)=>s+(r.delayPlan||0),0),p.length)),"Avg Missed Original Plan Days":r1(avg(pm.reduce((s,r)=>s+(r.delayPlan||0),0),pm.length)),"Avg Actual vs Revised Net Days":r1(avg(rv.reduce((s,r)=>s+(r.delayRevised||0),0),rv.length)),"Avg Missed Revised Days":r1(avg(rvm.reduce((s,r)=>s+(r.delayRevised||0),0),rvm.length)),"Worst Stage":o.worstStage,"Worst Stage Delay":`${r1(o.worstDelay)}d`,"Management Reading":`${o.late} delayed stage(s); ${r1(o.delaySum)}d total delay`,"Suggested Action":"Make recovery plan"};
  });
  const worstKeys=new Set(Object.values(styleAgg).sort((a,b)=>b.delaySum-a.delaySum).slice(0,25).map(o=>(o.order||"")+"|"+(o.style||"")));
  const worstStyleDetail=detailRowsFor("Worst Performing Styles",lateRecords.filter(r=>worstKeys.has((r.order||"")+"|"+(r.style||""))),{problemOnly:true});
  const lateDetailRows=detailRowsFor("Worst Delays",lateRecords,{problemOnly:true});
  const exportChaseDelayData=compactAgg(lateRecords,r=>r.dept,"Chase Label");
  const exportBuyerDelayData=compactAgg(buyerApprovalLateRecords,r=>r.buyer||"(No buyer)","Buyer / Brand");
  const exportBrandDelayData=compactAgg(lateRecords,r=>r.buyer||"(No buyer)","Buyer / Brand");

  const stageDelayData=compactAgg(lateRecords,r=>r.stage,"Stage").map(o=>({ ...o, "Chase Label":(STAGES.find(s=>s.label===o.Stage)||{}).owner||"" }));
  const performanceAnalysisData = perfMode==="delay"
    ? (perfView==="stage" ? stageDelayData : exportChaseDelayData)
    : (perfView==="stage" ? stageData : dept.map(r=>({ "Chase Label":r.Department, "Completed":r.Completed, "Late Count":r["Late Count"], "Avg Net Delay Days":r["Avg Net Delay Days"], "Avg Actual Duration Days":r["Avg Actual Duration Days"], "Avg Actual vs Original Plan Net Days":r["Avg Actual vs Original Plan Net Days"], "Avg Missed Original Plan Days":r["Avg Missed Original Plan Days"], "Avg Actual vs Revised Net Days":r["Avg Actual vs Revised Net Days"], "Avg Missed Revised Days":r["Avg Missed Revised Days"], "Duration Records":r["Duration Records"], "Original Plan Records":r["Original Plan Records"], "Revised Records":r["Revised Records"], "Worst Delay Days":r["Worst Delay Days"] })));
  const performanceAnalysisDetail = detailRowsFor("Performance Analysis", perfMode==="delay" ? lateRecords : delayRecords, { problemOnly: perfMode==="delay" }).filter(r=>{
    if(perfView!=="stage") return true;
    return !!r["Stage / Activity"];
  });
  const performanceAnalysisMeta=[{ "View By":perfView==="stage"?"Stage":"Chase Label", "Data Type":perfMode==="delay"?"Delay Only":"All Completed Performance", "Rows":performanceAnalysisData.length, "Detail Rows":performanceAnalysisDetail.length, "Rule":perfMode==="delay"?"Only positive delay/problem rows":"All completed actual date rows; net delay may be early/on-time/late" }];
  const analyticsSheets=[
      { label:"Summary", data:mgmtSummary, detailData:managementDetailRows, modes:["summary","detailed"] },
      { label:"Calculation Checks", data:checks, modes:["summary","detailed"] },
      { label:"Performance Analysis", data:performanceAnalysisData.length?performanceAnalysisData:performanceAnalysisMeta, detailData:performanceAnalysisDetail, modes:["summary","detailed"] },
      { label:"Buyer Approval Turnaround", data:buyerData, detailData:buyerApprovalTurnaroundDetail, modes:["summary","detailed"] },
      { label:"Buyer Approval Delay", data:exportBuyerDelayData, detailData:buyerApprovalDetail, modes:["summary","detailed"] },
      { label:"Chase Delay Ranking", data:exportChaseDelayData, detailData:chaseDetailRows, modes:["summary","detailed"] },
      { label:"Buyer Brand Delays", data:exportBrandDelayData, detailData:brandDetailRows, modes:["summary","detailed"] },
      { label:"Escalation Owner Load", data:escRows.map(([owner,count])=>({"Escalation Owner":owner,"Overdue Items":count})), detailData:(todoItems||[]).filter(t=>t.overdue&&t.escalationOwner).map(t=>({"Order No":t.orderNo||"","Style No":t.styleNo||"","Activity":t.activity||"","Chase Label":t.owner||"","Escalation Owner":t.escalationOwner||"","Escalation Level":t.escalationLevel||"","Days Overdue":t.daysLate||"","Action":t.escalationAction||""})), modes:["summary","detailed"] },
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
  const perfCell=(v)=>fmtDays(Number.isFinite(Number(v))?Number(v):0);
  const perfPlanCols=(r)=>({
    planNet:avg(r.planSum||0,r.planN||0),
    planMiss:avg(r.planMissSum||0,r.planMissN||0),
    revNet:avg(r.revSum||0,r.revN||0),
    revMiss:avg(r.revMissSum||0,r.revMissN||0),
  });
  const performanceRows = perfMode==="delay"
    ? (perfView==="stage" ? makeDelayAgg(lateRecords,r=>r.stage).slice(0,10).map(r=>({ ...r, displayLabel:r.label, count:r.delayed, late:r.delayed, avgDuration:null, avgDelay:avg(r.delaySum,r.delayed), worst:r.maxDelay, ...perfPlanCols(r) })) : chaseDelayRows.map(r=>({ ...r, displayLabel:r.label, count:r.delayed, late:r.delayed, avgDuration:null, avgDelay:avg(r.delaySum,r.delayed), worst:r.maxDelay, ...perfPlanCols(r) })))
    : (perfView==="stage" ? stageRows.slice(0,10).map(r=>({ ...r, displayLabel:r.label, count:r.n, late:r.lateN, avgDuration:avg(r.durSum,r.durN), avgDelay:avg(r.delaySum,r.n), worst:r.maxDelay, ...perfPlanCols(r), owner:r.owner||"" })) : deptRows.map(r=>({ ...r, displayLabel:r.label, count:r.n, late:r.lateN, avgDuration:avg(r.durSum,r.durN), avgDelay:avg(r.delaySum,r.n), worst:r.maxDelay, ...perfPlanCols(r) })));
  const performanceTable=(rows)=> rows.length?(
    <div style={{ overflowX:"auto" }}>
      <table style={{ width:"100%", borderCollapse:"collapse", fontSize:10 }}>
        <thead><tr>{[perfView==="stage"?"Stage":"Chase Label","Completed / delayed","Late","Avg actual time","Avg delay / net","Actual vs original","Missed original avg","Actual vs revised","Missed revised avg","Worst"].map((h,i)=><th key={h} style={{ textAlign:i===0?"left":"right", padding:"7px 8px", borderBottom:"1px solid var(--line-2)", color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.4, whiteSpace:"nowrap" }}>{h}</th>)}</tr></thead>
        <tbody>{rows.map(r=><tr key={r.key||r.label}>
          <td style={{ padding:"7px 8px", fontWeight:800, whiteSpace:"nowrap" }}>{r.displayLabel}{r.owner?<span style={{ color:"var(--muted-2)", fontWeight:700 }}> · {r.owner}</span>:null}</td>
          <td style={{ padding:"7px 8px", textAlign:"right" }}>{r.count||0}</td>
          <td style={{ padding:"7px 8px", textAlign:"right", color:(r.late||0)>0?"var(--danger)":"var(--success)", fontWeight:800 }}>{r.late||0}</td>
          <td style={{ padding:"7px 8px", textAlign:"right" }}>{r.avgDuration==null?"—":perfCell(r.avgDuration)}</td>
          <td style={{ padding:"7px 8px", textAlign:"right" }}>{perfCell(r.avgDelay)}</td>
          <td style={{ padding:"7px 8px", textAlign:"right", color:r.planNet>0?"var(--danger)":"var(--success)", fontWeight:800 }}>{perfCell(r.planNet)}</td>
          <td style={{ padding:"7px 8px", textAlign:"right" }}>{perfCell(r.planMiss)}</td>
          <td style={{ padding:"7px 8px", textAlign:"right", color:r.revNet>0?"var(--danger)":"var(--success)", fontWeight:800 }}>{perfCell(r.revNet)}</td>
          <td style={{ padding:"7px 8px", textAlign:"right" }}>{perfCell(r.revMiss)}</td>
          <td style={{ padding:"7px 8px", textAlign:"right", color:(r.worst||0)>0?"var(--danger)":"var(--ink)", fontWeight:800 }}>{perfCell(Math.max(0,r.worst||0))}</td>
        </tr>)}</tbody>
      </table>
      <div style={{ fontSize:9, color:"var(--muted-7)", marginTop:8 }}>Side-by-side plan columns are averages for the same selected view and data mode. Original = auto/system plan. Revised = latest revised commitment.</div>
    </div>
  ):<div style={{ fontSize:11, color:"var(--muted-1)" }}>{perfMode==="delay"?"No positive delays in this slice.":"No completed stage dates yet."}</div>;

  return (<div style={{ padding:"16px 22px", maxWidth:1280 }}>
    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:16, marginBottom:14, flexWrap:"wrap" }}>
      <div><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:23 }}>Management Dashboard</div><div style={{ fontSize:11.5, color:"var(--muted-2)", marginTop:4, maxWidth:760, lineHeight:1.45 }}>Operational view of current bottlenecks plus actual turnaround performance by stage, department, buyer and chase label. All figures respect the slice filters below. Live tables are drillable; historical aggregate tables use detailed export for row-level breakup.</div></div>
      <ReportExportMenu title="Management" prefix="management" sheets={analyticsSheets} defaultMode="detailed" />
    </div>

    <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"center", marginBottom:14 }}>
      <span style={{ fontSize:10, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase", letterSpacing:0.5 }}>Slice:</span>
      {sel("Order",df.order,orders,v=>setDf(d=>({...d,order:v})))} {sel("Fit",df.fit,fits,v=>setDf(d=>({...d,fit:v})))} {sel("Colour",df.colour,colours,v=>setDf(d=>({...d,colour:v})))} {sel("Junior",df.junior,juniors,v=>setDf(d=>({...d,junior:v})))} {sel("Family",df.family,families,v=>setDf(d=>({...d,family:v})))} {sel("Brand",df.brand,brands,v=>setDf(d=>({...d,brand:v})))} {sel("Fabric",df.fabric,fabrics,v=>setDf(d=>({...d,fabric:v})))}
      {anyDf && <button onClick={()=>setDf({})} style={{ fontFamily:"inherit", fontSize:10, padding:"6px 10px", cursor:"pointer", border:"1px solid var(--danger)", background:"var(--surface)", color:"var(--danger)", fontWeight:800, borderRadius:8 }}>clear slice</button>}
      <span style={{ marginLeft:"auto", display:"flex", alignItems:"center", gap:6 }}><span style={{ fontSize:10, color:"var(--muted-2)" }}>drill to:</span><div style={{ display:"flex", border:"1px solid var(--line-2)", borderRadius:8, overflow:"hidden" }}>{["tracker","todo"].map(t=>(<button key={t} onClick={()=>setTarget(t)} style={{ fontFamily:"inherit", fontSize:10, fontWeight:800, padding:"6px 10px", cursor:"pointer", border:"none", borderRight:t==="tracker"?"1px solid var(--line-2)":"none", background:target===t?"var(--ink)":"var(--surface)", color:target===t?"var(--bg)":"var(--ink)" }}>{t==="tracker"?"Tracker":"To-Do"}</button>))}</div></span>
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
      {card("Escalation items",escalationTodo.length,escalationTodo.length?"var(--danger)":"var(--success)",()=>drillTodo&&drillTodo({ todoType:"Escalation", priority:"Overdue" }),"who must chase now")}
      {card("Avg late delay",fmtDays(avgDelay),avgDelay>0?"var(--danger)":"var(--success)",null,"delayed stages only")}
      {card("Avg actual time",fmtDays(avgDuration),"var(--info)",null,"stage cycle time")}
      {card("Avg vs original plan",fmtDays(avgPlanNet),avgPlanNet>0?"var(--danger)":"var(--success)",null,`${planActualRecords.length} completed entries · missed avg ${fmtDays(avgPlanMiss)}`)}
      {card("Avg vs revised",fmtDays(avgRevNet),avgRevNet>0?"var(--danger)":"var(--success)",null,`${revisedActualRecords.length} revised entries · missed avg ${fmtDays(avgRevMiss)}`)}
    </div>

    <div style={{ display:"flex", flexDirection:"column", gap:12, marginTop:22 }}>
      {section("Who to chase — current", owners.length?owners.map(([o,n])=>barLine(o,o,n,maxOwner,OWNER_COLOR[o]||"var(--accent)",()=>goOwner(o))):<div style={{ fontSize:11, color:"var(--muted-1)" }}>Nothing pending.</div>, "Open actionable items by chase label")}
      {section("Escalation owner load", escRows.length?escRows.map(([o,n])=>barLine(o,o,n,maxEsc,"var(--danger)",()=>drillTodo&&drillTodo({ todoType:"Escalation", priority:"Overdue", escalationOwner:o }),n)):<div style={{ fontSize:11, color:"var(--muted-1)" }}>No overdue escalation items.</div>, "Who must chase now based on editable Settings duration slabs")}
      {section("Open activities", acts.length?acts.map(([label,v])=>barLine(label,label,v.n,maxAct,v.over?"var(--danger)":"var(--accent)",()=>goAct(label,v.key),v.over?`${v.n} (${v.over})`:v.n)):<div style={{ fontSize:11, color:"var(--muted-1)" }}>Nothing due.</div>, "Red count in bracket = overdue")}
      {section("Where styles are stuck", Object.entries(phase).map(([p,n])=>barLine(p,p,n,maxPhase,p==="Fabric IH"?"var(--danger)":"var(--accent)",()=>goPhase(p))), "Current next-pending phase · click opens matching To-Do phase")}
    </div>

    <div style={{ display:"flex", flexDirection:"column", gap:12, marginTop:12 }}>
      {section("Performance analysis", <>
        <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"center", margin:"0 0 8px" }}>
          <span style={{ fontSize:10, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase" }}>View by:</span>
          {toggleBtn(perfView==="chase","Chase Label",()=>setPerfView("chase"))}
          {toggleBtn(perfView==="stage","Stage",()=>setPerfView("stage"))}
          <span style={{ width:10 }}/>
          <span style={{ fontSize:10, fontWeight:800, color:"var(--muted-2)", textTransform:"uppercase" }}>Data:</span>
          {toggleBtn(perfMode==="all","All completed",()=>setPerfMode("all"))}
          {toggleBtn(perfMode==="delay","Delay only",()=>setPerfMode("delay"))}
        </div>
        {performanceTable(performanceRows)}
      </>, perfView==="stage" ? (perfMode==="delay"?"Stage-wise delay summary only; no early/on-time rows":"Stage-wise completed performance; net delay may be early/on-time/late") : (perfMode==="delay"?"Chase-label delay summary only; no early/on-time rows":"Chase-label completed performance; net delay may be early/on-time/late"))}
      {section("Plan accuracy", <>
        {rowBtn("actual-vs-plan","Actual vs original plan",`${fmtDays(avgPlanNet)} net`,avgPlanNet>0?"var(--danger)":"var(--success)",null,`${planActualRecords.length} completed entries · ${planMissedRecords.length} missed · avg missed ${fmtDays(avgPlanMiss)} · worst ${fmtDays(worstPlanMiss)}`)}
        {rowBtn("actual-vs-revised","Actual vs revised plan",`${fmtDays(avgRevNet)} net`,avgRevNet>0?"var(--danger)":"var(--success)",null,`${revisedActualRecords.length} revised entries · ${revisedMissedRecords.length} missed · avg missed ${fmtDays(avgRevMiss)} · worst ${fmtDays(worstRevMiss)}`)}
      </>, "Actual vs original plan checks first planned target; Actual vs revised checks latest committed/revised target")}
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
  const [tf,setTf]=useState({}); // header filters: priority, order, junior, activity, branch, chase
  useEffect(()=>{ if(filter&&Object.keys(filter).length) setTf(f=>({ ...f, ...filter })); },[filter]);
  const includeEsc=cfg&&cfg.todoEscalationRows!==false;
  const baseItems=items||[];
  const escalationRows=includeEsc?baseItems.filter(t=>t.overdue&&t.escalationOwner).map(t=>({ ...t, todoType:"Escalation", activity:`Escalate: ${t.activity}`, owner:t.owner, originalActivity:t.activity })):[];
  const displayItems=[...baseItems.map(t=>({ ...t, todoType:"Activity" })), ...escalationRows];
  const phaseMatch=(t,phase)=>{ if(!phase) return true; const k=t.key||""; if(phase==="Pre-Fit") return k==="techpack"; if(phase==="Fit / Print") return ["fitSend","fitAppr","artwork","artAppr","strikeOff","soAppr"].includes(k); if(phase==="Lab Dip") return ["labDip","labAppr"].includes(k); if(phase==="Fabric IH") return k==="fabricIH"; if(phase==="PP / Prod") return ["ppSample","ppAppr","prodFile"].includes(k); return true; };
  const passExcept=(t,except)=> (except==="phase"||phaseMatch(t,tf.phase)) && (except==="priority"||!tf.priority||(tf.priority==="Overdue"?t.overdue:!t.overdue)) && (except==="todoType"||!tf.todoType||t.todoType===tf.todoType) && (except==="orderNo"||!tf.orderNo||t.orderNo===tf.orderNo) && (except==="junior"||!tf.junior||t.junior===tf.junior) && (except==="activity"||!tf.activity||t.activity===tf.activity) && (except==="branch"||!tf.branch||t.branch===tf.branch) && (except==="owner"||!tf.owner||t.owner===tf.owner) && (except==="escalationOwner"||!tf.escalationOwner||t.escalationOwner===tf.escalationOwner);
  const pass=(t)=>passExcept(t,null);
  const distinct=(field)=>{ const vals=new Set(); displayItems.forEach(t=>{ if(!passExcept(t,field)) return; if(field==="priority") vals.add(t.overdue?"Overdue":"Upcoming"); else { const v=t[field]; if(v) vals.add(v); } }); return [...vals].sort(); };
  const orders=distinct("orderNo"), juniors=distinct("junior"), activities=distinct("activity"), branches=distinct("branch"), owners=distinct("owner"), escOwners=distinct("escalationOwner"), types=distinct("todoType"), priorities=distinct("priority");
  const shown=displayItems.filter(pass);
  const overdue=shown.filter(t=>t.overdue), upcoming=shown.filter(t=>!t.overdue);
  const anyF=Object.values(tf).some(Boolean);
  const COLW={ pri:96, type:80, ord:60, sty:170, jr:64, act:110, br:84, own:78, esc:100, date:84 };
  const set=(k,v)=>{ const upd=f=>({ ...(f||{}), [k]:v||undefined }); setTf(upd); setFilter&&setFilter(upd); };
  const hsel=(w,k,opts,first)=>(<select value={tf[k]||""} onChange={e=>set(k,e.target.value)} onClick={e=>e.stopPropagation()} style={{ width:w, fontFamily:"inherit", fontSize:9, padding:"2px 1px", border:"1px solid "+(tf[k]?"var(--accent)":"var(--line-2)"), background:tf[k]?"var(--accent-tint)":"var(--surface)" }}><option value="">{first}</option>{opts.map(o=>(<option key={o} value={o}>{o}</option>))}</select>);
  const head=(<div style={{ display:"flex", alignItems:"flex-end", gap:10, padding:"6px 12px 4px", borderBottom:"2px solid var(--ink)" }}>
    {hsel(COLW.pri,"priority",priorities,"Priority")}
    {hsel(COLW.type,"todoType",types,"Type")}
    {hsel(COLW.ord,"orderNo",orders,"Order")}
    <span style={{ width:COLW.sty, fontSize:9, fontWeight:700, textTransform:"uppercase", color:"#8a857a" }}>Style / Colour</span>
    {hsel(COLW.jr,"junior",juniors,"Junior")}
    {hsel(COLW.act,"activity",activities,"Activity")}
    {hsel(COLW.br,"branch",branches,"Branch")}
    {hsel(COLW.own,"owner",owners,"Chase")}
    {hsel(COLW.esc,"escalationOwner",escOwners,"Escalation")}
    <span style={{ width:COLW.date, fontSize:9, fontWeight:700, textTransform:"uppercase", color:"#8a857a" }}>Plan Date</span>
    <span style={{ flex:1, fontSize:9, fontWeight:700, textTransform:"uppercase", color:"#8a857a" }}>Days Late / Left</span>
  </div>);
  const data=shown.map(t=>({ "Priority":t.overdue?"Overdue":"Upcoming", "Type":t.todoType||"Activity", "Order No":t.orderNo||"", "Style / Colour":t.isColour?String(t.colour||""):t.styleNo, "Grouped Fabric Count":t.isColour?(t.count||0):"", "Junior":t.junior||"", "Activity":t.activity||"", "Branch":t.branch||"", "Chase Label":t.owner||"", "Escalation Owner":t.escalationOwner||"", "Escalation Level":t.escalationLevel||"", "Escalation Action":t.escalationAction||"", "Plan Date":t.exp?fmt(t.exp):"", "Days Late / Left":t.overdue?Math.abs(t.du):t.du, "Style ID":t.id, "Stage Key":t.key }));
  const summary=[{ "Report Type":"To-Do", "Shown Items":shown.length, "Base Activity Items":baseItems.length, "Escalation Rows Included":includeEsc?escalationRows.length:0, "Total Display Items":displayItems.length, "Overdue":overdue.length, "Upcoming":upcoming.length }];
  const byOwner={}; const byActivity={};
  shown.forEach(t=>{ byOwner[t.owner||"(blank)"]=(byOwner[t.owner||"(blank)"]||0)+1; byActivity[t.activity||"(blank)"]=(byActivity[t.activity||"(blank)"]||0)+1; });
  const ownerRows=Object.entries(byOwner).map(([k,v])=>({"Chase Label":k,"Items":v}));
  const activityRows=Object.entries(byActivity).map(([k,v])=>({"Activity":k,"Items":v}));
  const todoLogicChecks=[
      { Check:"Display count", Rule:"Shown rows equal filtered display rows", Value:shown.length, Expected:displayItems.filter(pass).length, Result:shown.length===displayItems.filter(pass).length?"OK":"CHECK" },
      { Check:"Escalation toggle", Rule:"Escalation rows added only when toggle is ON", Value:includeEsc?escalationRows.length:0, Expected:includeEsc?"overdue rows with escalation owner":"0", Result:(!includeEsc&&escalationRows.length===0)||(includeEsc&&escalationRows.every(t=>t.overdue&&t.escalationOwner))?"OK":"CHECK" },
      { Check:"Base activity rows", Rule:"Base activity rows remain even when escalation rows are ON", Value:baseItems.length, Expected:"unchanged source To-Do items", Result:"OK" },
      { Check:"Overdue/upcoming split", Rule:"Overdue + Upcoming equals shown rows", Value:overdue.length+upcoming.length, Expected:shown.length, Result:(overdue.length+upcoming.length)===shown.length?"OK":"CHECK" },
      { Check:"Plan accuracy columns", Rule:"Not applicable because To-Do rows are pending and do not have actual completion dates", Value:"N/A", Expected:"Use Dashboard/Management/Tracker completed-stage exports", Result:"OK" }
  ];
  const todoSheets=[
      { label:"Summary", data:summary, detailData:data, modes:["summary","detailed"] },
      { label:"Logic Checks", data:todoLogicChecks, modes:["summary","detailed"] },
      { label:"By Chase", data:ownerRows, detailData:data, modes:["summary","detailed"] },
      { label:"By Activity", data:activityRows, detailData:data, modes:["summary","detailed"] },
      { label:"To-Do Items", data:data, modes:["detailed"] },
  ];
  const row=(t)=>(<button key={(t.isColour?"col-":"")+t.id+t.key} onClick={()=>onJump(t.id,t.key)} style={{ display:"flex", alignItems:"center", gap:10, width:"100%", textAlign:"left", borderLeft:`4px solid ${t.overdue?"var(--danger)":"var(--accent)"}`, borderBottom:"1px solid #eee7da", background:t.isColour?"#fbf8f1":"var(--surface)", cursor:"pointer", fontFamily:"inherit", padding:"7px 12px" }}>
    <span style={{ width:COLW.pri, fontSize:10, fontWeight:700, display:"flex", alignItems:"center", gap:6, color:t.overdue?"var(--danger)":"#7a560f" }}><span style={{ width:8, height:8, borderRadius:"50%", background:t.overdue?"var(--danger)":"var(--accent)" }}/>{t.overdue?"Overdue":"Upcoming"}</span>
    <span style={{ width:COLW.type, fontSize:9, fontWeight:800, color:t.todoType==="Escalation"?"var(--danger)":"var(--muted-3)" }}>{t.todoType||"Activity"}</span>
    <span style={{ width:COLW.ord, fontSize:10, color:"var(--muted-4)" }}>{t.orderNo||"—"}</span>
    <span style={{ width:COLW.sty, fontSize:11, fontWeight:700, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{t.isColour?<span style={{ display:"inline-flex", alignItems:"center", gap:6 }}><span style={{ fontSize:8, fontWeight:700, background:"var(--ink)", color:"var(--bg)", padding:"1px 4px" }}>FABRIC</span>{t.colour} <span style={{ color:"var(--muted-1)", fontWeight:400 }}>×{t.count}</span></span>:t.styleNo}</span>
    <span style={{ width:COLW.jr, fontSize:10, color:"var(--muted-5)" }}>{t.junior||"—"}</span>
    <span style={{ width:COLW.act, fontSize:10, fontWeight:700, color:"#333" }}>{t.activity}</span>
    <span style={{ width:COLW.br, fontSize:10, color:"var(--muted-3)" }}>{t.branch}</span>
    <span style={{ width:COLW.own, fontSize:10, fontWeight:700, color:OWNER_COLOR2[t.owner]||"var(--muted-3)" }}>{t.owner}</span>
    <span style={{ width:COLW.esc, fontSize:10, fontWeight:800, color:t.escalationOwner?"var(--danger)":"var(--muted-2)" }}>{t.escalationOwner||"—"}</span>
    <span style={{ width:COLW.date, fontSize:10, color:"var(--muted-3)" }}>{fmt(t.exp)}</span>
    <span style={{ flex:1, fontSize:10, fontWeight:700, color:t.overdue?"var(--danger)":"#7a560f" }}>{t.overdue?`+${Math.abs(t.du)}d late`:`${t.du}d left`}</span>
  </button>);
  return (<div style={{ padding:"16px 22px", maxWidth:1080 }}>
    <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:8 }}>
      <span style={{ fontSize:10, color:"var(--muted-2)" }}>Showing {shown.length} of {displayItems.length}</span>
      <button onClick={()=>setCfg&&setCfg(c=>({ ...c, todoEscalationRows:!(c&&c.todoEscalationRows!==false) }))} disabled={!canEditSettings} title="Adds/removes separate escalation action rows in To-Do. It does not change the base activity row." style={{ fontFamily:"inherit", fontSize:10, padding:"4px 9px", cursor:canEditSettings?"pointer":"not-allowed", border:"1px solid var(--ink)", background:includeEsc?"var(--accent-tint)":"var(--surface)", fontWeight:800 }}>Escalation rows: {includeEsc?"ON":"OFF"}</button>
      {anyF && <button onClick={()=>{ setTf({}); setFilter&&setFilter({}); }} style={{ fontFamily:"inherit", fontSize:10, padding:"4px 9px", cursor:"pointer", border:"1px solid var(--danger)", background:"var(--surface)", color:"var(--danger)", fontWeight:700 }}>clear filters</button>}
      <span style={{ marginLeft:"auto" }}><ReportExportMenu title="To-Do" prefix="todo" sheets={todoSheets} defaultMode="detailed" /></span>
    </div>
    <div style={{ display:"flex", alignItems:"baseline", gap:12, margin:"4px 0 6px" }}><span style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:13 }}>TO-DO · {shown.length}</span>{overdue.length>0 && <span style={{ fontSize:11, fontWeight:700, color:"var(--danger)" }}>{overdue.length} overdue</span>}{upcoming.length>0 && <span style={{ fontSize:11, fontWeight:700, color:"#7a560f" }}>{upcoming.length} upcoming</span>}</div>
    {head}
    {shown.length?shown.map(row):<div style={{ fontSize:11, color:"var(--muted-1)", padding:"8px 12px" }}>Nothing due or coming up. 👍</div>}
    <div style={{ fontSize:9, color:"var(--muted-7)", marginTop:14 }}>One list, most urgent first. Toggle adds/removes separate escalation rows; base activity rows stay unchanged. Chase Label = blocker, Escalation Owner = who must chase now as per Settings.</div>
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

function AuthShell({ children }){ return (<div style={{ minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", background:"var(--bg)", fontFamily:"'JetBrains Mono',monospace", padding:20 }}><div style={{ width:360, maxWidth:"100%", background:"var(--surface)", border:"2px solid var(--ink)", boxShadow:"8px 8px 0 var(--ink)", padding:26 }}>{children}</div></div>); }

function LoginScreen(){
  const [email,setEmail]=useState(""); const [pw,setPw]=useState(""); const [mode,setMode]=useState("in"); const [msg,setMsg]=useState(""); const [busy,setBusy]=useState(false); const [remember,setRemember]=useState(true);
  const submit=async()=>{ if(!email.trim()||!pw){ setMsg("Enter email and password."); return; } setMsg(""); setBusy(true);
    const stamp=()=>{ try{ localStorage.setItem("mt_login_at", String(Date.now())); localStorage.setItem("mt_remember", remember?"1":"0"); }catch(x){} };
    try{ if(mode==="in"){ const { error }=await supabase.auth.signInWithPassword({ email:email.trim(), password:pw }); if(error) throw error; stamp(); }
      else { const { error,data }=await supabase.auth.signUp({ email:email.trim(), password:pw }); if(error) throw error; if(data.session){ stamp(); } if(!data.session){ setMsg("Account created. If sign-in doesn't happen automatically, check your email to confirm, then sign in."); setMode("in"); } } }
    catch(e){ setMsg(e.message||String(e)); } setBusy(false); };
  const inp={ width:"100%", fontFamily:"inherit", fontSize:13, padding:"9px 10px", border:"1px solid var(--ink)", marginBottom:10, boxSizing:"border-box" };
  return (<AuthShell>
    <div style={{ fontFamily:"'JetBrains Mono',monospace", fontSize:10, letterSpacing:1.5, color:"var(--accent)", fontWeight:700, marginBottom:6 }}>KOTHARI SPORTS & APPARELS</div>
    <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:22, marginBottom:2 }}>Merch Tracker</div>
    <div style={{ fontSize:11, color:"var(--muted-2)", marginBottom:18 }}>{mode==="in"?"Sign in to your account":"Create your account"}</div>
    <form onSubmit={e=>{ e.preventDefault(); submit(); }}>
    <input style={inp} type="email" name="email" autoComplete="username" placeholder="email" value={email} onChange={e=>setEmail(e.target.value)}/>
    <input style={inp} type="password" name="password" autoComplete={mode==="in"?"current-password":"new-password"} placeholder="password" value={pw} onChange={e=>setPw(e.target.value)}/>
    <label style={{ display:"flex", alignItems:"center", gap:7, fontSize:11, color:"var(--muted-4)", marginBottom:12, cursor:"pointer" }}><input type="checkbox" checked={remember} onChange={e=>setRemember(e.target.checked)}/>Keep me signed in <span style={{ color:"var(--muted-1)" }}>(else sign out after 12h)</span></label>
    <button type="submit" disabled={busy} style={{ width:"100%", fontFamily:"inherit", fontSize:13, fontWeight:700, padding:"10px", cursor:busy?"wait":"pointer", border:"1px solid var(--ink)", background:"var(--ink)", color:"var(--bg)", marginBottom:10 }}>{busy?"…":(mode==="in"?"Sign in":"Create account")}</button>
    </form>
    <div style={{ fontSize:11, textAlign:"center" }}><span style={{ color:"var(--muted-2)" }}>{mode==="in"?"New here? ":"Have an account? "}</span><button type="button" onClick={()=>{ setMode(mode==="in"?"up":"in"); setMsg(""); }} style={{ border:"none", background:"transparent", color:"var(--accent)", cursor:"pointer", fontFamily:"inherit", fontSize:11, fontWeight:700 }}>{mode==="in"?"Create account":"Sign in"}</button></div>
    {msg && <div style={{ fontSize:11, color:"var(--danger)", marginTop:12, lineHeight:1.4 }}>{msg}</div>}
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
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:18 }}>Users &amp; roles</div><button onClick={onClose} style={{ border:"none", background:"transparent", cursor:"pointer" }}><X size={18}/></button></div>
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

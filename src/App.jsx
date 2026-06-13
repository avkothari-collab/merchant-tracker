import React, { useState, useMemo, useRef, useEffect } from "react";
import { Check, Plus, Lock, Filter, X, Copy, ChevronUp, ChevronDown, CornerDownRight, Columns3, MessageSquare, RotateCcw, Droplet, Snowflake, Trash2, SkipForward } from "lucide-react";
import * as XLSX from "xlsx";
import { createPortal } from "react-dom";
import { supabase } from "./supabaseClient";

/* MERCH TRACKER — Excel-like entry grid PROTOTYPE (v13 Excel as smart base) */

const FONT = `@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@500;600;800&family=JetBrains+Mono:wght@400;500;700&display=swap');`;
const REL_GATE_DAYS = 30, FABRIC_CUTOFF_DAYS = 35, STYLE_W = 190;
const UPCOMING_DEFAULT = { fitSend:4, artwork:2, strikeOff:3, ppSample:4, fabricIH:15 }; // working days before a stage that it becomes "upcoming" in the To-Do list

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
const DEFAULT_CFG = { leads:Object.fromEntries(STAGES.map(s=>[s.key,s.lead])), rework:{...{ fitSend:4, artwork:2, strikeOff:3, labDip:7, ppSample:4 }}, fabricCutoff:FABRIC_CUTOFF_DAYS, relGate:REL_GATE_DAYS, upcoming:{...UPCOMING_DEFAULT} };
const OWNER_COLOR = { Merchant:"#1f6f54", CAD:"#2563a6", Buyer:"#b4531a", Designer:"#6d4aab", Mill:"#7a5a1e", Tamal:"#1f6f54", Rina:"#6d4aab" };

const ONE_DAY = 86400000;
function addWorkdays(date,n){ if(!date) return null; const d=new Date(date.getTime()); let a=0; const step=n>=0?1:-1, t=Math.abs(n); while(a<t){ d.setTime(d.getTime()+step*ONE_DAY); if(d.getDay()!==0) a++; } return d; }
function netWorkdays(a,b){ if(!a||!b) return null; let d=new Date(a.getTime()), end=new Date(b.getTime()), sign=1; if(end<d){ const tmp=d; d=end; end=tmp; sign=-1; } let n=0; const cur=new Date(d.getTime()); while(cur<end){ cur.setTime(cur.getTime()+ONE_DAY); if(cur.getDay()!==0) n++; } return sign*n; }
const fmt=(d)=> !d?"":d.toLocaleDateString("en-GB",{day:"2-digit",month:"short"});
const parse=(s)=>{ if(!s) return null; const m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(s); return m?new Date(Number(m[1]),Number(m[2])-1,Number(m[3])):new Date(s); };
const iso=(d)=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
const colLetter=(n)=>{ let s=""; n=Number(n)+1; if(n<1) return ""; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s; };
const letterToIndex=(s)=>{ s=String(s||"").toUpperCase(); if(!/^[A-Z]+$/.test(s)) return -1; let n=0; for(const ch of s) n=n*26+(ch.charCodeAt(0)-64); return n-1; };
const _now=new Date(); const TODAY=new Date(_now.getFullYear(),_now.getMonth(),_now.getDate()); // live current date (local midnight)

const SEED=[
  { id:1, orderNo:"T1", sampleFit:"fit 20", family:"DRESS/COLLAR SET", styleNo:"AW26IG7512A", colour:"BERRY RED, CLOUD DANCER", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"6-24M", qty:300, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:false, soReq:false, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-29", artwork:"2026-05-19", artAppr:"2026-05-25", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{fitSend:"2026-05-27", artwork:"2026-05-25", fabricIH:"2026-06-01"} },
  { id:2, orderNo:"T1", sampleFit:"fit 48", family:"SWEAT", styleNo:"AW26IG7508A", colour:"CLOUD DANCER, LAVENDER", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"6-24M", qty:300, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:false, soReq:false, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-22", fitSend:"2026-05-30", artwork:"2026-05-20", artAppr:"2026-05-25", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{artwork:"2026-05-27"} },
  { id:3, orderNo:"T1", sampleFit:"fit 6", family:"SLEEVELESS", styleNo:"HSAW26CGSS7158", colour:"BLACK, CLOUD DANCER", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"4-10YRS", qty:420, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:false, soReq:false, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-30", artwork:"2026-05-30", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{fitSend:"2026-05-27", artwork:"2026-05-25", fabricIH:"2026-06-05"} },
  { id:4, orderNo:"T1", sampleFit:"fit 6", family:"SLEEVELESS", styleNo:"HSAW26CGSS7159", colour:"LAVENDER, WHITE", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"4-10YRS", qty:535, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:false, soReq:false, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-30", artwork:"2026-05-30", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{fitSend:"2026-05-27", artwork:"2026-05-25", fabricIH:"2026-06-05"} },
  { id:5, orderNo:"T1", sampleFit:"fit 10", family:"SLEEVELESS", styleNo:"HSAW26CGSS7156", colour:"CLOUD DANCER, LILAC, SACHET PINK", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"4-10YRS", qty:535, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:false, soReq:false, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-27", artwork:"2026-05-30", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{fitSend:"2026-05-27", artwork:"2026-05-25", fabricIH:"2026-06-05"} },
  { id:6, orderNo:"T1", sampleFit:"fit 10", family:"SLEEVELESS", styleNo:"HSAW26CGSS7157", colour:"CLOUD DANCER, SACHET PINK", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"4-10YRS", qty:535, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:false, soReq:false, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-27", artwork:"2026-05-30", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{fitSend:"2026-05-27", artwork:"2026-05-25", fabricIH:"2026-06-01"} },
  { id:7, orderNo:"T1", sampleFit:"fit 1-3", family:"SWEAT", styleNo:"HSAW26CGSS7161", colour:"BLACK", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"4-10YRS", qty:535, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:false, soReq:false, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-27", artwork:"2026-05-30", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{fitSend:"2026-05-27", artwork:"2026-05-25", fabricIH:"2026-06-01"} },
  { id:8, orderNo:"T1", sampleFit:"fit 1-3", family:"SWEAT", styleNo:"HSAW26CGSS7163", colour:"BERRY RED, MOCHA", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"4-10YRS", qty:535, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:false, soReq:false, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-27", artwork:"2026-05-30", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{fitSend:"2026-05-27", artwork:"2026-05-25", fabricIH:"2026-06-01"} },
  { id:9, orderNo:"T1", sampleFit:"fit 11", family:"SWEAT", styleNo:"HSAW26CGSS7170", colour:"BERRY RED", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"4-10YRS", qty:535, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:false, soReq:false, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-27", artwork:"2026-05-30", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{fitSend:"2026-05-27", artwork:"2026-05-25", fabricIH:"2026-06-01"} },
  { id:10, orderNo:"T1", sampleFit:"fit 1-3", family:"SWEAT", styleNo:"HSAW26CGSS7184", colour:"CLOUD DANCER", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"4-10YRS", qty:535, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:false, soReq:false, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-27", artwork:"2026-05-25", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{fitSend:"2026-05-27", artwork:"2026-05-25", fabricIH:"2026-06-05"} },
  { id:11, orderNo:"T1", sampleFit:"fit 1-3", family:"SWEAT", styleNo:"HSAW26CGSS7534", colour:"CLOUD DANCER, WHITE SHERPA", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"4-10YRS", qty:315, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:false, soReq:false, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-27", artwork:"2026-05-30", artAppr:"2026-05-25", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{fitSend:"2026-05-27", artwork:"2026-05-25", fabricIH:"2026-06-05"} },
  { id:12, orderNo:"T1", sampleFit:"fit 9", family:"SWEAT", styleNo:"HSAW26CGSS7178", colour:"CLOUD DANCER", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"4-10YRS", qty:535, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:true, soReq:true, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-27", artwork:"2026-05-25", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{fitSend:"2026-05-27", artwork:"2026-05-25", fabricIH:"2026-06-05"} },
  { id:13, orderNo:"T1", sampleFit:"fit 16", family:"SWEAT", styleNo:"HSAW26CGSS7526", colour:"BLACK", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"4-10YRS", qty:315, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:true, soReq:true, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-12", fitAppr:"2026-05-20", artwork:"2026-05-19", artAppr:"2026-05-25", strikeOff:"2026-05-28", soAppr:"2026-06-01", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{artwork:"2026-05-25", fabricIH:"2026-06-05"} },
  { id:14, orderNo:"T1", sampleFit:"fit 4", family:"HOODIE", styleNo:"HSAW26CGSS7165", colour:"BERRY RED, CLOUD DANCER", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"4-10YRS", qty:535, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:true, soReq:true, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-28", artwork:"2026-05-25", artAppr:"2026-05-28", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{fitSend:"2026-05-27", artwork:"2026-05-25", fabricIH:"2026-06-01"} },
  { id:15, orderNo:"T1", sampleFit:"fit 9", family:"SWEAT", styleNo:"HSAW26CGSS7177", colour:"SALSA RED", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"4-10YRS", qty:642, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:true, soReq:true, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-30", artwork:"2026-05-25", artAppr:"2026-05-28", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{fitSend:"2026-05-27", artwork:"2026-05-25", fabricIH:"2026-06-01"} },
  { id:16, orderNo:"T1", sampleFit:"fit 6", family:"SLEEVELESS", styleNo:"HSAW26CGSS7155", colour:"BLACK, BRIGHT ORCHID", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"4-10YRS", qty:535, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:true, soReq:true, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-27", artwork:"2026-05-28", artAppr:"2026-05-28", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{fitSend:"2026-05-27", artwork:"2026-05-27", fabricIH:"2026-06-01"} },
  { id:17, orderNo:"T1", sampleFit:"fit 1-3", family:"SWEAT", styleNo:"HSAW26CGSS7175", colour:"YELLOW", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"4-10YRS", qty:315, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:true, soReq:true, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-27", artwork:"2026-05-25", artAppr:"2026-05-29", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{fitSend:"2026-05-27", artwork:"2026-05-25", fabricIH:"2026-06-01"} },
  { id:18, orderNo:"T1", sampleFit:"fit 1-3", family:"SWEAT", styleNo:"HSAW26CGSS7173", colour:"CLOUD DANCER", owner:"Tamal", setId:"", setRole:"SWEATSHIRT", age:"4-10YRS", qty:315, ordRec:"2026-05-18", delivery:"2026-06-25", fitReq:true, printReq:true, soReq:true, ppBypass:true, labDipReq:true, ppNeeded:true, remarks:"", actuals:{techpack:"2026-05-18", fitSend:"2026-05-27", artwork:"2026-05-25", artAppr:"2026-05-30", labDip:"2026-05-15", labAppr:"2026-05-20"}, revs:{fitSend:"2026-05-27", artwork:"2026-05-25", strikeOff:"2026-06-03", fabricIH:"2026-06-05"} },
];

const REJECTABLE=["fitAppr","artAppr","soAppr","labAppr","ppAppr"]; // approval stages that can be rejected
const SKIPPABLE_STAGES=["fitSend","fitAppr","artwork","artAppr","strikeOff","soAppr","labDip","labAppr","ppSample","ppAppr"]; // activities that can be waived/skipped
const APPR_OF_SEND={ fitSend:"fitAppr", artwork:"artAppr", strikeOff:"soAppr", labDip:"labAppr", ppSample:"ppAppr" }; // send/make stage -> the approval that can reject it
const REWORK_DAYS={ fitSend:4, artwork:2, strikeOff:3, labDip:7, ppSample:4 }; // working days added on rejection (redo+resend)
const applicableStages=(s)=> STAGES.filter(st=> st.flag===null || s[st.flag]);

function computeStyle(s, cfg){
  const ordRec=parse(s.ordRec), delivery=parse(s.delivery);
  const leadOf=(st)=>{ const v=cfg&&cfg.leads&&cfg.leads[st.key]; return v==null?st.lead:v; }; const rwOf=(st)=>{ const v=cfg&&cfg.rework&&cfg.rework[st.key]; return v==null?(REWORK_DAYS[st.key]||st.lead):v; }; const CUTD=(cfg&&cfg.fabricCutoff!=null)?cfg.fabricCutoff:FABRIC_CUTOFF_DAYS; const GATED=(cfg&&cfg.relGate!=null)?cfg.relGate:REL_GATE_DAYS;
  const cutoff=addWorkdays(delivery,-CUTD);
  const eff={}, plan={};
  const applies=(k)=>{ const st=STAGES.find(x=>x.key===k); return st.flag===null||s[st.flag]; };
  const actualOf=(k)=>parse(s.actuals[k]); const revOf=(k)=>parse(s.revs?.[k]); const rejOf=(k)=>parse(s.rejects?.[k]); const skipOf=(k)=>parse(s.skips?.[k]);
  STAGES.forEach(st=>{
    let p;
    if(st.cutoff){ const base=s.labDipReq?(eff["labAppr"]||eff["labDip"]||ordRec):ordRec; p=s.labDipReq?new Date(Math.max(addWorkdays(base,15)?.getTime()||0, cutoff.getTime())):cutoff; }
    else { let predEff; if(st.key==="prodFile") predEff = s.ppBypass ? eff["fabricIH"] : eff["ppAppr"]; else predEff = st.pred==="__ord"?ordRec:eff[st.pred]; if((st.key==="ppSample"||st.key==="prodFile") && s.fitReq && eff["fitAppr"]) predEff = new Date(Math.max((predEff&&predEff.getTime())||0, eff["fitAppr"].getTime())); p=addWorkdays(predEff||ordRec, leadOf(st)); }
    const apprK=APPR_OF_SEND[st.key]; const rejAppr = !!(apprK && rejOf(apprK) && !actualOf(apprK));
    const selfRej = REJECTABLE.includes(st.key) && rejOf(st.key) && !actualOf(st.key);
    if(rejAppr){ const rjd=rejOf(apprK); const auto=addWorkdays(rjd, rwOf(st)); const a=actualOf(st.key); const rv=revOf(st.key); plan[st.key]=auto;
      if(a && a>rjd) eff[st.key]=a; else if(rv && rv>=rjd) eff[st.key]=rv; else eff[st.key]=auto; } // redo: re-sent actual wins, else fresh revised, else rejection+rework days
    else if(selfRej){ const rjd=rejOf(st.key); const rv=revOf(st.key); plan[st.key]=p; eff[st.key]=(rv && rv>=rjd)?rv:p; } // rejected approval cascades off redone send
    else { plan[st.key]=p; eff[st.key]=actualOf(st.key)||skipOf(st.key)||revOf(st.key)||p; }
  });
  const stages=applicableStages(s).map(st=>{ const apprK=APPR_OF_SEND[st.key]; const apprRej=apprK?rejOf(apprK):null; const selfRejDate=REJECTABLE.includes(st.key)?rejOf(st.key):null; const rejAppr=!!(apprK&&apprRej&&!actualOf(apprK)); const a=actualOf(st.key); const skp=skipOf(st.key); const isSkip=!!skp&&!a; const rjd_=rejAppr?apprRej:null; const resent=rejAppr&&a&&a>rjd_; const rework=rejAppr&&!resent; const rejected=REJECTABLE.includes(st.key)&&!!selfRejDate&&!actualOf(st.key); const rjd=rework?rjd_:(rejected?selfRejDate:null); let rv=revOf(st.key); if(rjd&&rv&&rv<rjd) rv=null; const histReject = a ? ((apprRej&&a>=apprRej)?apprRej:(selfRejDate||null)) : null; return { ...st, actual:a, rev:rv, reject:rjd, histReject, rework:isSkip?false:rework, rejected:isSkip?false:rejected, skipped:isSkip, skip:skp, plan:plan[st.key], done: isSkip?true:(rework?false:!!a) }; });
  let nextPending=null, lastActual=null, lastActualKey=null;
  stages.forEach(r=>{ if(r.actual&&(!lastActual||r.actual>lastActual)){ lastActual=r.actual; lastActualKey=r.key; } if(!r.done&&!nextPending) nextPending=r; });
  const released=stages.every(r=>r.done);
  const idle=lastActual?Math.max(0,netWorkdays(lastActual,TODAY)):null;
  const get=(k)=>stages.find(r=>r.key===k); const done=(k)=>!!(get(k)&&get(k).done); const rejected=(k)=>{ const r=get(k); return !!(r&&r.rejected); }; const isSkipped=(k)=>{ const r=get(k); return !!(r&&r.skipped); };
  const fabricInHouse=done("fabricIH"); const fihA=actualOf("fabricIH"); const lateFIH=(k)=>{ const r=get(k); return !!(fihA && r && r.actual && r.actual>fihA); };
  const lastPlan=stages[stages.length-1]?.plan;
  const float=lastPlan?netWorkdays(lastPlan,delivery):null;
  let status="On Track", tone="ok";
  if(released){ status="Released"; tone="done"; }
  else if(nextPending&&nextPending.plan&&TODAY>nextPending.plan){ status=`Overdue ${Math.round((TODAY-nextPending.plan)/ONE_DAY)}d`; tone="late"; }
  else if(idle!==null&&idle>=7){ status=`Idle ${idle}d`; tone="warn"; }
  const dueText=(k)=>{ const r=get(k); if(!r||!r.plan) return "pending"; return TODAY>r.plan?`OVERDUE ${Math.round((TODAY-r.plan)/ONE_DAY)}d`:`due ${fmt(r.plan)}`; };
  const bs=(txt,tn)=>({txt,tone:tn});
  let fitBranch;
  if(!s.fitReq) fitBranch=bs("—","na"); else if(done("fitAppr")) fitBranch=bs(isSkipped("fitAppr")?"Fit Skipped":(lateFIH("fitAppr")?"Fit Approved · LATE":"Fit Approved"), isSkipped("fitAppr")?"ok":(lateFIH("fitAppr")?"late":"ok")); else if(rejected("fitAppr")) fitBranch=bs("Fit REJECTED · rework","late"); else if(fabricInHouse) fitBranch=bs("Not done before Fabric IH","late"); else if(done("fitSend")) fitBranch=bs(`Fit appr ${dueText("fitAppr")}`,TODAY>(get("fitAppr")?.plan||TODAY)?"late":"warn"); else fitBranch=bs(`Fit send ${dueText("fitSend")}`,TODAY>(get("fitSend")?.plan||TODAY)?"late":"warn");
  let printBranch; const printDone=s.soReq?done("soAppr"):done("artAppr"); const printComp=s.soReq?"soAppr":"artAppr";
  if(!s.printReq) printBranch=bs("—","na"); else if(printDone) printBranch=bs(isSkipped(printComp)?"Print Skipped":(lateFIH(printComp)?"Print Approved · LATE":"Print Approved"), isSkipped(printComp)?"ok":(lateFIH(printComp)?"late":"ok")); else if(rejected("artAppr")||rejected("soAppr")) printBranch=bs("Print REJECTED · rework","late"); else if(fabricInHouse) printBranch=bs("Not done before Fabric IH","late"); else if(!done("artwork")) printBranch=bs(`Artwork ${dueText("artwork")}`,TODAY>(get("artwork")?.plan||TODAY)?"late":"warn"); else if(!done("artAppr")) printBranch=bs(`Art appr ${dueText("artAppr")}`,"warn"); else if(s.soReq&&!done("strikeOff")) printBranch=bs(`S/O ${dueText("strikeOff")}`,"warn"); else printBranch=bs(`S/O appr ${dueText("soAppr")}`,"warn");
  let fabricBranch; const fabPlan=get("fabricIH")?.plan;
  const fabDue=fabPlan?(TODAY>fabPlan?`IH OVERDUE ${Math.round((TODAY-fabPlan)/ONE_DAY)}d`:`IH due ${fmt(fabPlan)}`):"IH —";
  const fabTone=fabPlan&&TODAY>fabPlan?"late":"warn";
  if(fabricInHouse) fabricBranch=bs("Bulk Fabric In-House","ok"); else if(rejected("labAppr")) fabricBranch=bs("Lab Dip REJECTED · rework","late"); else if(s.labDipReq&&done("labAppr")) fabricBranch=bs(`Lab Dip Appr | ${fabDue}`,fabTone); else if(s.labDipReq&&done("labDip")) fabricBranch=bs(`Lab dip sent, appr pending | ${fabDue}`,"warn"); else if(s.labDipReq) fabricBranch=bs(`Lab dip pending | ${fabDue}`,"warn"); else fabricBranch=bs(fabDue,fabTone);
  let ppBranch;
  if(!s.ppNeeded) ppBranch=bs("PP Not Required","na"); else if(done("ppAppr")) ppBranch=bs(isSkipped("ppAppr")?"PP Skipped":(lateFIH("ppAppr")?"PP Approved · LATE":"PP Approved"), isSkipped("ppAppr")?"ok":(lateFIH("ppAppr")?"late":"ok")); else if(rejected("ppAppr")) ppBranch=bs("PP REJECTED · rework","late"); else if(done("ppSample")) ppBranch=bs(`PP appr ${dueText("ppAppr")}`,"warn"); else if(fabricInHouse) ppBranch=bs(`PP sample ${dueText("ppSample")}`,"warn"); else ppBranch=bs("Awaiting bulk fabric","warn");
  // ---- Production File: a tracked activity; reflects PP bypass vs PP-approval gate ----
  let prodFileBranch;
  { const pfA=actualOf("prodFile"); const pfP=eff["prodFile"]; const pfDue=pfP?`due ${fmt(pfP)}`:"";
    if(pfA){ prodFileBranch=bs(`Released ${fmt(pfA)}`,"ok"); }
    else { const prodGate=addWorkdays(delivery,-CUTD); const overdue=pfP&&pfP<TODAY; const pastGate=pfP&&prodGate&&pfP>prodGate; const tn=(overdue||pastGate)?"late":"warn";
      if(s.ppBypass){ const ready=fabricInHouse||done("fabricIH"); prodFileBranch=ready?bs(`Bypass · ready ${pfDue}`,tn):bs(`Bypass · awaiting fabric`,tn); }
      else { const ready=done("ppAppr"); prodFileBranch=ready?bs(`Ready ${pfDue}`,tn):bs(`Awaiting PP appr`,tn); } } }
  let fabricCountdown;
  if(fabricInHouse) fabricCountdown={txt:"in-house",date:fihA,n:9e9,tone:"ok"}; else if(fabPlan){ const n=netWorkdays(TODAY,fabPlan); fabricCountdown={txt:n<0?`${-n}d over`:`${n}d`,date:fabPlan,n,tone:n<0?"late":n<=7?"warn":"ok"}; } else fabricCountdown={txt:"no plan",date:null,n:null,tone:"warn"};
  const releaseGate=addWorkdays(delivery,-GATED);
  let projRelease;
  if(released) projRelease=lastActual;
  else { let cur=eff["fabricIH"]; const chain = s.ppBypass ? ["prodFile"] : ["ppSample","ppAppr","prodFile"]; chain.forEach(k=>{ if(!applies(k)) return; const a=actualOf(k), r=revOf(k); const st=STAGES.find(x=>x.key===k); if(a) cur=a; else if(r) cur=r; else cur=addWorkdays(cur,leadOf(st)); }); projRelease=cur; }
  const gateGap=projRelease&&releaseGate?Math.round((releaseGate-projRelease)/ONE_DAY):null;
  const releaseOnTrack=projRelease&&releaseGate?projRelease<=releaseGate:true;
  const projTone=released?"done":(!releaseOnTrack?"late":(gateGap!=null&&gateGap<=5?"warn":"ok"));
  if(!released){ if(!releaseOnTrack){ if(!String(status).startsWith("Overdue")) status="Delivery risk"; tone="late"; } else if(tone==="ok" && gateGap!=null && gateGap<=5){ status=`Tight · ${gateGap}d`; tone="warn"; } }
  const total=stages.length, doneCount=stages.filter(r=>r.done).length;
  const pct=total?Math.round((doneCount/total)*100):0;
  const ownerToChase=released?"—":nextPending.owner;
  // ---- parallel chase: owners of every pending stage whose predecessor is done (actionable now) ----
  const stById=(k)=>STAGES.find(x=>x.key===k);
  const appl=(k)=>{ const st=stById(k); return st&&(st.flag===null||s[st.flag]); };
  const predDone=(st)=>{ if(st.cutoff) return s.labDipReq?done("labAppr"):true; if(st.pred==="__ord") return true; let p=st.pred; while(p&&p!=="__ord"&&!appl(p)) p=stById(p)?.pred; return (!p||p==="__ord")?true:done(p); };
  const chaseCount={};
  if(!released) STAGES.forEach(st=>{ if(appl(st.key)&&!done(st.key)&&predDone(st)) chaseCount[st.owner]=(chaseCount[st.owner]||0)+1; });
  const chaseOwners=Object.entries(chaseCount).map(([owner,count])=>({owner,count}));
  const frontierReady=(k)=>{ if(k==="ppSample") return fabricInHouse && (!s.fitReq || done("fitAppr")); if(k==="prodFile"){ const base = s.ppBypass ? fabricInHouse : done("ppAppr"); return base && (!s.fitReq || done("fitAppr")); } return true; };
  const frontier=new Set(); Object.entries(BRANCH_STAGES).forEach(([bk,keys])=>{ const nx=keys.find(k=>applies(k)&&!done(k)); if(!nx) return; if(fabricInHouse && !s.ppBypass && (bk==="fit"||bk==="print")) return; if(frontierReady(nx)) frontier.add(nx); });
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
const TONE_STYLE={ ok:{dot:"#1f6f54",bg:"#eef6f1",fg:"#16523d"}, warn:{dot:"#b4801a",bg:"#fbf4e6",fg:"#7a560f"}, late:{dot:"#c0392b",bg:"#fcecea",fg:"#8c241a"}, done:{dot:"#555",bg:"#efefea",fg:"#444"} };
const BR_TONE={ ok:{bg:"#eef6f1",fg:"#16523d"}, warn:{bg:"#fbf4e6",fg:"#7a560f"}, late:{bg:"#fcecea",fg:"#8c241a"}, na:{bg:"transparent",fg:"#c4c0b8"}, done:{bg:"#efefea",fg:"#555"} };
const FLAG_DEFS=[ {key:"fitReq",short:"FIT",title:"Fit sample required"}, {key:"printReq",short:"PRT",title:"Print required"}, {key:"soReq",short:"S/O",title:"Strike-off required"}, {key:"labDipReq",short:"LAB",title:"Lab dip required"}, {key:"ppBypass",short:"BYP",title:"PP bypass — Prod File flows straight from Fabric IH (not PP Appr)"}, {key:"ppNeeded",short:"PP",title:"PP sample required"} ];
const FILL_SWATCHES=["#fff7ec","#fde2e1","#e7f3ec","#e3edf9","#f3e8fa","#fff3bf",""];

const NORMH=(h)=> String(h||"").toLowerCase().replace(/[^a-z0-9]/g,"");
const HEADER_MAP={ samplefit:"sampleFit", family:"family", fit:"orderNo", orderno:"orderNo", order:"orderNo", tranche:"orderNo", styleno:"styleNo", style:"styleNo", colour:"colour", color:"colour", brand:"brand", buyer:"brand", juniorowner:"owner", junior:"owner", owner:"owner", merchant:"owner", setpackid:"setId", setid:"setId", setpackrole:"setRole", setrole:"setRole", agegroup:"age", age:"age", orderqty:"qty", qty:"qty", quantity:"qty", reserve3:"fabricType", fabrictype:"fabricType", fabric:"fabricType", construction:"fabricType", orderreceived:"ordRec", orderdate:"ordRec", received:"ordRec", deliverydate:"delivery", delivery:"delivery", fitreq:"fitReq", printreq:"printReq", soreq:"soReq", ppbypass:"ppBypass", labdipreq:"labDipReq", ppneeded:"ppNeeded" };
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
const TEXT_COLS=["orderNo","sampleFit","family","colour","owner","setId","setRole","remarks"];
const isEditableCol=(col)=> col==="__style"||col==="qty"||col==="ordRec"||col==="delivery"||TEXT_COLS.includes(col)||STAGE_KEYS.includes(col);
const isDateCol=(col)=> col==="ordRec"||col==="delivery"||STAGE_KEYS.includes(col);
const BRANCH_STAGES={ fit:["fitSend","fitAppr"], print:["artwork","artAppr","strikeOff","soAppr"], fabric:["labDip","labAppr","fabricIH"], pp:["ppSample","ppAppr"], prod:["prodFile"] };
const BRANCH_LABEL={ fit:"Fit", print:"Print", fabric:"Fabric", pp:"PP", prod:"Production" };
const BRANCH_OF={}; Object.entries(BRANCH_STAGES).forEach(([b,ks])=>ks.forEach(k=>{ BRANCH_OF[k]=BRANCH_LABEL[b]; }));
function branchTarget(s,c,branch){ const keys=BRANCH_STAGES[branch].filter(k=>{ const st=STAGES.find(x=>x.key===k); return st.flag===null||s[st.flag]; }); for(const k of keys){ if(!c.stages.find(r=>r.key===k)?.done) return k; } return keys[keys.length-1]; }

function CalPopup({ value, onPick, onClose, label }){
  const init=value?parse(value):new Date(TODAY);
  const [view,setView]=useState(new Date(init.getFullYear(),init.getMonth(),1));
  const y=view.getFullYear(), m=view.getMonth();
  const first=new Date(y,m,1).getDay(), days=new Date(y,m+1,0).getDate();
  const cells=[]; for(let i=0;i<first;i++) cells.push(null); for(let d=1;d<=days;d++) cells.push(d);
  const sel=value?parse(value):null;
  return (
    <div style={{ position:"absolute", zIndex:90, top:"100%", left:0, marginTop:2 }} onClick={(e)=>e.stopPropagation()}>
      <div style={{ background:"#fff", border:"1px solid #1a1a1a", boxShadow:"4px 4px 0 #1a1a1a", padding:10, width:228, fontFamily:"'JetBrains Mono', monospace" }}>
        {label && <div style={{ fontSize:9, fontWeight:700, color:"#d97706", marginBottom:6, textTransform:"uppercase", letterSpacing:0.5 }}>{label}</div>}
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}><button onClick={()=>setView(new Date(y,m-1,1))} style={navBtn}>‹</button><span style={{ fontWeight:700, fontSize:12 }}>{view.toLocaleDateString("en-GB",{month:"long",year:"numeric"})}</span><button onClick={()=>setView(new Date(y,m+1,1))} style={navBtn}>›</button></div>
        <div style={{ display:"grid", gridTemplateColumns:"repeat(7,1fr)", gap:2 }}>
          {["S","M","T","W","T","F","S"].map((d,i)=>(<div key={i} style={{ textAlign:"center", fontSize:9, color:"#999" }}>{d}</div>))}
          {cells.map((d,i)=>{ if(!d) return <div key={i}/>; const isSel=sel&&sel.getFullYear()===y&&sel.getMonth()===m&&sel.getDate()===d; const isToday=TODAY.getFullYear()===y&&TODAY.getMonth()===m&&TODAY.getDate()===d; const isSun=new Date(y,m,d).getDay()===0; return <button key={i} onClick={()=>{ onPick(`${y}-${String(m+1).padStart(2,"0")}-${String(d).padStart(2,"0")}`); onClose(); }} style={{ fontSize:11, padding:"5px 0", cursor:"pointer", border:"none", fontFamily:"inherit", background:isSel?"#d97706":"transparent", color:isSel?"#fff":isSun?"#c0392b":"#1a1a1a", fontWeight:isSel?700:400, border:(isToday&&!isSel)?"1px solid #1a1a1a":"1px solid transparent", boxSizing:"border-box" }}>{d}</button>; })}
        </div>
        <div style={{ display:"flex", gap:6, marginTop:8 }}><button onClick={()=>{ onPick(iso(TODAY)); onClose(); }} style={{ ...chip, flex:1 }}>Today</button>{value && <button onClick={()=>{ onPick(null); onClose(); }} style={{ ...chip, flex:1 }}>Clear</button>}</div>
      </div>
    </div>
  );
}
const navBtn={ border:"none", background:"transparent", cursor:"pointer", fontSize:18, lineHeight:1, padding:"0 6px", fontFamily:"inherit" };
const chip={ fontSize:10, padding:"4px 6px", border:"1px solid #1a1a1a", background:"#f4f0e8", cursor:"pointer", fontFamily:"'JetBrains Mono', monospace", fontWeight:600 };

function BranchPill({ b, onJump }){
  if(!b) return null; const t=BR_TONE[b.tone]||BR_TONE.na;
  if(b.tone==="na") return <span style={{ color:"#ccc", fontSize:10 }}>{b.txt}</span>;
  return (<span style={{ display:"inline-flex", flexDirection:"column", alignItems:"flex-start", gap:1, maxWidth:"100%" }}>
    <span onClick={(e)=>{ if(onJump){ e.stopPropagation(); onJump(); } }} title={`${b.txt}  ·  click → jump to enter`} style={{ display:"inline-flex", alignItems:"center", gap:3, background:t.bg, color:t.fg, padding:"2px 7px", fontSize:9.5, fontWeight:600, whiteSpace:"nowrap", maxWidth:"100%", overflow:"hidden", textOverflow:"ellipsis", cursor:onJump?"pointer":"default", textDecoration:onJump?"underline dotted":"none" }}>{b.txt}{onJump && <CornerDownRight size={9} style={{ flexShrink:0 }}/>}</span>
    {b.last && <span style={{ fontSize:8, color:"#9a958a", whiteSpace:"nowrap" }}>✓ {b.last.l} · {fmt(b.last.d)}</span>}
  </span>);
}

const PEER_COLORS=["#1f6f54","#2563a6","#b4531a","#6d4aab","#c0392b","#0e7490","#9d174d","#4d7c0f"];
const colorFor=(id)=>{ let h=0; const s=String(id); for(let i=0;i<s.length;i++) h=(h*31+s.charCodeAt(i))>>>0; return PEER_COLORS[h%PEER_COLORS.length]; };
const initials=(n)=>String(n||"?").trim().split(/\s+/).map(w=>w[0]||"").slice(0,2).join("").toUpperCase()||"?";
function PeerTag({ who }){ if(!who) return null; return (<span style={{ position:"absolute", inset:0, border:"2px solid "+who.color, pointerEvents:"none", zIndex:4, boxSizing:"border-box" }}><span style={{ position:"absolute", bottom:0, left:0, background:who.color, color:"#fff", fontSize:8, fontWeight:700, padding:"0 4px", whiteSpace:"nowrap", lineHeight:"12px", maxWidth:"100%", overflow:"hidden", textOverflow:"ellipsis" }}>{who.name}</span></span>); }
function MerchTracker({ me, onSignOut }){
  const [styles,setStyles]=useState([]); // loaded from Supabase on mount
  const role=(me&&me.role)||"junior";
  const [usersOpen,setUsersOpen]=useState(false);
  const [textScale,setTextScale]=useState(()=>{ try{ const v=parseFloat(localStorage.getItem("mt_textscale")); return (v&&v>=0.7&&v<=1.6)?v:1; }catch(e){ return 1; } });
  const bumpScale=(d)=>setTextScale(v=>{ const n=Math.min(1.6,Math.max(0.7,Math.round((v+d)*100)/100)); try{ localStorage.setItem("mt_textscale",String(n)); }catch(e){} return n; });
  const PF=(()=>{ try{ return JSON.parse(localStorage.getItem("mt_trackfilters")||"{}"); }catch(e){ return {}; } })();
  const [search,setSearch]=useState(PF.search||"");
  const [statusFilter,setStatusFilter]=useState(PF.statusFilter||"All");
  const [ownerFilter,setOwnerFilter]=useState(PF.ownerFilter||"All");
  const [archiveView,setArchiveView]=useState(PF.archiveView||"active");
  const [activityFilter,setActivityFilter]=useState(PF.activityFilter||null);
  const [viewSnap,setViewSnap]=useState(null); // saved tracker view before a drill, for one-click restore
  const scrollWrapRef=useRef(null);
  const [saved,setSaved]=useState(false);
  const [fillOpen,setFillOpen]=useState(false);
  const [colsOpen,setColsOpen]=useState(false);
  const [sel,setSel]=useState(null);      // anchor {id,col}
  const [focus,setFocus]=useState(null);   // range focus {id,col}
  const [editing,setEditing]=useState(null);
  const [editVal,setEditVal]=useState("");
  const [sort,setSort]=useState({ col:null, dir:1 });
  const [hidden,setHidden]=useState(()=>{ try{ return new Set(JSON.parse(localStorage.getItem("mt_hidden_cols")||"[]")); }catch(e){ return new Set(); } });
  useEffect(()=>{ try{ localStorage.setItem("mt_hidden_cols", JSON.stringify([...hidden])); }catch(e){} },[hidden]);
  const [specialClip,setSpecialClip]=useState(null);
  const [expOpen,setExpOpen]=useState(false); const [expMode,setExpMode]=useState("full"); const [expBuf,setExpBuf]=useState(2); const [expIncBuf,setExpIncBuf]=useState(false);
  const [frOpen,setFrOpen]=useState(false); const [frFind,setFrFind]=useState(""); const [frRepl,setFrRepl]=useState(""); const [frScope,setFrScope]=useState("filtered"); const [frCase,setFrCase]=useState(false);
  const [freezeN,setFreezeN]=useState(1);  // # leading columns frozen (incl style)
  const [colW,setColW]=useState({});  // per-column width overrides (drag to resize)
  const [fills,setFills]=useState({});
  const [notes,setNotes]=useState({});
  const [noteEditing,setNoteEditing]=useState(false);
  const [noteText,setNoteText]=useState("");
  const [clip,setClip]=useState(null);     // {values:2D,h,w}
  const [showAux,setShowAux]=useState(false); // toggle: reveal underlying auto/plan + revised dates in cells
  const [cfg,setCfg]=useState(DEFAULT_CFG); // editable calculation numbers (Settings tab)
  const [tab,setTab]=useState(PF.tab||"tracker");
  const [colFilters,setColFilters]=useState(PF.colFilters||{});
  const [filterCol,setFilterCol]=useState(null); // which header filter is open
  const [past,setPast]=useState([]); const [future,setFuture]=useState([]);
  const [filling,setFilling]=useState(false); const [fillFrom,setFillFrom]=useState(null); const [fillTo,setFillTo]=useState(null);
  const selectingRef=useRef(false); const [dragSel,setDragSel]=useState(false);
  useEffect(()=>{ const up=()=>{ if(selectingRef.current){ selectingRef.current=false; setDragSel(false); } }; window.addEventListener("mouseup",up); return ()=>window.removeEventListener("mouseup",up); },[]);
  const [newRow,setNewRow]=useState({ styleNo:"", orderNo:"", sampleFit:"", family:"", colour:"", owner:"", qty:"", ordRec:iso(TODAY), delivery:"", fitReq:true, printReq:false, soReq:false, ppBypass:false, labDipReq:true, ppNeeded:true });
  const [newError,setNewError]=useState("");
  const savedTimer=useRef();
  const gridRef=useRef();
  const firstRender=useRef(true);
  const loadedRef=useRef(false);
  const savedRef=useRef({ sty:{}, stg:{}, meta:{} }); // last-persisted snapshot, keyed per row, so we only write what changed
  const S2C={ orderNo:"order_no", styleNo:"style_no", sampleFit:"sample_fit", family:"family", colour:"colour", brand:"brand", fabricType:"fabric_type", owner:"owner", setId:"set_id", setRole:"set_role", age:"age", qty:"qty", ordRec:"order_date", delivery:"delivery_date", fitReq:"fit_req", printReq:"print_req", soReq:"so_req", ppBypass:"pp_bypass", labDipReq:"lab_dip_req", ppNeeded:"pp_needed", remarks:"remarks" };
  const styleToRow=(s)=>{ const r={ id:s.id }; Object.entries(S2C).forEach(([k,col])=>{ r[col]= k==="qty"?(Number(s[k])||0):(s[k]||null); }); r.archived=!!s.archived; return r; };
  const rowToStyle=(row,byId)=>({ id:row.id, orderNo:row.order_no||"", sampleFit:row.sample_fit||"", family:row.family||"", styleNo:row.style_no||"", colour:row.colour||"", brand:row.brand||"", fabricType:row.fabric_type||"", owner:row.owner||"", setId:row.set_id||"", setRole:row.set_role||"", age:row.age||"", qty:row.qty||0, ordRec:row.order_date||"", delivery:row.delivery_date||"", fitReq:!!row.fit_req, printReq:!!row.print_req, soReq:!!row.so_req, ppBypass:!!row.pp_bypass, labDipReq:!!row.lab_dip_req, ppNeeded:!!row.pp_needed, remarks:row.remarks||"", actuals:(byId[row.id]&&byId[row.id].actuals)||{}, revs:(byId[row.id]&&byId[row.id].revs)||{}, rejects:(byId[row.id]&&byId[row.id].rejects)||{}, skips:(byId[row.id]&&byId[row.id].skips)||{}, archived:!!row.archived });
  // LOAD everything from Supabase (also used by the Sync button)
  const loadShared=async()=>{ try{
    const [styRes, sdRes, cmRes] = await Promise.all([
      supabase.from("styles").select("*").order("id"),
      supabase.from("stage_dates").select("*"),
      supabase.from("cell_meta").select("*"),
    ]);
    if(styRes.error||!styRes.data){ console.error(styRes.error); return; }
    const byId={}; (sdRes.data||[]).forEach(r=>{ const e=(byId[r.style_id]=byId[r.style_id]||{actuals:{},revs:{},rejects:{},skips:{}}); if(r.actual_date) e.actuals[r.stage]=r.actual_date; if(r.revised_date) e.revs[r.stage]=r.revised_date; if(r.reject_date) e.rejects[r.stage]=r.reject_date; if(r.skip_date) e.skips[r.stage]=r.skip_date; });
    const appStyles=styRes.data.map(row=>rowToStyle(row,byId));
    setStyles(appStyles);
    const SR={ sty:{}, stg:{}, meta:{} };
    appStyles.forEach(s=>{ SR.sty[s.id]=JSON.stringify(styleToRow(s)); STAGE_KEYS.forEach(k=>{ SR.stg[s.id+":"+k]=JSON.stringify({ style_id:s.id, stage:k, revised_date:(s.revs&&s.revs[k])||null, actual_date:s.actuals[k]||null, reject_date:(s.rejects&&s.rejects[k])||null, skip_date:(s.skips&&s.skips[k])||null }); }); });
    (cmRes.data||[]).forEach(r=>{ if(r.fill||r.note) SR.meta[r.style_id+":"+r.col]=JSON.stringify({ style_id:r.style_id, col:r.col, fill:r.fill||null, note:r.note||null }); });
    savedRef.current=SR;
    try{ const cfgRes=await supabase.from("app_settings").select("data").eq("id","global").maybeSingle(); if(cfgRes&&cfgRes.data&&cfgRes.data.data){ const d=cfgRes.data.data; setCfg({ ...DEFAULT_CFG, ...d, leads:{...DEFAULT_CFG.leads,...(d.leads||{})}, rework:{...DEFAULT_CFG.rework,...(d.rework||{})}, upcoming:{...DEFAULT_CFG.upcoming,...(d.upcoming||{})} }); } }catch(e){ /* settings table optional */ }
    const f={}, n={}; (cmRes.data||[]).forEach(r=>{ if(r.fill) f[`${r.style_id}:${r.col}`]=r.fill; if(r.note) n[`${r.style_id}:${r.col}`]=r.note; });
    setFills(f); setNotes(n); loadedRef.current=true; flash();
  }catch(e){ console.error("load failed",e); } };
  useEffect(()=>{ loadShared(); },[]);
  const editingRef=useRef(null); useEffect(()=>{ editingRef.current=editing; },[editing]);
  const presRef=useRef(null); const [peers,setPeers]=useState([]);
  useEffect(()=>{ if(!me) return; const ch=supabase.channel("merch-presence",{ config:{ presence:{ key:String(me.id) } } });
    ch.on("presence",{ event:"sync" },()=>{ const st=ch.presenceState(); const arr=[]; Object.keys(st).forEach(k=>{ const m=st[k]&&st[k][0]; if(m&&String(m.id)!==String(me.id)) arr.push(m); }); setPeers(arr); });
    ch.subscribe(async(status)=>{ if(status==="SUBSCRIBED"){ presRef.current=ch; try{ await ch.track({ id:me.id, name:me.name||me.email, role:me.role, cell:null }); }catch(e){} } });
    return ()=>{ try{ supabase.removeChannel(ch); }catch(e){} presRef.current=null; }; },[me]);
  useEffect(()=>{ const ch=presRef.current; if(ch&&me){ try{ ch.track({ id:me.id, name:me.name||me.email, role:me.role, cell: sel?{ id:sel.id, col:sel.col }:null }); }catch(e){} } },[sel]);
  const peerCell={}; peers.forEach(p=>{ if(p.cell) peerCell[p.cell.id+":"+p.cell.col]={ name:p.name, color:colorFor(p.id) }; });
  const peerOn=(id,col)=> peerCell[id+":"+col]||null;
  useEffect(()=>{ const ch=supabase.channel("merch-live")
    .on("postgres_changes",{ event:"*", schema:"public", table:"stage_dates" },(p)=>{ const del=p.eventType==="DELETE"; const n=(p.new&&Object.keys(p.new).length)?p.new:p.old; if(!n||!n.style_id){ setRemoteChanged(true); return; } const sid=n.style_id, stg=n.stage, key=sid+":"+stg; const row=JSON.stringify({ style_id:sid, stage:stg, revised_date:n.revised_date||null, actual_date:n.actual_date||null, reject_date:n.reject_date||null, skip_date:n.skip_date||null }); if(!del && savedRef.current.stg[key]===row) return; const ed=editingRef.current; if(ed&&ed.id===sid&&ed.col===stg){ setRemoteChanged(true); return; } if(del) delete savedRef.current.stg[key]; else savedRef.current.stg[key]=row; setStyles(prev=>{ let found=false; const next=prev.map(s=>{ if(s.id!==sid) return s; found=true; const ns={...s, actuals:{...s.actuals}, revs:{...(s.revs||{})}, rejects:{...(s.rejects||{})}, skips:{...(s.skips||{})} }; if(del){ delete ns.actuals[stg]; delete ns.revs[stg]; delete ns.rejects[stg]; delete ns.skips[stg]; } else { if(n.actual_date) ns.actuals[stg]=n.actual_date; else delete ns.actuals[stg]; if(n.revised_date) ns.revs[stg]=n.revised_date; else delete ns.revs[stg]; if(n.reject_date) ns.rejects[stg]=n.reject_date; else delete ns.rejects[stg]; if(n.skip_date) ns.skips[stg]=n.skip_date; else delete ns.skips[stg]; } return ns; }); if(!found){ setRemoteChanged(true); return prev; } return next; }); })
    .on("postgres_changes",{ event:"*", schema:"public", table:"styles" },(p)=>{ const del=p.eventType==="DELETE"; const n=(p.new&&p.new.id)?p.new:p.old; if(!n||!n.id){ setRemoteChanged(true); return; } const sid=n.id; if(del){ delete savedRef.current.sty[sid]; setStyles(prev=>prev.filter(s=>s.id!==sid)); return; } if(savedRef.current.sty[sid]===JSON.stringify(n)) return; const ed=editingRef.current; if(ed&&ed.id===sid){ setRemoteChanged(true); return; } savedRef.current.sty[sid]=JSON.stringify(n); setStyles(prev=>{ const idx=prev.findIndex(s=>s.id===sid); if(idx===-1) return [...prev, rowToStyle(n,{})]; const cur=prev[idx]; const merged={ ...rowToStyle(n,{}), actuals:cur.actuals, revs:cur.revs, rejects:cur.rejects, skips:cur.skips }; const copy=prev.slice(); copy[idx]=merged; return copy; }); })
    .subscribe();
    return ()=>{ try{ supabase.removeChannel(ch); }catch(e){} }; },[]);
  // SAVE everything to Supabase shortly after any change (debounced)
  useEffect(()=>{ if(firstRender.current){ firstRender.current=false; return; } if(!loadedRef.current) return; const t=setTimeout(async()=>{ try{ setSaveState("saving");
    const SR=savedRef.current;
    // ---- only upsert rows that actually changed since last save: protects other users' concurrent edits ----
    const styRows=styles.map(styleToRow); const styChanged=styRows.filter(r=>SR.sty[r.id]!==JSON.stringify(r));
    if(styChanged.length){ const up1=await supabase.from("styles").upsert(styChanged); if(up1.error) throw up1.error; }
    const stgChanged=[]; styles.forEach(s=> STAGE_KEYS.forEach(k=>{ const row={ style_id:s.id, stage:k, revised_date:(s.revs&&s.revs[k])||null, actual_date:s.actuals[k]||null, reject_date:(s.rejects&&s.rejects[k])||null, skip_date:(s.skips&&s.skips[k])||null }; const key=s.id+":"+k; const j=JSON.stringify(row); if(SR.stg[key]!==j) stgChanged.push({row,key,j}); }));
    if(stgChanged.length){ const up2=await supabase.from("stage_dates").upsert(stgChanged.map(x=>x.row),{ onConflict:"style_id,stage" }); if(up2.error) throw up2.error; }
    const keys=new Set([...Object.keys(fills),...Object.keys(notes)]); const metaChanged=[]; keys.forEach(key=>{ const i=key.indexOf(":"); const row={ style_id:Number(key.slice(0,i)), col:key.slice(i+1), fill:fills[key]||null, note:notes[key]||null }; const j=JSON.stringify(row); if(SR.meta[key]!==j) metaChanged.push({row,key,j}); });
    if(metaChanged.length){ const up3=await supabase.from("cell_meta").upsert(metaChanged.map(x=>x.row),{ onConflict:"style_id,col" }); if(up3.error) throw up3.error; }
    styChanged.forEach(r=>{ SR.sty[r.id]=JSON.stringify(r); }); stgChanged.forEach(x=>{ SR.stg[x.key]=x.j; }); metaChanged.forEach(x=>{ SR.meta[x.key]=x.j; });
    setSaveState("saved"); flash();
  }catch(e){ console.error("save failed",e); setSaveState("error"); } },700); return ()=>clearTimeout(t); },[styles,fills,notes]);

  const [saveState,setSaveState]=useState("idle"); // idle | saving | saved | error
  useEffect(()=>{ if(!loadedRef.current) return; const t=setTimeout(()=>{ supabase.from("app_settings").upsert({ id:"global", data:cfg }).then(()=>{}).catch(()=>{}); },600); return ()=>clearTimeout(t); },[cfg]);
  const [remoteChanged,setRemoteChanged]=useState(false); // another user wrote data
  const flash=()=>{ setSaved(true); clearTimeout(savedTimer.current); savedTimer.current=setTimeout(()=>setSaved(false),1200); };
  const setField=(id,field,val)=>{ pushHistory(); setStyles(prev=>prev.map(s=>{ if(s.id!==id) return s; if(STAGE_KEYS.includes(field)) return { ...s, actuals:{ ...s.actuals, [field]: val||undefined } }; if(field==="qty") return { ...s, qty:Number(val)||0 }; return { ...s, [field]:val }; })); flash(); };
  const setRev=(id,key,val)=>{ pushHistory(); setStyles(prev=>prev.map(s=> s.id===id?{...s,revs:{...(s.revs||{}),[key]:val||undefined}}:s)); flash(); };
  const setReject=(id,key,val)=>{ pushHistory(); setStyles(prev=>prev.map(s=> s.id===id?{...s,rejects:{...(s.rejects||{}),[key]:val||undefined}}:s)); flash(); };
  const setSkip=(id,key,val)=>{ pushHistory(); setStyles(prev=>prev.map(s=>{ if(s.id!==id) return s; const skips={...(s.skips||{})}; const ap=APPR_OF_SEND[key]; if(val){ skips[key]=val; if(ap) skips[ap]=val; } else { skips[key]=undefined; if(ap) skips[ap]=undefined; } return {...s,skips}; })); flash(); };
  const toggleFlag=(id,flag)=>{ pushHistory(); setStyles(prev=>prev.map(s=>s.id===id?{...s,[flag]:!s[flag]}:s)); flash(); };
  const [bulkOpen,setBulkOpen]=useState(false);
  const [bulkResult,setBulkResult]=useState(null); // {inserts, updates, unchanged, sheetName} | {error}
  const toBool=(v)=>{ const s=String(v||"").trim().toLowerCase(); return s==="y"||s==="yes"||s==="true"||s==="1"; };
  const toISO=(v)=>{ if(v==null||v==="") return ""; if(v instanceof Date && !isNaN(v)) return iso(v); const d=parse(String(v)); if(d) return iso(d); const dd=new Date(v); return isNaN(dd)?"":iso(dd); };
  const downloadTemplate=()=>{ const headers=["Style No","Order No","Sample Fit","Family","Colour","Brand","Fabric Type","Junior Owner","Set-Pack ID","Set-Pack Role","Age Group","Order Qty","Order Received","Delivery Date","Fit Req?","Print Req?","S/O Req?","PP Bypass?","Lab Dip Req?","PP Needed?"]; const example=["HSAW26EXAMPLE01","T1","fit 1-3","SWEAT","BLACK","Hopscotch","TERRY","Tamal","","SWEATSHIRT","4-10YRS",400,"2026-05-18","2026-06-25","Y","N","N","Y","Y","Y"]; const ws=XLSX.utils.aoa_to_sheet([headers,example]); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Styles"); XLSX.writeFile(wb,"merch_tracker_upload_template.xlsx"); };
  const parseUpload=async(file)=>{ try{ const buf=await file.arrayBuffer(); const wb=XLSX.read(buf,{cellDates:true}); const sn=wb.SheetNames.includes("Tracker")?"Tracker":wb.SheetNames[0]; const aoa=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:false,cellDates:true,defval:""}); let hr=-1; for(let i=0;i<Math.min(aoa.length,6);i++){ if((aoa[i]||[]).some(c=>{ const n=NORMH(c); return n==="styleno"||n==="style"; })){ hr=i; break; } } if(hr<0) hr=0; const headers=(aoa[hr]||[]).map(NORMH); const fi={}; headers.forEach((h,i)=>{ const f=HEADER_MAP[h]; if(f&&fi[f]==null) fi[f]=i; }); if(fi.styleNo==null){ setBulkResult({ error:"Couldn't find a 'Style No' column. Use the template headers." }); return; }
    const recs=[]; for(let i=hr+1;i<aoa.length;i++){ const row=aoa[i]||[]; const styleNo=String(row[fi.styleNo]??"").trim(); if(!styleNo) continue; const rec={ styleNo }; ["orderNo","sampleFit","family","colour","brand","fabricType","owner","setId","setRole","age"].forEach(f=>{ if(fi[f]!=null){ const v=String(row[fi[f]]??"").trim(); if(v) rec[f]=v; } }); if(fi.qty!=null){ const q=String(row[fi.qty]??"").replace(/[^0-9.]/g,""); if(q) rec.qty=Number(q)||0; } if(fi.ordRec!=null){ const d=toISO(row[fi.ordRec]); if(d) rec.ordRec=d; } if(fi.delivery!=null){ const d=toISO(row[fi.delivery]); if(d) rec.delivery=d; } ["fitReq","printReq","soReq","ppBypass","labDipReq","ppNeeded"].forEach(f=>{ if(fi[f]!=null && String(row[fi[f]]??"").trim()!=="") rec[f]=toBool(row[fi[f]]); }); recs.push(rec); }
    const byNo={}; styles.forEach(s=> byNo[String(s.styleNo).trim().toLowerCase()]=s); const inserts=[], updates=[]; let unchanged=0;
    recs.forEach(rec=>{ const ex=byNo[rec.styleNo.toLowerCase()]; if(ex){ const chg={}; Object.keys(rec).forEach(f=>{ if(f==="styleNo") return; const nv=rec[f]; if(nv===undefined||nv==="") return; if(String(ex[f]??"")!==String(nv)) chg[f]=nv; }); if(Object.keys(chg).length) updates.push({ id:ex.id, styleNo:rec.styleNo, chg }); else unchanged++; } else inserts.push(rec); });
    setBulkResult({ inserts, updates, unchanged, sheetName:sn, total:recs.length });
  }catch(e){ console.error("parse failed",e); setBulkResult({ error:"Couldn't read that file: "+(e.message||e) }); } };
  const applyBulk=async()=>{ if(!bulkResult||bulkResult.error) return; pushHistory(); const { inserts, updates }=bulkResult;
    if(updates.length){ const m={}; updates.forEach(u=>{ m[u.id]={...(m[u.id]||{}),...u.chg}; }); setStyles(prev=>prev.map(s=> m[s.id]?{...s,...m[s.id]}:s)); }
    if(inserts.length){ const rows=inserts.map(rec=>{ const s={ orderNo:rec.orderNo||"NEW", styleNo:rec.styleNo, sampleFit:rec.sampleFit||"", family:rec.family||"", colour:rec.colour||"", brand:rec.brand||"", fabricType:rec.fabricType||"", owner:rec.owner||"", setId:rec.setId||"", setRole:rec.setRole||"", age:rec.age||"", qty:rec.qty||0, ordRec:rec.ordRec||iso(TODAY), delivery:rec.delivery||rec.ordRec||iso(TODAY), fitReq:rec.fitReq??true, printReq:rec.printReq??false, soReq:rec.soReq??false, ppBypass:rec.ppBypass??false, labDipReq:rec.labDipReq??true, ppNeeded:rec.ppNeeded??true, remarks:"" }; const r=styleToRow(s); delete r.id; return r; });
      try{ const { data, error }=await supabase.from("styles").insert(rows).select(); if(error) throw error; if(data) setStyles(prev=>[...prev, ...data.map(d=>rowToStyle(d,{}))]); }catch(e){ console.error("bulk insert failed",e); alert("New styles failed to insert (existing updates were applied): "+(e.message||e)); } }
    setBulkOpen(false); setBulkResult(null); flash(); };
  const BUYER_STAGES=["fitAppr","artAppr","soAppr","labAppr","ppAppr"];
  const runExport=(mode,buf,incBuf)=>{ try{ let data, name, sheet="Tracker"; const B=Math.abs(buf||0);
    if(mode==="buyer"){ const bs=STAGES.filter(st=>BUYER_STAGES.includes(st.key)); data=rows.map(({s,c})=>{ const o={ "Style No":s.styleNo, "Family":s.family, "Colour":s.colour, "Brand":s.brand, "Qty":s.qty, "Delivery":s.delivery, "Status":c.status }; bs.forEach(st=>{ const r=(c.stages||[]).find(x=>x.key===st.key); if(!r){ o[st.label]="\u2014"; return; } const a=(s.actuals&&s.actuals[st.key])?parse(s.actuals[st.key]):null; o[st.label]= a?fmt(a):(r.skipped?"WAIVED":((r.rev||r.plan)?fmt(r.rev||r.plan):"")); }); if(incBuf){ STAGES.filter(st=>!BUYER_STAGES.includes(st.key)).forEach(st=>{ const r=(c.stages||[]).find(x=>x.key===st.key); const lab=st.label+" (plan)"; if(!r){ o[lab]=""; return; } if(r.skipped){ o[lab]="WAIVED"; return; } const a=(s.actuals&&s.actuals[st.key])?parse(s.actuals[st.key]):null; if(a){ o[lab]=fmt(a); return; } const tgt=r.rev||r.plan; o[lab]= tgt?fmt(B?addWorkdays(tgt,-B):tgt):""; }); } o["Proj. Release"]=c.projRelease?fmt(c.projRelease):""; return o; }); name=incBuf?"buyer_plan_buf"+B+"d":"buyer_summary"; sheet="Buyer"; }
    else if(mode==="internal"){ data=rows.map(({s,c})=>{ const o={ "Order No":s.orderNo, "Style No":s.styleNo, "Sample Fit":s.sampleFit, "Family":s.family, "Colour":s.colour, "Qty":s.qty, "Delivery":s.delivery, "Status":c.status }; STAGES.forEach(st=>{ const r=(c.stages||[]).find(x=>x.key===st.key); if(!r){ o[st.label]=""; return; } if(r.skipped){ o[st.label]="WAIVED"; return; } const a=(s.actuals&&s.actuals[st.key])?parse(s.actuals[st.key]):null; if(a){ o[st.label]=fmt(a); return; } const tgt=r.rev||r.plan; if(!tgt){ o[st.label]=r.rejected?"REJECTED":""; return; } const internal=!BUYER_STAGES.includes(st.key); const d=(internal&&B)?addWorkdays(tgt,-B):tgt; o[st.label]=fmt(d); }); return o; }); name="internal_plan_buf"+B+"d"; }
    else { data=rows.map(({s,c})=>{ const o={ "Order No":s.orderNo, "Style No":s.styleNo, "Sample Fit":s.sampleFit, "Family":s.family, "Colour":s.colour, "Brand":s.brand, "Fabric Type":s.fabricType, "Junior":s.owner, "Qty":s.qty, "Order Date":s.ordRec, "Delivery":s.delivery, "Status":c.status }; STAGES.forEach(st=>{ const r=(c.stages||[]).find(x=>x.key===st.key); o[st.label]= (s.actuals&&s.actuals[st.key])? fmt(parse(s.actuals[st.key])) : (r&&r.skipped?"WAIVED":(r&&r.rejected?"REJECTED":"")); }); return o; }); name="merch_tracker"; }
    const ws=XLSX.utils.json_to_sheet(data); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,sheet); XLSX.writeFile(wb,name+"_"+iso(TODAY)+".xlsx"); }catch(e){ alert("Export failed: "+(e.message||e)); } };
  const archiveFiltered=(val)=>{ const ids=new Set(rows.map(r=>r.s.id)); if(!ids.size) return; if(!window.confirm(`${val?"Archive":"Restore"} ${ids.size} style(s)? Archived styles are hidden from the active sheet — this is reversible.`)) return; pushHistory(); setStyles(prev=>prev.map(s=> ids.has(s.id)?{...s,archived:val}:s)); flash(); };
  const deleteStyle=async(id)=>{ if(!window.confirm("Delete this style row? This removes it for everyone and cannot be undone.")) return; pushHistory(); setStyles(prev=>prev.filter(s=>s.id!==id)); flash(); try{ await supabase.from("stage_dates").delete().eq("style_id",id); await supabase.from("cell_meta").delete().eq("style_id",id); await supabase.from("styles").delete().eq("id",id); }catch(e){ console.error("delete failed",e); } };

  const computed=useMemo(()=>styles.map(s=>({s,c:computeStyle(s,cfg)})),[styles,cfg]);
  const todoItems=useMemo(()=>{ const out=[]; const fabByCol={}; computed.forEach(({s,c})=>{ if(c.released) return; const front=c.frontier?[...c.frontier]:[]; front.forEach(key=>{ const r=(c.stages||[]).find(x=>x.key===key); if(!r||r.done) return; const exp=r.rev||r.plan; if(!exp) return; const du=netWorkdays(TODAY,exp); const overdue=TODAY>exp; const win=(cfg.upcoming&&cfg.upcoming[key]!=null)?cfg.upcoming[key]:null; const include = overdue || (win!=null && du<=win); if(!include) return; const branch=BRANCH_OF[key]||""; if(key==="fabricIH"){ const cols=String(s.colour||"").split(/[,/]/).map(x=>x.trim()).filter(Boolean); (cols.length?cols:["(no colour)"]).forEach(col=>{ let cur=fabByCol[col]; if(!cur){ cur=fabByCol[col]={ colour:col, key, label:r.label, owner:r.owner, branch, exp, du, overdue, anyStyle:s.id, anyOrder:s.orderNo, anyJunior:s.owner, count:0 }; } cur.count++; if(exp<cur.exp){ cur.exp=exp; cur.du=du; cur.overdue=overdue; cur.anyStyle=s.id; cur.anyOrder=s.orderNo; cur.anyJunior=s.owner; } }); } else { out.push({ id:s.id, orderNo:s.orderNo, styleNo:s.styleNo, junior:s.owner, colour:s.colour, key, activity:r.label, branch, owner:r.owner, exp, du, overdue }); } }); }); Object.values(fabByCol).forEach(f=> out.push({ id:f.anyStyle, orderNo:f.anyOrder, styleNo:f.colour, junior:f.anyJunior, colour:f.colour, key:f.key, activity:f.label, branch:f.branch, owner:f.owner, exp:f.exp, du:f.du, overdue:f.overdue, isColour:true, count:f.count })); out.sort((a,b)=> (a.overdue!==b.overdue)?(a.overdue?-1:1):((a.exp&&b.exp)?(a.exp-b.exp):0)); return out; },[computed,cfg]);
  const [todoFilter,setTodoFilter]=useState(PF.todoFilter||{});
  useEffect(()=>{ try{ localStorage.setItem("mt_trackfilters", JSON.stringify({ search, statusFilter, ownerFilter, archiveView, activityFilter, colFilters, tab, todoFilter })); }catch(e){} },[search,statusFilter,ownerFilter,archiveView,activityFilter,colFilters,tab,todoFilter]);
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
    if(STAGE_KEYS.includes(col)){ const r=(cc.stages||[]).find(x=>x.key===col); if(!r) return "— n/a"; if(r.done) return fmt(r.actual); if(r.rework) return "↻ Redo & resend"; if(r.rejected) return "✕ Rejected"; return "● Pending"; }
    return "";
  };
  const passCol=(s,c,col,allowed)=>{ if(!allowed||allowed.length===0) return true; if(col==="chase"){ const owners=(c.chaseOwners||[]).map(o=>o.owner); if(owners.length===0) return allowed.includes("(Blanks)"); return owners.some(o=>allowed.includes(o)); } return allowed.includes(valueFor(s,c,col)); };
  const anyFilter = statusFilter!=="All"||ownerFilter!=="All"||!!search||Object.keys(colFilters).length>0||!!activityFilter;
  const resetFilters=()=>{ setStatusFilter("All"); setOwnerFilter("All"); setSearch(""); setColFilters({}); setActivityFilter(null); };
  const snapCurrent=()=>setViewSnap({ statusFilter, ownerFilter, search, colFilters, activityFilter });
  const clearAllFilters=()=>{ resetFilters(); setViewSnap(null); };
  const restoreView=()=>{ if(!viewSnap) return; setStatusFilter(viewSnap.statusFilter); setOwnerFilter(viewSnap.ownerFilter); setSearch(viewSnap.search); setColFilters(viewSnap.colFilters); setActivityFilter(viewSnap.activityFilter); setViewSnap(null); };
  const applyDrill=(spec)=>{ snapCurrent(); setStatusFilter(spec.status||"All"); setOwnerFilter(spec.owner||"All"); setSearch(spec.search||""); setColFilters(spec.colFilters||{}); setActivityFilter(spec.activity||null); setTab("tracker"); };
  const filtered=computed.filter(({s,c})=>{ const q=search.toLowerCase(); const ownerMatch=(c.chaseOwners||[]).some(o=>o.owner.toLowerCase().includes(q)); const matchQ=!q||s.styleNo.toLowerCase().includes(q)||s.colour.toLowerCase().includes(q)||s.family.toLowerCase().includes(q)||s.sampleFit.toLowerCase().includes(q)||s.orderNo.toLowerCase().includes(q)||ownerMatch; const matchS=statusFilter==="All"||(statusFilter==="At Risk"&&(c.tone==="late"||c.tone==="warn"))||(statusFilter==="On Track"&&c.tone==="ok")||(statusFilter==="Released"&&c.released); const matchF=Object.entries(colFilters).every(([col,allowed])=> passCol(s,c,col,allowed)); const matchO=ownerFilter==="All"||(c.chaseOwners||[]).some(o=>o.owner===ownerFilter); const matchA=!activityFilter||(c.frontier&&c.frontier.has(activityFilter)); const matchArch=archiveView==="all"?true:(archiveView==="archived"?!!s.archived:!s.archived); return matchQ&&matchS&&matchF&&matchO&&matchA&&matchArch; });
  const toneRank={ late:0, warn:1, ok:2, done:3, na:4 };
  const fitNum=(s)=>{ const m=String(s.sampleFit).match(/\d+/); return m?Number(m[0]):Infinity; };
  const sortVal=(col,{s,c})=>{ switch(col){ case "__style": return s.styleNo.toLowerCase(); case "orderNo": return (s.orderNo||"~").toLowerCase(); case "sampleFit": return fitNum(s); case "family": return s.family.toLowerCase(); case "colour": return s.colour.toLowerCase(); case "owner": return (s.owner||"").toLowerCase(); case "setId": return (s.setId||"~").toLowerCase(); case "setRole": return (s.setRole||"").toLowerCase(); case "qty": return s.qty; case "ordRec": return s.ordRec?new Date(s.ordRec).getTime():Infinity; case "delivery": return s.delivery?new Date(s.delivery).getTime():Infinity; case "overall": return toneRank[c.tone]; case "fit": return toneRank[c.fitBranch.tone]; case "print": return toneRank[c.printBranch.tone]; case "fabric": return toneRank[c.fabricBranch.tone]; case "pp": return toneRank[c.ppBranch.tone]; case "prod": return toneRank[c.prodFileBranch.tone]; case "fabricCD": return c.fabricCountdown.n==null?Infinity:c.fabricCountdown.n; case "proj": return c.projRelease?c.projRelease.getTime():Infinity; case "pct": return c.pct; case "chase": return (c.chaseOwners||[]).length; case "float": return c.float==null?Infinity:c.float; case "idle": return c.idle==null?-1:c.idle; case "remarks": return (s.remarks||"~").toLowerCase(); default: { const a=s.actuals[col]; return a?new Date(a).getTime():Infinity; } } };
  const rows=useMemo(()=>{ if(!sort.col) return filtered; return [...filtered].sort((A,B)=>{ const a=sortVal(sort.col,A), b=sortVal(sort.col,B); return a<b?-sort.dir:a>b?sort.dir:0; }); },[filtered,sort]);
  const clickHeader=(col)=>{ finishEditing(); setSort(p=> p.col===col?{col,dir:-p.dir}:{col,dir:1}); };

  const visInfo=INFO_COLS.filter(c=>!hidden.has(c.key));
  const visStages=STAGES.filter(s=>!hidden.has(s.key));
  const remarksVis=!hidden.has("remarks");
  const navCols=["__style", ...visInfo.map(c=>c.key), ...visStages.map(s=>s.key), ...(remarksVis?["remarks"]:[])];
  const totalCols=navCols.length;
  const maxFreeze=1+visInfo.length;

  // ---- freeze: cumulative left offsets for frozen leading columns ----
  const frozenCols=navCols.slice(0, Math.min(freezeN, maxFreeze));
  const widthOf=(col)=> colW[col] ?? (col==="__style"?STYLE_W : col==="remarks"?REMARK_COL.w : (INFO_W[col]!==undefined?INFO_W[col]:84));
  const onResize=(col,w)=> setColW(p=>({ ...p, [col]:Math.max(40, Math.round(w)) }));
  const leftOf=(col)=>{ if(col==="__style") return 0; let x=widthOf("__style"); for(const c of frozenCols){ if(c===col) return x; if(c!=="__style") x+=widthOf(c); } return x; };
  const isFrozen=(col)=> frozenCols.includes(col);
  const freezeStyle=(col,bg)=> isFrozen(col)?{ position:"sticky", left:leftOf(col), zIndex: col==="__style"?5:4, background:bg }:{};

  const ownerOfCol=(col)=>{ const st=STAGES.find(s=>s.key===col); if(st) return st.owner; const ic=INFO_COLS.find(c=>c.key===col); if(ic&&ic.owner) return ic.owner; return "Merchant"; };
  const isStageCol=(col)=>STAGE_KEYS.includes(col);
  const fmtTyped=(isoStr)=>{ const d=parse(isoStr); return d?`${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`:""; };
  const parseTyped=(strv)=>{ const t=(strv||"").trim(); if(!t) return ""; let m=/^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t); if(m){ const y=+m[1],mo=+m[2],d=+m[3]; if(mo<1||mo>12||d<1||d>31) return false; return `${y}-${String(mo).padStart(2,"0")}-${String(d).padStart(2,"0")}`; } m=/^(\d{1,2})[\/\-.](\d{1,2})(?:[\/\-.](\d{2,4}))?$/.exec(t); if(m){ let d=+m[1],mo=+m[2],y=m[3]?+m[3]:TODAY.getFullYear(); if(y<100) y+=2000; if(mo<1||mo>12||d<1||d>31) return false; return `${y}-${String(mo).padStart(2,"0")}-${String(d).padStart(2,"0")}`; } return false; };
  const [calOpen,setCalOpen]=useState(false);
  const beginDate=(id,col,mode,initialChar)=>{ if(!canEdit(role,col,mode)) return; setSel({id,col}); setFocus(null); setEditing({id,col,mode}); setCalOpen(false); const s=styles.find(x=>x.id===id); const cur= mode==="rev"?(s&&s.revs&&s.revs[col]):mode==="reject"?(s&&s.rejects&&s.rejects[col]):(isStageCol(col)?(s&&s.actuals[col]):(s&&s[col])); setEditVal(initialChar!=null?initialChar:(cur?fmtTyped(cur):"")); };
  const commitDate=()=>{ if(!editing) return; const r=parseTyped(editVal); if(r!==false){ const val=r===""?null:r; if(editing.mode==="rev") setRev(editing.id,editing.col,val); else if(editing.mode==="reject") setReject(editing.id,editing.col,val); else setField(editing.id,editing.col,val); } setEditing(null); setCalOpen(false); };
  const dateEditor=(id,col,mode)=>{ const s=styles.find(x=>x.id===id); const stored= mode==="rev"?(s&&s.revs&&s.revs[col]):mode==="reject"?(s&&s.rejects&&s.rejects[col]):(isStageCol(col)?(s&&s.actuals[col]):(s&&s[col])); const colLabel=col==="ordRec"?"Order Date":col==="delivery"?"Delivery Date":((STAGES.find(x=>x.key===col)||{}).label||col); const modeLabel=mode==="rev"?"REVISED":mode==="reject"?"REJECTED":"ACTUAL"; const mc=mode==="rev"?"#d97706":mode==="reject"?"#c0392b":"#1d4ed8"; return (<span onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:1, left:1, zIndex:80, display:"flex", flexDirection:"column", gap:1, background:"#fff", border:"1px solid "+mc, padding:"2px 4px", boxShadow:"2px 2px 0 rgba(0,0,0,0.18)" }}><span style={{ fontSize:8, fontWeight:700, color:mc, textTransform:"uppercase", letterSpacing:0.3, whiteSpace:"nowrap" }}>{colLabel} · {modeLabel}</span><span style={{ display:"flex", alignItems:"center", gap:2, position:"relative" }}><input autoFocus onFocus={e=>{ if((e.target.value||"").length>2) e.target.select(); }} value={editVal} placeholder="dd/mm/yyyy" onChange={e=>setEditVal(e.target.value.replace(/[^0-9\/\-. ]/g,""))} onKeyDown={e=>{ e.stopPropagation(); if(e.key==="Enter") commitDate(); else if(e.key==="Escape"){ setEditing(null); setCalOpen(false); } }} onBlur={()=>{ if(!calOpen) commitDate(); }} style={{ width:80, fontFamily:"inherit", fontSize:11, border:"none", outline:"none" }}/><button onMouseDown={e=>e.preventDefault()} onClick={e=>{ e.stopPropagation(); setCalOpen(o=>!o); }} title="calendar" style={{ border:"none", background:"transparent", cursor:"pointer", padding:0, lineHeight:0, fontSize:12 }}>📅</button>{calOpen && <CalPopup label={colLabel+" · "+modeLabel} value={stored} onClose={()=>setCalOpen(false)} onPick={(d)=>{ if(mode==="rev") setRev(id,col,d); else if(mode==="reject") setReject(id,col,d); else setField(id,col,d); setEditing(null); setCalOpen(false); }}/>}</span></span>); };
    const startEdit=(id,col,initialChar)=>{ if(!isEditableCol(col)) return; if(!canEdit(role,col,"actual")) return; if(isDateCol(col)){ beginDate(id,col,"actual"); return; } const s=styles.find(x=>x.id===id); setEditing({id,col,mode:"text"}); if(col==="qty") setEditVal(initialChar??String(s.qty)); else if(col==="__style") setEditVal(initialChar??s.styleNo); else setEditVal(initialChar??(s[col]||"")); };
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
  const scrollToCell=(id,col)=>{ requestAnimationFrame(()=>{ const el=document.getElementById(`cell-${id}-${col}`); if(!el) return; el.scrollIntoView({ block:"nearest" }); const wrap=scrollWrapRef.current; if(wrap){ const cr=el.getBoundingClientRect(), wr=wrap.getBoundingClientRect(); const frozen=STYLE_W+6; if(cr.left < wr.left+frozen){ wrap.scrollLeft -= (wr.left+frozen-cr.left)+8; } else if(cr.right > wr.right){ wrap.scrollLeft += (cr.right-wr.right)+8; } } }); };
  const selectRow=(id)=>{ setSel({id,col:navCols[0]}); setFocus({id,col:navCols[navCols.length-1]}); };
  const selectAll=()=>{ if(!rows.length) return; setSel({id:rows[0].s.id,col:navCols[0]}); setFocus({id:rows[rows.length-1].s.id,col:navCols[navCols.length-1]}); };
  const moveFocus=(dr,dc)=>{ if(!sel) return; const f=focus||sel; let r=rowIndex(f.id)+dr, c=colIndex(f.col)+dc; r=Math.min(Math.max(r,0),rows.length-1); c=Math.min(Math.max(c,0),navCols.length-1); if(rows[r]){ setFocus({id:rows[r].s.id,col:navCols[c]}); scrollToCell(rows[r].s.id,navCols[c]); } };

  const snap=()=>({ styles, fills, notes });
  const pushHistory=()=>{ setPast(p=>[...p.slice(-60), snap()]); setFuture([]); };
  const applySnap=(d)=>{ setStyles(d.styles); setFills(d.fills); setNotes(d.notes); };
  const undo=()=>{ if(!past.length) return; const prev=past[past.length-1]; setFuture(f=>[...f, snap()]); setPast(p=>p.slice(0,-1)); applySnap(prev); flash(); };
  const redo=()=>{ if(!future.length) return; const nx=future[future.length-1]; setPast(p=>[...p, snap()]); setFuture(f=>f.slice(0,-1)); applySnap(nx); flash(); };
  const getVal=(s,col)=>{ if(col==="__style") return s.styleNo; if(col==="qty") return String(s.qty); if(STAGE_KEYS.includes(col)) return s.actuals[col]||""; if(col==="ordRec"||col==="delivery"||TEXT_COLS.includes(col)) return s[col]||""; return null; };
  const doCopy=()=>{ const R=rect(); if(!R) return; const values=[]; let any=false; for(let r=R.r1;r<=R.r2;r++){ const row=[]; for(let c=R.c1;c<=R.c2;c++){ const v=rows[r]?getVal(rows[r].s,navCols[c]):null; if(v!=null&&v!=="") any=true; row.push(v); } values.push(row); } if(any){ setClip({ values, h:values.length, w:values[0].length }); flash(); } };
  const canPasteCell=(s,col)=>{ if(!isEditableCol(col)) return false; if(!canEdit(role,col,"actual")) return false; if(STAGE_KEYS.includes(col)){ const st=STAGES.find(x=>x.key===col); if(!(st.flag===null||s[st.flag])) return false; } return true; };
  const doPaste=()=>{ if(!clip||!sel) return; pushHistory(); const R=rect(); const changes={}; const put=(id,col,val)=>{ (changes[id]=changes[id]||{})[col]=val; };
    if(clip.h===1&&clip.w===1){ const v=clip.values[0][0]; for(let r=R.r1;r<=R.r2;r++){ for(let c=R.c1;c<=R.c2;c++){ const row=rows[r]; const col=navCols[c]; if(row&&canPasteCell(row.s,col)) put(row.s.id,col,v); } } }
    else { for(let i=0;i<clip.h;i++){ for(let j=0;j<clip.w;j++){ const r=R.r1+i, c=R.c1+j; const row=rows[r]; const col=navCols[c]; if(row&&col&&canPasteCell(row.s,col)) put(row.s.id,col,clip.values[i][j]); } } }
    setStyles(prev=>prev.map(s=>{ const ch=changes[s.id]; if(!ch) return s; let ns={...s, actuals:{...s.actuals}}; Object.entries(ch).forEach(([col,val])=>{ if(STAGE_KEYS.includes(col)) ns.actuals[col]=val||undefined; else if(col==="qty") ns.qty=Number(val)||0; else if(col==="__style") ns.styleNo=val; else ns[col]=val; }); return ns; })); flash(); };

  // batch write a {id:{col:val}} change map into styles
  const writeChanges=(changes)=>{ setStyles(prev=>prev.map(s=>{ const ch=changes[s.id]; if(!ch) return s; let ns={...s, actuals:{...s.actuals}, revs:{...(s.revs||{})}}; Object.entries(ch).forEach(([col,val])=>{ if(STAGE_KEYS.includes(col)) ns.actuals[col]=val||undefined; else if(col==="qty") ns.qty=Number(val)||0; else if(col==="__style") ns.styleNo=val; else ns[col]=val; }); return ns; })); flash(); };
  const coerce=(col,raw)=>{ if(raw==null) return ""; const v=String(raw).trim(); if(isDateCol(col)){ const pt=parseTyped(v); if(pt!==false) return pt; const d=new Date(v); return isNaN(d)?"":iso(d); } return v; };
  const clearRange=()=>{ const R=rect(); if(!R) return; pushHistory(); const ch={}; for(let r=R.r1;r<=R.r2;r++){ for(let cc=R.c1;cc<=R.c2;cc++){ const row=rows[r]; const col=navCols[cc]; if(row&&canPasteCell(row.s,col)) (ch[row.s.id]=ch[row.s.id]||{})[col]= STAGE_KEYS.includes(col)?null:(col==="qty"?0:""); } } writeChanges(ch); };
  const applyFillHandle=()=>{ if(!fillFrom||!fillTo) return; const aR=rowIndex(fillFrom.id), aC=colIndex(fillFrom.col), tR=rowIndex(fillTo.id), tC=colIndex(fillTo.col); const r1=Math.min(aR,tR), r2=Math.max(aR,tR), c1=Math.min(aC,tC), c2=Math.max(aC,tC); const srcRow=rows[aR]; if(!srcRow) return; pushHistory(); const ch={}; for(let r=r1;r<=r2;r++){ for(let cc=c1;cc<=c2;cc++){ const row=rows[r]; const col=navCols[cc]; if(!row) continue; if(r===aR&&cc===aC) continue; const srcVal=getVal(srcRow.s, navCols[cc]); if(canPasteCell(row.s,col)) (ch[row.s.id]=ch[row.s.id]||{})[col]=srcVal; } } writeChanges(ch); };
  const fitAllCols=()=>{ const cols=[...visInfo.map(c=>c.key), ...visStages.map(s=>s.key)]; const upd={}; cols.forEach(col=>{ let max=String((INFO_COLS.find(x=>x.key===col)||{}).label||(STAGES.find(x=>x.key===col)||{}).label||col).length; rows.forEach(({s})=>{ const v=getVal(s,col); if(v!=null) max=Math.max(max,String(isDateCol(col)?fmt(parse(v)):v).length); }); upd[col]=Math.max(48,Math.min(320,max*7+26)); }); setColW(p=>({...p,...upd})); flash(); };
  const autoFit=(col)=>{ let max=String(col==="__style"?"Style No":(INFO_COLS.find(x=>x.key===col)?.label||STAGES.find(x=>x.key===col)?.label||col)).length; rows.forEach(({s})=>{ const v=getVal(s,col); if(v!=null) max=Math.max(max,String(isDateCol(col)?fmt(parse(v)):v).length); }); setColW(p=>({ ...p, [col]:Math.max(48, Math.min(320, max*7+26)) })); };
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
  const findReplace=(preview)=>{ const f=frFind; if(!f) return 0; let cells=[]; if(frScope==="selected"){ const R=rect(); if(!R) return 0; for(let r=R.r1;r<=R.r2;r++) for(let cc=R.c1;cc<=R.c2;cc++){ const row=rows[r], col=navCols[cc]; if(row&&col&&FR_COLS.includes(col)) cells.push([row.s,col]); } } else { for(const row of rows) for(const col of FR_COLS) cells.push([row.s,col]); } const esc=f.replace(/[.*+?^${}()|[\]\\]/g,"\\$&"); const re=new RegExp(esc, frCase?"g":"gi"); const repStr=frRepl.replace(/\$/g,"$$$$"); const needle=frCase?f:f.toLowerCase(); let count=0; const ch={}; const seen=new Set(); for(const [s,col] of cells){ const k=s.id+":"+(col==="__style"?"styleNo":col); if(seen.has(k)) continue; seen.add(k); const cur=frGet(s,col); if(!cur) continue; const hay=frCase?cur:cur.toLowerCase(); if(hay.indexOf(needle)===-1) continue; const nv=cur.replace(re,repStr); if(nv===cur) continue; count++; if(!preview && canPasteCell(s,col)) (ch[s.id]=ch[s.id]||{})[col]=nv; } if(!preview && Object.keys(ch).length){ pushHistory(); writeChanges(ch); } return count; };
  const copySpecial=()=>{ if(!sel||!isStageCol(sel.col)){ alert("Select a stage date cell first. Copy-special grabs that cell full state: actual + revised + rejected + skip."); return; } const s=styles.find(x=>x.id===sel.id); if(!s) return; setSpecialClip({ actual:s.actuals[sel.col]||null, rev:(s.revs&&s.revs[sel.col])||null, reject:(s.rejects&&s.rejects[sel.col])||null, skip:(s.skips&&s.skips[sel.col])||null }); flash(); };
  const pasteSpecial=()=>{ if(!sel||!isStageCol(sel.col)||!specialClip) return; const col=sel.col; pushHistory(); setStyles(prev=>prev.map(s=>{ if(s.id!==sel.id) return s; const ns={...s, actuals:{...s.actuals}, revs:{...(s.revs||{})}, rejects:{...(s.rejects||{})}, skips:{...(s.skips||{})} }; if(canEdit(role,col,"actual")){ if(specialClip.actual) ns.actuals[col]=specialClip.actual; else delete ns.actuals[col]; } if(canEditRev(role)){ if(specialClip.rev) ns.revs[col]=specialClip.rev; else delete ns.revs[col]; } if(canEditReject(role,col)){ if(specialClip.reject) ns.rejects[col]=specialClip.reject; else delete ns.rejects[col]; } if(MERCH_ROLES.includes(role)){ if(specialClip.skip) ns.skips[col]=specialClip.skip; else delete ns.skips[col]; } return ns; })); flash(); };
  const copySelection=async()=>{ const R=rect(); if(!R) return; const lines=[]; for(let r=R.r1;r<=R.r2;r++){ const cells=[]; for(let cc=R.c1;cc<=R.c2;cc++){ let v=rows[r]?getVal(rows[r].s,navCols[cc]):""; if(isDateCol(navCols[cc])&&v) v=fmtTyped(v); cells.push(v??""); } lines.push(cells.join("\t")); } const tsv=lines.join("\n"); setClip({ values:lines.map(l=>l.split("\t")), h:lines.length, w:lines[0].split("\t").length }); try{ await navigator.clipboard.writeText(tsv); }catch(err){} flash(); };
  const cellKey=(id,col)=>`${id}:${col}`;
  const applyFill=(color)=>{ if(!sel) return; pushHistory(); const R=rect(); setFills(p=>{ const n={...p}; for(let r=R.r1;r<=R.r2;r++){ for(let c=R.c1;c<=R.c2;c++){ if(!rows[r]) continue; const k=`${rows[r].s.id}:${navCols[c]}`; if(color==="") delete n[k]; else n[k]=color; } } return n; }); flash(); };
  const saveNote=()=>{ if(!sel) return; pushHistory(); setNotes(p=>{ const n={...p}; const k=cellKey(sel.id,sel.col); if(noteText.trim()==="") delete n[k]; else n[k]=noteText.trim(); return n; }); setNoteEditing(false); setNoteText(""); flash(); };
  const beginNote=()=>{ if(!sel) return; setNoteText(notes[cellKey(sel.id,sel.col)]||""); setNoteEditing(true); };

  const distinctFor=(col)=>{ const set=new Set(); computed.forEach(({s,c})=>{ const passOthers=Object.entries(colFilters).every(([cc,allowed])=> cc===col || passCol(s,c,cc,allowed)); if(!passOthers) return; if(col==="chase"){ const owners=(c.chaseOwners||[]).map(o=>o.owner); if(owners.length===0) set.add("(Blanks)"); else owners.forEach(o=>set.add(o)); } else set.add(valueFor(s,c,col)); }); return [...set].sort((a,b)=> a==="(Blanks)"?1:b==="(Blanks)"?-1:(a>b?1:a<b?-1:0)); };
  const filterProps=(col)=>({ filterActive: !!colFilters[col], filterOpen: filterCol===col, filterValues: filterCol===col?distinctFor(col):null, filterAllowed: colFilters[col]||null,
    onToggleFilter:()=>{ finishEditing(); setFilterCol(p=>p===col?null:col); },
    onSetFilter:(arr)=>setColFilters(f=>{ const n={...f}; if(!arr) delete n[col]; else n[col]=arr; return n; }),
    onCloseFilter:()=>setFilterCol(null) });
  const funnel=useMemo(()=>{ const b={ "Pre-Fit":0,"Fit/Print":0,"Lab Dip":0,"Fabric IH":0,"PP":0,"Released":0 }; computed.forEach(({c})=>{ if(c.released) b["Released"]++; else { const k=c.nextPending.key; if(k==="techpack") b["Pre-Fit"]++; else if(["fitSend","fitAppr","artwork","artAppr","strikeOff","soAppr"].includes(k)) b["Fit/Print"]++; else if(["labDip","labAppr"].includes(k)) b["Lab Dip"]++; else if(k==="fabricIH") b["Fabric IH"]++; else b["PP"]++; } }); return b; },[computed]);

  const requiredMissing=()=>{ const m=[]; if(!newRow.styleNo.trim()) m.push("Style No"); if(!newRow.orderNo.trim()) m.push("Order No"); if(!newRow.ordRec) m.push("Order Date"); if(!newRow.delivery) m.push("Delivery Date"); return m; };
  const addNewStyle=async()=>{ const miss=requiredMissing(); if(miss.length){ setNewError("Required: "+miss.join(", ")); return; } setNewError(""); pushHistory();
    const base={ order_no:newRow.orderNo||"", style_no:newRow.styleNo.trim(), sample_fit:newRow.sampleFit||"", family:newRow.family||"", colour:newRow.colour||"", brand:newRow.brand||"", fabric_type:newRow.fabricType||"", owner:newRow.owner||"", set_id:"", set_role:"", age:"", qty:Number(newRow.qty)||0, order_date:newRow.ordRec||iso(TODAY), delivery_date:newRow.delivery||"2026-07-15", fit_req:newRow.fitReq, print_req:newRow.printReq, so_req:newRow.soReq, pp_bypass:newRow.ppBypass, lab_dip_req:newRow.labDipReq, pp_needed:newRow.ppNeeded, remarks:"" };
    try{ const { data, error }=await supabase.from("styles").insert(base).select().single(); if(error||!data) throw error||new Error("no row"); setStyles(prev=>[...prev, rowToStyle(data,{})]); }catch(e){ console.error("create failed",e); }
    setNewRow({ styleNo:"", orderNo:"", sampleFit:"", family:"", colour:"", owner:"", qty:"", ordRec:iso(TODAY), delivery:"", fitReq:true, printReq:false, soReq:false, ppBypass:false, labDipReq:true, ppNeeded:true }); flash(); };

  const isAnchor=(id,col)=> sel&&sel.id===id&&sel.col===col;
  const inRange=(id,col)=> selKeys.has(`${id}:${col}`);
  const inFill=(id,col)=>{ if(!filling||!fillFrom||!fillTo) return false; const aR=rowIndex(fillFrom.id),aC=colIndex(fillFrom.col),tR=rowIndex(fillTo.id),tC=colIndex(fillTo.col); const r=rowIndex(id),cc=colIndex(col); return r>=Math.min(aR,tR)&&r<=Math.max(aR,tR)&&cc>=Math.min(aC,tC)&&cc<=Math.max(aC,tC); };
  const bgFor=(id,col,base)=>{ const f=fills[cellKey(id,col)]; if(f) return f; if(inFill(id,col)) return "#def0e0"; if(focus&&inRange(id,col)&&!isAnchor(id,col)) return "#e3edfb"; return base; };
  const FillHandle=({id,col})=> isAnchor(id,col)&&!editing&&canPasteCell(styles.find(x=>x.id===id),col)? <span onMouseDown={(e)=>{ e.stopPropagation(); e.preventDefault(); setFilling(true); setFillFrom({id,col}); setFillTo({id,col}); }} title="drag to fill" style={{ position:"absolute", right:-2, bottom:-2, width:7, height:7, background:"#1d4ed8", cursor:"crosshair", zIndex:6 }}/> : null;
  const ringFor=(id,col)=> isAnchor(id,col)?"inset 0 0 0 2px #1d4ed8":null;
  const NoteTri=({k})=> notes[k] ? <span title={notes[k]} style={{ position:"absolute", top:0, right:0, width:0, height:0, borderTop:"7px solid #c0392b", borderLeft:"7px solid transparent" }}/> : null;

  const renderEditable=(s,col)=>{
    const k=cellKey(s.id,col.key);
    const val=col.key==="qty"?s.qty:s[col.key];
    const editingThis=editing&&editing.id===s.id&&editing.col===col.key;
    const bg=bgFor(s.id,col.key,"#fff");
    return (
      <td key={col.key} id={`cell-${s.id}-${col.key}`} onClick={(e)=>onCellClick(e,s.id,col.key)} onDoubleClick={(e)=>{ e.stopPropagation(); startEdit(s.id,col.key); }}
        style={{ border:"1px solid #ddd", padding:"6px 9px", whiteSpace: col.key==="remarks"?"normal":"nowrap", boxShadow:ringFor(s.id,col.key), cursor:"cell", maxWidth:col.w, overflow:"hidden", textOverflow:"ellipsis", fontSize: col.key==="remarks"?10:11, color: col.key==="remarks"?"#a15":"#1a1a1a", position:"relative", background:bg, ...freezeStyle(col.key,bg) }}>
        {editingThis ? (<input autoFocus onFocus={e=>{ if((e.target.value||"").length>1) e.target.select(); }} value={editVal} onClick={e=>e.stopPropagation()} onChange={e=>setEditVal(col.key==="qty"?e.target.value.replace(/[^0-9]/g,""):e.target.value)} onBlur={commitText} style={{ width:Math.max(40,col.w-16), fontFamily:"inherit", fontSize:11, border:"1px solid #1d4ed8", outline:"none", padding:"1px 3px" }}/>) : (val===""||val==null ? <span style={{color:"#ccc"}}>—</span> : String(val))}
        <PeerTag who={peerOn(s.id,col.key)}/><NoteTri k={k}/><FillHandle id={s.id} col={col.key}/>
      </td>
    );
  };

  return (
    <div ref={gridRef} tabIndex={0} onKeyDown={onKeyDown} onCopy={handleCopy} onPaste={handlePaste} onMouseDown={(e)=>{ if(editing) return; if(e.target.closest && (e.target.closest("input")||e.target.closest("button")||e.target.closest("th"))) return; const td=e.target.closest && e.target.closest('td[id^="cell-"]'); if(!td) return; const m=td.id.match(/^cell-(\d+)-(.+)$/); if(!m||e.shiftKey) return; e.preventDefault(); setSel({ id:Number(m[1]), col:m[2] }); setFocus(null); selectingRef.current=true; setDragSel(true); }} onMouseUp={()=>{ if(filling){ applyFillHandle(); setFilling(false); setFillFrom(null); setFillTo(null); } if(selectingRef.current){ selectingRef.current=false; setDragSel(false); } }} onMouseOver={(e)=>{ const td=e.target.closest && e.target.closest('td[id^="cell-"]'); if(!td) return; const m=td.id.match(/^cell-(\d+)-(.+)$/); if(!m) return; if(filling){ setFillTo({ id:Number(m[1]), col:m[2] }); return; } if(selectingRef.current){ setFocus({ id:Number(m[1]), col:m[2] }); } }} onClick={()=>{ finishEditing(); setFillOpen(false); setColsOpen(false); setFilterCol(null); setFrOpen(false); setExpOpen(false); }}
      style={{ minHeight:"100vh", background:"#f4f0e8", fontFamily:"'JetBrains Mono', monospace", color:"#1a1a1a", paddingBottom:80, outline:"none" }}>
      <style>{FONT}</style>

      <div style={{ background:"#1a1a1a", color:"#f4f0e8", padding:"14px 22px", display:"flex", alignItems:"center", justifyContent:"space-between", borderBottom:"3px solid #d97706" }}>
        <div style={{ display:"flex", alignItems:"baseline", gap:12 }}><span style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:20, letterSpacing:-0.5 }}>MERCH<span style={{color:"#d97706"}}>·</span>TRACKER</span><span style={{ fontSize:10, color:"#9a958c", letterSpacing:1 }}>PRE-PRODUCTION · SPREADSHEET GRID · PROTOTYPE</span></div>
        <div style={{ display:"flex", alignItems:"center", gap:12 }}><span style={{ fontSize:11, color: saveState==="error"?"#e8746b":saveState==="saving"?"#d9b46a":saveState==="saved"?"#7fd1a8":"#6a665e" }}>{saveState==="error"?"⚠ save failed":saveState==="saving"?"… saving":saveState==="saved"?"● saved to cloud":"○ connected"}</span><span style={{ fontSize:11, color:"#cfc9bf", whiteSpace:"nowrap" }}>{(me&&(me.name||me.email))||""} · <b style={{ color:"#d97706" }}>{(ROLES[role]||{}).label||role}</b></span>{peers.length>0 && (<span style={{ display:"flex", alignItems:"center", gap:3 }} title={peers.map(p=>p.name).join(", ")}>{peers.slice(0,6).map(p=>(<span key={p.id} title={(p.name||"")+" · "+((ROLES[p.role]||{}).label||p.role)} style={{ width:18, height:18, borderRadius:9, background:colorFor(p.id), color:"#fff", fontSize:8, fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center" }}>{initials(p.name)}</span>))}<span style={{ fontSize:9, color:"#9b958a", marginLeft:2 }}>{peers.length} here</span></span>)}{canManageUsers(role) && <button onClick={(e)=>{ e.stopPropagation(); setUsersOpen(true); }} style={{ fontFamily:"inherit", fontSize:10, padding:"5px 9px", cursor:"pointer", border:"1px solid #4a463e", background:"transparent", color:"#cfc9bf" }}>Users</button>}<button onClick={(e)=>{ e.stopPropagation(); onSignOut&&onSignOut(); }} style={{ fontFamily:"inherit", fontSize:10, padding:"5px 9px", cursor:"pointer", border:"1px solid #4a463e", background:"transparent", color:"#cfc9bf" }}>Sign out</button></div>
      </div>

      <div style={{ display:"flex", gap:0, padding:"0 22px", background:"#1a1a1a", borderBottom:"1px solid #3a362e" }}>
        {[["tracker","Tracker"],["dashboard","Dashboard"],["todo","To-Do"],["settings","Settings"]].map(([k,lab])=>(<button key={k} onClick={(e)=>{ e.stopPropagation(); setTab(k); }} style={{ fontFamily:"'Archivo',sans-serif", fontWeight:700, fontSize:12, letterSpacing:0.3, padding:"9px 16px", cursor:"pointer", border:"none", borderBottom:tab===k?"3px solid #d97706":"3px solid transparent", background:"transparent", color:tab===k?"#f4f0e8":"#9a958c" }}>{lab}{k==="todo"&&todoItems.length?` · ${todoItems.length}`:""}</button>))}
      </div>

      {usersOpen && <UsersPanel onClose={()=>setUsersOpen(false)}/>}
      {bulkOpen && (<div onClick={()=>{ setBulkOpen(false); setBulkResult(null); }} style={{ position:"fixed", inset:0, background:"rgba(26,26,26,0.55)", zIndex:200, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
        <div onClick={e=>e.stopPropagation()} style={{ background:"#f4f0e8", border:"2px solid #1a1a1a", boxShadow:"8px 8px 0 #1a1a1a", width:560, maxWidth:"100%", maxHeight:"86vh", overflowY:"auto", padding:22, fontFamily:"'JetBrains Mono',monospace" }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:6 }}><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:18 }}>Bulk upload styles</div><button onClick={()=>{ setBulkOpen(false); setBulkResult(null); }} style={{ border:"none", background:"transparent", cursor:"pointer" }}><X size={18}/></button></div>
          <p style={{ fontSize:11, color:"#666", lineHeight:1.5 }}>Upload an Excel (.xlsx) or CSV. Rows are matched by <b>Style No</b> — existing styles are updated, new ones are added. Nothing is saved until you confirm.</p>
          <button onClick={downloadTemplate} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#fff", marginBottom:12 }}>⬇ Download template</button>
          <div style={{ marginBottom:12 }}><input type="file" accept=".xlsx,.xls,.csv" onChange={e=>{ const f=e.target.files&&e.target.files[0]; if(f) parseUpload(f); }} style={{ fontFamily:"inherit", fontSize:11 }}/></div>
          {bulkResult && bulkResult.error && <div style={{ fontSize:11, color:"#c0392b", fontWeight:700, marginBottom:12 }}>{bulkResult.error}</div>}
          {bulkResult && !bulkResult.error && (<div style={{ marginBottom:12 }}>
            <div style={{ fontSize:12, marginBottom:8 }}>Read <b>{bulkResult.total}</b> rows from sheet "<b>{bulkResult.sheetName}</b>":</div>
            <div style={{ display:"flex", gap:10, marginBottom:10 }}>
              <div style={{ flex:1, border:"1px solid #1a1a1a", background:"#fff", padding:"8px 10px" }}><div style={{ fontSize:22, fontWeight:800, color:"#1f6f54", fontFamily:"'Archivo',sans-serif" }}>{bulkResult.inserts.length}</div><div style={{ fontSize:9, color:"#888", textTransform:"uppercase" }}>new</div></div>
              <div style={{ flex:1, border:"1px solid #1a1a1a", background:"#fff", padding:"8px 10px" }}><div style={{ fontSize:22, fontWeight:800, color:"#d97706", fontFamily:"'Archivo',sans-serif" }}>{bulkResult.updates.length}</div><div style={{ fontSize:9, color:"#888", textTransform:"uppercase" }}>updated</div></div>
              <div style={{ flex:1, border:"1px solid #1a1a1a", background:"#fff", padding:"8px 10px" }}><div style={{ fontSize:22, fontWeight:800, color:"#999", fontFamily:"'Archivo',sans-serif" }}>{bulkResult.unchanged}</div><div style={{ fontSize:9, color:"#888", textTransform:"uppercase" }}>unchanged</div></div>
            </div>
            {bulkResult.updates.length>0 && <div style={{ maxHeight:140, overflowY:"auto", border:"1px solid #ddd", background:"#fff", padding:8, fontSize:10 }}><div style={{ fontWeight:700, marginBottom:4 }}>Changes to existing styles:</div>{bulkResult.updates.slice(0,40).map(u=>(<div key={u.id} style={{ padding:"2px 0", borderBottom:"1px solid #f0ece3" }}><b>{u.styleNo}</b> — {Object.keys(u.chg).join(", ")}</div>))}{bulkResult.updates.length>40 && <div style={{ color:"#999" }}>…and {bulkResult.updates.length-40} more</div>}</div>}
          </div>)}
          {bulkResult && !bulkResult.error && (bulkResult.inserts.length+bulkResult.updates.length>0) && <button onClick={applyBulk} style={{ fontFamily:"inherit", fontSize:12, padding:"9px 16px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#1f6f54", color:"#fff", fontWeight:700 }}>Confirm — apply {bulkResult.inserts.length+bulkResult.updates.length} change(s)</button>}
        </div>
      </div>)}

      {tab==="tracker" && (<>
      <div style={{ display:"flex", padding:"12px 22px 0", flexWrap:"wrap" }}>
        {Object.entries(funnel).map(([k,v],i,arr)=>(<div key={k} style={{ flex:1, minWidth:90, background:"#fff", border:"1px solid #1a1a1a", borderRight:i===arr.length-1?"1px solid #1a1a1a":"none", padding:"8px 10px" }}><div style={{ fontSize:22, fontWeight:700, lineHeight:1, fontFamily:"'Archivo',sans-serif", color:k==="Released"?"#1f6f54":k==="Fabric IH"?"#c0392b":"#1a1a1a" }}>{v}</div><div style={{ fontSize:9, color:"#888", marginTop:3, letterSpacing:0.5, textTransform:"uppercase" }}>{k}</div></div>))}
      </div>

      <div style={{ display:"flex", gap:10, alignItems:"center", padding:"12px 22px 6px", flexWrap:"wrap" }}>
        <div style={{ display:"flex", alignItems:"center", gap:6, background:"#fff", border:"1px solid #1a1a1a", padding:"5px 9px" }}><Filter size={13}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="search style / colour / fit / order…" onClick={e=>e.stopPropagation()} style={{ border:"none", outline:"none", fontFamily:"inherit", fontSize:12, width:180, background:"transparent" }}/></div>
        <div style={{ display:"flex", border:"1px solid #1a1a1a" }}>{["All","At Risk","On Track","Released"].map(f=>(<button key={f} onClick={(e)=>{ e.stopPropagation(); setStatusFilter(f); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"none", borderRight:f!=="Released"?"1px solid #1a1a1a":"none", background:statusFilter===f?"#1a1a1a":"#fff", color:statusFilter===f?"#f4f0e8":"#1a1a1a" }}>{f}</button>))}</div>
        <div style={{ display:"flex", border:"1px solid #1a1a1a" }}>{["All","Merchant","CAD","Buyer","Designer","Mill"].map(f=>(<button key={f} onClick={(e)=>{ e.stopPropagation(); setOwnerFilter(f); }} title="chase owner" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 9px", cursor:"pointer", border:"none", borderRight:f!=="Mill"?"1px solid #1a1a1a":"none", background:ownerFilter===f?"#2563a6":"#fff", color:ownerFilter===f?"#fff":"#1a1a1a" }}>{f==="All"?"Owner":f}</button>))}</div>
        <div style={{ display:"flex", border:"1px solid #1a1a1a" }}>{[["active","Active"],["all","All"],["archived","Archived"]].map(([v,lab])=>(<button key={v} onClick={(e)=>{ e.stopPropagation(); setArchiveView(v); }} title="archived styles are hidden from the live sheet" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 9px", cursor:"pointer", border:"none", borderRight:v!=="archived"?"1px solid #1a1a1a":"none", background:archiveView===v?"#5a6650":"#fff", color:archiveView===v?"#fff":"#1a1a1a" }}>{lab}</button>))}</div>
        {canAdmin(role) && archiveView!=="archived" && anyFilter && <button onClick={(e)=>{ e.stopPropagation(); archiveFiltered(true); }} title="archive the styles currently shown (e.g. a finished season)" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid #5a6650", background:"#fff", color:"#5a6650", fontWeight:700 }}>Archive these ({rows.length})</button>}
        {canAdmin(role) && archiveView==="archived" && rows.length>0 && <button onClick={(e)=>{ e.stopPropagation(); archiveFiltered(false); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid #1f6f54", background:"#fff", color:"#1f6f54", fontWeight:700 }}>Restore these ({rows.length})</button>}
        {canMaster(role) && <button onClick={(e)=>{ e.stopPropagation(); setBulkResult(null); setBulkOpen(true); }} title="bulk upload / update styles from Excel" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#1a1a1a", color:"#f4f0e8", fontWeight:700, display:"flex", alignItems:"center", gap:6 }}><Plus size={13}/> Upload styles</button>}
        <span style={{ position:"relative" }}><button onClick={(e)=>{ e.stopPropagation(); setExpOpen(o=>!o); }} title="export options" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid #1a1a1a", background:expOpen?"#1a1a1a":"#fff", color:expOpen?"#f4f0e8":"#1a1a1a", fontWeight:700, display:"flex", alignItems:"center", gap:6 }}>⬇ Export</button>{expOpen && (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:0, marginTop:4, zIndex:75, background:"#fff", border:"1px solid #1a1a1a", boxShadow:"4px 4px 0 #1a1a1a", padding:12, width:288 }}><div style={{ fontSize:11, fontWeight:700, marginBottom:8 }}>Export {rows.length} filtered styles</div>{[["full","Full \u2014 all columns (actual dates)"],["internal","Internal plan (with buffer)"],["buyer","Buyer view \u2014 key details (printable)"]].map(([v,lbl])=>(<label key={v} style={{ display:"flex", gap:6, fontSize:10, padding:"3px 0", cursor:"pointer", alignItems:"flex-start" }}><input type="radio" checked={expMode===v} onChange={()=>setExpMode(v)} style={{ marginTop:1 }}/>{lbl}</label>))}{expMode==="buyer" && (<label style={{ display:"flex", gap:6, fontSize:10, padding:"4px 0", cursor:"pointer", alignItems:"flex-start", borderTop:"1px solid #eee", marginTop:4, paddingTop:6 }}><input type="checkbox" checked={expIncBuf} onChange={e=>setExpIncBuf(e.target.checked)} style={{ marginTop:1 }}/>Include internal buffered plan dates <span style={{ color:"#999" }}>(for your team only)</span></label>)}{(expMode==="internal"||(expMode==="buyer"&&expIncBuf)) && (<div style={{ fontSize:10, color:"#555", margin:"6px 0 4px", lineHeight:1.5, background:"#f7f4ee", border:"1px solid #e6e0d4", padding:"6px 8px" }}>Pull internal plan dates earlier by <input type="number" min={0} max={30} value={expBuf} onChange={e=>setExpBuf(Math.max(0,Math.min(30,Number(e.target.value)||0)))} style={{ width:40, fontFamily:"inherit", fontSize:11, padding:"2px 4px", border:"1px solid #1a1a1a", margin:"0 4px" }}/> working days.<div style={{ color:"#999", marginTop:4 }}>Buyer approvals (Fit / Art / S-O / Lab / PP Appr) and all actual dates stay unchanged.</div></div>)}<div style={{ display:"flex", gap:8, marginTop:10 }}><button onClick={()=>{ runExport(expMode,expBuf,expIncBuf); setExpOpen(false); }} style={{ flex:1, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:6, cursor:"pointer", border:"1px solid #1a1a1a", background:"#d97706" }}>⬇ Export .xlsx</button><button onClick={()=>setExpOpen(false)} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#f4f0e8" }}><X size={12}/></button></div></div>)}</span>
        <button onClick={(e)=>{ e.stopPropagation(); copySpecial(); }} title="Copy a stage cell full state (actual + revised + rejected + skip)" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#fff", fontWeight:700 }}>⧉ Copy✦</button>
        <button disabled={!specialClip} onClick={(e)=>{ e.stopPropagation(); pasteSpecial(); }} title="Paste the copied full state into the selected stage cell" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:specialClip?"pointer":"not-allowed", border:"1px solid #1a1a1a", background:specialClip?"#1f6f54":"#fff", color:specialClip?"#fff":"#bbb", fontWeight:700, opacity:specialClip?1:0.6 }}>Paste✦</button>
        <span style={{ position:"relative" }}><button onClick={(e)=>{ e.stopPropagation(); setFrOpen(o=>!o); }} title="Find & replace text in cells" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid #1a1a1a", background:frOpen?"#1a1a1a":"#fff", color:frOpen?"#f4f0e8":"#1a1a1a", fontWeight:700 }}>⌕ Find/Replace</button>{frOpen && (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:0, marginTop:4, zIndex:75, background:"#fff", border:"1px solid #1a1a1a", boxShadow:"4px 4px 0 #1a1a1a", padding:12, width:260 }}><div style={{ fontSize:11, fontWeight:700, marginBottom:8 }}>Find &amp; replace</div><label style={{ fontSize:10, color:"#888" }}>Find</label><input autoFocus value={frFind} onChange={e=>setFrFind(e.target.value)} placeholder="text to find" style={{ width:"100%", boxSizing:"border-box", fontFamily:"inherit", fontSize:11, padding:5, marginBottom:6, border:"1px solid #1a1a1a" }}/><label style={{ fontSize:10, color:"#888" }}>Replace with</label><input value={frRepl} onChange={e=>setFrRepl(e.target.value)} placeholder="(leave blank to delete)" style={{ width:"100%", boxSizing:"border-box", fontFamily:"inherit", fontSize:11, padding:5, marginBottom:8, border:"1px solid #1a1a1a" }}/><div style={{ display:"flex", gap:12, fontSize:10, marginBottom:6 }}><label style={{ display:"flex", gap:4, cursor:"pointer" }}><input type="radio" checked={frScope==="filtered"} onChange={()=>setFrScope("filtered")}/>All filtered</label><label style={{ display:"flex", gap:4, cursor:"pointer" }}><input type="radio" checked={frScope==="selected"} onChange={()=>setFrScope("selected")}/>Selected cells</label></div><label style={{ fontSize:10, display:"flex", gap:5, marginBottom:8, cursor:"pointer" }}><input type="checkbox" checked={frCase} onChange={e=>setFrCase(e.target.checked)}/>Match case</label><div style={{ fontSize:10, color:"#666", marginBottom:8, minHeight:13 }}>{frScope==="selected" && !rect() ? "Select cells in the grid first." : (frFind ? findReplace(true)+" cell(s) will change" : "Applies to text fields (style, colour, brand, owner, remarks…)")}</div><div style={{ display:"flex", gap:8 }}><button disabled={!frFind} onClick={(e)=>{ e.stopPropagation(); const n=findReplace(false); flash(); }} style={{ flex:1, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:6, cursor:frFind?"pointer":"not-allowed", border:"1px solid #1a1a1a", background:frFind?"#d97706":"#ccc", opacity:frFind?1:0.6 }}>Replace all</button><button onClick={()=>setFrOpen(false)} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#f4f0e8" }}><X size={12}/></button></div></div>)}</span>
        <select value={activityFilter||""} onChange={e=>{ setActivityFilter(e.target.value||null); }} onClick={e=>e.stopPropagation()} title="show only styles whose current pending action is this activity" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 8px", cursor:"pointer", border:"1px solid "+(activityFilter?"#d97706":"#1a1a1a"), background:activityFilter?"#fff7ec":"#fff", fontWeight:activityFilter?700:400 }}><option value="">Activity: all</option>{STAGES.map(st=>(<option key={st.key} value={st.key}>{st.label}</option>))}</select>
        {anyFilter && <button onClick={(e)=>{ e.stopPropagation(); clearAllFilters(); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid #c0392b", background:"#fff", color:"#c0392b", fontWeight:700, display:"flex", alignItems:"center", gap:5 }}><X size={12}/> clear filters</button>}
        {viewSnap && <button onClick={(e)=>{ e.stopPropagation(); restoreView(); }} title="go back to the view you had before drilling in" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#fff", display:"flex", alignItems:"center", gap:5 }}><RotateCcw size={12}/> restore view</button>}
        <div style={{ position:"relative" }}><button onClick={(e)=>{ e.stopPropagation(); finishEditing(); setFillOpen(o=>!o); setColsOpen(false); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#d97706", color:"#1a1a1a", fontWeight:700, display:"flex", alignItems:"center", gap:6 }}><Copy size={13}/> Fill date → {rows.length}</button>{fillOpen && (<FillPanel count={rows.length} role={role} onClose={()=>setFillOpen(false)} onApply={(key,val,mode)=>{ if(mode==="revised"){ if(!canEditRev(role)){ setFillOpen(false); return; } setStyles(prev=>prev.map(s=>rows.some(r=>r.s.id===s.id)?{...s,revs:{...(s.revs||{}),[key]:val||undefined}}:s)); flash(); setFillOpen(false); return; } if(!canEdit(role,key,"actual")){ setFillOpen(false); return; } const top=(key==="ordRec"||key==="delivery"); setStyles(prev=>prev.map(s=>rows.some(r=>r.s.id===s.id)?(top?{...s,[key]:val||""}:{...s,actuals:{...s.actuals,[key]:val||undefined}}):s)); flash(); setFillOpen(false); }}/>)}</div>
        <div style={{ display:"flex", border:"1px solid #1a1a1a" }} title="data text size"><button onClick={(e)=>{ e.stopPropagation(); bumpScale(-0.1); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 9px", cursor:"pointer", border:"none", borderRight:"1px solid #1a1a1a", background:"#fff", fontWeight:700 }}>A−</button><button onClick={(e)=>{ e.stopPropagation(); bumpScale(0.1); }} style={{ fontFamily:"inherit", fontSize:13, padding:"6px 9px", cursor:"pointer", border:"none", background:"#fff", fontWeight:700 }}>A+</button></div>
        <div style={{ position:"relative" }}><button onClick={(e)=>{ e.stopPropagation(); finishEditing(); setColsOpen(o=>!o); setFillOpen(false); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#fff", display:"flex", alignItems:"center", gap:6 }}><Columns3 size={13}/> Columns {hidden.size>0?`(${hidden.size} hidden)`:""}</button>
          {colsOpen && (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:0, marginTop:4, zIndex:70, background:"#fff", border:"1px solid #1a1a1a", boxShadow:"4px 4px 0 #1a1a1a", padding:10, width:230, maxHeight:300, overflowY:"auto" }}><div style={{ fontSize:10, fontWeight:700, marginBottom:6 }}>Show / hide columns</div><button onClick={(e)=>{ e.stopPropagation(); fitAllCols(); }} style={{ ...chip, width:"100%", marginBottom:8 }}>↔ Auto-fit widths to content</button>{[...INFO_COLS,{key:"remarks",label:"Remarks / Delays"},...STAGES].map(col=>(<label key={col.key} style={{ display:"flex", alignItems:"center", gap:6, fontSize:10, padding:"2px 0", cursor:"pointer" }}><input type="checkbox" checked={!hidden.has(col.key)} onChange={()=>setHidden(p=>{ const n=new Set(p); n.has(col.key)?n.delete(col.key):n.add(col.key); return n; })}/>{col.label}</label>))}<div style={{ fontSize:8, color:"#999", marginTop:8, lineHeight:1.4 }}>Your column view is saved on this device.</div>{hidden.size>0 && <button onClick={()=>setHidden(new Set())} style={{ ...chip, marginTop:6, width:"100%" }}>Reset to default (show all)</button>}</div>)}
        </div>
        {/* FREEZE control */}
        <div style={{ display:"flex", alignItems:"center", gap:5, border:"1px solid #1a1a1a", background:"#fff", padding:"4px 8px" }}>
          <Snowflake size={13} color="#2563a6"/><span style={{ fontSize:10, color:"#666" }}>freeze</span>
          <select value={freezeN} onClick={e=>e.stopPropagation()} onChange={e=>setFreezeN(Number(e.target.value))} style={{ fontFamily:"inherit", fontSize:10, border:"none", outline:"none", background:"transparent", cursor:"pointer" }}>
            <option value={1}>Style only</option>
            {visInfo.slice(0,6).map((c,i)=><option key={c.key} value={i+2}>thru {c.label}</option>)}
          </select>
        </div>
        <div style={{ display:"flex", border:"1px solid #1a1a1a" }}>
          <button onClick={(e)=>{ e.stopPropagation(); undo(); }} disabled={!past.length} title="Undo (Ctrl/Cmd+Z)" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 9px", cursor:past.length?"pointer":"not-allowed", border:"none", borderRight:"1px solid #1a1a1a", background:"#fff", opacity:past.length?1:0.4 }}>↶</button>
          <button onClick={(e)=>{ e.stopPropagation(); redo(); }} disabled={!future.length} title="Redo (Ctrl/Cmd+Shift+Z)" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 9px", cursor:future.length?"pointer":"not-allowed", border:"none", background:"#fff", opacity:future.length?1:0.4 }}>↷</button>
        </div>
        <button onClick={(e)=>{ e.stopPropagation(); setRemoteChanged(false); loadShared(); }} title="reload shared data (pull latest edits)" style={{ fontFamily:"inherit", fontSize:11, fontWeight:remoteChanged?700:400, padding:"6px 11px", cursor:"pointer", border:"1px solid #1a1a1a", background:remoteChanged?"#d97706":"#fff", display:"flex", alignItems:"center", gap:6 }}><RotateCcw size={13}/> {remoteChanged?"Sync · new changes":"Sync"}</button>
        <button onClick={(e)=>{ e.stopPropagation(); setShowAux(v=>!v); }} title="show every cell's auto + revised dates alongside the actual" style={{ fontFamily:"inherit", fontSize:11, fontWeight:showAux?700:400, padding:"6px 11px", cursor:"pointer", border:"1px solid #1a1a1a", background:showAux?"#d97706":"#fff", color:"#1a1a1a", display:"flex", alignItems:"center", gap:6 }}>{showAux?"✓ auto+rev shown":"show auto+rev"}</button>
        <span style={{ fontSize:10, color:"#999", marginLeft:"auto" }}>{sort.col?<>sorted by <b>{sort.col==="__style"?"Style":(INFO_COLS.find(c=>c.key===sort.col)?.label||STAGES.find(s=>s.key===sort.col)?.label||(sort.col==="remarks"?"Remarks":sort.col))}</b> {sort.dir>0?"↑":"↓"}</>:"shift-click / shift-arrows = range · Ctrl/Cmd C & V = copy/paste"}</span>
      </div>

      <div style={{ display:"flex", gap:10, alignItems:"center", padding:"0 22px 10px", flexWrap:"wrap", fontSize:10, color:"#777" }}>
        <button onClick={(e)=>{ e.stopPropagation(); selectAll(); }} title="select all" style={{ fontFamily:"inherit", fontSize:10, padding:"4px 8px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#fff" }}>⌖ all</button>
        <button onClick={(e)=>{ e.stopPropagation(); copySelection(); }} disabled={!sel} title="copy selected cell(s)" style={{ fontFamily:"inherit", fontSize:10, padding:"4px 9px", cursor:sel?"pointer":"not-allowed", border:"1px solid #1a1a1a", background:"#fff", display:"inline-flex", alignItems:"center", gap:5, opacity:sel?1:0.4 }}><Copy size={12}/> copy</button>
        <Droplet size={13}/><span>fill:</span>
        {FILL_SWATCHES.map((sw,i)=>(<button key={i} onClick={(e)=>{ e.stopPropagation(); applyFill(sw); }} disabled={!sel} title={sw===""?"clear fill":sw} style={{ width:18, height:18, cursor:sel?"pointer":"not-allowed", border:"1px solid #1a1a1a", background:sw===""?"#fff":sw, position:"relative", opacity:sel?1:0.4 }}>{sw===""?<X size={11} style={{position:"absolute",top:2,left:2}}/>:null}</button>))}
        <span style={{ marginLeft:10, position:"relative" }}>
          <button onClick={(e)=>{ e.stopPropagation(); beginNote(); }} disabled={!sel} style={{ fontFamily:"inherit", fontSize:10, padding:"4px 9px", cursor:sel?"pointer":"not-allowed", border:"1px solid #1a1a1a", background:"#fff", display:"inline-flex", alignItems:"center", gap:5, opacity:sel?1:0.4 }}><MessageSquare size={12}/> {sel&&notes[cellKey(sel.id,sel.col)]?"edit comment":"add comment"}</button>
          {noteEditing && sel && (<span onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:0, marginTop:4, zIndex:70, background:"#fff", border:"1px solid #1a1a1a", boxShadow:"4px 4px 0 #1a1a1a", padding:8, width:220, display:"block" }}><textarea autoFocus value={noteText} onChange={e=>setNoteText(e.target.value)} placeholder="comment on this cell…" style={{ width:"100%", height:50, fontFamily:"inherit", fontSize:11, border:"1px solid #ccc", outline:"none", resize:"none" }}/><span style={{ display:"flex", gap:6, marginTop:6 }}><button onClick={saveNote} style={{ ...chip, flex:1, background:"#d97706" }}>Save</button><button onClick={()=>{ setNoteEditing(false); setNoteText(""); }} style={chip}>Cancel</button></span></span>)}
        </span>
        <span style={{ display:"inline-flex", alignItems:"center", gap:4, marginLeft:6 }} title="type a cell like A9 and press Enter to jump there"><span style={{ fontSize:9, color:"#888" }}>Go to</span><input value={nameBox} onChange={e=>setNameBox(e.target.value)} onKeyDown={e=>{ if(e.key==="Enter") gotoCell(nameBox); }} onFocus={e=>e.target.select()} placeholder="A9" style={{ width:54, fontFamily:"inherit", fontSize:11, fontWeight:700, textTransform:"uppercase", padding:"3px 6px", border:"1px solid #1a1a1a", textAlign:"center", color:"#d97706" }}/></span><span style={{ marginLeft:8 }}>{clip?<span style={{color:"#2563a6"}}>📋 {clip.h}×{clip.w} copied — select & Ctrl/Cmd+V to paste</span>:(sel?<>selected: <b style={{ color:"#d97706" }}>{cellRef(sel)}{focus&&(focus.id!==sel.id||focus.col!==sel.col)?(":"+cellRef(focus)):""}</b> · {styles.find(s=>s.id===sel.id)?.styleNo} · {sel.col==="__style"?"Style":(INFO_COLS.find(c=>c.key===sel.col)?.label||STAGES.find(s=>s.key===sel.col)?.label||sel.col)}</>:"click a cell to format / comment")}</span>
      </div>

      <div ref={scrollWrapRef} style={{ overflow:"auto", padding:"0 22px", maxHeight:"calc(100vh - 210px)" }}>
        <table role="grid" aria-label="Pre-production tracker grid. Arrow keys to move, Escape to exit, Tab to leave the grid." style={{ borderCollapse:"separate", borderSpacing:0, zoom:textScale, fontSize:11, tableLayout:"fixed", userSelect:dragSel?"none":"auto" }}>
          <colgroup>
            <col style={{ width:widthOf("__style") }}/>
            {visInfo.map(c=><col key={c.key} style={{ width:widthOf(c.key) }}/>)}
            {visStages.map(st=><col key={st.key} style={{ width:widthOf(st.key) }}/>)}
            {remarksVis && <col style={{ width:widthOf("remarks") }}/>}
          </colgroup>
          <thead><tr role="row">
            <Th col="__style" label="Style No" sort={sort} onSort={clickHeader} width={widthOf("__style")} letter={colLetter(colIndex("__style"))} onResize={onResize} onAutoFit={autoFit} scale={textScale} {...filterProps("__style")} sticky left={0} z={6}/>
            {visInfo.map(c=><Th key={c.key} col={c.key} label={c.label} sort={sort} onSort={clickHeader} width={widthOf(c.key)} letter={colLetter(colIndex(c.key))} onResize={onResize} onAutoFit={autoFit} scale={textScale} {...filterProps(c.key)} sticky={isFrozen(c.key)} left={isFrozen(c.key)?leftOf(c.key):undefined} z={5}/>)}
            {visStages.map(st=><Th key={st.key} col={st.key} label={st.label} sort={sort} onSort={clickHeader} width={widthOf(st.key)} letter={colLetter(colIndex(st.key))} onResize={onResize} onAutoFit={autoFit} scale={textScale} {...filterProps(st.key)}/>)}
            {remarksVis && <Th col="remarks" label={REMARK_COL.label} sort={sort} onSort={clickHeader} width={widthOf("remarks")} letter={colLetter(colIndex("remarks"))} onResize={onResize} onAutoFit={autoFit} scale={textScale} {...filterProps("remarks")}/>}
          </tr></thead>
          <tbody>
            {rows.map(({s,c},rowIdx)=>{ const t=TONE_STYLE[c.tone]; const sk=cellKey(s.id,"__style"); const styBg=bgFor(s.id,"__style","#fff"); return (
              <tr key={s.id} role="row">
                <td id={`cell-${s.id}-__style`} onClick={(e)=>onCellClick(e,s.id,"__style")} onDoubleClick={(e)=>{ e.stopPropagation(); startEdit(s.id,"__style"); }} style={{ position:"sticky", left:0, zIndex:5, background:styBg, border:"1px solid #ddd", padding:"6px 9px", overflow:"hidden", boxShadow:ringFor(s.id,"__style"), cursor:"cell" }}>
                  {editing&&editing.id===s.id&&editing.col==="__style" ? (<input autoFocus onFocus={e=>{ if((e.target.value||"").length>1) e.target.select(); }} value={editVal} onClick={e=>e.stopPropagation()} onChange={e=>setEditVal(e.target.value)} onBlur={commitText} style={{ width:150, fontFamily:"inherit", fontSize:11, fontWeight:700, border:"1px solid #1d4ed8", outline:"none", padding:"1px 3px" }}/>) : <div style={{ fontWeight:700, display:"flex", alignItems:"center", gap:6 }}><span onClick={(e)=>{ e.stopPropagation(); selectRow(s.id); }} title="select row" style={{ fontSize:8, color:"#bbb", cursor:"pointer", minWidth:14 }}>{rowIdx+1}</span>{s.styleNo}</div>}
                  <div style={{ fontSize:9, color:"#999", whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis", maxWidth:188, marginTop:1 }}>{s.colour}</div>
                  <div style={{ display:"flex", gap:3, marginTop:4, flexWrap:"wrap" }}>{FLAG_DEFS.map(f=>{ const on=!!s[f.key]; return (<button key={f.key} title={f.title} onClick={(e)=>{ e.stopPropagation(); if(canMaster(role)) toggleFlag(s.id,f.key); }} style={{ fontFamily:"inherit", fontSize:8.5, fontWeight:700, letterSpacing:0.3, padding:"2px 5px", cursor:!canMaster(role)?"not-allowed":"pointer", lineHeight:1.3, border:`1px solid ${on?"#1a1a1a":"#cfcabf"}`, background:on?"#1a1a1a":"transparent", color:on?"#f4f0e8":"#bbb", opacity:!canMaster(role)?0.5:1 }}>{f.short}</button>); })}</div>
                  <NoteTri k={sk}/>
                  {canAdmin(role) && <button title="delete row" onClick={(e)=>{ e.stopPropagation(); deleteStyle(s.id); }} style={{ position:"absolute", top:2, right:2, zIndex:6, border:"none", background:"transparent", cursor:"pointer", padding:0, lineHeight:0, color:"#cbb4ac" }}><Trash2 size={11}/></button>}
                </td>

                {visInfo.map(col=>{
                  if(col.kind==="text"||col.kind==="num") return renderEditable(s,col);
                  const k=cellKey(s.id,col.key);
                  if(col.kind==="date"){ const bg=bgFor(s.id,col.key,"#fff"); return (<td key={col.key} id={`cell-${s.id}-${col.key}`} onClick={(e)=>onCellClick(e,s.id,col.key)} onDoubleClick={(e)=>{ e.stopPropagation(); if(canEdit(role,col.key,"actual")) beginDate(s.id,col.key,"actual"); }} style={{ border:"1px solid #ddd", padding:"6px 9px", whiteSpace:"nowrap", boxShadow:ringFor(s.id,col.key), cursor:"cell", position:"relative", overflow:(editing&&editing.id===s.id&&editing.col===col.key)?"visible":"hidden", background:bg, ...freezeStyle(col.key,bg) }}>{fmt(parse(s[col.key]))||<span style={{color:"#ccc"}}>—</span>}{editing&&editing.id===s.id&&editing.col===col.key && dateEditor(s.id,col.key,editing.mode)}<PeerTag who={peerOn(s.id,col.key)}/><NoteTri k={k}/><FillHandle id={s.id} col={col.key}/></td>); }
                  let content=null;
                  if(col.kind==="branch"){ const b=col.branch==="fit"?c.fitBranch:col.branch==="print"?c.printBranch:col.branch==="fabric"?c.fabricBranch:col.branch==="pp"?c.ppBranch:c.prodFileBranch; const canJump=b.tone!=="na"&&!c.released; content=<BranchPill b={b} onJump={canJump?()=>jumpToEnter(s.id,branchTarget(s,c,col.branch)):null}/>; }
                  else if(col.key==="overall") content=(<span style={{ display:"inline-flex", flexDirection:"column", gap:2, alignItems:"flex-start" }}><span style={{ display:"inline-flex", alignItems:"center", gap:5, background:t.bg, color:t.fg, padding:"2px 7px", fontSize:10, fontWeight:700 }}><span style={{ width:6,height:6,borderRadius:"50%", background:t.dot }}/>{c.status}</span>{c.lastActual && <span style={{ fontSize:8.5, color:"#9a958a", whiteSpace:"nowrap" }}>last: {fmt(c.lastActual)}{c.lastActualKey?` · ${(STAGES.find(x=>x.key===c.lastActualKey)||{}).label||""}`:""}</span>}</span>);
                  else if(col.key==="fabricCD"){ const fc=c.fabricCountdown; content=(<span style={{ display:"inline-flex", flexDirection:"column", lineHeight:1.1 }}><span style={{ fontWeight:700, color:(BR_TONE[fc.tone]||BR_TONE.na).fg }}>{fc.txt}</span>{fc.date && <span style={{ fontSize:8, color:"#9a958a" }}>{fmt(fc.date)}</span>}</span>); }
                  else if(col.key==="proj") content=<span title={`release gate (30wd before delivery): ${fmt(c.releaseGate)}`} style={{ fontWeight:600, color:c.projTone==="late"?"#c0392b":c.projTone==="warn"?"#7a560f":c.projTone==="done"?"#888":"#1f6f54" }}>{fmt(c.projRelease)}{c.projTone==="late"&&!c.released?" ⚠":c.projTone==="ok"?" ✓":""}</span>;
                  else if(col.key==="pct") content=(<div style={{ display:"flex", alignItems:"center", gap:5 }}><div style={{ flex:1, height:6, background:"#eee", position:"relative", minWidth:34 }}><div style={{ position:"absolute", left:0, top:0, bottom:0, width:`${c.pct}%`, background:c.pct===100?"#1f6f54":"#d97706" }}/></div><span style={{ fontSize:9, color:"#666", width:26, textAlign:"right" }}>{c.pct}%</span></div>);
                  else if(col.key==="chase") content=(!c.chaseOwners||c.chaseOwners.length===0)?<span style={{color:"#ccc"}}>—</span>:(<span style={{ display:"flex", gap:3, flexWrap:"wrap" }}>{c.chaseOwners.map(o=>(<span key={o.owner} title={`${o.owner}: ${o.count} open ${o.count>1?"branches":"branch"}`} style={{ fontSize:9, fontWeight:700, padding:"2px 5px", background:(OWNER_COLOR[o.owner]||"#888")+"22", color:OWNER_COLOR[o.owner]||"#555", whiteSpace:"nowrap" }}>{o.owner}{o.count>1?`×${o.count}`:""}</span>))}</span>);
                  else if(col.key==="float") content=<span style={{ fontWeight:700, color:c.float<0?"#c0392b":"#1f6f54" }}>{c.float==null?"—":`${c.float>0?"+":""}${c.float}d`}</span>;
                  else if(col.key==="idle") content=<span style={{ color:c.idle>=7?"#c0392b":"#888" }}>{c.idle==null?"—":`${c.idle}d`}</span>;
                  const bg=bgFor(s.id,col.key,"#fff");
                  return <td key={col.key} id={`cell-${s.id}-${col.key}`} onClick={(e)=>onCellClick(e,s.id,col.key)} style={{ border:"1px solid #ddd", padding:"6px 9px", whiteSpace:"nowrap", boxShadow:ringFor(s.id,col.key), cursor:"default", background:bg, position:"relative", overflow:"hidden", ...freezeStyle(col.key,bg) }}>{content}<PeerTag who={peerOn(s.id,col.key)}/><NoteTri k={k}/><FillHandle id={s.id} col={col.key}/></td>;
                })}

                {visStages.map(st=>{
                  const applies=st.flag===null||s[st.flag];
                  const cs=c.stages.find(x=>x.key===st.key);
                  const isNext=applies && c.frontier && c.frontier.has(st.key);
                  const editable=applies&&canEdit(role,st.key,"actual"); const canRev=applies&&canEditRev(role); const canRej=applies&&canEditReject(role,st.key); const canSkp=applies&&MERCH_ROLES.includes(role);
                  const k=cellKey(s.id,st.key);
                  if(!applies){ const bg=bgFor(s.id,st.key,"#f3f1ec"); return <td key={st.key} id={`cell-${s.id}-${st.key}`} onClick={(e)=>onCellClick(e,s.id,st.key)} style={{ border:"1px solid #ddd", background:bg, color:"#ccc", textAlign:"center", padding:"6px 9px", boxShadow:ringFor(s.id,st.key), position:"relative", overflow:"hidden" }}>—<NoteTri k={k}/></td>; }
                  const hasRev=cs&&cs.rev&&!cs.done;
                  const bg=bgFor(s.id,st.key,(cs&&cs.skipped)?"#f1ead9":(cs&&(cs.rework||cs.rejected))?"#fcecea":(cs&&cs.actual&&cs.histReject?"#fff4f2":(isNext?"#fff7ec":"#fff")));
                  return (
                    <td key={st.key} id={`cell-${s.id}-${st.key}`} onClick={(e)=>onCellClick(e,s.id,st.key)} onDoubleClick={(e)=>{ e.stopPropagation(); if(editable) beginDate(s.id,st.key,"actual"); }}
                      style={{ border:"1px solid #ddd", padding:0, position:"relative", overflow:(editing&&editing.id===s.id&&editing.col===st.key)?"visible":"hidden", background:bg, boxShadow:ringFor(s.id,st.key)||(isNext?"inset 0 0 0 2px #d97706":null), cursor:editable?"cell":"default" }}>
                      <div style={{ minHeight:38, padding:"4px 8px", fontSize:11, color:cs.actual?"#1a1a1a":"#bbb" }}>
                        {showAux && cs.plan && <span style={{ display:"block", fontSize:8, color:"#bcb6a8", lineHeight:1.3 }}>auto {fmt(cs.plan)}{cs.rev?` · rev ${fmt(cs.rev)}`:""}</span>}
                        {cs.skipped ? (
                          <span style={{ display:"flex", flexDirection:"column", lineHeight:1.25 }}>
                            <span style={{ fontSize:9, color:"#8a6d3b", fontWeight:700, display:"flex", alignItems:"center", gap:3 }}><SkipForward size={9}/>SKIPPED</span>
                            <span style={{ fontSize:8, color:"#9a958a" }}>{fmt(cs.skip)}</span>
                          </span>
                        ) : cs.rework ? (
                          <span style={{ display:"flex", flexDirection:"column", lineHeight:1.25 }}>
                            <span style={{ fontSize:9, color:"#b03020", fontWeight:700, display:"flex", alignItems:"center", gap:3 }}><X size={9}/>REDO &amp; RESEND</span>
                            <span style={{ fontSize:9, color:"#7a560f" }}>{hasRev?"→ rev ":"→ "}{fmt(cs.rev||cs.plan)}</span>
                            {editable && <span style={{ fontSize:9, color:"#d97706", fontWeight:700 }}>▸ enter resend</span>}
                          </span>
                        ) : cs.actual ? (<span style={{ display:"flex", flexDirection:"column", lineHeight:1.25 }}><span style={{ display:"flex", alignItems:"center", gap:4 }}><Check size={11} color={OWNER_COLOR[st.owner]}/>{fmt(cs.actual)}</span>{cs.histReject && <span style={{ fontSize:8, color:"#b03020", fontWeight:700 }}>↻ was REJ {fmt(cs.histReject)}</span>}</span>) : cs.rejected ? (
                          <span style={{ display:"flex", flexDirection:"column", lineHeight:1.25 }}>
                            <span style={{ fontSize:9, color:"#b03020", fontWeight:700, display:"flex", alignItems:"center", gap:3 }}><X size={9}/>REJECTED</span>
                            <span style={{ fontSize:9, color:"#b03020" }}>rej {fmt(cs.reject)}</span>
                            <span style={{ fontSize:9, color:"#7a560f" }}>re-appr → {fmt(cs.rev||cs.plan)}</span>
                          </span>
                        ) : (
                          <span style={{ display:"flex", flexDirection:"column", lineHeight:1.2 }}>
                            <span style={{ fontSize:9, color:hasRev?"#6d4aab":isNext?"#d97706":"#c4c0b8" }}>{hasRev?"rev":st.cutoff?"cutoff":"plan"} {fmt(hasRev?cs.rev:cs.plan)}</span>
                            {editable?<span style={{ fontSize:9, color:isNext?"#d97706":"#c4c0b8", fontWeight:isNext?700:400 }}>{isNext?"▸ enter":st.cutoff?"log arrival":"—"}</span>:<span style={{ fontSize:9, color:"#ccc", display:"flex", alignItems:"center", gap:3 }}><Lock size={8}/>locked</span>}
                          </span>
                        )}
                      </div>
                      {canRev && !cs.skipped && (!cs.actual || cs.rework) && (<button title="set revised plan date" onClick={(e)=>{ e.stopPropagation(); beginDate(s.id,st.key,"rev"); }} style={{ position:"absolute", top:3, right:3, border:"none", background:"transparent", cursor:"pointer", padding:0, lineHeight:1, display:"flex" }}><RotateCcw size={11} color="#6d4aab"/></button>)}
                      {canRej && !cs.skipped && !cs.actual && REJECTABLE.includes(st.key) && (<button title={cs.rejected?"clear rejection (remove rework)":"mark REJECTED (log rejection date)"} onClick={(e)=>{ e.stopPropagation(); if(cs.rejected) setReject(s.id,st.key,null); else beginDate(s.id,st.key,"reject"); }} style={{ position:"absolute", top:3, right:20, border:"none", background:cs.rejected?"#b03020":"transparent", borderRadius:2, cursor:"pointer", padding:cs.rejected?2:0, lineHeight:1, display:"flex" }}><X size={cs.rejected?9:11} color={cs.rejected?"#fff":"#b03020"}/></button>)}
                      {canSkp && SKIPPABLE_STAGES.includes(st.key) && !cs.actual && (<button title={cs.skipped?"un-skip (restore this activity)":"skip this activity (waive — counts as resolved, not done)"} onClick={(e)=>{ e.stopPropagation(); if(cs.skipped){ setSkip(s.id,st.key,null); } else if(window.confirm(`Skip / waive "${st.label}" for ${s.styleNo}?\n\nIt will count as RESOLVED (not done) and drop off the to-do. You can un-skip later.`)){ setSkip(s.id,st.key,iso(TODAY)); } }} style={{ position:"absolute", bottom:3, right:3, border:"none", background:cs.skipped?"#8a6d3b":"transparent", borderRadius:2, cursor:"pointer", padding:cs.skipped?2:0, lineHeight:1, display:"flex" }}><SkipForward size={cs.skipped?9:12} color={cs.skipped?"#fff":"#b8a98a"}/></button>)}
                      {cs.rework && canRej && (<button title="clear rework (un-reject the approval)" onClick={(e)=>{ e.stopPropagation(); setReject(s.id, APPR_OF_SEND[st.key], null); }} style={{ position:"absolute", top:3, right:20, border:"none", background:"#b03020", borderRadius:2, cursor:"pointer", padding:2, lineHeight:1, display:"flex" }}><X size={9} color="#fff"/></button>)}
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
                <td style={{ position:"sticky", left:0, zIndex:5, background:"#fbf7ee", border:"1px dashed #d97706", padding:"6px 9px" }}><div style={{ display:"flex", alignItems:"center", gap:5 }}><Plus size={13} color="#d97706"/><input value={newRow.styleNo} onClick={e=>e.stopPropagation()} onChange={e=>setNewRow(n=>({...n,styleNo:e.target.value}))} onKeyDown={e=>e.key==="Enter"&&addNewStyle()} placeholder="new style no… *" style={{ border:"none", outline:"none", background:"transparent", fontFamily:"inherit", fontSize:11, fontWeight:700, width:120 }}/></div></td>
                <td colSpan={totalCols-1} style={ndCell}>
                  <span style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" }}>
                    <input value={newRow.orderNo} onClick={e=>e.stopPropagation()} onChange={e=>setNewRow(n=>({...n,orderNo:e.target.value}))} placeholder="order# *" style={ndInput(54)}/>
                    <input value={newRow.sampleFit} onClick={e=>e.stopPropagation()} onChange={e=>setNewRow(n=>({...n,sampleFit:e.target.value}))} placeholder="fit#" style={ndInput(50)}/>
                    <input value={newRow.family} onClick={e=>e.stopPropagation()} onChange={e=>setNewRow(n=>({...n,family:e.target.value}))} placeholder="family" style={ndInput(110)}/>
                    <input value={newRow.colour} onClick={e=>e.stopPropagation()} onChange={e=>setNewRow(n=>({...n,colour:e.target.value}))} placeholder="colour" style={ndInput(130)}/>
                    <input value={newRow.owner} onClick={e=>e.stopPropagation()} onChange={e=>setNewRow(n=>({...n,owner:e.target.value}))} placeholder="owner" style={ndInput(56)}/>
                    <input value={newRow.qty} onClick={e=>e.stopPropagation()} onChange={e=>setNewRow(n=>({...n,qty:e.target.value.replace(/[^0-9]/g,"")}))} placeholder="qty" style={ndInput(40)}/>
                    <span style={{ color:"#999", fontSize:9 }}>ord</span><input type="date" value={newRow.ordRec} onClick={e=>e.stopPropagation()} onChange={e=>setNewRow(n=>({...n,ordRec:e.target.value}))} style={ndInput(80)}/>
                    <span style={{ color:"#999", fontSize:9 }}>del</span><input type="date" value={newRow.delivery} onClick={e=>e.stopPropagation()} onChange={e=>setNewRow(n=>({...n,delivery:e.target.value}))} style={ndInput(80)}/>
                    {FLAG_DEFS.map(f=>{ const on=!!newRow[f.key]; return (<button key={f.key} title={f.title} onClick={(e)=>{ e.stopPropagation(); setNewRow(n=>({...n,[f.key]:!n[f.key]})); }} style={{ fontFamily:"inherit", fontSize:8, fontWeight:700, padding:"2px 4px", cursor:"pointer", lineHeight:1.3, border:`1px solid ${on?"#1a1a1a":"#cfcabf"}`, background:on?"#1a1a1a":"transparent", color:on?"#f4f0e8":"#bbb" }}>{f.short}</button>); })}
                    <button onClick={(e)=>{ e.stopPropagation(); addNewStyle(); }} style={{ fontFamily:"inherit", fontSize:11, fontWeight:700, padding:"5px 14px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#d97706", color:"#1a1a1a" }}>+ Create (Enter)</button>{newError && <span style={{ fontSize:10, color:"#c0392b", fontWeight:700, marginLeft:8 }}>{newError}</span>}
                  </span>
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      <div style={{ padding:"16px 22px", display:"flex", gap:18, flexWrap:"wrap", fontSize:10, color:"#777" }}>
        <span style={{ display:"flex", alignItems:"center", gap:5 }}><span style={{ width:14,height:14, boxShadow:"inset 0 0 0 2px #1d4ed8", display:"inline-block" }}/> active cell · <span style={{ width:14,height:14, background:"#e3edfb", display:"inline-block", marginLeft:4 }}/> range (shift-click / shift-arrows)</span>
        <span><b>Excel keys:</b> Ctrl/Cmd C/V copy-paste (works with Excel too) · Ctrl/Cmd Z / Shift+Z undo-redo · F2 edit · Del clears range · drag the blue corner to fill down · double-click a header edge to auto-fit · click a row number to select the row.</span>
        <span style={{ display:"flex", alignItems:"center", gap:5 }}><Snowflake size={11} color="#2563a6"/> freeze leading columns</span>
        <span style={{ display:"flex", alignItems:"center", gap:5 }}><RotateCcw size={11} color="#6d4aab"/> set REVISED plan (incl. Fabric IH) — plans cascade from it</span>
        <span style={{ display:"flex", alignItems:"center", gap:5 }}><span style={{ width:0,height:0, borderTop:"8px solid #c0392b", borderLeft:"8px solid transparent", display:"inline-block" }}/> comment</span>
      </div>
      </>)}

      {tab==="dashboard" && <DashboardView computed={computed} todoItems={todoItems} applyDrill={applyDrill} drillTodo={(obj)=>{ setTodoFilter(obj); setTab("todo"); }}/>}
      {tab==="todo" && <TodoView items={todoItems} filter={todoFilter} setFilter={setTodoFilter} onJump={(id,key)=>{ snapCurrent(); resetFilters(); setTab("tracker"); requestAnimationFrame(()=>setTimeout(()=>jumpToEnter(id,key),60)); }}/>}
      {tab==="settings" && <SettingsView cfg={cfg} setCfg={setCfg} canEdit={canAdmin(role)}/>}
    </div>
  );
}

function Th({ col, label, sort, onSort, sticky, left, z, width, onResize, onAutoFit, scale, letter, filterActive, filterOpen, filterValues, filterAllowed, onToggleFilter, onSetFilter, onCloseFilter }){
  const active=sort.col===col;
  const startDrag=(e)=>{ e.preventDefault(); e.stopPropagation(); const sx=e.clientX, sw=width||80; const sc=scale||1; const move=(ev)=>onResize&&onResize(col, sw+(ev.clientX-sx)/sc); const up=()=>{ window.removeEventListener("mousemove",move); window.removeEventListener("mouseup",up); }; window.addEventListener("mousemove",move); window.addEventListener("mouseup",up); };
  return (<th role="columnheader" aria-sort={active?(sort.dir>0?"ascending":"descending"):"none"} style={{ position:"sticky", top:0, left:sticky?left:undefined, zIndex:sticky?(z||5):3, background:active?"#d97706":"#1a1a1a", color:active?"#1a1a1a":"#f4f0e8", padding:"8px 9px", textAlign:"left", fontWeight:600, fontSize:9.5, letterSpacing:0.4, textTransform:"uppercase", whiteSpace:"nowrap", overflow:"visible", border:"1px solid #3a362e", userSelect:"none" }}>
    <span style={{ display:"flex", alignItems:"center", gap:3 }}>
      <span onClick={(e)=>{ e.stopPropagation(); onSort(col); }} title="click to sort" style={{ display:"inline-flex", alignItems:"center", gap:3, cursor:"pointer", flex:1, overflow:"hidden", textOverflow:"ellipsis" }}>{letter && <span style={{ fontSize:8, fontWeight:700, opacity:0.5, marginRight:4 }}>{letter}</span>}{label}{active?(sort.dir>0?<ChevronUp size={11}/>:<ChevronDown size={11}/>):null}</span>
      <span onClick={(e)=>{ e.stopPropagation(); onToggleFilter&&onToggleFilter(); }} title={filterActive?"filter ON — click to edit/clear":"filter"} style={{ cursor:"pointer", display:"inline-flex", padding:"0 1px", color: filterActive?(active?"#1a1a1a":"#f4b942"):(active?"#7a4a08":"#cfc9bf") }}><Filter size={filterActive?12:10} fill={filterActive?"currentColor":"none"}/></span>
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
    <div onClick={e=>e.stopPropagation()} style={{ position:"fixed", top:pos?pos.top:-9999, left:pos?pos.left:-9999, zIndex:300, background:"#fff", color:"#1a1a1a", border:"1px solid #1a1a1a", boxShadow:"4px 4px 0 #1a1a1a", padding:8, width:210, textTransform:"none", letterSpacing:0, fontWeight:400, maxHeight:"80vh", overflowY:"auto" }}>
      <input autoFocus value={q} onClick={e=>e.stopPropagation()} onKeyDown={e=>e.stopPropagation()} onChange={e=>setQ(e.target.value)} placeholder="search values…" style={{ width:"100%", fontFamily:"inherit", fontSize:11, padding:"4px 6px", border:"1px solid #ccc", outline:"none", marginBottom:6 }}/>
      <label style={{ display:"flex", alignItems:"center", gap:6, fontSize:10, fontWeight:700, padding:"3px 0", cursor:"pointer", borderBottom:"1px solid #eee", marginBottom:4 }}><input ref={masterRef} type="checkbox" checked={allOn} onChange={q?selectResults:toggleAll}/>{q?"(Select matches)":"(Select All)"}</label>
      <div style={{ maxHeight:180, overflowY:"auto" }}>
        {shown.map(v=>(<div key={v} style={{ display:"flex", alignItems:"center", gap:6, fontSize:10, padding:"2px 0" }}><input type="checkbox" checked={isOn(v)} onChange={()=>toggle(v)} style={{ cursor:"pointer" }}/><span onClick={()=>toggle(v)} style={{ flex:1, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis", cursor:"pointer" }}>{v}</span><button onClick={(e)=>{ e.stopPropagation(); onSet([v]); }} title="show only this" style={{ fontSize:8, border:"1px solid #ccc", background:"#f4f0e8", cursor:"pointer", padding:"1px 5px", color:"#666" }}>only</button></div>))}
        {shown.length===0 && <div style={{ fontSize:10, color:"#999", padding:"4px 0" }}>no matches</div>}
      </div>
      {noneOn && <div style={{ fontSize:9, color:"#c0392b", marginTop:4 }}>Nothing selected — no rows shown.</div>}
      <div style={{ display:"flex", gap:6, marginTop:6 }}>
        <button onClick={()=>onSet(null)} style={{ ...chip, flex:1, fontSize:9 }}>Clear filter</button>
        <button onClick={onClose} style={{ ...chip, flex:1, fontSize:9, background:"#1a1a1a", color:"#f4f0e8" }}>Done</button>
      </div>
    </div>
  );
  return (<><span ref={anchorRef} style={{ position:"absolute", width:0, height:0 }}/>{createPortal(menu, document.body)}</>);
}
function FillPanel({ count, role, onApply, onClose }){
  const [mode,setMode]=useState("actual");
  const stageOpts=STAGES.map(s=>({key:s.key,label:s.label+(s.cutoff?" (Fabric IH)":"")}));
  const opts = mode==="revised" ? (canEditRev(role)?stageOpts:[]) : [{key:"ordRec",label:"Order Date"},{key:"delivery",label:"Delivery Date"}].concat(stageOpts).filter(o=>canEdit(role,o.key,"actual"));
  const [key,setKey]=useState(opts[0]?opts[0].key:"labAppr"); const [val,setVal]=useState(iso(TODAY));
  useEffect(()=>{ if(opts.length && !opts.some(o=>o.key===key)) setKey(opts[0].key); },[mode]);
  const accent = mode==="revised"?"#d97706":"#1d4ed8";
  const tabBtn=(m,label)=>(<button onClick={()=>setMode(m)} style={{ flex:1, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:"5px 0", cursor:"pointer", border:"1px solid #1a1a1a", background:mode===m?(m==="revised"?"#d97706":"#1d4ed8"):"#fff", color:mode===m?"#fff":"#1a1a1a" }}>{label}</button>);
  return (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:0, marginTop:4, zIndex:70, background:"#fff", border:"1px solid #1a1a1a", boxShadow:"4px 4px 0 #1a1a1a", padding:12, width:280 }}><div style={{ fontSize:11, fontWeight:700, marginBottom:8 }}>Set one date across {count} filtered styles</div><div style={{ display:"flex", marginBottom:8 }}>{tabBtn("actual","Actual")}{tabBtn("revised","Revised")}</div>{opts.length===0 ? <div style={{ fontSize:10, color:"#c0392b", marginBottom:8 }}>Your role cannot set revised dates.</div> : (<><label style={{ fontSize:10, color:"#888" }}>Stage</label><select value={key} onChange={e=>setKey(e.target.value)} style={{ width:"100%", fontFamily:"inherit", fontSize:11, padding:5, marginBottom:8, border:"1px solid #1a1a1a" }}>{opts.map(o=><option key={o.key} value={o.key}>{o.label}</option>)}</select><label style={{ fontSize:10, color:"#888" }}>{mode==="revised"?"Revised plan date":"Actual date"}</label><input type="date" value={val} onChange={e=>setVal(e.target.value)} style={{ width:"100%", fontFamily:"inherit", fontSize:11, padding:5, marginBottom:10, border:"1px solid #1a1a1a" }}/></>)}<div style={{ display:"flex", gap:8 }}><button disabled={opts.length===0} onClick={()=>onApply(key,val,mode)} style={{ flex:1, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:6, cursor:opts.length?"pointer":"not-allowed", border:"1px solid #1a1a1a", background:opts.length?accent:"#ccc", color:"#fff", opacity:opts.length?1:0.6 }}>Apply {mode} → {count}</button><button onClick={onClose} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#f4f0e8" }}><X size={12}/></button></div></div>);
}
const ndCell={ border:"1px dashed #e8dcc2", padding:"6px 9px", whiteSpace:"nowrap" };
const ndInput=(w)=>({ border:"none", outline:"none", background:"transparent", fontFamily:"'JetBrains Mono', monospace", fontSize:10, width:w });

/* ========================= DASHBOARD ========================= */
function DashboardView({ computed, todoItems, applyDrill, drillTodo }){
  const [target,setTarget]=useState("tracker"); // where bar/owner/activity drills go
  const [df,setDf]=useState(()=>{ try{ return JSON.parse(localStorage.getItem("mt_dashfilter")||"{}"); }catch(e){ return {}; } });
  useEffect(()=>{ try{ localStorage.setItem("mt_dashfilter", JSON.stringify(df)); }catch(e){} },[df]);
  const matchDf=(st,except)=> (except==="order"||!df.order||st.orderNo===df.order) && (except==="fit"||!df.fit||st.sampleFit===df.fit) && (except==="junior"||!df.junior||st.owner===df.junior) && (except==="family"||!df.family||st.family===df.family) && (except==="brand"||!df.brand||st.brand===df.brand) && (except==="fabric"||!df.fabric||st.fabricType===df.fabric) && (except==="colour"||!df.colour||String(st.colour||"").split(/[,/]/).map(x=>x.trim()).includes(df.colour));
  const distinctC=(key,fn)=>{ const s=new Set(); computed.forEach(({s:st})=>{ if(!matchDf(st,key)) return; fn(st).forEach(v=>{ if(v) s.add(v); }); }); return [...s].sort(); };
  const orders=distinctC("order",s=>[s.orderNo]); const fits=distinctC("fit",s=>[s.sampleFit]); const juniors=distinctC("junior",s=>[s.owner]); const families=distinctC("family",s=>[s.family]); const brands=distinctC("brand",s=>[s.brand]); const fabrics=distinctC("fabric",s=>[s.fabricType]);
  const colours=distinctC("colour",s=>String(s.colour||"").split(/[,/]/).map(x=>x.trim()));
  const fc=computed.filter(({s})=> (!df.order||s.orderNo===df.order) && (!df.fit||s.sampleFit===df.fit) && (!df.junior||s.owner===df.junior) && (!df.family||s.family===df.family) && (!df.brand||s.brand===df.brand) && (!df.fabric||s.fabricType===df.fabric) && (!df.colour||String(s.colour||"").split(/[,/]/).map(x=>x.trim()).includes(df.colour)) );
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
  const phase={ "Pre-Fit":0,"Fit / Print":0,"Lab Dip":0,"Fabric IH":0,"PP / Prod":0 };
  fc.forEach(({c})=>{ if(c.released) return; const k=c.nextPending&&c.nextPending.key; if(k==="techpack") phase["Pre-Fit"]++; else if(["fitSend","fitAppr","artwork","artAppr","strikeOff","soAppr"].includes(k)) phase["Fit / Print"]++; else if(["labDip","labAppr"].includes(k)) phase["Lab Dip"]++; else if(k==="fabricIH") phase["Fabric IH"]++; else phase["PP / Prod"]++; });
  const maxPhase=Math.max(1,...Object.values(phase));
  const OWNER_COLOR2={ Merchant:"#1f6f54", CAD:"#2563a6", Buyer:"#b4531a", Designer:"#6d4aab", Mill:"#7a5a1e" };
  // splice carried into drills so the tracker shows the same slice
  const spliceCols=()=>{ const cf={}; if(df.order) cf.orderNo=[df.order]; if(df.fit) cf.sampleFit=[df.fit]; if(df.junior) cf.owner=[df.junior]; if(df.family) cf.family=[df.family]; if(df.brand) cf.brand=[df.brand]; if(df.fabric) cf.fabricType=[df.fabric]; return cf; };
  const spliceSearch=()=> df.colour||"";
  const goOwner=(o)=>{ if(target==="todo") drillTodo({ owner:o }); else applyDrill({ owner:o, colFilters:spliceCols(), search:spliceSearch() }); };
  const goAct=(label,key)=>{ if(target==="todo") drillTodo({ activity:label }); else applyDrill({ activity:key, colFilters:spliceCols(), search:spliceSearch() }); };
  const goStatus=(st,extra)=>applyDrill({ status:st, colFilters:{...spliceCols(),...(extra||{})}, search:spliceSearch() });
  const card=(label,val,color,onClick)=>(<button onClick={onClick} disabled={!onClick} style={{ flex:1, minWidth:130, textAlign:"left", background:"#fff", border:"1px solid #1a1a1a", padding:"14px 16px", cursor:onClick?"pointer":"default", fontFamily:"inherit" }}><div style={{ fontSize:28, fontWeight:800, fontFamily:"'Archivo',sans-serif", color, lineHeight:1 }}>{val}</div><div style={{ fontSize:10, color:"#888", marginTop:5, letterSpacing:0.5, textTransform:"uppercase" }}>{label}{onClick?" ›":""}</div></button>);
  const sel=(label,val,opts,onChange)=>(<select value={val||""} onChange={e=>onChange(e.target.value||undefined)} style={{ fontFamily:"inherit", fontSize:11, padding:"5px 7px", border:"1px solid #1a1a1a", background:val?"#fff7ec":"#fff", maxWidth:150 }}><option value="">{label}: all</option>{opts.map(o=>(<option key={o} value={o}>{o}</option>))}</select>);
  const anyDf=Object.values(df).some(Boolean);
  const bar=(items,maxV,colorFn,labelW,onClick,fmtR)=> items.length===0?<div style={{ fontSize:11, color:"#999" }}>Nothing pending.</div>:items.map(([k,v])=>{ const n=typeof v==="number"?v:v.n; return (
    <button key={k} onClick={()=>onClick(k,v)} style={{ display:"flex", alignItems:"center", gap:8, width:"100%", border:"none", background:"transparent", cursor:"pointer", fontFamily:"inherit", padding:"4px 0" }}>
      <span style={{ width:labelW, fontSize:10, fontWeight:700, color:colorFn(k), textAlign:"left" }}>{k}</span>
      <span style={{ flex:1, height:16, background:"#f0ece3", position:"relative" }}><span style={{ position:"absolute", left:0, top:0, bottom:0, width:`${(n/maxV)*100}%`, background:colorFn(k,v) }}/></span>
      <span style={{ width:54, textAlign:"right", fontSize:10, fontWeight:700 }}>{fmtR?fmtR(v):n}</span>
    </button>); });
  return (<div style={{ padding:"16px 22px", maxWidth:1140 }}>
    {/* splice filter bar */}
    <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"center", marginBottom:14 }}>
      <span style={{ fontSize:10, fontWeight:700, color:"#888", textTransform:"uppercase", letterSpacing:0.5 }}>Slice:</span>
      {sel("Order",df.order,orders,v=>setDf(d=>({...d,order:v})))}
      {sel("Fit",df.fit,fits,v=>setDf(d=>({...d,fit:v})))}
      {sel("Colour",df.colour,colours,v=>setDf(d=>({...d,colour:v})))}
      {sel("Junior",df.junior,juniors,v=>setDf(d=>({...d,junior:v})))}
      {sel("Family",df.family,families,v=>setDf(d=>({...d,family:v})))}
      {sel("Brand",df.brand,brands,v=>setDf(d=>({...d,brand:v})))}
      {sel("Fabric",df.fabric,fabrics,v=>setDf(d=>({...d,fabric:v})))}
      {anyDf && <button onClick={()=>setDf({})} style={{ fontFamily:"inherit", fontSize:10, padding:"5px 9px", cursor:"pointer", border:"1px solid #c0392b", background:"#fff", color:"#c0392b", fontWeight:700 }}>clear slice</button>}
      <span style={{ marginLeft:"auto", display:"flex", alignItems:"center", gap:6 }}><span style={{ fontSize:10, color:"#888" }}>drill to:</span><div style={{ display:"flex", border:"1px solid #1a1a1a" }}>{["tracker","todo"].map(t=>(<button key={t} onClick={()=>setTarget(t)} style={{ fontFamily:"inherit", fontSize:10, fontWeight:700, padding:"5px 10px", cursor:"pointer", border:"none", borderRight:t==="tracker"?"1px solid #1a1a1a":"none", background:target===t?"#1a1a1a":"#fff", color:target===t?"#f4f0e8":"#1a1a1a" }}>{t==="tracker"?"Tracker":"To-Do"}</button>))}</div></span>
    </div>

    <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
      {card("Styles (slice)",total,"#1a1a1a")}
      {card("On track",onTrack,"#1f6f54",()=>goStatus("On Track"))}
      {card("At risk",atRisk,"#c0392b",()=>goStatus("At Risk"))}
      {card("Delivery risk",delRisk,"#c0392b",()=>goStatus("All",{ overall:["Delivery risk"] }))}
      {card("Released",released,"#1f6f54",()=>goStatus("Released"))}
      {card("Overdue activities",overdueAct,"#c0392b")}
    </div>

    <div style={{ display:"flex", gap:18, flexWrap:"wrap", marginTop:22 }}>
      <div style={{ flex:1, minWidth:320, background:"#fff", border:"1px solid #1a1a1a", padding:16 }}>
        <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:13, marginBottom:12 }}>WHO TO CHASE — by owner</div>
        {bar(owners, maxOwner, (o)=>OWNER_COLOR2[o]||"#888", 64, (o)=>goOwner(o))}
        <div style={{ fontSize:9, color:"#aaa", marginTop:8 }}>Click to open in {target==="todo"?"To-Do":"Tracker"}.</div>
      </div>
      <div style={{ flex:1, minWidth:320, background:"#fff", border:"1px solid #1a1a1a", padding:16 }}>
        <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:13, marginBottom:12 }}>OPEN ACTIVITIES</div>
        {acts.length===0?<div style={{ fontSize:11, color:"#999" }}>Nothing due.</div>:acts.map(([label,v])=>(
          <button key={label} onClick={()=>goAct(label,v.key)} style={{ display:"flex", alignItems:"center", gap:8, width:"100%", border:"none", background:"transparent", cursor:"pointer", fontFamily:"inherit", padding:"4px 0" }}>
            <span style={{ width:84, fontSize:10, fontWeight:700, color:"#444", textAlign:"left" }}>{label}</span>
            <span style={{ flex:1, height:16, background:"#f0ece3", position:"relative" }}><span style={{ position:"absolute", left:0, top:0, bottom:0, width:`${(v.n/maxAct)*100}%`, background:v.over?"#c0392b":"#d97706" }}/></span>
            <span style={{ width:54, textAlign:"right", fontSize:10, fontWeight:700 }}>{v.n}{v.over?<span style={{ color:"#c0392b" }}> ({v.over})</span>:null}</span>
          </button>))}
        <div style={{ fontSize:9, color:"#aaa", marginTop:8 }}>Click to open in {target==="todo"?"To-Do":"Tracker"}. Red = overdue.</div>
      </div>
      <div style={{ flex:1, minWidth:320, background:"#fff", border:"1px solid #1a1a1a", padding:16 }}>
        <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:13, marginBottom:12 }}>WHERE STYLES ARE STUCK</div>
        {Object.entries(phase).map(([p,n])=>(
          <div key={p} style={{ display:"flex", alignItems:"center", gap:8, padding:"4px 0" }}>
            <span style={{ width:80, fontSize:10, fontWeight:700, color:"#555" }}>{p}</span>
            <span style={{ flex:1, height:16, background:"#f0ece3", position:"relative" }}><span style={{ position:"absolute", left:0, top:0, bottom:0, width:`${(n/maxPhase)*100}%`, background:p==="Fabric IH"?"#c0392b":"#d97706" }}/></span>
            <span style={{ width:28, textAlign:"right", fontSize:11, fontWeight:700 }}>{n}</span>
          </div>))}
      </div>
    </div>
  </div>);
}

/* ========================= TO-DO ========================= */
function TodoView({ items, filter, setFilter, onJump }){
  const OWNER_COLOR2={ Merchant:"#1f6f54", CAD:"#2563a6", Buyer:"#b4531a", Designer:"#6d4aab", Mill:"#7a5a1e" };
  const [tf,setTf]=useState({}); // header filters: priority, order, junior, activity, branch, owner
  useEffect(()=>{ if(filter&&Object.keys(filter).length) setTf(f=>({ ...f, ...filter })); },[filter]);
  const distinct=(field)=>{ const s=new Set(); items.forEach(t=>{ const v=t[field]; if(v) s.add(v); }); return [...s].sort(); };
  const orders=distinct("orderNo"), juniors=distinct("junior"), activities=distinct("activity"), branches=distinct("branch"), owners=distinct("owner");
  const pass=(t)=> (!tf.priority||(tf.priority==="Overdue"?t.overdue:!t.overdue)) && (!tf.orderNo||t.orderNo===tf.orderNo) && (!tf.junior||t.junior===tf.junior) && (!tf.activity||t.activity===tf.activity) && (!tf.branch||t.branch===tf.branch) && (!tf.owner||t.owner===tf.owner);
  const shown=items.filter(pass);
  const overdue=shown.filter(t=>t.overdue), upcoming=shown.filter(t=>!t.overdue);
  const anyF=Object.values(tf).some(Boolean);
  const COLW={ pri:96, ord:60, sty:170, jr:64, act:92, br:84, own:70, date:84 };
  const set=(k,v)=>setTf(f=>({ ...f, [k]:v||undefined }));
  const hsel=(w,k,opts,first)=>(<select value={tf[k]||""} onChange={e=>set(k,e.target.value)} onClick={e=>e.stopPropagation()} style={{ width:w, fontFamily:"inherit", fontSize:9, padding:"2px 1px", border:"1px solid "+(tf[k]?"#d97706":"#ccc"), background:tf[k]?"#fff7ec":"#fff" }}><option value="">{first}</option>{opts.map(o=>(<option key={o} value={o}>{o}</option>))}</select>);
  const head=(<div style={{ display:"flex", alignItems:"flex-end", gap:10, padding:"6px 12px 4px", borderBottom:"2px solid #1a1a1a" }}>
    {hsel(COLW.pri,"priority",["Overdue","Upcoming"],"Priority")}
    {hsel(COLW.ord,"orderNo",orders,"Order")}
    <span style={{ width:COLW.sty, fontSize:9, fontWeight:700, textTransform:"uppercase", color:"#8a857a" }}>Style / Colour</span>
    {hsel(COLW.jr,"junior",juniors,"Junior")}
    {hsel(COLW.act,"activity",activities,"Activity")}
    {hsel(COLW.br,"branch",branches,"Branch")}
    {hsel(COLW.own,"owner",owners,"Owner")}
    <span style={{ width:COLW.date, fontSize:9, fontWeight:700, textTransform:"uppercase", color:"#8a857a" }}>Plan Date</span>
    <span style={{ flex:1, fontSize:9, fontWeight:700, textTransform:"uppercase", color:"#8a857a" }}>Days Late / Left</span>
  </div>);
  const row=(t)=>(<button key={(t.isColour?"col-":"")+t.id+t.key} onClick={()=>onJump(t.id,t.key)} style={{ display:"flex", alignItems:"center", gap:10, width:"100%", textAlign:"left", borderLeft:`4px solid ${t.overdue?"#c0392b":"#d97706"}`, borderBottom:"1px solid #eee7da", background:t.isColour?"#fbf8f1":"#fff", cursor:"pointer", fontFamily:"inherit", padding:"7px 12px" }}>
    <span style={{ width:COLW.pri, fontSize:10, fontWeight:700, display:"flex", alignItems:"center", gap:6, color:t.overdue?"#c0392b":"#7a560f" }}><span style={{ width:8, height:8, borderRadius:"50%", background:t.overdue?"#c0392b":"#d97706" }}/>{t.overdue?"Overdue":"Upcoming"}</span>
    <span style={{ width:COLW.ord, fontSize:10, color:"#555" }}>{t.orderNo||"—"}</span>
    <span style={{ width:COLW.sty, fontSize:11, fontWeight:700, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{t.isColour?<span style={{ display:"inline-flex", alignItems:"center", gap:6 }}><span style={{ fontSize:8, fontWeight:700, background:"#1a1a1a", color:"#f4f0e8", padding:"1px 4px" }}>FABRIC</span>{t.colour} <span style={{ color:"#999", fontWeight:400 }}>×{t.count}</span></span>:t.styleNo}</span>
    <span style={{ width:COLW.jr, fontSize:10, color:"#444" }}>{t.junior||"—"}</span>
    <span style={{ width:COLW.act, fontSize:10, fontWeight:700, color:"#333" }}>{t.activity}</span>
    <span style={{ width:COLW.br, fontSize:10, color:"#666" }}>{t.branch}</span>
    <span style={{ width:COLW.own, fontSize:10, fontWeight:700, color:OWNER_COLOR2[t.owner]||"#666" }}>{t.owner}</span>
    <span style={{ width:COLW.date, fontSize:10, color:"#666" }}>{fmt(t.exp)}</span>
    <span style={{ flex:1, fontSize:10, fontWeight:700, color:t.overdue?"#c0392b":"#7a560f" }}>{t.overdue?`+${Math.abs(t.du)}d late`:`${t.du}d left`}</span>
  </button>);
  return (<div style={{ padding:"16px 22px", maxWidth:1080 }}>
    <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:8 }}>
      <span style={{ fontSize:10, color:"#888" }}>Showing {shown.length} of {items.length}</span>
      {anyF && <button onClick={()=>{ setTf({}); setFilter&&setFilter({}); }} style={{ fontFamily:"inherit", fontSize:10, padding:"4px 9px", cursor:"pointer", border:"1px solid #c0392b", background:"#fff", color:"#c0392b", fontWeight:700 }}>clear filters</button>}
    </div>
    <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:13, color:"#c0392b", margin:"4px 0 6px" }}>OVERDUE · {overdue.length}</div>
    {head}
    {overdue.length?overdue.map(row):<div style={{ fontSize:11, color:"#999", padding:"8px 12px" }}>Nothing overdue. 👍</div>}
    <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:13, color:"#7a560f", margin:"20px 0 6px" }}>UPCOMING · {upcoming.length}</div>
    {head}
    {upcoming.length?upcoming.map(row):<div style={{ fontSize:11, color:"#999", padding:"8px 12px" }}>Nothing coming up in the watch windows.</div>}
    <div style={{ fontSize:9, color:"#aaa", marginTop:14 }}>Filter any column above · fabric grouped by unique colour (×N = styles needing it now) · click a row to jump to it in the Tracker.</div>
  </div>);
}

/* ========================= SETTINGS ========================= */
function SettingsView({ cfg, setCfg, canEdit }){
  const num=(v)=> v==null?"":v;
  const setLead=(k,val)=> setCfg(c=>({ ...c, leads:{ ...c.leads, [k]: val===""?undefined:Math.max(0,Number(val)||0) } }));
  const setRew=(k,val)=> setCfg(c=>({ ...c, rework:{ ...c.rework, [k]: val===""?undefined:Math.max(0,Number(val)||0) } }));
  const setUpc=(k,val)=> setCfg(c=>({ ...c, upcoming:{ ...c.upcoming, [k]: val===""?undefined:Math.max(0,Number(val)||0) } }));
  const setTop=(k,val)=> setCfg(c=>({ ...c, [k]: val===""?undefined:Math.max(0,Number(val)||0) }));
  const inp=(value,onChange)=>(<input type="number" min="0" disabled={!canEdit} value={num(value)} onChange={e=>onChange(e.target.value)} style={{ width:58, fontFamily:"inherit", fontSize:12, padding:"4px 6px", border:"1px solid #bbb", outline:"none", background:canEdit?"#fff":"#f3f1ec" }}/>);
  const box={ background:"#fff", border:"1px solid #1a1a1a", padding:16, minWidth:300, flex:1 };
  const head={ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:13, marginBottom:4 };
  const sub={ fontSize:10, color:"#999", marginBottom:12 };
  const rworkLabels={ fitSend:"Fit (redo & resend)", artwork:"Artwork", strikeOff:"Strike-off", labDip:"Lab Dip", ppSample:"PP Sample" };
  const upcLabels={ fitSend:"Fit Send", artwork:"Artwork", strikeOff:"Strike-off", ppSample:"PP Sample", fabricIH:"Fabric In-House" };
  return (<div style={{ padding:"18px 22px", maxWidth:1100 }}>
    {!canEdit && <div style={{ fontSize:11, color:"#c0392b", marginBottom:12 }}>Your role cannot edit these — Management / Senior Merchant only.</div>}
    <div style={{ display:"flex", gap:18, flexWrap:"wrap" }}>
      <div style={box}>
        <div style={head}>STAGE LEAD TIMES</div><div style={sub}>Working days each stage takes after its predecessor.</div>
        {STAGES.map(st=>(<div key={st.key} style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"3px 0", fontSize:11 }}><span>{st.label}</span>{inp(cfg.leads[st.key], v=>setLead(st.key,v))}</div>))}
      </div>
      <div style={{ display:"flex", flexDirection:"column", gap:18, flex:1, minWidth:300 }}>
        <div style={box}>
          <div style={head}>REWORK DAYS (on rejection)</div><div style={sub}>Working days added on a rejection, before re-send.</div>
          {Object.keys(rworkLabels).map(k=>(<div key={k} style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"3px 0", fontSize:11 }}><span>{rworkLabels[k]}</span>{inp(cfg.rework[k], v=>setRew(k,v))}</div>))}
        </div>
        <div style={box}>
          <div style={head}>DELIVERY GATES</div><div style={sub}>Working days before delivery.</div>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"3px 0", fontSize:11 }}><span>Fabric cut-off (Fabric IH by)</span>{inp(cfg.fabricCutoff, v=>setTop("fabricCutoff",v))}</div>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"3px 0", fontSize:11 }}><span>Release gate (prod must start by)</span>{inp(cfg.relGate, v=>setTop("relGate",v))}</div>
        </div>
      </div>
      <div style={box}>
        <div style={head}>TO-DO WATCH WINDOWS</div><div style={sub}>Working days before due that an activity becomes "upcoming".</div>
        {Object.keys(upcLabels).map(k=>(<div key={k} style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"3px 0", fontSize:11 }}><span>{upcLabels[k]}</span>{inp(cfg.upcoming[k], v=>setUpc(k,v))}</div>))}
      </div>
    </div>
    <button disabled={!canEdit} onClick={()=>setCfg(DEFAULT_CFG)} style={{ marginTop:16, fontFamily:"inherit", fontSize:11, padding:"7px 14px", cursor:canEdit?"pointer":"not-allowed", border:"1px solid #1a1a1a", background:"#fff", opacity:canEdit?1:0.5 }}>Reset to defaults</button>
    <div style={{ fontSize:9, color:"#aaa", marginTop:10 }}>Changes save automatically and apply to every user's calculations.</div>
  </div>);
}

// ============================ AUTH LAYER ============================
function Splash({ text }){ return (<div style={{ minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", background:"#f4f0e8", fontFamily:"'JetBrains Mono',monospace", color:"#6a665e", fontSize:13 }}>{text}</div>); }

function AuthShell({ children }){ return (<div style={{ minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", background:"#f4f0e8", fontFamily:"'JetBrains Mono',monospace", padding:20 }}><div style={{ width:360, maxWidth:"100%", background:"#fff", border:"2px solid #1a1a1a", boxShadow:"8px 8px 0 #1a1a1a", padding:26 }}>{children}</div></div>); }

function LoginScreen(){
  const [email,setEmail]=useState(""); const [pw,setPw]=useState(""); const [mode,setMode]=useState("in"); const [msg,setMsg]=useState(""); const [busy,setBusy]=useState(false); const [remember,setRemember]=useState(true);
  const submit=async()=>{ if(!email.trim()||!pw){ setMsg("Enter email and password."); return; } setMsg(""); setBusy(true);
    const stamp=()=>{ try{ localStorage.setItem("mt_login_at", String(Date.now())); localStorage.setItem("mt_remember", remember?"1":"0"); }catch(x){} };
    try{ if(mode==="in"){ const { error }=await supabase.auth.signInWithPassword({ email:email.trim(), password:pw }); if(error) throw error; stamp(); }
      else { const { error,data }=await supabase.auth.signUp({ email:email.trim(), password:pw }); if(error) throw error; if(data.session){ stamp(); } if(!data.session){ setMsg("Account created. If sign-in doesn't happen automatically, check your email to confirm, then sign in."); setMode("in"); } } }
    catch(e){ setMsg(e.message||String(e)); } setBusy(false); };
  const inp={ width:"100%", fontFamily:"inherit", fontSize:13, padding:"9px 10px", border:"1px solid #1a1a1a", marginBottom:10, boxSizing:"border-box" };
  return (<AuthShell>
    <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:22, marginBottom:2 }}>Merch Tracker</div>
    <div style={{ fontSize:11, color:"#888", marginBottom:18 }}>{mode==="in"?"Sign in to your account":"Create your account"}</div>
    <input style={inp} type="email" placeholder="email" value={email} onChange={e=>setEmail(e.target.value)} onKeyDown={e=>{ if(e.key==="Enter") submit(); }}/>
    <input style={inp} type="password" placeholder="password" value={pw} onChange={e=>setPw(e.target.value)} onKeyDown={e=>{ if(e.key==="Enter") submit(); }}/>
    <label style={{ display:"flex", alignItems:"center", gap:7, fontSize:11, color:"#555", marginBottom:12, cursor:"pointer" }}><input type="checkbox" checked={remember} onChange={e=>setRemember(e.target.checked)}/>Keep me signed in <span style={{ color:"#999" }}>(else sign out after 12h)</span></label>
    <button disabled={busy} onClick={submit} style={{ width:"100%", fontFamily:"inherit", fontSize:13, fontWeight:700, padding:"10px", cursor:busy?"wait":"pointer", border:"1px solid #1a1a1a", background:"#1a1a1a", color:"#f4f0e8", marginBottom:10 }}>{busy?"…":(mode==="in"?"Sign in":"Create account")}</button>
    <div style={{ fontSize:11, textAlign:"center" }}><span style={{ color:"#888" }}>{mode==="in"?"New here? ":"Have an account? "}</span><button onClick={()=>{ setMode(mode==="in"?"up":"in"); setMsg(""); }} style={{ border:"none", background:"transparent", color:"#d97706", cursor:"pointer", fontFamily:"inherit", fontSize:11, fontWeight:700 }}>{mode==="in"?"Create account":"Sign in"}</button></div>
    {msg && <div style={{ fontSize:11, color:"#c0392b", marginTop:12, lineHeight:1.4 }}>{msg}</div>}
  </AuthShell>);
}

function PendingScreen({ email, onSignOut }){
  return (<AuthShell>
    <div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:18, marginBottom:8 }}>Awaiting access</div>
    <div style={{ fontSize:12, color:"#555", lineHeight:1.55, marginBottom:16 }}>You're signed in as <b>{email}</b>, but your account hasn't been given a role yet. Ask a Management user to set your role in the <b>Users</b> panel, then refresh this page.</div>
    <button onClick={onSignOut} style={{ fontFamily:"inherit", fontSize:12, padding:"8px 14px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#fff" }}>Sign out</button>
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
    <div onClick={e=>e.stopPropagation()} style={{ background:"#f4f0e8", border:"2px solid #1a1a1a", boxShadow:"8px 8px 0 #1a1a1a", width:560, maxWidth:"100%", maxHeight:"86vh", overflowY:"auto", padding:22, fontFamily:"'JetBrains Mono',monospace" }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}><div style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:18 }}>Users &amp; roles</div><button onClick={onClose} style={{ border:"none", background:"transparent", cursor:"pointer" }}><X size={18}/></button></div>
      <p style={{ fontSize:11, color:"#666", lineHeight:1.55, marginBottom:12 }}>Each person signs in with their own email + password; their access follows the role you set here. Set anyone to "pending" to suspend access.</p>
      {list===null ? <div style={{ fontSize:12, color:"#888" }}>Loading…</div> : list.length===0 ? <div style={{ fontSize:12, color:"#888" }}>No users yet — have your team sign up from the login screen.</div> : (
        <div style={{ border:"1px solid #1a1a1a", background:"#fff" }}>{list.map((u,i)=>(<div key={u.id} style={{ display:"flex", alignItems:"center", gap:10, padding:"8px 10px", borderBottom:i<list.length-1?"1px solid #eee7da":"none" }}>
          <div style={{ flex:1, minWidth:0 }}><input value={u.name||""} placeholder="(set name)" onChange={e=>editName(u.id,e.target.value)} onBlur={e=>{ e.target.style.border="1px solid transparent"; e.target.style.background="transparent"; saveName(u.id,e.target.value); }} onKeyDown={e=>{ if(e.key==="Enter") e.target.blur(); }} style={{ fontFamily:"inherit", fontSize:12, fontWeight:700, width:"100%", border:"1px solid transparent", background:"transparent", padding:"2px 4px", boxSizing:"border-box" }} onFocus={e=>{ e.target.style.border="1px solid #1a1a1a"; e.target.style.background="#fff"; }} onMouseLeave={e=>{}} />
<div style={{ fontSize:10, color:"#999", whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis", padding:"0 4px" }}>{u.email}</div></div>
          <select disabled={busy} value={u.role||"pending"} onChange={e=>setR(u.id,e.target.value)} style={{ fontFamily:"inherit", fontSize:11, padding:"5px 7px", border:"1px solid #1a1a1a", background:u.role==="pending"?"#fbeaea":"#fff7ec" }}>{ROLE_OPTS.map(r=>(<option key={r} value={r}>{r==="pending"?"— pending —":(ROLES[r]||{}).label||r}</option>))}</select>
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
  return <MerchTracker me={profile} onSignOut={signOut}/>;
}

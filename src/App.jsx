import React, { useState, useMemo, useRef, useEffect } from "react";
import { Check, Plus, Lock, Filter, X, Copy, ChevronUp, ChevronDown, CornerDownRight, Columns3, MessageSquare, RotateCcw, Droplet, Snowflake } from "lucide-react";
import { supabase } from "./supabaseClient";

/* MERCH TRACKER — Excel-like entry grid PROTOTYPE (v13 Excel as smart base) */

const FONT = `@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@500;600;800&family=JetBrains+Mono:wght@400;500;700&display=swap');`;
const REL_GATE_DAYS = 30, FABRIC_CUTOFF_DAYS = 35, STYLE_W = 190;

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
const OWNER_COLOR = { Merchant:"#1f6f54", CAD:"#2563a6", Buyer:"#b4531a", Designer:"#6d4aab", Mill:"#7a5a1e", Tamal:"#1f6f54", Rina:"#6d4aab" };

const ONE_DAY = 86400000;
function addWorkdays(date,n){ if(!date) return null; const d=new Date(date.getTime()); let a=0; const step=n>=0?1:-1, t=Math.abs(n); while(a<t){ d.setTime(d.getTime()+step*ONE_DAY); if(d.getDay()!==0) a++; } return d; }
function netWorkdays(a,b){ if(!a||!b) return null; let d=new Date(a.getTime()), end=new Date(b.getTime()), sign=1; if(end<d){ const tmp=d; d=end; end=tmp; sign=-1; } let n=0; const cur=new Date(d.getTime()); while(cur<end){ cur.setTime(cur.getTime()+ONE_DAY); if(cur.getDay()!==0) n++; } return sign*n; }
const fmt=(d)=> !d?"":d.toLocaleDateString("en-GB",{day:"2-digit",month:"short"});
const parse=(s)=>{ if(!s) return null; const m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(s); return m?new Date(Number(m[1]),Number(m[2])-1,Number(m[3])):new Date(s); };
const iso=(d)=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
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

const applicableStages=(s)=> STAGES.filter(st=> st.flag===null || s[st.flag]);

function computeStyle(s){
  const ordRec=parse(s.ordRec), delivery=parse(s.delivery);
  const cutoff=addWorkdays(delivery,-FABRIC_CUTOFF_DAYS);
  const eff={}, plan={};
  const applies=(k)=>{ const st=STAGES.find(x=>x.key===k); return st.flag===null||s[st.flag]; };
  const actualOf=(k)=>parse(s.actuals[k]); const revOf=(k)=>parse(s.revs?.[k]);
  STAGES.forEach(st=>{
    let p;
    if(st.cutoff){ const base=s.labDipReq?(eff["labAppr"]||eff["labDip"]||ordRec):ordRec; p=s.labDipReq?new Date(Math.max(addWorkdays(base,15)?.getTime()||0, cutoff.getTime())):cutoff; }
    else { let predEff; if(st.key==="prodFile") predEff = s.ppBypass ? eff["fabricIH"] : eff["ppAppr"]; else predEff = st.pred==="__ord"?ordRec:eff[st.pred]; p=addWorkdays(predEff||ordRec, st.lead); }
    plan[st.key]=p; eff[st.key]=actualOf(st.key)||revOf(st.key)||p;
  });
  const stages=applicableStages(s).map(st=>({ ...st, actual:actualOf(st.key), rev:revOf(st.key), plan:plan[st.key], done:!!actualOf(st.key) }));
  let nextPending=null, lastActual=null;
  stages.forEach(r=>{ if(r.actual&&(!lastActual||r.actual>lastActual)) lastActual=r.actual; if(!r.done&&!nextPending) nextPending=r; });
  const released=stages.every(r=>r.done);
  const idle=lastActual?Math.max(0,netWorkdays(lastActual,TODAY)):null;
  const get=(k)=>stages.find(r=>r.key===k); const done=(k)=>!!(get(k)&&get(k).done);
  const fabricInHouse=done("fabricIH");
  const lastPlan=stages[stages.length-1]?.plan;
  const float=lastPlan?netWorkdays(lastPlan,delivery):null;
  let status="On Track", tone="ok";
  if(released){ status="Released"; tone="done"; }
  else if(nextPending&&nextPending.plan&&TODAY>nextPending.plan){ status=`Overdue ${Math.round((TODAY-nextPending.plan)/ONE_DAY)}d`; tone="late"; }
  else if(idle!==null&&idle>=7){ status=`Idle ${idle}d`; tone="warn"; }
  const dueText=(k)=>{ const r=get(k); if(!r||!r.plan) return "pending"; return TODAY>r.plan?`OVERDUE ${Math.round((TODAY-r.plan)/ONE_DAY)}d`:`due ${fmt(r.plan)}`; };
  const bs=(txt,tn)=>({txt,tone:tn});
  let fitBranch;
  if(!s.fitReq) fitBranch=bs("—","na"); else if(done("fitAppr")) fitBranch=bs("Fit Approved","ok"); else if(fabricInHouse) fitBranch=bs("Not done before Fabric IH","late"); else if(done("fitSend")) fitBranch=bs(`Fit appr ${dueText("fitAppr")}`,TODAY>(get("fitAppr")?.plan||TODAY)?"late":"warn"); else fitBranch=bs(`Fit send ${dueText("fitSend")}`,TODAY>(get("fitSend")?.plan||TODAY)?"late":"warn");
  let printBranch; const printDone=s.soReq?done("soAppr"):done("artAppr");
  if(!s.printReq) printBranch=bs("—","na"); else if(printDone) printBranch=bs("Print Approved","ok"); else if(fabricInHouse) printBranch=bs("Not done before Fabric IH","late"); else if(!done("artwork")) printBranch=bs(`Artwork ${dueText("artwork")}`,TODAY>(get("artwork")?.plan||TODAY)?"late":"warn"); else if(!done("artAppr")) printBranch=bs(`Art appr ${dueText("artAppr")}`,"warn"); else if(s.soReq&&!done("strikeOff")) printBranch=bs(`S/O ${dueText("strikeOff")}`,"warn"); else printBranch=bs(`S/O appr ${dueText("soAppr")}`,"warn");
  let fabricBranch; const fabPlan=get("fabricIH")?.plan;
  const fabDue=fabPlan?(TODAY>fabPlan?`IH OVERDUE ${Math.round((TODAY-fabPlan)/ONE_DAY)}d`:`IH due ${fmt(fabPlan)}`):"IH —";
  const fabTone=fabPlan&&TODAY>fabPlan?"late":"warn";
  if(fabricInHouse) fabricBranch=bs("Bulk Fabric In-House","ok"); else if(s.labDipReq&&done("labAppr")) fabricBranch=bs(`Lab Dip Appr | ${fabDue}`,fabTone); else if(s.labDipReq&&done("labDip")) fabricBranch=bs(`Lab dip sent, appr pending | ${fabDue}`,"warn"); else if(s.labDipReq) fabricBranch=bs(`Lab dip pending | ${fabDue}`,"warn"); else fabricBranch=bs(fabDue,fabTone);
  let ppBranch;
  if(!s.ppNeeded) ppBranch=bs("PP Not Required","na"); else if(done("ppAppr")) ppBranch=bs("PP Approved","ok"); else if(done("ppSample")) ppBranch=bs(`PP appr ${dueText("ppAppr")}`,"warn"); else if(fabricInHouse) ppBranch=bs(`PP sample ${dueText("ppSample")}`,"warn"); else ppBranch=bs("Awaiting bulk fabric","warn");
  let fabricCountdown;
  if(fabricInHouse) fabricCountdown={txt:"in-house",n:9e9,tone:"ok"}; else if(fabPlan){ const n=netWorkdays(TODAY,fabPlan); fabricCountdown={txt:n<0?`${-n}d over`:`${n}d`,n,tone:n<0?"late":n<=7?"warn":"ok"}; } else fabricCountdown={txt:"—",n:null,tone:"na"};
  const releaseGate=addWorkdays(delivery,-REL_GATE_DAYS);
  let projRelease;
  if(released) projRelease=lastActual;
  else { let cur=eff["fabricIH"]; const chain = s.ppBypass ? ["prodFile"] : ["ppSample","ppAppr","prodFile"]; chain.forEach(k=>{ if(!applies(k)) return; const a=actualOf(k), r=revOf(k); const st=STAGES.find(x=>x.key===k); if(a) cur=a; else if(r) cur=r; else cur=addWorkdays(cur,st.lead); }); projRelease=cur; }
  const gateGap=projRelease&&releaseGate?Math.round((releaseGate-projRelease)/ONE_DAY):null;
  const releaseOnTrack=projRelease&&releaseGate?projRelease<=releaseGate:true;
  const projTone=released?"done":(!releaseOnTrack?"late":(gateGap!=null&&gateGap<=5?"warn":"ok"));
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
  return { stages, nextPending, status, tone, idle, float, released, fitBranch, printBranch, fabricBranch, ppBranch, fabricCountdown, projRelease, projTone, releaseGate, releaseOnTrack, pct, ownerToChase, chaseOwners };
}

const ROLES={ Merchant:{label:"Merchant",canEdit:()=>true}, Junior:{label:"Junior",canEdit:(o)=>o==="Merchant"||o==="CAD"}, Viewer:{label:"Viewer (Buyer)",canEdit:()=>false} };
const TONE_STYLE={ ok:{dot:"#1f6f54",bg:"#eef6f1",fg:"#16523d"}, warn:{dot:"#b4801a",bg:"#fbf4e6",fg:"#7a560f"}, late:{dot:"#c0392b",bg:"#fcecea",fg:"#8c241a"}, done:{dot:"#555",bg:"#efefea",fg:"#444"} };
const BR_TONE={ ok:{bg:"#eef6f1",fg:"#16523d"}, warn:{bg:"#fbf4e6",fg:"#7a560f"}, late:{bg:"#fcecea",fg:"#8c241a"}, na:{bg:"transparent",fg:"#c4c0b8"}, done:{bg:"#efefea",fg:"#555"} };
const FLAG_DEFS=[ {key:"fitReq",short:"FIT",title:"Fit sample required"}, {key:"printReq",short:"PRT",title:"Print required"}, {key:"soReq",short:"S/O",title:"Strike-off required"}, {key:"labDipReq",short:"LAB",title:"Lab dip required"}, {key:"ppBypass",short:"BYP",title:"PP bypass — Prod File flows straight from Fabric IH (not PP Appr)"}, {key:"ppNeeded",short:"PP",title:"PP sample required"} ];
const FILL_SWATCHES=["#fff7ec","#fde2e1","#e7f3ec","#e3edf9","#f3e8fa","#fff3bf",""];

const INFO_COLS=[
  { key:"orderNo",   label:"Order No",   kind:"text", w:60,  owner:"Merchant" },
  { key:"sampleFit", label:"Sample Fit", kind:"text", w:72,  owner:"Merchant" },
  { key:"family",    label:"Family",     kind:"text", w:130, owner:"Merchant" },
  { key:"colour",    label:"Colour",     kind:"text", w:150, owner:"Merchant" },
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
const BRANCH_STAGES={ fit:["fitSend","fitAppr"], print:["artwork","artAppr","strikeOff","soAppr"], fabric:["labDip","labAppr","fabricIH"], pp:["ppSample","ppAppr"] };
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
          {cells.map((d,i)=>{ if(!d) return <div key={i}/>; const isSel=sel&&sel.getFullYear()===y&&sel.getMonth()===m&&sel.getDate()===d; const isToday=TODAY.getFullYear()===y&&TODAY.getMonth()===m&&TODAY.getDate()===d; const isSun=new Date(y,m,d).getDay()===0; return <button key={i} onClick={()=>{ onPick(`${y}-${String(m+1).padStart(2,"0")}-${String(d).padStart(2,"0")}`); onClose(); }} style={{ fontSize:11, padding:"5px 0", cursor:"pointer", border:"none", fontFamily:"inherit", background:isSel?"#d97706":isToday?"#1a1a1a":"transparent", color:isSel||isToday?"#fff":isSun?"#c0392b":"#1a1a1a", fontWeight:isSel||isToday?700:400 }}>{d}</button>; })}
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
  return (<span onClick={(e)=>{ if(onJump){ e.stopPropagation(); onJump(); } }} title={`${b.txt}  ·  click → jump to enter`} style={{ display:"inline-flex", alignItems:"center", gap:3, background:t.bg, color:t.fg, padding:"2px 7px", fontSize:9.5, fontWeight:600, whiteSpace:"nowrap", maxWidth:"100%", overflow:"hidden", textOverflow:"ellipsis", cursor:onJump?"pointer":"default", textDecoration:onJump?"underline dotted":"none" }}>{b.txt}{onJump && <CornerDownRight size={9} style={{ flexShrink:0 }}/>}</span>);
}

export default function App(){
  const [styles,setStyles]=useState([]); // loaded from Supabase on mount
  const [role,setRole]=useState("Merchant");
  const [search,setSearch]=useState("");
  const [statusFilter,setStatusFilter]=useState("All");
  const [saved,setSaved]=useState(false);
  const [fillOpen,setFillOpen]=useState(false);
  const [colsOpen,setColsOpen]=useState(false);
  const [sel,setSel]=useState(null);      // anchor {id,col}
  const [focus,setFocus]=useState(null);   // range focus {id,col}
  const [editing,setEditing]=useState(null);
  const [editVal,setEditVal]=useState("");
  const [sort,setSort]=useState({ col:null, dir:1 });
  const [hidden,setHidden]=useState(new Set());
  const [freezeN,setFreezeN]=useState(1);  // # leading columns frozen (incl style)
  const [colW,setColW]=useState({});  // per-column width overrides (drag to resize)
  const [fills,setFills]=useState({});
  const [notes,setNotes]=useState({});
  const [noteEditing,setNoteEditing]=useState(false);
  const [noteText,setNoteText]=useState("");
  const [clip,setClip]=useState(null);     // {values:2D,h,w}
  const [colFilters,setColFilters]=useState({}); // col -> array of allowed display values
  const [filterCol,setFilterCol]=useState(null); // which header filter is open
  const [past,setPast]=useState([]); const [future,setFuture]=useState([]);
  const [filling,setFilling]=useState(false); const [fillFrom,setFillFrom]=useState(null); const [fillTo,setFillTo]=useState(null);
  const selectingRef=useRef(false); const [dragSel,setDragSel]=useState(false);
  const [newRow,setNewRow]=useState({ styleNo:"", orderNo:"", sampleFit:"", family:"", colour:"", owner:"", qty:"", ordRec:iso(TODAY), delivery:"", fitReq:true, printReq:false, soReq:false, ppBypass:false, labDipReq:true, ppNeeded:true });
  const [newError,setNewError]=useState("");
  const savedTimer=useRef();
  const gridRef=useRef();
  const firstRender=useRef(true);
  const loadedRef=useRef(false);
  const S2C={ orderNo:"order_no", styleNo:"style_no", sampleFit:"sample_fit", family:"family", colour:"colour", owner:"owner", setId:"set_id", setRole:"set_role", age:"age", qty:"qty", ordRec:"order_date", delivery:"delivery_date", fitReq:"fit_req", printReq:"print_req", soReq:"so_req", ppBypass:"pp_bypass", labDipReq:"lab_dip_req", ppNeeded:"pp_needed", remarks:"remarks" };
  const styleToRow=(s)=>{ const r={ id:s.id }; Object.entries(S2C).forEach(([k,col])=>{ r[col]= k==="qty"?(Number(s[k])||0):(s[k]||null); }); return r; };
  const rowToStyle=(row,byId)=>({ id:row.id, orderNo:row.order_no||"", sampleFit:row.sample_fit||"", family:row.family||"", styleNo:row.style_no||"", colour:row.colour||"", owner:row.owner||"", setId:row.set_id||"", setRole:row.set_role||"", age:row.age||"", qty:row.qty||0, ordRec:row.order_date||"", delivery:row.delivery_date||"", fitReq:!!row.fit_req, printReq:!!row.print_req, soReq:!!row.so_req, ppBypass:!!row.pp_bypass, labDipReq:!!row.lab_dip_req, ppNeeded:!!row.pp_needed, remarks:row.remarks||"", actuals:(byId[row.id]&&byId[row.id].actuals)||{}, revs:(byId[row.id]&&byId[row.id].revs)||{} });
  // LOAD everything from Supabase (also used by the Sync button)
  const loadShared=async()=>{ try{
    const [styRes, sdRes, cmRes] = await Promise.all([
      supabase.from("styles").select("*").order("id"),
      supabase.from("stage_dates").select("*"),
      supabase.from("cell_meta").select("*"),
    ]);
    if(styRes.error||!styRes.data){ console.error(styRes.error); return; }
    const byId={}; (sdRes.data||[]).forEach(r=>{ const e=(byId[r.style_id]=byId[r.style_id]||{actuals:{},revs:{}}); if(r.actual_date) e.actuals[r.stage]=r.actual_date; if(r.revised_date) e.revs[r.stage]=r.revised_date; });
    setStyles(styRes.data.map(row=>rowToStyle(row,byId)));
    const f={}, n={}; (cmRes.data||[]).forEach(r=>{ if(r.fill) f[`${r.style_id}:${r.col}`]=r.fill; if(r.note) n[`${r.style_id}:${r.col}`]=r.note; });
    setFills(f); setNotes(n); loadedRef.current=true; flash();
  }catch(e){ console.error("load failed",e); } };
  useEffect(()=>{ loadShared(); },[]);
  // SAVE everything to Supabase shortly after any change (debounced)
  useEffect(()=>{ if(firstRender.current){ firstRender.current=false; return; } if(!loadedRef.current) return; const t=setTimeout(async()=>{ try{ setSaveState("saving");
    const up1=await supabase.from("styles").upsert(styles.map(styleToRow)); if(up1.error) throw up1.error;
    const stageRows=[]; styles.forEach(s=> STAGE_KEYS.forEach(k=> stageRows.push({ style_id:s.id, stage:k, revised_date:(s.revs&&s.revs[k])||null, actual_date:s.actuals[k]||null })));
    if(stageRows.length){ const up2=await supabase.from("stage_dates").upsert(stageRows,{ onConflict:"style_id,stage" }); if(up2.error) throw up2.error; }
    const keys=new Set([...Object.keys(fills),...Object.keys(notes)]); const metaRows=[...keys].map(key=>{ const i=key.indexOf(":"); return { style_id:Number(key.slice(0,i)), col:key.slice(i+1), fill:fills[key]||null, note:notes[key]||null }; });
    if(metaRows.length){ const up3=await supabase.from("cell_meta").upsert(metaRows,{ onConflict:"style_id,col" }); if(up3.error) throw up3.error; }
    setSaveState("saved"); flash();
  }catch(e){ console.error("save failed",e); setSaveState("error"); } },700); return ()=>clearTimeout(t); },[styles,fills,notes]);

  const [saveState,setSaveState]=useState("idle"); // idle | saving | saved | error
  const flash=()=>{ setSaved(true); clearTimeout(savedTimer.current); savedTimer.current=setTimeout(()=>setSaved(false),1200); };
  const setField=(id,field,val)=>{ pushHistory(); setStyles(prev=>prev.map(s=>{ if(s.id!==id) return s; if(STAGE_KEYS.includes(field)) return { ...s, actuals:{ ...s.actuals, [field]: val||undefined } }; if(field==="qty") return { ...s, qty:Number(val)||0 }; return { ...s, [field]:val }; })); flash(); };
  const setRev=(id,key,val)=>{ pushHistory(); setStyles(prev=>prev.map(s=> s.id===id?{...s,revs:{...(s.revs||{}),[key]:val||undefined}}:s)); flash(); };
  const toggleFlag=(id,flag)=>{ pushHistory(); setStyles(prev=>prev.map(s=>s.id===id?{...s,[flag]:!s[flag]}:s)); flash(); };

  const computed=useMemo(()=>styles.map(s=>({s,c:computeStyle(s)})),[styles]);
  const valueFor=(s,cc,col)=>{
    if(col==="__style") return s.styleNo||"";
    if(["orderNo","sampleFit","family","colour","owner","setId","setRole","remarks"].includes(col)) return s[col]||"(Blanks)";
    if(col==="qty") return String(s.qty);
    if(col==="ordRec"||col==="delivery") return fmt(parse(s[col]))||"(Blanks)";
    if(col==="overall") return cc.status;
    if(col==="fit") return cc.fitBranch.txt; if(col==="print") return cc.printBranch.txt; if(col==="fabric") return cc.fabricBranch.txt; if(col==="pp") return cc.ppBranch.txt;
    if(col==="fabricCD") return cc.fabricCountdown.txt;
    if(col==="proj") return fmt(cc.projRelease)||"(Blanks)";
    if(col==="pct") return cc.pct+"%";
    if(col==="chase") return (cc.chaseOwners||[]).map(o=>o.owner).join(", ")||"(Blanks)";
    if(col==="float") return cc.float==null?"(Blanks)":String(cc.float);
    if(col==="idle") return cc.idle==null?"(Blanks)":String(cc.idle);
    if(STAGE_KEYS.includes(col)){ const a=s.actuals[col]; return a?fmt(parse(a)):"(Blanks)"; }
    return "";
  };
  const filtered=computed.filter(({s,c})=>{ const q=search.toLowerCase(); const ownerMatch=(c.chaseOwners||[]).some(o=>o.owner.toLowerCase().includes(q)); const matchQ=!q||s.styleNo.toLowerCase().includes(q)||s.colour.toLowerCase().includes(q)||s.family.toLowerCase().includes(q)||s.sampleFit.toLowerCase().includes(q)||s.orderNo.toLowerCase().includes(q)||ownerMatch; const matchS=statusFilter==="All"||(statusFilter==="At Risk"&&(c.tone==="late"||c.tone==="warn"))||(statusFilter==="On Track"&&c.tone==="ok")||(statusFilter==="Released"&&c.released); const matchF=Object.entries(colFilters).every(([col,allowed])=> !allowed || allowed.includes(valueFor(s,c,col))); return matchQ&&matchS&&matchF; });
  const toneRank={ late:0, warn:1, ok:2, done:3, na:4 };
  const fitNum=(s)=>{ const m=String(s.sampleFit).match(/\d+/); return m?Number(m[0]):Infinity; };
  const sortVal=(col,{s,c})=>{ switch(col){ case "__style": return s.styleNo.toLowerCase(); case "orderNo": return (s.orderNo||"~").toLowerCase(); case "sampleFit": return fitNum(s); case "family": return s.family.toLowerCase(); case "colour": return s.colour.toLowerCase(); case "owner": return (s.owner||"").toLowerCase(); case "setId": return (s.setId||"~").toLowerCase(); case "setRole": return (s.setRole||"").toLowerCase(); case "qty": return s.qty; case "ordRec": return s.ordRec?new Date(s.ordRec).getTime():Infinity; case "delivery": return s.delivery?new Date(s.delivery).getTime():Infinity; case "overall": return toneRank[c.tone]; case "fit": return toneRank[c.fitBranch.tone]; case "print": return toneRank[c.printBranch.tone]; case "fabric": return toneRank[c.fabricBranch.tone]; case "pp": return toneRank[c.ppBranch.tone]; case "fabricCD": return c.fabricCountdown.n==null?Infinity:c.fabricCountdown.n; case "proj": return c.projRelease?c.projRelease.getTime():Infinity; case "pct": return c.pct; case "chase": return (c.chaseOwners||[]).length; case "float": return c.float==null?Infinity:c.float; case "idle": return c.idle==null?-1:c.idle; case "remarks": return (s.remarks||"~").toLowerCase(); default: { const a=s.actuals[col]; return a?new Date(a).getTime():Infinity; } } };
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
  const beginDate=(id,col,mode,initialChar)=>{ if(!ROLES[role].canEdit(ownerOfCol(col))) return; setSel({id,col}); setFocus(null); setEditing({id,col,mode}); setCalOpen(false); const s=styles.find(x=>x.id===id); const cur= mode==="rev"?(s&&s.revs&&s.revs[col]):(isStageCol(col)?(s&&s.actuals[col]):(s&&s[col])); setEditVal(initialChar!=null?initialChar:(cur?fmtTyped(cur):"")); };
  const commitDate=()=>{ if(!editing) return; const r=parseTyped(editVal); if(r!==false){ const val=r===""?null:r; if(editing.mode==="rev") setRev(editing.id,editing.col,val); else setField(editing.id,editing.col,val); } setEditing(null); setCalOpen(false); };
  const dateEditor=(id,col,mode)=>{ const s=styles.find(x=>x.id===id); const stored= mode==="rev"?(s&&s.revs&&s.revs[col]):(isStageCol(col)?(s&&s.actuals[col]):(s&&s[col])); return (<span onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:1, left:1, zIndex:80, display:"flex", alignItems:"center", gap:2, background:"#fff", border:"1px solid #1d4ed8", padding:"2px 3px" }}><input autoFocus value={editVal} placeholder="dd/mm/yyyy" onChange={e=>setEditVal(e.target.value)} onKeyDown={e=>{ e.stopPropagation(); if(e.key==="Enter") commitDate(); else if(e.key==="Escape"){ setEditing(null); setCalOpen(false); } }} onBlur={()=>{ if(!calOpen) commitDate(); }} style={{ width:80, fontFamily:"inherit", fontSize:11, border:"none", outline:"none" }}/><button onMouseDown={e=>e.preventDefault()} onClick={e=>{ e.stopPropagation(); setCalOpen(o=>!o); }} title="calendar" style={{ border:"none", background:"transparent", cursor:"pointer", padding:0, lineHeight:0, fontSize:12 }}>📅</button>{calOpen && <CalPopup label={mode==="rev"?"set REVISED plan":"set actual"} value={stored} onClose={()=>setCalOpen(false)} onPick={(d)=>{ if(mode==="rev") setRev(id,col,d); else setField(id,col,d); setEditing(null); setCalOpen(false); }}/>}</span>); };
    const startEdit=(id,col,initialChar)=>{ if(!isEditableCol(col)) return; if(!ROLES[role].canEdit(ownerOfCol(col))) return; if(isDateCol(col)){ beginDate(id,col,"actual"); return; } const s=styles.find(x=>x.id===id); setEditing({id,col,mode:"text"}); if(col==="qty") setEditVal(initialChar??String(s.qty)); else if(col==="__style") setEditVal(initialChar??s.styleNo); else setEditVal(initialChar??(s[col]||"")); };
  const commitText=()=>{ if(!editing) return; const f=editing.col==="__style"?"styleNo":editing.col; if(editing.mode==="text"||editing.mode===undefined){ if(!isDateCol(editing.col)) setField(editing.id,f,editVal); } setEditing(null); };
  const finishEditing=()=>{ if(!editing) return; if(editing.mode==="actual"||editing.mode==="rev") commitDate(); else commitText(); };

  // ---- selection range ----
  const rowIndex=(id)=>rows.findIndex(r=>r.s.id===id);
  const colIndex=(col)=>navCols.indexOf(col);
  const rect=()=>{ if(!sel) return null; const aR=rowIndex(sel.id), aC=colIndex(sel.col); const f=focus||sel; const fR=rowIndex(f.id), fC=colIndex(f.col); return { r1:Math.min(aR,fR), r2:Math.max(aR,fR), c1:Math.min(aC,fC), c2:Math.max(aC,fC) }; };
  const selKeys=useMemo(()=>{ const R=rect(); const set=new Set(); if(!R) return set; for(let r=R.r1;r<=R.r2;r++){ for(let c=R.c1;c<=R.c2;c++){ if(rows[r]) set.add(`${rows[r].s.id}:${navCols[c]}`); } } return set; },[sel,focus,rows,navCols]);
  const onCellClick=(e,id,col)=>{ e.stopPropagation(); if(gridRef.current) gridRef.current.focus({preventScroll:true}); if(editing){ if(editing.id===id&&editing.col===col) return; finishEditing(); } if(e.shiftKey&&sel){ setFocus({id,col}); scrollToCell(id,col); return; } if(sel&&sel.id===id&&sel.col===col&&!editing&&isEditableCol(col)&&ROLES[role].canEdit(ownerOfCol(col))){ startEdit(id,col); return; } setSel({id,col}); setFocus(null); };

  const moveAnchor=(dr,dc)=>{ if(!sel) return; let r=rowIndex(sel.id)+dr, c=colIndex(sel.col)+dc; r=Math.min(Math.max(r,0),rows.length-1); c=Math.min(Math.max(c,0),navCols.length-1); if(rows[r]){ setSel({id:rows[r].s.id,col:navCols[c]}); setFocus(null); scrollToCell(rows[r].s.id,navCols[c]); } };
  const scrollToCell=(id,col)=>{ requestAnimationFrame(()=>{ const el=document.getElementById(`cell-${id}-${col}`); if(el) el.scrollIntoView({ inline:"nearest", block:"nearest" }); }); };
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
  const canPasteCell=(s,col)=>{ if(!isEditableCol(col)) return false; if(!ROLES[role].canEdit(ownerOfCol(col))) return false; if(STAGE_KEYS.includes(col)){ const st=STAGES.find(x=>x.key===col); if(!(st.flag===null||s[st.flag])) return false; } return true; };
  const doPaste=()=>{ if(!clip||!sel) return; pushHistory(); const R=rect(); const changes={}; const put=(id,col,val)=>{ (changes[id]=changes[id]||{})[col]=val; };
    if(clip.h===1&&clip.w===1){ const v=clip.values[0][0]; for(let r=R.r1;r<=R.r2;r++){ for(let c=R.c1;c<=R.c2;c++){ const row=rows[r]; const col=navCols[c]; if(row&&canPasteCell(row.s,col)) put(row.s.id,col,v); } } }
    else { for(let i=0;i<clip.h;i++){ for(let j=0;j<clip.w;j++){ const r=R.r1+i, c=R.c1+j; const row=rows[r]; const col=navCols[c]; if(row&&col&&canPasteCell(row.s,col)) put(row.s.id,col,clip.values[i][j]); } } }
    setStyles(prev=>prev.map(s=>{ const ch=changes[s.id]; if(!ch) return s; let ns={...s, actuals:{...s.actuals}}; Object.entries(ch).forEach(([col,val])=>{ if(STAGE_KEYS.includes(col)) ns.actuals[col]=val||undefined; else if(col==="qty") ns.qty=Number(val)||0; else if(col==="__style") ns.styleNo=val; else ns[col]=val; }); return ns; })); flash(); };

  // batch write a {id:{col:val}} change map into styles
  const writeChanges=(changes)=>{ setStyles(prev=>prev.map(s=>{ const ch=changes[s.id]; if(!ch) return s; let ns={...s, actuals:{...s.actuals}, revs:{...(s.revs||{})}}; Object.entries(ch).forEach(([col,val])=>{ if(STAGE_KEYS.includes(col)) ns.actuals[col]=val||undefined; else if(col==="qty") ns.qty=Number(val)||0; else if(col==="__style") ns.styleNo=val; else ns[col]=val; }); return ns; })); flash(); };
  const coerce=(col,raw)=>{ if(raw==null) return ""; const v=String(raw).trim(); if(isDateCol(col)){ const pt=parseTyped(v); if(pt!==false) return pt; const d=new Date(v); return isNaN(d)?"":iso(d); } return v; };
  const clearRange=()=>{ const R=rect(); if(!R) return; pushHistory(); const ch={}; for(let r=R.r1;r<=R.r2;r++){ for(let cc=R.c1;cc<=R.c2;cc++){ const row=rows[r]; const col=navCols[cc]; if(row&&canPasteCell(row.s,col)) (ch[row.s.id]=ch[row.s.id]||{})[col]= STAGE_KEYS.includes(col)?null:(col==="qty"?0:""); } } writeChanges(ch); };
  const applyFillHandle=()=>{ if(!fillFrom||!fillTo) return; const aR=rowIndex(fillFrom.id), aC=colIndex(fillFrom.col), tR=rowIndex(fillTo.id), tC=colIndex(fillTo.col); const r1=Math.min(aR,tR), r2=Math.max(aR,tR), c1=Math.min(aC,tC), c2=Math.max(aC,tC); const srcRow=rows[aR]; if(!srcRow) return; pushHistory(); const ch={}; for(let r=r1;r<=r2;r++){ for(let cc=c1;cc<=c2;cc++){ const row=rows[r]; const col=navCols[cc]; if(!row) continue; if(r===aR&&cc===aC) continue; const srcVal=getVal(srcRow.s, navCols[cc]); if(canPasteCell(row.s,col)) (ch[row.s.id]=ch[row.s.id]||{})[col]=srcVal; } } writeChanges(ch); };
  const autoFit=(col)=>{ let max=String(col==="__style"?"Style No":(INFO_COLS.find(x=>x.key===col)?.label||STAGES.find(x=>x.key===col)?.label||col)).length; rows.forEach(({s})=>{ const v=getVal(s,col); if(v!=null) max=Math.max(max,String(isDateCol(col)?fmt(parse(v)):v).length); }); setColW(p=>({ ...p, [col]:Math.max(48, Math.min(320, max*7+26)) })); };
  const onKeyDown=(e)=>{
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
  const jumpToEnter=(id,stageKey)=>{ const st=STAGES.find(x=>x.key===stageKey); setSel({id,col:stageKey}); setFocus(null); requestAnimationFrame(()=>{ const el=document.getElementById(`cell-${id}-${stageKey}`); if(el) el.scrollIntoView({ behavior:"smooth", inline:"center", block:"nearest" }); }); if(st&&ROLES[role].canEdit(st.owner)) setTimeout(()=>beginDate(id,stageKey,"actual"),260); };

  const handleCopy=(e)=>{ const R=rect(); if(!R) return; const lines=[]; let any=false; for(let r=R.r1;r<=R.r2;r++){ const cells=[]; for(let cc=R.c1;cc<=R.c2;cc++){ let v=rows[r]?getVal(rows[r].s,navCols[cc]):""; if(isDateCol(navCols[cc])&&v) v=fmtTyped(v); if(v) any=true; cells.push(v??""); } lines.push(cells.join("\t")); } if(any){ const tsv=lines.join("\n"); try{ e.clipboardData.setData("text/plain",tsv); e.preventDefault(); }catch(err){} setClip({ values:lines.map(l=>l.split("\t")), h:lines.length, w:lines[0].split("\t").length }); flash(); } };
  const handlePaste=(e)=>{ if(!sel) return; let txt=""; try{ txt=e.clipboardData.getData("text/plain"); }catch(err){} if(!txt){ doPaste(); return; } e.preventDefault(); const grid=txt.replace(/\r/g,"").replace(/\n$/,"").split("\n").map(l=>l.split("\t")); pushHistory(); const aR=rowIndex(sel.id), aC=colIndex(sel.col); const ch={}; if(grid.length===1&&grid[0].length===1){ const R=rect(); for(let r=R.r1;r<=R.r2;r++){ for(let cc=R.c1;cc<=R.c2;cc++){ const row=rows[r]; const col=navCols[cc]; if(row&&canPasteCell(row.s,col)) (ch[row.s.id]=ch[row.s.id]||{})[col]=coerce(col,grid[0][0]); } } } else { for(let i=0;i<grid.length;i++){ for(let j=0;j<grid[i].length;j++){ const row=rows[aR+i]; const col=navCols[aC+j]; if(row&&col&&canPasteCell(row.s,col)) (ch[row.s.id]=ch[row.s.id]||{})[col]=coerce(col,grid[i][j]); } } } writeChanges(ch); };
  const cellKey=(id,col)=>`${id}:${col}`;
  const applyFill=(color)=>{ if(!sel) return; pushHistory(); const R=rect(); setFills(p=>{ const n={...p}; for(let r=R.r1;r<=R.r2;r++){ for(let c=R.c1;c<=R.c2;c++){ if(!rows[r]) continue; const k=`${rows[r].s.id}:${navCols[c]}`; if(color==="") delete n[k]; else n[k]=color; } } return n; }); flash(); };
  const saveNote=()=>{ if(!sel) return; pushHistory(); setNotes(p=>{ const n={...p}; const k=cellKey(sel.id,sel.col); if(noteText.trim()==="") delete n[k]; else n[k]=noteText.trim(); return n; }); setNoteEditing(false); setNoteText(""); flash(); };
  const beginNote=()=>{ if(!sel) return; setNoteText(notes[cellKey(sel.id,sel.col)]||""); setNoteEditing(true); };

  const distinctFor=(col)=>{ const set=new Set(); computed.forEach(({s,c})=>set.add(valueFor(s,c,col))); return [...set].sort((a,b)=> a==="(Blanks)"?1:b==="(Blanks)"?-1:(a>b?1:a<b?-1:0)); };
  const filterProps=(col)=>({ filterActive: !!colFilters[col], filterOpen: filterCol===col, filterValues: filterCol===col?distinctFor(col):null, filterAllowed: colFilters[col]||null,
    onToggleFilter:()=>{ finishEditing(); setFilterCol(p=>p===col?null:col); },
    onSetFilter:(arr)=>setColFilters(f=>{ const n={...f}; if(!arr) delete n[col]; else n[col]=arr; return n; }),
    onCloseFilter:()=>setFilterCol(null) });
  const funnel=useMemo(()=>{ const b={ "Pre-Fit":0,"Fit/Print":0,"Lab Dip":0,"Fabric IH":0,"PP":0,"Released":0 }; computed.forEach(({c})=>{ if(c.released) b["Released"]++; else { const k=c.nextPending.key; if(k==="techpack") b["Pre-Fit"]++; else if(["fitSend","fitAppr","artwork","artAppr","strikeOff","soAppr"].includes(k)) b["Fit/Print"]++; else if(["labDip","labAppr"].includes(k)) b["Lab Dip"]++; else if(k==="fabricIH") b["Fabric IH"]++; else b["PP"]++; } }); return b; },[computed]);

  const requiredMissing=()=>{ const m=[]; if(!newRow.styleNo.trim()) m.push("Style No"); if(!newRow.orderNo.trim()) m.push("Order No"); if(!newRow.ordRec) m.push("Order Date"); if(!newRow.delivery) m.push("Delivery Date"); return m; };
  const addNewStyle=async()=>{ const miss=requiredMissing(); if(miss.length){ setNewError("Required: "+miss.join(", ")); return; } setNewError(""); pushHistory();
    const base={ order_no:newRow.orderNo||"", style_no:newRow.styleNo.trim(), sample_fit:newRow.sampleFit||"", family:newRow.family||"", colour:newRow.colour||"", owner:newRow.owner||"", set_id:"", set_role:"", age:"", qty:Number(newRow.qty)||0, order_date:newRow.ordRec||iso(TODAY), delivery_date:newRow.delivery||"2026-07-15", fit_req:newRow.fitReq, print_req:newRow.printReq, so_req:newRow.soReq, pp_bypass:newRow.ppBypass, lab_dip_req:newRow.labDipReq, pp_needed:newRow.ppNeeded, remarks:"" };
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
        {editingThis ? (<input autoFocus value={editVal} onClick={e=>e.stopPropagation()} onChange={e=>setEditVal(col.key==="qty"?e.target.value.replace(/[^0-9]/g,""):e.target.value)} onBlur={commitText} style={{ width:Math.max(40,col.w-16), fontFamily:"inherit", fontSize:11, border:"1px solid #1d4ed8", outline:"none", padding:"1px 3px" }}/>) : (val===""||val==null ? <span style={{color:"#ccc"}}>—</span> : String(val))}
        <NoteTri k={k}/><FillHandle id={s.id} col={col.key}/>
      </td>
    );
  };

  return (
    <div ref={gridRef} tabIndex={0} onKeyDown={onKeyDown} onCopy={handleCopy} onPaste={handlePaste} onMouseDown={(e)=>{ if(editing) return; if(e.target.closest && (e.target.closest("input")||e.target.closest("button")||e.target.closest("th"))) return; const td=e.target.closest && e.target.closest('td[id^="cell-"]'); if(!td) return; const m=td.id.match(/^cell-(\d+)-(.+)$/); if(!m||e.shiftKey) return; e.preventDefault(); setSel({ id:Number(m[1]), col:m[2] }); setFocus(null); selectingRef.current=true; setDragSel(true); }} onMouseUp={()=>{ if(filling){ applyFillHandle(); setFilling(false); setFillFrom(null); setFillTo(null); } if(selectingRef.current){ selectingRef.current=false; setDragSel(false); } }} onMouseOver={(e)=>{ const td=e.target.closest && e.target.closest('td[id^="cell-"]'); if(!td) return; const m=td.id.match(/^cell-(\d+)-(.+)$/); if(!m) return; if(filling){ setFillTo({ id:Number(m[1]), col:m[2] }); return; } if(selectingRef.current){ setFocus({ id:Number(m[1]), col:m[2] }); } }} onClick={()=>{ finishEditing(); setFillOpen(false); setColsOpen(false); setFilterCol(null); }}
      style={{ minHeight:"100vh", background:"#f4f0e8", fontFamily:"'JetBrains Mono', monospace", color:"#1a1a1a", paddingBottom:80, outline:"none" }}>
      <style>{FONT}</style>

      <div style={{ background:"#1a1a1a", color:"#f4f0e8", padding:"14px 22px", display:"flex", alignItems:"center", justifyContent:"space-between", borderBottom:"3px solid #d97706" }}>
        <div style={{ display:"flex", alignItems:"baseline", gap:12 }}><span style={{ fontFamily:"'Archivo',sans-serif", fontWeight:800, fontSize:20, letterSpacing:-0.5 }}>MERCH<span style={{color:"#d97706"}}>·</span>TRACKER</span><span style={{ fontSize:10, color:"#9a958c", letterSpacing:1 }}>PRE-PRODUCTION · SPREADSHEET GRID · PROTOTYPE</span></div>
        <div style={{ display:"flex", alignItems:"center", gap:14 }}><span style={{ fontSize:11, color: saveState==="error"?"#e8746b":saveState==="saving"?"#d9b46a":saveState==="saved"?"#7fd1a8":"#6a665e" }}>{saveState==="error"?"⚠ save failed":saveState==="saving"?"… saving":saveState==="saved"?"● saved to cloud":"○ connected"}</span><div style={{ display:"flex", border:"1px solid #4a463e" }}>{Object.keys(ROLES).map(r=>(<button key={r} onClick={(e)=>{ e.stopPropagation(); setRole(r); }} style={{ fontFamily:"inherit", fontSize:10, padding:"5px 9px", cursor:"pointer", border:"none", background:role===r?"#d97706":"transparent", color:role===r?"#1a1a1a":"#cfc9bf", fontWeight:role===r?700:400 }}>{ROLES[r].label}</button>))}</div></div>
      </div>

      <div style={{ display:"flex", padding:"12px 22px 0", flexWrap:"wrap" }}>
        {Object.entries(funnel).map(([k,v],i,arr)=>(<div key={k} style={{ flex:1, minWidth:90, background:"#fff", border:"1px solid #1a1a1a", borderRight:i===arr.length-1?"1px solid #1a1a1a":"none", padding:"8px 10px" }}><div style={{ fontSize:22, fontWeight:700, lineHeight:1, fontFamily:"'Archivo',sans-serif", color:k==="Released"?"#1f6f54":k==="Fabric IH"?"#c0392b":"#1a1a1a" }}>{v}</div><div style={{ fontSize:9, color:"#888", marginTop:3, letterSpacing:0.5, textTransform:"uppercase" }}>{k}</div></div>))}
      </div>

      <div style={{ display:"flex", gap:10, alignItems:"center", padding:"12px 22px 6px", flexWrap:"wrap" }}>
        <div style={{ display:"flex", alignItems:"center", gap:6, background:"#fff", border:"1px solid #1a1a1a", padding:"5px 9px" }}><Filter size={13}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="search style / colour / fit / order…" onClick={e=>e.stopPropagation()} style={{ border:"none", outline:"none", fontFamily:"inherit", fontSize:12, width:180, background:"transparent" }}/></div>
        <div style={{ display:"flex", border:"1px solid #1a1a1a" }}>{["All","At Risk","On Track","Released"].map(f=>(<button key={f} onClick={(e)=>{ e.stopPropagation(); setStatusFilter(f); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"none", borderRight:f!=="Released"?"1px solid #1a1a1a":"none", background:statusFilter===f?"#1a1a1a":"#fff", color:statusFilter===f?"#f4f0e8":"#1a1a1a" }}>{f}</button>))}</div>
        <div style={{ position:"relative" }}><button onClick={(e)=>{ e.stopPropagation(); finishEditing(); setFillOpen(o=>!o); setColsOpen(false); }} disabled={role==="Viewer"} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:role==="Viewer"?"not-allowed":"pointer", border:"1px solid #1a1a1a", background:"#d97706", color:"#1a1a1a", fontWeight:700, display:"flex", alignItems:"center", gap:6, opacity:role==="Viewer"?0.4:1 }}><Copy size={13}/> Fill date → {rows.length}</button>{fillOpen && role!=="Viewer" && (<FillPanel count={rows.length} onClose={()=>setFillOpen(false)} onApply={(key,val)=>{ setStyles(prev=>prev.map(s=>rows.some(r=>r.s.id===s.id)?{...s,actuals:{...s.actuals,[key]:val||undefined}}:s)); flash(); setFillOpen(false); }}/>)}</div>
        <div style={{ position:"relative" }}><button onClick={(e)=>{ e.stopPropagation(); finishEditing(); setColsOpen(o=>!o); setFillOpen(false); }} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#fff", display:"flex", alignItems:"center", gap:6 }}><Columns3 size={13}/> Columns {hidden.size>0?`(${hidden.size} hidden)`:""}</button>
          {colsOpen && (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:0, marginTop:4, zIndex:70, background:"#fff", border:"1px solid #1a1a1a", boxShadow:"4px 4px 0 #1a1a1a", padding:10, width:230, maxHeight:300, overflowY:"auto" }}><div style={{ fontSize:10, fontWeight:700, marginBottom:6 }}>Show / hide columns</div>{[...INFO_COLS,{key:"remarks",label:"Remarks / Delays"},...STAGES].map(col=>(<label key={col.key} style={{ display:"flex", alignItems:"center", gap:6, fontSize:10, padding:"2px 0", cursor:"pointer" }}><input type="checkbox" checked={!hidden.has(col.key)} onChange={()=>setHidden(p=>{ const n=new Set(p); n.has(col.key)?n.delete(col.key):n.add(col.key); return n; })}/>{col.label}</label>))}{hidden.size>0 && <button onClick={()=>setHidden(new Set())} style={{ ...chip, marginTop:6, width:"100%" }}>Show all</button>}</div>)}
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
        <button onClick={(e)=>{ e.stopPropagation(); loadShared(); }} title="reload shared data (pull latest edits)" style={{ fontFamily:"inherit", fontSize:11, padding:"6px 11px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#fff", display:"flex", alignItems:"center", gap:6 }}><RotateCcw size={13}/> Sync</button>
        <span style={{ fontSize:10, color:"#999", marginLeft:"auto" }}>{sort.col?<>sorted by <b>{sort.col==="__style"?"Style":(INFO_COLS.find(c=>c.key===sort.col)?.label||STAGES.find(s=>s.key===sort.col)?.label||(sort.col==="remarks"?"Remarks":sort.col))}</b> {sort.dir>0?"↑":"↓"}</>:"shift-click / shift-arrows = range · Ctrl/Cmd C & V = copy/paste"}</span>
      </div>

      <div style={{ display:"flex", gap:10, alignItems:"center", padding:"0 22px 10px", flexWrap:"wrap", fontSize:10, color:"#777" }}>
        <button onClick={(e)=>{ e.stopPropagation(); selectAll(); }} title="select all" style={{ fontFamily:"inherit", fontSize:10, padding:"4px 8px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#fff" }}>⌖ all</button><Droplet size={13}/><span>fill:</span>
        {FILL_SWATCHES.map((sw,i)=>(<button key={i} onClick={(e)=>{ e.stopPropagation(); applyFill(sw); }} disabled={!sel} title={sw===""?"clear fill":sw} style={{ width:18, height:18, cursor:sel?"pointer":"not-allowed", border:"1px solid #1a1a1a", background:sw===""?"#fff":sw, position:"relative", opacity:sel?1:0.4 }}>{sw===""?<X size={11} style={{position:"absolute",top:2,left:2}}/>:null}</button>))}
        <span style={{ marginLeft:10, position:"relative" }}>
          <button onClick={(e)=>{ e.stopPropagation(); beginNote(); }} disabled={!sel} style={{ fontFamily:"inherit", fontSize:10, padding:"4px 9px", cursor:sel?"pointer":"not-allowed", border:"1px solid #1a1a1a", background:"#fff", display:"inline-flex", alignItems:"center", gap:5, opacity:sel?1:0.4 }}><MessageSquare size={12}/> {sel&&notes[cellKey(sel.id,sel.col)]?"edit comment":"add comment"}</button>
          {noteEditing && sel && (<span onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:0, marginTop:4, zIndex:70, background:"#fff", border:"1px solid #1a1a1a", boxShadow:"4px 4px 0 #1a1a1a", padding:8, width:220, display:"block" }}><textarea autoFocus value={noteText} onChange={e=>setNoteText(e.target.value)} placeholder="comment on this cell…" style={{ width:"100%", height:50, fontFamily:"inherit", fontSize:11, border:"1px solid #ccc", outline:"none", resize:"none" }}/><span style={{ display:"flex", gap:6, marginTop:6 }}><button onClick={saveNote} style={{ ...chip, flex:1, background:"#d97706" }}>Save</button><button onClick={()=>{ setNoteEditing(false); setNoteText(""); }} style={chip}>Cancel</button></span></span>)}
        </span>
        <span style={{ marginLeft:8 }}>{clip?<span style={{color:"#2563a6"}}>📋 {clip.h}×{clip.w} copied — select & Ctrl/Cmd+V to paste</span>:(sel?<>selected: <b>{styles.find(s=>s.id===sel.id)?.styleNo} · {sel.col==="__style"?"Style":(INFO_COLS.find(c=>c.key===sel.col)?.label||STAGES.find(s=>s.key===sel.col)?.label||sel.col)}</b>{focus?" (range)":""}</>:"click a cell to format / comment")}</span>
      </div>

      <div style={{ overflowX:"auto", padding:"0 22px" }}>
        <table role="grid" aria-label="Pre-production tracker grid. Arrow keys to move, Escape to exit, Tab to leave the grid." style={{ borderCollapse:"separate", borderSpacing:0, fontSize:11, tableLayout:"fixed", userSelect:dragSel?"none":"auto" }}>
          <colgroup>
            <col style={{ width:widthOf("__style") }}/>
            {visInfo.map(c=><col key={c.key} style={{ width:widthOf(c.key) }}/>)}
            {visStages.map(st=><col key={st.key} style={{ width:widthOf(st.key) }}/>)}
            {remarksVis && <col style={{ width:widthOf("remarks") }}/>}
          </colgroup>
          <thead><tr role="row">
            <Th col="__style" label="Style No" sort={sort} onSort={clickHeader} width={widthOf("__style")} onResize={onResize} onAutoFit={autoFit} {...filterProps("__style")} sticky left={0} z={6}/>
            {visInfo.map(c=><Th key={c.key} col={c.key} label={c.label} sort={sort} onSort={clickHeader} width={widthOf(c.key)} onResize={onResize} onAutoFit={autoFit} {...filterProps(c.key)} sticky={isFrozen(c.key)} left={isFrozen(c.key)?leftOf(c.key):undefined} z={5}/>)}
            {visStages.map(st=><Th key={st.key} col={st.key} label={st.label} sort={sort} onSort={clickHeader} width={widthOf(st.key)} onResize={onResize} onAutoFit={autoFit} {...filterProps(st.key)}/>)}
            {remarksVis && <Th col="remarks" label={REMARK_COL.label} sort={sort} onSort={clickHeader} width={widthOf("remarks")} onResize={onResize} onAutoFit={autoFit} {...filterProps("remarks")}/>}
          </tr></thead>
          <tbody>
            {rows.map(({s,c},rowIdx)=>{ const t=TONE_STYLE[c.tone]; const sk=cellKey(s.id,"__style"); const styBg=bgFor(s.id,"__style","#fff"); return (
              <tr key={s.id} role="row">
                <td id={`cell-${s.id}-__style`} onClick={(e)=>onCellClick(e,s.id,"__style")} onDoubleClick={(e)=>{ e.stopPropagation(); startEdit(s.id,"__style"); }} style={{ position:"sticky", left:0, zIndex:5, background:styBg, border:"1px solid #ddd", padding:"6px 9px", overflow:"hidden", boxShadow:ringFor(s.id,"__style"), cursor:"cell" }}>
                  {editing&&editing.id===s.id&&editing.col==="__style" ? (<input autoFocus value={editVal} onClick={e=>e.stopPropagation()} onChange={e=>setEditVal(e.target.value)} onBlur={commitText} style={{ width:150, fontFamily:"inherit", fontSize:11, fontWeight:700, border:"1px solid #1d4ed8", outline:"none", padding:"1px 3px" }}/>) : <div style={{ fontWeight:700, display:"flex", alignItems:"center", gap:6 }}><span onClick={(e)=>{ e.stopPropagation(); selectRow(s.id); }} title="select row" style={{ fontSize:8, color:"#bbb", cursor:"pointer", minWidth:14 }}>{rowIdx+1}</span>{s.styleNo}</div>}
                  <div style={{ fontSize:9, color:"#999", whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis", maxWidth:188, marginTop:1 }}>{s.colour}</div>
                  <div style={{ display:"flex", gap:3, marginTop:4, flexWrap:"wrap" }}>{FLAG_DEFS.map(f=>{ const on=!!s[f.key]; return (<button key={f.key} title={f.title} onClick={(e)=>{ e.stopPropagation(); if(role!=="Viewer") toggleFlag(s.id,f.key); }} style={{ fontFamily:"inherit", fontSize:8.5, fontWeight:700, letterSpacing:0.3, padding:"2px 5px", cursor:role==="Viewer"?"not-allowed":"pointer", lineHeight:1.3, border:`1px solid ${on?"#1a1a1a":"#cfcabf"}`, background:on?"#1a1a1a":"transparent", color:on?"#f4f0e8":"#bbb", opacity:role==="Viewer"?0.5:1 }}>{f.short}</button>); })}</div>
                  <NoteTri k={sk}/>
                </td>

                {visInfo.map(col=>{
                  if(col.kind==="text"||col.kind==="num") return renderEditable(s,col);
                  const k=cellKey(s.id,col.key);
                  if(col.kind==="date"){ const bg=bgFor(s.id,col.key,"#fff"); return (<td key={col.key} id={`cell-${s.id}-${col.key}`} onClick={(e)=>onCellClick(e,s.id,col.key)} onDoubleClick={(e)=>{ e.stopPropagation(); if(role!=="Viewer") beginDate(s.id,col.key,"actual"); }} style={{ border:"1px solid #ddd", padding:"6px 9px", whiteSpace:"nowrap", boxShadow:ringFor(s.id,col.key), cursor:"cell", position:"relative", overflow:(editing&&editing.id===s.id&&editing.col===col.key)?"visible":"hidden", background:bg, ...freezeStyle(col.key,bg) }}>{fmt(parse(s[col.key]))||<span style={{color:"#ccc"}}>—</span>}{editing&&editing.id===s.id&&editing.col===col.key && dateEditor(s.id,col.key,editing.mode)}<NoteTri k={k}/><FillHandle id={s.id} col={col.key}/></td>); }
                  let content=null;
                  if(col.kind==="branch"){ const b=col.branch==="fit"?c.fitBranch:col.branch==="print"?c.printBranch:col.branch==="fabric"?c.fabricBranch:c.ppBranch; const canJump=b.tone!=="na"&&!c.released; content=<BranchPill b={b} onJump={canJump?()=>jumpToEnter(s.id,branchTarget(s,c,col.branch)):null}/>; }
                  else if(col.key==="overall") content=<span style={{ display:"inline-flex", alignItems:"center", gap:5, background:t.bg, color:t.fg, padding:"2px 7px", fontSize:10, fontWeight:700 }}><span style={{ width:6,height:6,borderRadius:"50%", background:t.dot }}/>{c.status}</span>;
                  else if(col.key==="fabricCD") content=<span style={{ fontWeight:700, color:BR_TONE[c.fabricCountdown.tone].fg }}>{c.fabricCountdown.txt}</span>;
                  else if(col.key==="proj") content=<span title={`release gate (30wd before delivery): ${fmt(c.releaseGate)}`} style={{ fontWeight:600, color:c.projTone==="late"?"#c0392b":c.projTone==="warn"?"#7a560f":c.projTone==="done"?"#888":"#1f6f54" }}>{fmt(c.projRelease)}{c.projTone==="late"&&!c.released?" ⚠":c.projTone==="ok"?" ✓":""}</span>;
                  else if(col.key==="pct") content=(<div style={{ display:"flex", alignItems:"center", gap:5 }}><div style={{ flex:1, height:6, background:"#eee", position:"relative", minWidth:34 }}><div style={{ position:"absolute", left:0, top:0, bottom:0, width:`${c.pct}%`, background:c.pct===100?"#1f6f54":"#d97706" }}/></div><span style={{ fontSize:9, color:"#666", width:26, textAlign:"right" }}>{c.pct}%</span></div>);
                  else if(col.key==="chase") content=(!c.chaseOwners||c.chaseOwners.length===0)?<span style={{color:"#ccc"}}>—</span>:(<span style={{ display:"flex", gap:3, flexWrap:"wrap" }}>{c.chaseOwners.map(o=>(<span key={o.owner} title={`${o.owner}: ${o.count} open ${o.count>1?"branches":"branch"}`} style={{ fontSize:9, fontWeight:700, padding:"2px 5px", background:(OWNER_COLOR[o.owner]||"#888")+"22", color:OWNER_COLOR[o.owner]||"#555", whiteSpace:"nowrap" }}>{o.owner}{o.count>1?`×${o.count}`:""}</span>))}</span>);
                  else if(col.key==="float") content=<span style={{ fontWeight:700, color:c.float<0?"#c0392b":"#1f6f54" }}>{c.float==null?"—":`${c.float>0?"+":""}${c.float}d`}</span>;
                  else if(col.key==="idle") content=<span style={{ color:c.idle>=7?"#c0392b":"#888" }}>{c.idle==null?"—":`${c.idle}d`}</span>;
                  const bg=bgFor(s.id,col.key,"#fff");
                  return <td key={col.key} id={`cell-${s.id}-${col.key}`} onClick={(e)=>onCellClick(e,s.id,col.key)} style={{ border:"1px solid #ddd", padding:"6px 9px", whiteSpace:"nowrap", boxShadow:ringFor(s.id,col.key), cursor:"default", background:bg, position:"relative", overflow:"hidden", ...freezeStyle(col.key,bg) }}>{content}<NoteTri k={k}/><FillHandle id={s.id} col={col.key}/></td>;
                })}

                {visStages.map(st=>{
                  const applies=st.flag===null||s[st.flag];
                  const cs=c.stages.find(x=>x.key===st.key);
                  const isNext=c.nextPending&&c.nextPending.key===st.key&&applies;
                  const editable=applies&&ROLES[role].canEdit(st.owner);
                  const k=cellKey(s.id,st.key);
                  if(!applies){ const bg=bgFor(s.id,st.key,"#f3f1ec"); return <td key={st.key} id={`cell-${s.id}-${st.key}`} onClick={(e)=>onCellClick(e,s.id,st.key)} style={{ border:"1px solid #ddd", background:bg, color:"#ccc", textAlign:"center", padding:"6px 9px", boxShadow:ringFor(s.id,st.key), position:"relative", overflow:"hidden" }}>—<NoteTri k={k}/></td>; }
                  const hasRev=cs&&cs.rev&&!cs.done;
                  const bg=bgFor(s.id,st.key,isNext?"#fff7ec":"#fff");
                  return (
                    <td key={st.key} id={`cell-${s.id}-${st.key}`} onClick={(e)=>onCellClick(e,s.id,st.key)} onDoubleClick={(e)=>{ e.stopPropagation(); if(editable) beginDate(s.id,st.key,"actual"); }}
                      style={{ border:"1px solid #ddd", padding:0, position:"relative", overflow:(editing&&editing.id===s.id&&editing.col===st.key)?"visible":"hidden", background:bg, boxShadow:ringFor(s.id,st.key)||(isNext?"inset 0 0 0 2px #d97706":null), cursor:editable?"cell":"default" }}>
                      <div style={{ minHeight:38, padding:"4px 8px", fontSize:11, color:cs.actual?"#1a1a1a":"#bbb" }}>
                        {cs.actual ? (<span style={{ display:"flex", alignItems:"center", gap:4 }}><Check size={11} color={OWNER_COLOR[st.owner]}/>{fmt(cs.actual)}</span>) : (
                          <span style={{ display:"flex", flexDirection:"column", lineHeight:1.2 }}>
                            <span style={{ fontSize:9, color:hasRev?"#6d4aab":isNext?"#d97706":"#c4c0b8" }}>{hasRev?"rev":st.cutoff?"cutoff":"plan"} {fmt(hasRev?cs.rev:cs.plan)}</span>
                            {editable?<span style={{ fontSize:9, color:isNext?"#d97706":"#c4c0b8", fontWeight:isNext?700:400 }}>{isNext?"▸ enter":st.cutoff?"log arrival":"—"}</span>:<span style={{ fontSize:9, color:"#ccc", display:"flex", alignItems:"center", gap:3 }}><Lock size={8}/>locked</span>}
                          </span>
                        )}
                      </div>
                      {editable && !cs.actual && (<button title="set revised plan date" onClick={(e)=>{ e.stopPropagation(); beginDate(s.id,st.key,"rev"); }} style={{ position:"absolute", top:1, right:1, border:"none", background:"transparent", cursor:"pointer", padding:1, lineHeight:0 }}><RotateCcw size={10} color="#6d4aab"/></button>)}
                      {editing&&editing.id===s.id&&editing.col===st.key&&editable && dateEditor(s.id,st.key,editing.mode)}
                      <NoteTri k={k}/><FillHandle id={s.id} col={st.key}/>
                    </td>
                  );
                })}

                {remarksVis && renderEditable(s,REMARK_COL)}
              </tr>
            ); })}

            {role!=="Viewer" && (
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
    </div>
  );
}

function Th({ col, label, sort, onSort, sticky, left, z, width, onResize, onAutoFit, filterActive, filterOpen, filterValues, filterAllowed, onToggleFilter, onSetFilter, onCloseFilter }){
  const active=sort.col===col;
  const startDrag=(e)=>{ e.preventDefault(); e.stopPropagation(); const sx=e.clientX, sw=width||80; const move=(ev)=>onResize&&onResize(col, sw+(ev.clientX-sx)); const up=()=>{ window.removeEventListener("mousemove",move); window.removeEventListener("mouseup",up); }; window.addEventListener("mousemove",move); window.addEventListener("mouseup",up); };
  return (<th role="columnheader" aria-sort={active?(sort.dir>0?"ascending":"descending"):"none"} style={{ position:"sticky", top:0, left:sticky?left:undefined, zIndex:sticky?(z||5):3, background:active?"#d97706":"#1a1a1a", color:active?"#1a1a1a":"#f4f0e8", padding:"8px 9px", textAlign:"left", fontWeight:600, fontSize:9.5, letterSpacing:0.4, textTransform:"uppercase", whiteSpace:"nowrap", overflow:"visible", border:"1px solid #3a362e", userSelect:"none" }}>
    <span style={{ display:"flex", alignItems:"center", gap:3 }}>
      <span onClick={(e)=>{ e.stopPropagation(); onSort(col); }} title="click to sort" style={{ display:"inline-flex", alignItems:"center", gap:3, cursor:"pointer", flex:1, overflow:"hidden", textOverflow:"ellipsis" }}>{label}{active?(sort.dir>0?<ChevronUp size={11}/>:<ChevronDown size={11}/>):null}</span>
      <span onClick={(e)=>{ e.stopPropagation(); onToggleFilter&&onToggleFilter(); }} title="filter" style={{ cursor:"pointer", display:"inline-flex", padding:"0 1px", color: filterActive?"#1a1a1a":(active?"#7a4a08":"#9a958c") }}><Filter size={10} fill={filterActive?"currentColor":"none"}/></span>
    </span>
    {filterOpen && <FilterMenu values={filterValues||[]} allowed={filterAllowed} onSet={onSetFilter} onClose={onCloseFilter}/>}
    <span onMouseDown={startDrag} onDoubleClick={(e)=>{ e.stopPropagation(); onAutoFit&&onAutoFit(col); }} onClick={(e)=>e.stopPropagation()} title="drag to resize · double-click to auto-fit" style={{ position:"absolute", top:0, right:0, bottom:0, width:7, cursor:"col-resize", zIndex:2 }}/>
  </th>);
}
function FilterMenu({ values, allowed, onSet, onClose }){
  const [q,setQ]=useState("");
  const masterRef=useRef(null); const menuRef=useRef(null); const [flip,setFlip]=useState(false);
  const isOn=(v)=> !allowed || allowed.includes(v);
  const allOn = !allowed;                       // no filter = every value shown
  const noneOn = allowed && allowed.length===0; // nothing selected = grid empty
  const shown=values.filter(v=> v.toLowerCase().includes(q.toLowerCase()));
  const toggle=(v)=>{ const cur = allowed? new Set(allowed): new Set(values); if(cur.has(v)) cur.delete(v); else cur.add(v); const arr=[...cur]; onSet(arr.length===values.length? null : arr); };
  const toggleAll=()=> onSet(allOn? [] : null);  // checked -> deselect all ; otherwise select all
  useEffect(()=>{ if(masterRef.current) masterRef.current.indeterminate = !allOn && !noneOn; },[allOn,noneOn]);
  useEffect(()=>{ const el=menuRef.current; if(el){ const r=el.getBoundingClientRect(); if(r.right>window.innerWidth-8) setFlip(true); else if(r.left<8) setFlip(false); } },[]);
  return (
    <div ref={menuRef} onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:flip?"auto":0, right:flip?0:"auto", marginTop:2, zIndex:95, background:"#fff", color:"#1a1a1a", border:"1px solid #1a1a1a", boxShadow:"4px 4px 0 #1a1a1a", padding:8, width:210, textTransform:"none", letterSpacing:0, fontWeight:400 }}>
      <input autoFocus value={q} onChange={e=>setQ(e.target.value)} placeholder="search values…" style={{ width:"100%", fontFamily:"inherit", fontSize:11, padding:"4px 6px", border:"1px solid #ccc", outline:"none", marginBottom:6 }}/>
      <label style={{ display:"flex", alignItems:"center", gap:6, fontSize:10, fontWeight:700, padding:"3px 0", cursor:"pointer", borderBottom:"1px solid #eee", marginBottom:4 }}><input ref={masterRef} type="checkbox" checked={allOn} onChange={toggleAll}/>(Select All)</label>
      <div style={{ maxHeight:180, overflowY:"auto" }}>
        {shown.map(v=>(<label key={v} style={{ display:"flex", alignItems:"center", gap:6, fontSize:10, padding:"2px 0", cursor:"pointer" }}><input type="checkbox" checked={isOn(v)} onChange={()=>toggle(v)}/><span style={{ whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{v}</span></label>))}
        {shown.length===0 && <div style={{ fontSize:10, color:"#999", padding:"4px 0" }}>no matches</div>}
      </div>
      {noneOn && <div style={{ fontSize:9, color:"#c0392b", marginTop:4 }}>Nothing selected — no rows shown.</div>}
      <div style={{ display:"flex", gap:6, marginTop:6 }}>
        <button onClick={()=>onSet(null)} style={{ ...chip, flex:1, fontSize:9 }}>Clear filter</button>
        <button onClick={onClose} style={{ ...chip, flex:1, fontSize:9, background:"#1a1a1a", color:"#f4f0e8" }}>Done</button>
      </div>
    </div>
  );
}
function FillPanel({ count, onApply, onClose }){
  const [key,setKey]=useState("labAppr"); const [val,setVal]=useState(iso(TODAY));
  return (<div onClick={e=>e.stopPropagation()} style={{ position:"absolute", top:"100%", left:0, marginTop:4, zIndex:70, background:"#fff", border:"1px solid #1a1a1a", boxShadow:"4px 4px 0 #1a1a1a", padding:12, width:280 }}><div style={{ fontSize:11, fontWeight:700, marginBottom:8 }}>Set one date across {count} filtered styles</div><label style={{ fontSize:10, color:"#888" }}>Stage</label><select value={key} onChange={e=>setKey(e.target.value)} style={{ width:"100%", fontFamily:"inherit", fontSize:11, padding:5, marginBottom:8, border:"1px solid #1a1a1a" }}>{STAGES.filter(s=>!s.cutoff).map(s=><option key={s.key} value={s.key}>{s.label}</option>)}</select><label style={{ fontSize:10, color:"#888" }}>Date</label><input type="date" value={val} onChange={e=>setVal(e.target.value)} style={{ width:"100%", fontFamily:"inherit", fontSize:11, padding:5, marginBottom:10, border:"1px solid #1a1a1a" }}/><div style={{ display:"flex", gap:8 }}><button onClick={()=>onApply(key,val)} style={{ flex:1, fontFamily:"inherit", fontSize:11, fontWeight:700, padding:6, cursor:"pointer", border:"1px solid #1a1a1a", background:"#d97706" }}>Apply to all filtered</button><button onClick={onClose} style={{ fontFamily:"inherit", fontSize:11, padding:"6px 10px", cursor:"pointer", border:"1px solid #1a1a1a", background:"#f4f0e8" }}><X size={12}/></button></div></div>);
}
const ndCell={ border:"1px dashed #e8dcc2", padding:"6px 9px", whiteSpace:"nowrap" };
const ndInput=(w)=>({ border:"none", outline:"none", background:"transparent", fontFamily:"'JetBrains Mono', monospace", fontSize:10, width:w });

import { useState, useCallback, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";
import wintLogo from "./assets/wint-logo.png";

const SUPABASE_URL = "https://ofalmwbegelirbfbsjqi.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9mYWxtd2JlZ2VsaXJiZmJzanFpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg1NTEwODMsImV4cCI6MjA5NDEyNzA4M30.ySOyXBKrOu37uGO7Np1S5zZs2WDSouMkzm0Mz4ofca8";
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

const ADMIN_EMAIL = "admin@wintwealth.com";
const ADMIN_PASS  = "Wint@2024";
const PASS_MARK   = 70;

const SEED_SCHEDULE = [
  { day:1, topics:"Wint Wealth, Reason to issue debt, NBFC", quiz_url:"", reading_links:"" },
  { day:2, topics:"Interest Rates, Bonds, Asset Classes, Liquidity, Role of Brokers", quiz_url:"", reading_links:"" },
  { day:3, topics:"Bonds, Listing, OBPP Regulations, Curated Bonds, CRAR", quiz_url:"", reading_links:"" },
  { day:4, topics:"ROI, Coupon Rate/YTM, XIRR, XIRR Calculation", quiz_url:"", reading_links:"" },
  { day:5, topics:"Sign up, KYC, Payment, After payment process, NSDL App, Sell Anytime Feature, DDPI, Order Cancelation, Flexi tenure", quiz_url:"", reading_links:"" },
  { day:6, topics:"Portfolio, Dashboard, Investments, Repayments, Past holdings, FD Section", quiz_url:"", reading_links:"" },
  { day:7, topics:"Reports, Referral Program, Family Account, Account Deletion", quiz_url:"", reading_links:"" },
  { day:8, topics:"Taxation, Capital Gain/Loss, Form 15G/H", quiz_url:"", reading_links:"" },
  { day:9, topics:"Repayments, SIP", quiz_url:"", reading_links:"" },
];

const inits = (n) => n.split(" ").map((w) => w[0]).join("").toUpperCase().slice(0,2);
const scoreColor = (p) => p>=70?"#16a34a":p>=50?"#d97706":"#e53e3e";
const scoreBg    = (p) => p>=70?"#dcfce7":p>=50?"#fffbeb":"#fff5f5";
const G = { green:"#6cfe5f", greenDark:"#3db832", greenPale:"#f2fff1", black:"#111111", white:"#ffffff", gray:"#f5f5f5", grayMid:"#e8e8e8", grayText:"#666666", grayFaint:"#aaaaaa", red:"#e53e3e", redLight:"#fff5f5", amber:"#d97706", amberLight:"#fffbeb" };

function irScore(quiz, email) {
  const r = quiz.results[email]; if (!r) return null;
  const earned = r.reduce((s,q)=>s+q.marks,0);
  const max    = r.reduce((s,q)=>s+q.maxMarks,0);
  const pct    = max>0?Math.round(earned/max*100):0;
  return { earned, max, pct, correct:r.filter(q=>q.correct).length, total:r.length };
}

function parseWorkbook(wb, fileName) {
  const sheetName = wb.SheetNames.find(s=>/scoring/i.test(s))||wb.SheetNames[0];
  const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{defval:"",header:1});
  const headerRowIdx = rawRows.findIndex(r=>r.some(c=>/^(name|ir.?name|email|ir.?email)$/i.test(String(c).trim())));
  if (headerRowIdx<0) throw new Error("Cannot find header row with Name/Email columns.");
  const hdrs     = rawRows[headerRowIdx].map(c=>String(c||"").trim());
  const dataRows = rawRows.slice(headerRowIdx+1).filter(r=>r.some(c=>c!==null&&c!==undefined&&String(c).trim()!==""));
  const idx = (patterns) => hdrs.findIndex(h=>patterns.some(p=>new RegExp(p,"i").test(h)));
  const nameIdx   = idx(["^(name|ir.?name)$"]);
  const emailIdx  = idx(["^(email|ir.?email)$"]);
  const qTextIdx  = idx(["^question$"]);
  const ansIdx    = idx(["answer.?given|ir.?answer|^answer$"]);
  const modelIdx  = idx(["correct.?answer|model.?answer"]);
  const statusIdx = idx(["correct\\?|^status$|^result$"]);
  const marksIdx  = idx(["^score.?given$|^marks.?given$|^score$|^marks$"]);
  const maxIdx    = idx(["^max.?marks$"]);
  const keyIdx    = idx(["key.?point"]);
  if (nameIdx<0||emailIdx<0) throw new Error("Cannot find Name/Email columns.");
  const questions={}; const results={}; const irOrder=[];
  let curName="", curEmail="";
  dataRows.forEach(row=>{
    const name  = String(row[nameIdx]||"").trim();
    const email = String(row[emailIdx]||"").trim().toLowerCase();
    if (name) curName=name; if (email) curEmail=email;
    if (!curEmail) return;
    if (!results[curEmail]) { results[curEmail]=[]; irOrder.push({name:curName,email:curEmail}); }
    const qText = qTextIdx>=0?String(row[qTextIdx]||"").trim():"";
    const ans   = ansIdx>=0?String(row[ansIdx]||"").trim():"-";
    const model = modelIdx>=0?String(row[modelIdx]||"").trim():"-";
    const keyPt = keyIdx>=0?String(row[keyIdx]||"").trim():"";
    const maxMk = maxIdx>=0?parseFloat(row[maxIdx])||1:1;
    const marks = marksIdx>=0?parseFloat(row[marksIdx])||0:0;
    let ok=false;
    if (statusIdx>=0) { const sv=String(row[statusIdx]||"").trim().toLowerCase(); ok=["yes","correct","1","true","right","y"].includes(sv); }
    else { ok=ans&&model&&ans.toLowerCase()===model.toLowerCase(); }
    const qNum=results[curEmail].length+1;
    if (!questions[qNum]) questions[qNum]={text:qText||"Question "+qNum,modelAnswer:model,keyPoint:keyPt,maxMarks:maxMk};
    results[curEmail].push({answer:ans,correct:ok,marks,maxMarks:maxMk});
  });
  if (!Object.keys(results).length) throw new Error("No IR data found.");
  const title=fileName.replace(/\.(xlsx?|csv)$/i,"").replace(/_/g," ");
  const date=new Date().toLocaleDateString("en-IN",{day:"numeric",month:"short",year:"numeric"});
  return {id:"q_"+Date.now(),title,date,questions,results,irOrder};
}

/* ── Rich Text Editor ── */
const RichEditor = ({ value, onChange, day }) => {
  const ref = useRef(null);
  const imgInputRef = useRef(null);
  const [imgUploading, setImgUploading] = useState(false);
  useEffect(()=>{
    if (ref.current && ref.current.innerHTML !== value) ref.current.innerHTML = value||"";
  },[]);
  const exec = (cmd,val=null) => {
    ref.current.focus();
    document.execCommand(cmd,false,val);
    onChange(ref.current.innerHTML);
  };
  const uploadImage = async (file) => {
    setImgUploading(true);
    try {
      const ext = file.name.split(".").pop();
      const path = `day${day||"misc"}/img_${Date.now()}.${ext}`;
      const { error } = await supabase.storage.from("training-docs").upload(path, file, { upsert: true });
      if (error) throw error;
      const { data } = supabase.storage.from("training-docs").getPublicUrl(path);
      const url = data.publicUrl;
      ref.current.focus();
      document.execCommand("insertImage", false, url);
      const imgs = ref.current.querySelectorAll("img");
      imgs.forEach(img => { img.style.maxWidth = "100%"; img.style.height = "auto"; img.style.borderRadius = "8px"; img.style.margin = "8px 0"; });
      onChange(ref.current.innerHTML);
    } catch(e) { alert("Image upload failed: " + e.message); }
    setImgUploading(false);
  };
  const tbStyle = {
    padding:"5px 10px",
    border:`1px solid ${G.grayMid}`,
    borderRadius:6,
    fontSize:12,
    fontWeight:700,
    cursor:"pointer",
    background:G.white,
    fontFamily:"inherit",
    color:G.black,
    lineHeight:"normal",
    display:"inline-block",
    minWidth:32,
    textAlign:"center",
  };
  return (
    <div style={{border:`1.5px solid ${G.grayMid}`,borderRadius:10,overflow:"hidden"}}>
      <div style={{padding:"8px 10px",background:G.gray,borderBottom:`1px solid ${G.grayMid}`,display:"flex",flexWrap:"wrap",gap:4,alignItems:"center"}}>
        <button type="button" onMouseDown={e=>{e.preventDefault();exec("bold");}} style={{...tbStyle,fontWeight:900}}>B</button>
        <button type="button" onMouseDown={e=>{e.preventDefault();exec("italic");}} style={{...tbStyle,fontStyle:"italic"}}>I</button>
        <button type="button" onMouseDown={e=>{e.preventDefault();exec("underline");}} style={{...tbStyle,textDecoration:"underline"}}>U</button>
        <button type="button" onMouseDown={e=>{e.preventDefault();exec("formatBlock","h2");}} style={tbStyle}>H1</button>
        <button type="button" onMouseDown={e=>{e.preventDefault();exec("formatBlock","h3");}} style={tbStyle}>H2</button>
        <button type="button" onMouseDown={e=>{e.preventDefault();exec("insertUnorderedList");}} style={tbStyle}>Bullets</button>
        <button type="button" onMouseDown={e=>{e.preventDefault();exec("insertOrderedList");}} style={tbStyle}>Numbers</button>
        <button type="button" onMouseDown={e=>{e.preventDefault();exec("removeFormat");}} style={{...tbStyle,color:G.red}}>Clear</button>
        <div style={{width:1,height:20,background:G.grayMid,margin:"0 2px"}}/>
        <button type="button" onMouseDown={e=>{e.preventDefault();imgInputRef.current.click();}} style={{...tbStyle,background:"#f0fff0",color:G.greenDark,border:`1px solid ${G.green}`,cursor:imgUploading?"wait":"pointer"}}>
          {imgUploading?"Uploading...":"+ Image"}
        </button>
        <input ref={imgInputRef} type="file" accept="image/*" style={{display:"none"}} onChange={e=>{if(e.target.files[0])uploadImage(e.target.files[0]);e.target.value="";}}/>
      </div>
      <div
        ref={ref}
        contentEditable
        suppressContentEditableWarning
        onInput={()=>onChange(ref.current.innerHTML)}
        style={{minHeight:240,padding:"14px 16px",outline:"none",fontSize:14,lineHeight:1.8,fontFamily:"inherit",textAlign:"left",color:"#111111",caretColor:"#111111"}}
      />
    </div>
  );
};

/* ── UI Components ── */
const Btn = ({ children, variant="outline", size="md", onClick, style={}, disabled }) => {
  const base={fontFamily:"inherit",fontWeight:600,cursor:disabled?"not-allowed":"pointer",border:"none",borderRadius:10,transition:"all .15s",opacity:disabled?.5:1};
  const sizes={sm:{padding:"6px 13px",fontSize:12},md:{padding:"8px 18px",fontSize:13},lg:{padding:"11px 24px",fontSize:14}};
  const variants={green:{background:G.green,color:G.black},ghost:{background:"rgba(255,255,255,.1)",color:"#fff",border:"1px solid rgba(255,255,255,.18)"},outline:{background:"transparent",color:G.black,border:`1.5px solid ${G.grayMid}`},danger:{background:G.redLight,color:G.red,border:"1px solid #fecaca"},amber:{background:G.amberLight,color:G.amber,border:"1px solid #fde68a"}};
  return <button style={{...base,...sizes[size],...variants[variant],...style}} onClick={onClick} disabled={disabled}>{children}</button>;
};
const Card=({children,style={}})=><div style={{background:G.white,border:`1.5px solid ${G.grayMid}`,borderRadius:16,overflow:"hidden",marginBottom:20,...style}}>{children}</div>;
const CardHeader=({title,right})=><div style={{padding:"16px 20px",borderBottom:`1.5px solid ${G.grayMid}`,display:"flex",alignItems:"center",justifyContent:"space-between"}}><span style={{fontSize:14,fontWeight:700}}>{title}</span>{right}</div>;
const CardBody=({children,style={}})=><div style={{padding:"16px 20px",...style}}>{children}</div>;
const Metric=({label,value,note,accent})=>(
  <div style={{background:accent?G.greenPale:G.white,border:`1.5px solid ${accent?G.green:G.grayMid}`,borderRadius:16,padding:"18px 20px"}}>
    <div style={{fontSize:11,fontWeight:700,color:G.grayFaint,letterSpacing:".6px",textTransform:"uppercase",marginBottom:8}}>{label}</div>
    <div style={{fontSize:32,fontWeight:700,letterSpacing:"-1.5px",lineHeight:1}}>{value}</div>
    {note&&<div style={{fontSize:11,color:G.grayText,marginTop:4}}>{note}</div>}
  </div>
);
const NavItem=({icon,label,active,onClick})=><div onClick={onClick} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 20px",fontSize:13,fontWeight:500,color:active?G.green:"rgba(255,255,255,.45)",cursor:"pointer",borderLeft:`2px solid ${active?G.green:"transparent"}`,background:active?"rgba(108,254,95,.07)":"transparent",transition:"all .12s"}}>{icon}{label}</div>;
const Pill=({pct,label})=><span style={{fontSize:12,fontWeight:700,padding:"3px 10px",borderRadius:100,fontFamily:"monospace",background:scoreBg(pct),color:scoreColor(pct),flexShrink:0}}>{label}</span>;
const ProgBar=({pct})=><div style={{height:5,borderRadius:3,background:G.grayMid,overflow:"hidden",flex:1,maxWidth:110}}><div style={{height:"100%",borderRadius:3,width:`${pct}%`,background:scoreColor(pct),transition:"width .4s"}}/></div>;
const StatusBadge=({status})=>{
  const colors={Confirmed:{bg:"#dcfce7",color:"#16a34a"},Rejected:{bg:G.redLight,color:G.red},Pending:{bg:"#fffbeb",color:G.amber}};
  const c=colors[status]||colors.Pending;
  return <span style={{fontSize:11,fontWeight:700,padding:"3px 10px",borderRadius:100,background:c.bg,color:c.color}}>{status}</span>;
};

const ICONS = {
  overview: <svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>,
  history:  <svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>,
  review:   <svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>,
  schedule: <svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>,
  session:  <svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>,
  upload:   <svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><polyline points="16 16 12 12 8 16"/><line x1="12" y1="12" x2="12" y2="21"/><path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/></svg>,
  irs:      <svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>,
  batch:    <svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>,
  quizzes:  <svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4h6v2"/></svg>,
  sessions: <svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>,
  answers:  <svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>,
  sched_a:  <svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>,
};

export default function App() {
  const [db, setDb]               = useState({irs:[],quizzes:[],sessions:[],schedule:[],dayContent:{},tasks:[]});
  const [user, setUser]           = useState(null);
  const [loginMode, setLoginMode] = useState("ir");
  const [email, setEmail]         = useState("");
  const [pass, setPass]           = useState("");
  const [loginErr, setLoginErr]   = useState("");
  const [loading, setLoading]     = useState(true);
  const [adminPage, setAdminPage] = useState("upload");
  const [uploadStatus, setUploadStatus] = useState(null);
  const [showAddModal, setShowAddModal] = useState(false);
  const [newName, setNewName]     = useState("");
  const [newEmail, setNewEmail]   = useState("");
  const [irPage, setIrPage]       = useState("overview");
  const [selectedQuiz, setSelectedQuiz] = useState(null);
  const [sessionTopic, setSessionTopic] = useState("");
  const [sessionSubmitting, setSessionSubmitting] = useState(false);
  const [sessionMsg, setSessionMsg]     = useState(null);
  const [completedDays, setCompletedDays] = useState({});
  const [activeDay, setActiveDay]       = useState(null);
  const [editingDay, setEditingDay]     = useState(null);
  const [editForm, setEditForm]         = useState({});
  const [editContent, setEditContent]   = useState({html:"",questions:[]});
  const [schedSaving, setSchedSaving]   = useState(false);
  const [taskForm, setTaskForm]         = useState({ir_email:"",title:"",notes:""});
  const [taskSaving, setTaskSaving]     = useState(false);
  const [taskDay, setTaskDay]           = useState(null);
  const [irAnswers, setIrAnswers]       = useState({});
  const [answersSaving, setAnswersSaving] = useState(false);
  const [allAnswers, setAllAnswers]     = useState([]);

  useEffect(()=>{
    async function loadData() {
      try {
        const [
          {data:irs},{data:quizzes},{data:results},{data:sessions},
          {data:schedule},{data:progress},{data:answers},{data:tasks}
        ] = await Promise.all([
          supabase.from("irs").select("*"),
          supabase.from("quizzes").select("*"),
          supabase.from("results").select("*"),
          supabase.from("session_requests").select("*").order("created_at",{ascending:false}),
          supabase.from("training_schedule").select("*").order("day"),
          supabase.from("tasks").select("*").order("created_at",{ascending:false}),
          supabase.from("training_progress").select("*"),
          supabase.from("ir_answers").select("*").order("submitted_at",{ascending:false}),
        ]);
        // Load day_content separately - only fetch day and questions (not full html) to avoid large payload
        const {data:dayContent, error:dcError} = await supabase.from("day_content").select("day,questions");
        if (dcError) console.error("day_content load error:", dcError);
        // Load full content only when needed (on demand per day)
        const quizzesWithResults=(quizzes||[]).map(q=>({
          ...q,questions:q.questions,
          results:Object.fromEntries((results||[]).filter(r=>r.quiz_id===q.id).map(r=>[r.ir_email,r.answers])),
          irOrder:[],
        }));
        const sched=(schedule&&schedule.length>0)?schedule:SEED_SCHEDULE;
        const dcMap={};
        (dayContent||[]).forEach(d=>{dcMap[d.day]={html:d.content_html||"",questions:d.questions||[]};});
        setDb({irs:irs||[],quizzes:quizzesWithResults,sessions:sessions||[],schedule:sched,dayContent:dcMap,tasks:tasks||[]});
        const prog={};
        (progress||[]).forEach(p=>{prog[p.ir_email+"_"+p.day]=true;});
        setCompletedDays(prog);
        setAllAnswers(answers||[]);
      } catch(e){console.error("Load error:",e);}
      finally{setLoading(false);}
    }
    loadData();
  },[]);

  const doLogin = async () => {
    setLoginErr("");
    if (loginMode==="admin") {
      if (email.trim().toLowerCase()===ADMIN_EMAIL&&pass===ADMIN_PASS) setUser({role:"admin",name:"Trainer",email:ADMIN_EMAIL});
      else setLoginErr("Incorrect credentials.");
    } else {
      const trimmedEmail=email.trim().toLowerCase();
      if (!trimmedEmail.endsWith("@wintwealth.com")) {setLoginErr("Only @wintwealth.com emails are allowed.");return;}
      let ir=db.irs.find(i=>i.email.toLowerCase()===trimmedEmail);
      if (!ir) {
        const name=trimmedEmail.split("@")[0].replace(/[._]/g," ").split(" ").map(w=>w.charAt(0).toUpperCase()+w.slice(1)).join(" ");
        await supabase.from("irs").upsert({name,email:trimmedEmail},{onConflict:"email"});
        ir={name,email:trimmedEmail};
        setDb(p=>({...p,irs:[...p.irs,ir]}));
      }
      setUser({role:"ir",...ir}); setIrPage("overview");
    }
  };

  const logout=()=>{setUser(null);setEmail("");setPass("");setLoginErr("");};

  const handleFile=useCallback(async(file)=>{
    try {
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:"array"});
      const quiz=parseWorkbook(wb,file.name);
      const newIrs=[...db.irs];
      for (const {name,email} of quiz.irOrder) {
        if (!newIrs.find(i=>i.email===email)) {
          await supabase.from("irs").upsert({name,email},{onConflict:"email"});
          newIrs.push({name,email});
        }
      }
      await supabase.from("quizzes").upsert({id:quiz.id,title:quiz.title,date:quiz.date,questions:quiz.questions});
      for (const [ir_email,answers] of Object.entries(quiz.results)) {
        await supabase.from("results").upsert({quiz_id:quiz.id,ir_email,answers},{onConflict:"quiz_id,ir_email"});
      }
      setDb(prev=>({...prev,irs:newIrs,quizzes:[...prev.quizzes,quiz]}));
      setUploadStatus({ok:true,title:quiz.title,nIRs:Object.keys(quiz.results).length,nQs:Object.keys(quiz.questions).length});
    } catch(e){setUploadStatus({ok:false,msg:e.message});}
  },[db.irs]);

  const isDayUnlocked=(day)=>{if(day===1)return true;return completedDays[user?.email+"_"+(day-1)];};
  const isDayComplete=(day)=>completedDays[user?.email+"_"+day];

  const markDayComplete=async(day)=>{
    const key=user.email+"_"+day;
    await supabase.from("training_progress").upsert({ir_email:user.email,day},{onConflict:"ir_email,day"});
    setCompletedDays(p=>({...p,[key]:true}));
  };

  const saveScheduleDay=async(dayObj)=>{
    setSchedSaving(true);
    await supabase.from("training_schedule").upsert(dayObj,{onConflict:"day"});
    setDb(p=>({...p,schedule:p.schedule.map(d=>d.day===dayObj.day?dayObj:d)}));
    setEditingDay(null);setEditForm({});
    setSchedSaving(false);
  };

  const saveDayContent=async(day,html,questions)=>{
    await supabase.from("day_content").upsert({day,content_html:html,questions},{onConflict:"day"});
    setDb(p=>({...p,dayContent:{...p.dayContent,[day]:{html,questions}}}));
  };

  const addNewDay=async()=>{
    const nextDay=Math.max(...db.schedule.map(d=>d.day),0)+1;
    const newDay={day:nextDay,topics:"",quiz_url:"",reading_links:""};
    await supabase.from("training_schedule").upsert(newDay,{onConflict:"day"});
    setDb(p=>({...p,schedule:[...p.schedule,newDay]}));
    setEditingDay(nextDay);setEditForm(newDay);
    setEditContent({html:"",questions:[]});
  };

  const submitIRAnswers=async(day,answers)=>{
    setAnswersSaving(true);
    await supabase.from("ir_answers").upsert({ir_email:user.email,day,answers},{onConflict:"ir_email,day"});
    await markDayComplete(day);
    setAnswersSaving(false);
  };

  const inputStyle={width:"100%",padding:"10px 13px",border:`1.5px solid ${G.grayMid}`,borderRadius:10,fontFamily:"inherit",fontSize:14,outline:"none",boxSizing:"border-box"};
  const taStyle={...inputStyle,resize:"vertical",minHeight:80};

  const myQuizzes=user?.role==="ir"?db.quizzes.filter(q=>q.results[user.email]):[];
  const myScores=myQuizzes.map(q=>irScore(q,user?.email)).filter(Boolean);
  const avgPct=myScores.length?Math.round(myScores.reduce((s,x)=>s+x.pct,0)/myScores.length):0;
  const totalEarned=myScores.reduce((s,x)=>s+x.earned,0);
  const totalMax=myScores.reduce((s,x)=>s+x.max,0);
  const activeQuiz=selectedQuiz?db.quizzes.find(q=>q.id===selectedQuiz):myQuizzes[0];

  if (loading) return (
    <div style={{fontFamily:"Helvetica Neue,Helvetica,Arial,sans-serif",background:G.black,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{textAlign:"center",color:"#fff"}}>
        <img src={wintLogo} alt="Wint" style={{height:40,marginBottom:24,opacity:.9}}/>
        <div style={{fontSize:14,color:"rgba(255,255,255,.5)"}}>Loading portal...</div>
      </div>
    </div>
  );

  if (!user) return (
    <div style={{fontFamily:"Helvetica Neue,Helvetica,Arial,sans-serif",background:G.white,color:G.black,minHeight:"100vh"}}>
      <div style={{minHeight:"100vh",display:"grid",gridTemplateColumns:"1fr 1fr"}}>
        <div style={{background:G.black,display:"flex",flexDirection:"column",justifyContent:"center",padding:"60px 56px",position:"relative",overflow:"hidden"}}>
          <div style={{position:"absolute",bottom:-80,left:-40,width:300,height:300,borderRadius:"50%",background:"rgba(108,254,95,.06)",pointerEvents:"none"}}/>
          <img src={wintLogo} alt="Wint" style={{height:38,marginBottom:44,objectFit:"contain",objectPosition:"left"}}/>
          <h1 style={{fontSize:36,fontWeight:700,color:"#fff",lineHeight:1.2,marginBottom:16,letterSpacing:"-.8px"}}>Track your <span style={{color:G.green}}>quiz results</span> and grow faster.</h1>
          <p style={{fontSize:14,color:"rgba(255,255,255,.45)",lineHeight:1.8,maxWidth:300,fontWeight:400}}>Wint Training Portal gives IRs instant access to their scores, question-by-question feedback, and progress over time.</p>
        </div>
        <div style={{display:"flex",alignItems:"center",justifyContent:"center",padding:"48px 40px"}}>
          <div style={{width:"100%",maxWidth:380}}>
            <div style={{display:"flex",background:G.gray,borderRadius:10,padding:4,marginBottom:28,gap:4}}>
              {["ir","admin"].map(m=>(
                <div key={m} onClick={()=>{setLoginMode(m);setLoginErr("");}} style={{flex:1,textAlign:"center",padding:8,fontSize:13,fontWeight:600,cursor:"pointer",borderRadius:8,background:loginMode===m?G.black:"transparent",color:loginMode===m?"#fff":G.grayText,transition:"all .15s"}}>
                  {m==="ir"?"IR Login":"Admin Login"}
                </div>
              ))}
            </div>
            <h2 style={{fontSize:24,fontWeight:700,marginBottom:4,letterSpacing:"-.4px"}}>{loginMode==="admin"?"Admin sign in":"Welcome back"}</h2>
            <p style={{fontSize:13,color:G.grayText,marginBottom:24,fontWeight:400}}>{loginMode==="admin"?"Access admin panel.":"Enter your @wintwealth.com email."}</p>
            <div style={{marginBottom:14}}>
              <label style={{display:"block",fontSize:11,fontWeight:700,color:G.grayText,marginBottom:5,letterSpacing:".5px",textTransform:"uppercase"}}>Email</label>
              <input value={email} onChange={e=>setEmail(e.target.value)} onKeyDown={e=>e.key==="Enter"&&doLogin()} type="email" placeholder="yourname@wintwealth.com" style={inputStyle}/>
            </div>
            {loginMode==="admin"&&(
              <div style={{marginBottom:14}}>
                <label style={{display:"block",fontSize:11,fontWeight:700,color:G.grayText,marginBottom:5,letterSpacing:".5px",textTransform:"uppercase"}}>Password</label>
                <input value={pass} onChange={e=>setPass(e.target.value)} onKeyDown={e=>e.key==="Enter"&&doLogin()} type="password" placeholder="password" style={inputStyle}/>
              </div>
            )}
            {loginErr&&<div style={{fontSize:12,color:G.red,background:G.redLight,border:"1px solid #fecaca",borderRadius:10,padding:"8px 12px",marginBottom:10}}>{loginErr}</div>}
            <Btn variant="green" size="lg" style={{width:"100%"}} onClick={doLogin}>Sign in</Btn>
          </div>
        </div>
      </div>
    </div>
  );

  const Nav=()=>(
    <>
    <style>{`:root{color-scheme:light}*{color-scheme:light}textarea,input{color:#111111!important;background:#ffffff!important;-webkit-text-fill-color:#111111!important}textarea::placeholder,input::placeholder{color:#aaaaaa!important;-webkit-text-fill-color:#aaaaaa!important}`}</style>
    <nav style={{background:G.black,height:58,display:"flex",alignItems:"center",justifyContent:"space-between",padding:"0 28px",position:"sticky",top:0,zIndex:100}}>
      <div style={{display:"flex",alignItems:"center",gap:12}}>
        <img src={wintLogo} alt="Wint" style={{height:26,objectFit:"contain"}}/>
        <div style={{width:1,height:20,background:"rgba(255,255,255,.2)"}}/>
        <div style={{fontSize:11,fontWeight:700,letterSpacing:"1.8px",textTransform:"uppercase",color:"rgba(255,255,255,.45)"}}>Training Portal</div>
      </div>
      <div style={{display:"flex",alignItems:"center",gap:8}}>
        <span style={{fontSize:12,color:"rgba(255,255,255,.45)"}}>{user.name}</span>
        <Btn variant="ghost" size="sm" onClick={logout}>Sign out</Btn>
      </div>
    </nav>
    </>
  );

  const Layout=({sidebar,children})=>(
    <div style={{fontFamily:"Helvetica Neue,Helvetica,Arial,sans-serif",background:G.white,color:G.black,minHeight:"100vh"}}>
      <Nav/>
      <div style={{display:"grid",gridTemplateColumns:"230px 1fr",minHeight:"calc(100vh - 58px)"}}>
        <aside style={{background:G.black,display:"flex",flexDirection:"column"}}>
          <div style={{padding:"20px 20px 16px",borderBottom:"1px solid rgba(255,255,255,.08)"}}>
            <div style={{width:38,height:38,borderRadius:"50%",background:user.role==="admin"?"rgba(255,255,255,.12)":G.green,color:user.role==="admin"?"rgba(255,255,255,.8)":G.black,fontSize:13,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>{inits(user.name)}</div>
            <div style={{fontSize:14,fontWeight:600,color:"#fff",marginTop:10}}>{user.name}</div>
            <div style={{fontSize:11,color:"rgba(255,255,255,.35)",marginTop:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{user.email}</div>
            <div style={{marginTop:7,display:"inline-block",fontSize:10,fontWeight:700,letterSpacing:".8px",textTransform:"uppercase",padding:"2px 8px",borderRadius:100,background:user.role==="admin"?"rgba(255,255,255,.1)":G.green,color:user.role==="admin"?"rgba(255,255,255,.6)":G.black}}>{user.role==="admin"?"Admin":"IR"}</div>
          </div>
          <div style={{fontSize:10,fontWeight:700,letterSpacing:"1.2px",textTransform:"uppercase",color:"rgba(255,255,255,.22)",padding:"20px 20px 6px"}}>Menu</div>
          {sidebar}
          <div style={{marginTop:"auto",padding:"16px 20px",borderTop:"1px solid rgba(255,255,255,.08)"}}>
            <Btn variant="ghost" size="sm" style={{width:"100%",color:"rgba(255,255,255,.6)"}} onClick={logout}>Sign out</Btn>
          </div>
        </aside>
        <main style={{padding:32,background:G.white,overflowY:"auto"}}>{children}</main>
      </div>
    </div>
  );

  /* IR TRAINING SCHEDULE PAGE */
  const IRSchedulePage=()=>{
    const completedCount=db.schedule.filter(d=>isDayComplete(d.day)).length;
    const [localAnswers, setLocalAnswers]=useState({});
    const [submitted, setSubmitted]=useState(false);
    const [fullDayContent, setFullDayContent]=useState({});
    useEffect(()=>{
      if (!activeDay) return;
      async function loadDayContent() {
        const {data,error}=await supabase.from("day_content").select("*").eq("day",activeDay).single();
        if (data) setFullDayContent(p=>({...p,[activeDay]:{html:data.content_html||"",questions:data.questions||[]}}));
        if (error&&error.code!=="PGRST116") console.error("Content load error:",error);
      }
      loadDayContent();
    },[activeDay]);

    if (activeDay) {
      const day=db.schedule.find(d=>d.day===activeDay);
      const unlocked=isDayUnlocked(day.day);
      const done=isDayComplete(day.day);
      const content=fullDayContent[day.day]||db.dayContent[day.day]||{html:"",questions:[]};
      const links=day.reading_links?day.reading_links.split("\n").filter(l=>l.trim()):[];
      const allAnswered=content.questions.length===0||content.questions.every((_,i)=>(localAnswers[i]||"").trim().length>0);

      return (
        <div>
          <button onClick={()=>{setActiveDay(null);setSubmitted(false);}} style={{background:"none",border:"none",cursor:"pointer",fontSize:13,color:G.grayText,marginBottom:20,display:"flex",alignItems:"center",gap:6,fontFamily:"inherit",fontWeight:600,padding:0}}>
            <svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><polyline points="15 18 9 12 15 6"/></svg>
            Back to schedule
          </button>
          <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:24}}>
            <div style={{width:44,height:44,borderRadius:12,background:done?G.green:unlocked?G.black:"#ccc",color:done?G.black:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,fontWeight:700,flexShrink:0}}>
              {done?"✓":day.day}
            </div>
            <div>
              <div style={{fontSize:20,fontWeight:700}}>Day {day.day}</div>
              <div style={{fontSize:13,color:G.grayText}}>{day.topics}</div>
            </div>
          </div>

          {!unlocked&&(
            <div style={{background:G.gray,border:`1.5px solid ${G.grayMid}`,borderRadius:16,padding:"32px",textAlign:"center"}}>
              <div style={{fontSize:32,marginBottom:8}}>&#x1F512;</div>
              <div style={{fontSize:15,fontWeight:700,marginBottom:4}}>Day {day.day-1} required</div>
              <div style={{fontSize:13,color:G.grayText}}>Complete Day {day.day-1} to unlock this day.</div>
            </div>
          )}

          {unlocked&&(
            <>
              {/* Reading content */}
              {content.html?(
                <Card style={{marginBottom:20}}>
                  <CardHeader title="Reading Material"/>
                  <CardBody>
                    <div dangerouslySetInnerHTML={{__html:content.html}} style={{fontSize:14,lineHeight:1.8,color:G.black,textAlign:"left"}}/>
                  </CardBody>
                </Card>
              ):(
                <Card style={{marginBottom:20}}>
                  <CardBody><div style={{textAlign:"center",padding:"24px 0",color:G.grayText,fontSize:13}}>Reading material not added yet.</div></CardBody>
                </Card>
              )}

              {/* Additional links */}
              {links.length>0&&(
                <Card style={{marginBottom:20}}>
                  <CardHeader title="Additional Resources"/>
                  <CardBody>
                    {links.map((line,i)=>{
                      // Extract URL from line if present
                      const urlMatch = line.match(/(https?:\/\/[^\s]+)/);
                      const url = urlMatch ? urlMatch[1] : null;
                      const label = line.replace(url||"","").replace(/^[0-9]+\.\s*/,"").trim() || url;
                      return (
                        <div key={i} style={{padding:"10px 0",borderBottom:i<links.length-1?`1px solid ${G.grayMid}`:"none",display:"flex",alignItems:"flex-start",gap:10}}>
                          <div style={{width:22,height:22,borderRadius:"50%",background:G.greenPale,color:G.greenDark,fontSize:11,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,marginTop:1}}>{i+1}</div>
                          <div style={{flex:1}}>
                            {label&&!url&&<div style={{fontSize:13,color:G.black,lineHeight:1.6}}>{label}</div>}
                            {label&&url&&label!==url&&<div style={{fontSize:13,color:G.black,marginBottom:4,lineHeight:1.6,fontWeight:500}}>{label}</div>}
                            {url&&(
                              <a href={url} target="_blank" rel="noopener noreferrer" style={{fontSize:12,color:G.greenDark,textDecoration:"none",display:"inline-flex",alignItems:"center",gap:5,fontWeight:600,wordBreak:"break-all"}}>
                                <svg width="12" height="12" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                                Open link
                              </a>
                            )}
                          </div>
                        </div>
                      );
                    })}
                  </CardBody>
                </Card>
              )}

              {/* Quiz */}
              <Card>
                <CardHeader title="Quiz"/>
                <CardBody>
                  {done?(
                    <div style={{display:"flex",alignItems:"center",gap:12}}>
                      <div style={{width:36,height:36,borderRadius:"50%",background:"#dcfce7",color:"#16a34a",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>&#x2713;</div>
                      <div><div style={{fontSize:13,fontWeight:700,color:"#16a34a"}}>Day {day.day} completed!</div><div style={{fontSize:12,color:G.grayText}}>Well done.</div></div>
                    </div>
                  ):day.quiz_url?(
                    <div>
                      <p style={{fontSize:13,color:G.grayText,marginBottom:16}}>Click to open the quiz in a new tab. Once done, come back and click "I have completed the quiz" to unlock the next day.</p>
                      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                        <a href={day.quiz_url} target="_blank" rel="noopener noreferrer" style={{textDecoration:"none"}}><Btn variant="green" size="md">Open Quiz</Btn></a>
                        <Btn variant="outline" size="md" onClick={()=>markDayComplete(day.day)}>I have completed the quiz</Btn>
                      </div>
                    </div>
                  ):(
                    <div style={{textAlign:"center",padding:"16px 0",color:G.grayText,fontSize:13}}>Quiz link not added yet.</div>
                  )}
                </CardBody>
              </Card>
              {/* Questions section */}
              {content.questions.length>0&&(
                <Card style={{marginBottom:20}}>
                  <CardHeader title="Questions"/>
                  <CardBody>
                    {done?(
                      <div>
                        <div style={{display:"flex",alignItems:"center",gap:10,padding:"8px 0 16px",color:"#16a34a"}}>
                          <svg width="20" height="20" fill="none" stroke="currentColor" strokeWidth="2.5" viewBox="0 0 24 24"><polyline points="20 6 9 17 4 12"/></svg>
                          <span style={{fontSize:13,fontWeight:600}}>You have submitted your answers for this day.</span>
                        </div>
                        {(() => {
                          const saved = allAnswers.find ? allAnswers.find(a=>a.ir_email===user.email&&a.day===day.day) : null;
                          if (!saved) return null;
                          return content.questions.map((q,i)=>(
                            <div key={i} style={{marginBottom:16,padding:"12px 14px",background:G.gray,borderRadius:10,border:`1px solid ${G.grayMid}`}}>
                              <div style={{fontSize:12,fontWeight:700,color:G.grayText,marginBottom:6,textAlign:"left"}}>Q{i+1}. {q}</div>
                              <div style={{fontSize:13,color:G.black,lineHeight:1.6,textAlign:"left",whiteSpace:"pre-wrap"}}>{saved.answers[i]||"-"}</div>
                            </div>
                          ));
                        })()}
                      </div>
                    ):(
                      <>
                        {content.questions.map((q,i)=>(
                          <div key={i} style={{marginBottom:20}}>
                            <div style={{fontSize:13,fontWeight:600,marginBottom:8,textAlign:"left"}}>Q{i+1}. {q}</div>
                            <textarea
                              value={localAnswers[i]||""}
                              onChange={e=>setLocalAnswers(p=>({...p,[i]:e.target.value}))}
                              onCopy={e=>e.preventDefault()}
                              onPaste={e=>e.preventDefault()}
                              onCut={e=>e.preventDefault()}
                              placeholder="Type your answer here..."
                              rows={3}
                              style={{...taStyle,minHeight:80,fontSize:13,textAlign:"left",color:G.black}}
                            />
                            {!(localAnswers[i]||"").trim()&&<div style={{fontSize:11,color:G.amber,marginTop:3}}>Required</div>}
                          </div>
                        ))}
                        {submitted&&!allAnswered&&<div style={{fontSize:12,color:G.red,marginBottom:10}}>Please answer all questions before submitting.</div>}
                        <Btn variant="green" size="md" disabled={answersSaving} onClick={async()=>{
                          setSubmitted(true);
                          if (!allAnswered) return;
                          await submitIRAnswers(day.day, localAnswers);
                        }}>{answersSaving?"Saving...":"Submit answers"}</Btn>
                      </>
                    )}
                  </CardBody>
                </Card>
              )}

              {/* My Tasks */}
              {(()=>{
                const myTasks = db.tasks.filter(t=>(t.ir_email===user.email||t.ir_email==="all")&&t.day===day.day);
                if (!myTasks.length) return null;
                return (
                  <Card style={{marginTop:20}}>
                    <CardHeader title="My Tasks"/>
                    <CardBody>
                      {myTasks.map((task,i)=>(
                        <div key={task.id} style={{display:"flex",alignItems:"flex-start",gap:12,padding:"12px 0",borderBottom:i<myTasks.length-1?`1px solid ${G.grayMid}`:"none"}}>
                          <div style={{flex:1}}>
                            <div style={{fontSize:13,fontWeight:600,marginBottom:task.notes?4:0}}>{task.title}</div>
                            {task.notes&&<div style={{fontSize:12,color:G.grayText,lineHeight:1.6}}>{task.notes}</div>}
                          </div>
                          <select value={task.status} onChange={async e=>{
                            const newStatus=e.target.value;
                            await supabase.from("tasks").update({status:newStatus}).eq("id",task.id);
                            setDb(p=>({...p,tasks:p.tasks.map(t=>t.id===task.id?{...t,status:newStatus}:t)}));
                          }} style={{padding:"4px 8px",border:`1.5px solid ${G.grayMid}`,borderRadius:8,fontFamily:"inherit",fontSize:12,background:task.status==="Completed"?"#dcfce7":task.status==="In Progress"?G.amberLight:G.white,color:task.status==="Completed"?"#16a34a":task.status==="In Progress"?G.amber:G.grayText,cursor:"pointer",fontWeight:600}}>
                            <option value="Pending">Pending</option>
                            <option value="In Progress">In Progress</option>
                            <option value="Completed">Completed</option>
                          </select>
                        </div>
                      ))}
                    </CardBody>
                  </Card>
                );
              })()}
            </>
          )}
        </div>
      );
    }

    return (
      <>
        <div style={{marginBottom:28,display:"flex",alignItems:"flex-start",justifyContent:"space-between"}}>
          <div>
            <div style={{fontSize:22,fontWeight:700,letterSpacing:"-.5px"}}>Training Schedule</div>
            <div style={{fontSize:13,color:G.grayText,marginTop:3}}>{completedCount} of {db.schedule.length} days completed</div>
          </div>
          <div style={{background:G.greenPale,border:`1.5px solid ${G.green}`,borderRadius:12,padding:"8px 16px",fontSize:13,fontWeight:600,color:G.greenDark}}>{Math.round(completedCount/db.schedule.length*100)||0}% complete</div>
        </div>
        <div style={{height:8,borderRadius:4,background:G.grayMid,overflow:"hidden",marginBottom:28}}>
          <div style={{height:"100%",borderRadius:4,background:G.green,width:`${Math.round(completedCount/db.schedule.length*100)||0}%`,transition:"width .4s"}}/>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {db.schedule.map(day=>{
            const unlocked=isDayUnlocked(day.day);
            const done=isDayComplete(day.day);
            return (
              <div key={day.day} onClick={()=>unlocked&&setActiveDay(day.day)}
                style={{display:"flex",alignItems:"center",gap:16,padding:"16px 20px",background:G.white,border:`1.5px solid ${done?G.green:unlocked?G.grayMid:"#e0e0e0"}`,borderRadius:14,cursor:unlocked?"pointer":"default",transition:"all .15s",opacity:unlocked?1:.7}}>
                <div style={{width:40,height:40,borderRadius:10,background:done?G.green:unlocked?G.black:G.grayMid,color:done?G.black:unlocked?"#fff":G.grayFaint,display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,fontWeight:700,flexShrink:0}}>
                  {done?"✓":day.day}
                </div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontSize:13,fontWeight:700,marginBottom:2}}>Day {day.day}</div>
                  <div style={{fontSize:12,color:G.grayText,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{day.topics||"Topics not set"}</div>
                </div>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  {done&&<span style={{fontSize:11,fontWeight:700,padding:"2px 8px",borderRadius:100,background:"#dcfce7",color:"#16a34a"}}>Done</span>}
                  {!done&&unlocked&&<span style={{fontSize:11,fontWeight:700,padding:"2px 8px",borderRadius:100,background:G.amberLight,color:G.amber}}>In progress</span>}
                  {!unlocked&&<span style={{fontSize:18}}>&#x1F512;</span>}
                  {unlocked&&<svg width="16" height="16" fill="none" stroke={G.grayFaint} strokeWidth="2" viewBox="0 0 24 24"><polyline points="9 18 15 12 9 6"/></svg>}
                </div>
              </div>
            );
          })}
        </div>
      </>
    );
  };

  /* IR DASHBOARD */
  if (user.role==="ir") {
    const QuizList=({quizzes,limit})=>{
      const list=limit?quizzes.slice(0,limit):quizzes;
      if (!list.length) return <div style={{textAlign:"center",padding:40,color:G.grayText}}>No quiz results yet.</div>;
      return list.map((q,i)=>{
        const s=irScore(q,user.email);
        return (
          <div key={q.id} style={{display:"flex",alignItems:"center",gap:14,padding:"12px 0",borderBottom:i<list.length-1?`1px solid ${G.grayMid}`:"none"}}>
            <div style={{fontSize:11,color:G.grayFaint,fontFamily:"monospace",width:22}}>{String(i+1).padStart(2,"0")}</div>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontSize:13,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{q.title}</div>
              <div style={{fontSize:11,color:G.grayFaint}}>{q.date}</div>
            </div>
            <ProgBar pct={s.pct}/>
            <Pill pct={s.pct} label={`${s.earned}/${s.max} - ${s.pct}%`}/>
          </div>
        );
      });
    };
    return (
      <Layout sidebar={["overview","history","review","schedule","session"].map(p=>(
        <NavItem key={p} icon={ICONS[p]} label={p==="review"?"Question Review":p==="session"?"Request Session":p==="schedule"?"Training Schedule":p.charAt(0).toUpperCase()+p.slice(1)} active={irPage===p} onClick={()=>{setIrPage(p);setActiveDay(null);}}/>
      ))}>
        {irPage==="overview"&&(
          <>
            <div style={{marginBottom:28}}><div style={{fontSize:22,fontWeight:700,letterSpacing:"-.5px"}}>Good day, {user.name.split(" ")[0]} &#x1F44B;</div><div style={{fontSize:13,color:G.grayText,marginTop:3}}>Your quiz performance at a glance.</div></div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:24}}>
              <Metric accent label="Avg Score" value={`${avgPct}%`} note={`${myScores.filter(s=>s.pct>=PASS_MARK).length}/${myScores.length} passed`}/>
              <Metric label="Total Marks" value={totalEarned} note={`out of ${totalMax}`}/>
              <Metric label="Quizzes" value={myQuizzes.length} note={`of ${db.quizzes.length} total`}/>
            </div>
            <Card><CardHeader title="Recent Quizzes"/><CardBody><QuizList quizzes={myQuizzes} limit={3}/></CardBody></Card>
          </>
        )}
        {irPage==="history"&&(
          <><div style={{marginBottom:28}}><div style={{fontSize:22,fontWeight:700}}>Quiz History</div><div style={{fontSize:13,color:G.grayText}}>All your attempts and scores.</div></div><Card><CardBody><QuizList quizzes={myQuizzes}/></CardBody></Card></>
        )}
        {irPage==="review"&&(
          <>
            <div style={{marginBottom:28}}><div style={{fontSize:22,fontWeight:700}}>Question Review</div><div style={{fontSize:13,color:G.grayText}}>Your answer, model answer, and marks.</div></div>
            <Card>
              <CardHeader title="Select Quiz" right={
                <select value={activeQuiz?.id||""} onChange={e=>setSelectedQuiz(e.target.value)} style={{padding:"6px 10px",border:`1.5px solid ${G.grayMid}`,borderRadius:10,fontFamily:"inherit",fontSize:13,background:G.white,cursor:"pointer"}}>
                  {myQuizzes.map(q=><option key={q.id} value={q.id}>{q.title}</option>)}
                </select>
              }/>
              <CardBody>
                {activeQuiz&&activeQuiz.results[user.email]?(
                  <ul style={{listStyle:"none",padding:0,margin:0}}>
                    {Object.keys(activeQuiz.questions).map((k,i)=>{
                      const q=activeQuiz.questions[k]; const ans=activeQuiz.results[user.email][i]||{};
                      return (
                        <li key={k} style={{display:"flex",gap:12,alignItems:"flex-start",padding:"14px 0",borderBottom:i<Object.keys(activeQuiz.questions).length-1?`1px solid ${G.grayMid}`:"none"}}>
                          <div style={{width:24,height:24,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,background:ans.correct?"#dcfce7":G.redLight,color:ans.correct?"#16a34a":G.red}}>
                            {ans.correct?<svg width="12" height="12" fill="none" stroke="currentColor" strokeWidth="3" viewBox="0 0 24 24"><polyline points="20 6 9 17 4 12"/></svg>:<svg width="12" height="12" fill="none" stroke="currentColor" strokeWidth="3" viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>}
                          </div>
                          <div style={{flex:1}}>
                            <div style={{fontSize:13,fontWeight:600,lineHeight:1.5,marginBottom:6}}>Q{k}. {q.text}</div>
                            <div style={{display:"flex",gap:7,flexWrap:"wrap",alignItems:"center"}}>
                              <span style={{fontSize:11,fontWeight:600,padding:"2px 8px",borderRadius:4,fontFamily:"monospace",background:G.gray,color:G.grayText,border:`1px solid ${G.grayMid}`}}>Your: {ans.answer||"-"}</span>
                              {!ans.correct&&<><span style={{fontSize:11,fontWeight:600,padding:"2px 8px",borderRadius:4,background:G.redLight,color:G.red}}>Wrong</span><span style={{fontSize:11,fontWeight:600,padding:"2px 8px",borderRadius:4,background:"#dcfce7",color:"#16a34a"}}>Correct: {q.modelAnswer}</span></>}
                              {ans.correct&&<span style={{fontSize:11,fontWeight:600,padding:"2px 8px",borderRadius:4,background:"#dcfce7",color:"#16a34a"}}>Correct</span>}
                              <span style={{fontSize:11,color:G.grayFaint,fontWeight:600,marginLeft:"auto"}}>{ans.marks}/{ans.maxMarks} marks</span>
                            </div>
                            {q.keyPoint&&<div style={{fontSize:11,color:G.grayText,marginTop:6,background:"#fffbeb",border:"1px solid #fde68a",borderRadius:6,padding:"4px 8px"}}>Tip: {q.keyPoint}</div>}
                          </div>
                        </li>
                      );
                    })}
                  </ul>
                ):<div style={{textAlign:"center",padding:40,color:G.grayText}}>No data for this quiz.</div>}
              </CardBody>
            </Card>
          </>
        )}
        {irPage==="schedule"&&<IRSchedulePage/>}
        {irPage==="session"&&(
          <>
            <div style={{marginBottom:28}}><div style={{fontSize:22,fontWeight:700}}>Request a Session</div><div style={{fontSize:13,color:G.grayText}}>Need help? Submit a request and your trainer will get back to you.</div></div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
              <Card style={{marginBottom:0}}>
                <CardHeader title="New Request"/>
                <CardBody>
                  <div style={{marginBottom:14}}>
                    <label style={{display:"block",fontSize:11,fontWeight:700,color:G.grayText,marginBottom:5,textTransform:"uppercase",letterSpacing:".5px"}}>Your Name</label>
                    <input value={user.name} disabled style={{...inputStyle,background:G.gray,color:G.grayText}}/>
                  </div>
                  <div style={{marginBottom:14}}>
                    <label style={{display:"block",fontSize:11,fontWeight:700,color:G.grayText,marginBottom:5,textTransform:"uppercase",letterSpacing:".5px"}}>Your Email</label>
                    <input value={user.email} disabled style={{...inputStyle,background:G.gray,color:G.grayText}}/>
                  </div>
                  <div style={{marginBottom:20}}>
                    <label style={{display:"block",fontSize:11,fontWeight:700,color:G.grayText,marginBottom:5,textTransform:"uppercase",letterSpacing:".5px"}}>Topic</label>
                    <textarea value={sessionTopic} onChange={e=>setSessionTopic(e.target.value)} placeholder="e.g. I need help understanding credit risk scoring..." rows={4} style={taStyle}/>
                  </div>
                  {sessionMsg&&<div style={{marginBottom:12,padding:"8px 12px",borderRadius:10,fontSize:12,fontWeight:600,background:sessionMsg.ok?"#dcfce7":G.redLight,color:sessionMsg.ok?"#16a34a":G.red}}>{sessionMsg.ok?"Request submitted! Your trainer will get back to you soon.":"Error: "+sessionMsg.text}</div>}
                  <Btn variant="green" size="md" style={{width:"100%"}} disabled={sessionSubmitting||!sessionTopic.trim()} onClick={async()=>{
                    setSessionSubmitting(true);setSessionMsg(null);
                    try {
                      const {error}=await supabase.from("session_requests").insert({ir_name:user.name,ir_email:user.email,topic:sessionTopic.trim(),status:"Pending"});
                      if (error) throw error;
                      const newSession={ir_name:user.name,ir_email:user.email,topic:sessionTopic.trim(),status:"Pending",created_at:new Date().toISOString()};
                      setDb(p=>({...p,sessions:[newSession,...p.sessions]}));
                      setSessionTopic("");setSessionMsg({ok:true});
                    } catch(e){setSessionMsg({ok:false,text:e.message});}
                    setSessionSubmitting(false);
                  }}>{sessionSubmitting?"Submitting...":"Submit Request"}</Btn>
                </CardBody>
              </Card>
              <Card style={{marginBottom:0}}>
                <CardHeader title="My Requests"/>
                <div style={{padding:"8px 0"}}>
                  {db.sessions.filter(s=>s.ir_email===user.email).length===0?(
                    <div style={{textAlign:"center",padding:"32px 20px",color:G.grayText,fontSize:13}}>No requests yet.</div>
                  ):db.sessions.filter(s=>s.ir_email===user.email).map((s,i)=>(
                    <div key={i} style={{padding:"12px 20px",borderBottom:`1px solid ${G.grayMid}`}}>
                      <div style={{fontSize:13,fontWeight:600,marginBottom:6}}>{s.topic}</div>
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <span style={{fontSize:11,color:G.grayFaint}}>{new Date(s.created_at).toLocaleDateString("en-IN",{day:"numeric",month:"short",year:"numeric"})}</span>
                        <StatusBadge status={s.status}/>
                      </div>
                    </div>
                  ))}
                </div>
              </Card>
            </div>
          </>
        )}
      </Layout>
    );
  }

  /* ADMIN DASHBOARD */
  const allScores=[];
  db.irs.forEach(ir=>db.quizzes.forEach(q=>{const s=irScore(q,ir.email);if(s)allScores.push(s.pct);}));
  const batchAvg=allScores.length?Math.round(allScores.reduce((a,b)=>a+b,0)/allScores.length):0;
  const passRate=allScores.length?Math.round(allScores.filter(p=>p>=PASS_MARK).length/allScores.length*100):0;

  return (
    <Layout sidebar={[["upload","Upload Results"],["sched_a","Training Schedule"],["answers","IR Answers"],["irs","IR List"],["batch","Batch Overview"],["quizzes","Manage Quizzes"],["sessions","Session Requests"]].map(([p,l])=>(
      <NavItem key={p} icon={ICONS[p]} label={l} active={adminPage===p} onClick={()=>setAdminPage(p)}/>
    ))}>

      {adminPage==="upload"&&(
        <>
          <div style={{marginBottom:28}}><div style={{fontSize:22,fontWeight:700}}>Upload Quiz Results</div><div style={{fontSize:13,color:G.grayText}}>Upload your filled Excel sheet.</div></div>
          <div onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();if(e.dataTransfer.files[0])handleFile(e.dataTransfer.files[0]);}} onClick={()=>document.getElementById("file-inp").click()}
            style={{border:`2px dashed ${G.grayMid}`,borderRadius:16,padding:"52px 32px",textAlign:"center",cursor:"pointer",background:G.gray}}>
            <div style={{width:60,height:60,borderRadius:"50%",background:G.green,color:G.black,display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 16px",fontSize:26}}>&#x1F4CA;</div>
            <div style={{fontSize:18,fontWeight:700,marginBottom:6}}>Drop your Excel file here</div>
            <div style={{fontSize:13,color:G.grayText,maxWidth:380,margin:"0 auto 20px"}}>Upload your Scoring Sheet Excel. Name, Email, answers, marks auto-detected.</div>
            <Btn variant="green" size="sm">Choose File (.xlsx)</Btn>
          </div>
          <input id="file-inp" type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>e.target.files[0]&&handleFile(e.target.files[0])}/>
          {uploadStatus&&(
            <div style={{marginTop:20,border:`1.5px solid ${uploadStatus.ok?G.green:"#fecaca"}`,background:uploadStatus.ok?G.greenPale:G.redLight,borderRadius:16,padding:"20px 24px"}}>
              {uploadStatus.ok?(<>
                <div style={{width:44,height:44,borderRadius:"50%",background:G.green,color:G.black,display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,marginBottom:12}}>&#x2713;</div>
                <div style={{fontSize:16,fontWeight:700,marginBottom:4}}>Upload successful!</div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginTop:14}}>
                  <Metric label="Quiz" value="" note={uploadStatus.title}/>
                  <Metric label="IRs" value={uploadStatus.nIRs}/>
                  <Metric label="Questions" value={uploadStatus.nQs}/>
                </div>
              </>):<div style={{fontSize:13,color:G.red}}><strong>Upload failed:</strong> {uploadStatus.msg}</div>}
            </div>
          )}
        </>
      )}

      {adminPage==="sched_a"&&(
        <>
          <div style={{marginBottom:28,display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
            <div><div style={{fontSize:22,fontWeight:700}}>Training Schedule</div><div style={{fontSize:13,color:G.grayText}}>Manage daily topics, content, questions, and quiz links.</div></div>
            <Btn variant="green" size="sm" onClick={addNewDay}>+ Add Day</Btn>
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            {db.schedule.map(day=>(
              <Card key={day.day} style={{marginBottom:0}}>
                <div style={{padding:"16px 20px"}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8,marginBottom:editingDay===day.day?16:0}}>
                    <div style={{display:"flex",alignItems:"center",gap:12}}>
                      <div style={{width:36,height:36,borderRadius:10,background:G.black,color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:700,flexShrink:0}}>D{day.day}</div>
                      <div>
                        <div style={{fontSize:13,fontWeight:700}}>Day {day.day}</div>
                        <div style={{fontSize:12,color:G.grayText,maxWidth:400,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{day.topics||"No topics set"}</div>
                      </div>
                    </div>
                    <div style={{display:"flex",gap:8,alignItems:"center"}}>
                      {db.dayContent[day.day]?.html&&<span style={{fontSize:11,fontWeight:600,padding:"2px 8px",borderRadius:100,background:"#dcfce7",color:"#16a34a"}}>Content</span>}
                      {db.dayContent[day.day]?.questions?.length>0&&<span style={{fontSize:11,fontWeight:600,padding:"2px 8px",borderRadius:100,background:G.amberLight,color:G.amber}}>{db.dayContent[day.day].questions.length} Q</span>}
                      {day.quiz_url&&<span style={{fontSize:11,fontWeight:600,padding:"2px 8px",borderRadius:100,background:"#ede9fe",color:"#7c3aed"}}>Quiz</span>}
                      <Btn variant="outline" size="sm" onClick={()=>{
                        if (editingDay===day.day){setEditingDay(null);setEditForm({});setEditContent({html:"",questions:[]});}
                        else{
                          setEditingDay(day.day);
                          setEditForm({...day});
                          const existing=db.dayContent[day.day]||{html:"",questions:[]};
                          setEditContent({html:existing.html||"",questions:Array.isArray(existing.questions)?existing.questions:[]});
                        }
                      }}>{editingDay===day.day?"Cancel":"Edit"}</Btn>
                    </div>
                  </div>

                  {editingDay===day.day&&(
                    <div style={{borderTop:`1px solid ${G.grayMid}`,paddingTop:16}}>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
                        <div>
                          <label style={{display:"block",fontSize:11,fontWeight:700,color:G.grayText,marginBottom:5,textTransform:"uppercase",letterSpacing:".5px"}}>Topics</label>
                          <textarea value={editForm.topics||""} onChange={e=>setEditForm(p=>({...p,topics:e.target.value}))} rows={3} style={taStyle} placeholder="e.g. Bonds, Interest Rates"/>
                        </div>
                        <div>
                          <label style={{display:"block",fontSize:11,fontWeight:700,color:G.grayText,marginBottom:5,textTransform:"uppercase",letterSpacing:".5px"}}>Quiz Link (Google Form URL)</label>
                          <input value={editForm.quiz_url||""} onChange={e=>setEditForm(p=>({...p,quiz_url:e.target.value}))} style={inputStyle} placeholder="https://forms.google.com/..."/>
                        </div>
                      </div>

                      <div style={{marginBottom:16}}>
                        <label style={{display:"block",fontSize:11,fontWeight:700,color:G.grayText,marginBottom:5,textTransform:"uppercase",letterSpacing:".5px"}}>Reading Material Links (one per line)</label>
                        <textarea value={editForm.reading_links||""} onChange={e=>setEditForm(p=>({...p,reading_links:e.target.value}))} rows={3} style={taStyle} placeholder="1. Article title: https://link.com\n2. https://youtube.com/watch?v=...\n3. Read about XYZ: https://article.com"/>
                      </div>

                      <div style={{marginBottom:16}}>
                        <label style={{display:"block",fontSize:11,fontWeight:700,color:G.grayText,marginBottom:8,textTransform:"uppercase",letterSpacing:".5px"}}>Reading Material Content (rich text)</label>
                        <RichEditor value={editContent.html} onChange={html=>setEditContent(p=>({...p,html}))} day={day.day}/>
                      </div>

                      <div style={{marginBottom:16}}>
                        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
                          <label style={{fontSize:11,fontWeight:700,color:G.grayText,textTransform:"uppercase",letterSpacing:".5px"}}>Questions (copy-paste disabled for IRs)</label>
                          <Btn variant="outline" size="sm" onClick={()=>setEditContent(p=>({...p,questions:[...p.questions,""]}))}>+ Add Question</Btn>
                        </div>
                        {editContent.questions.map((q,i)=>(
                          <div key={i} style={{display:"flex",gap:8,marginBottom:8,alignItems:"flex-start"}}>
                            <div style={{fontSize:12,fontWeight:700,color:G.grayFaint,paddingTop:10,width:24,flexShrink:0}}>Q{i+1}</div>
                            <input value={q} onChange={e=>setEditContent(p=>({...p,questions:p.questions.map((qq,j)=>j===i?e.target.value:qq)}))} style={{...inputStyle,flex:1}} placeholder={`Question ${i+1}`}/>
                            <Btn variant="danger" size="sm" onClick={()=>setEditContent(p=>({...p,questions:p.questions.filter((_,j)=>j!==i)}))}>Remove</Btn>
                          </div>
                        ))}
                        {editContent.questions.length===0&&<div style={{fontSize:12,color:G.grayFaint}}>No questions added yet. Click "+ Add Question" to add one.</div>}
                      </div>

                      <div style={{display:"flex",gap:8}}>
                        <Btn variant="green" size="sm" disabled={schedSaving} onClick={async()=>{
                          setSchedSaving(true);
                          // Save content first before schedule closes the panel
                          const htmlToSave = editContent.html;
                          const questionsToSave = editContent.questions;
                          await supabase.from("day_content").upsert({day:day.day,content_html:htmlToSave,questions:questionsToSave},{onConflict:"day"});
                          setDb(p=>({...p,dayContent:{...p.dayContent,[day.day]:{html:htmlToSave,questions:questionsToSave}}}));
                          await supabase.from("training_schedule").upsert({...day,...editForm},{onConflict:"day"});
                          setDb(p=>({...p,schedule:p.schedule.map(d=>d.day===day.day?{...day,...editForm}:d)}));
                          setEditingDay(null);setEditForm({});
                          setSchedSaving(false);
                        }}>{schedSaving?"Saving...":"Save Day"}</Btn>
                      </div>

                      {/* Tasks for this day */}
                      <div style={{marginTop:20,paddingTop:16,borderTop:`1px solid ${G.grayMid}`}}>
                        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
                          <label style={{fontSize:11,fontWeight:700,color:G.grayText,textTransform:"uppercase",letterSpacing:".5px"}}>Tasks for Day {day.day}</label>
                        </div>
                        <div style={{marginBottom:10}}>
                          <label style={{display:"block",fontSize:11,fontWeight:700,color:G.grayText,marginBottom:4,textTransform:"uppercase",letterSpacing:".5px"}}>Task Title</label>
                          <input value={taskDay===day.day?taskForm.title:""} onChange={e=>{setTaskDay(day.day);setTaskForm(p=>({...p,title:e.target.value}));}} style={inputStyle} placeholder="e.g. Read NBFC regulations"/>
                        </div>
                        <div style={{marginBottom:10}}>
                          <label style={{display:"block",fontSize:11,fontWeight:700,color:G.grayText,marginBottom:4,textTransform:"uppercase",letterSpacing:".5px"}}>Notes (optional)</label>
                          <textarea value={taskDay===day.day?taskForm.notes:""} onChange={e=>{setTaskDay(day.day);setTaskForm(p=>({...p,notes:e.target.value}));}} style={taStyle} placeholder="Any additional instructions..."/>
                        </div>
                        <Btn variant="green" size="sm" disabled={taskSaving||!taskForm.title} onClick={async()=>{
                          if (taskDay!==day.day||!taskForm.title) return;
                          setTaskSaving(true);
                          const {data,error}=await supabase.from("tasks").insert({day:day.day,ir_email:"all",title:taskForm.title,notes:taskForm.notes,status:"Pending"}).select();
                          if (!error&&data) {
                            setDb(p=>({...p,tasks:[data[0],...p.tasks]}));
                            setTaskForm({ir_email:"all",title:"",notes:""});
                            setTaskDay(null);
                          }
                          setTaskSaving(false);
                        }}>{taskSaving?"Adding...":"Add Task"}</Btn>

                        {/* Existing tasks for this day */}
                        {db.tasks.filter(t=>t.day===day.day).length>0&&(
                          <div style={{marginTop:14}}>
                            <div style={{fontSize:11,fontWeight:700,color:G.grayText,textTransform:"uppercase",letterSpacing:".5px",marginBottom:8}}>Tasks Added</div>
                            {db.tasks.filter(t=>t.day===day.day).map((task,i)=>(
                              <div key={task.id} style={{display:"flex",alignItems:"flex-start",gap:10,padding:"10px 12px",background:G.gray,borderRadius:10,marginBottom:6}}>
                                <div style={{flex:1}}>
                                  <div style={{fontSize:12,fontWeight:600}}>{task.title}</div>
                                  {task.notes&&<div style={{fontSize:11,color:G.grayText,marginTop:2}}>{task.notes}</div>}
                                </div>
                                <Btn variant="danger" size="sm" onClick={async()=>{
                                  await supabase.from("tasks").delete().eq("id",task.id);
                                  setDb(p=>({...p,tasks:p.tasks.filter(t=>t.id!==task.id)}));
                                }}>Remove</Btn>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>
                  )}
                </div>
              </Card>
            ))}
          </div>
        </>
      )}

      {adminPage==="answers"&&(
        <>
          <div style={{marginBottom:28,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
            <div><div style={{fontSize:22,fontWeight:700}}>IR Answers</div><div style={{fontSize:13,color:G.grayText}}>All answers submitted by IRs for each day.</div></div>
            <Btn variant="green" size="sm" onClick={()=>{
              const rows=allAnswers.flatMap(a=>{
                const qs=db.dayContent[a.day]?.questions||[];
                return qs.map((q,i)=>({
                  "IR Email":a.ir_email,
                  "Day":a.day,
                  "Question":q,
                  "Answer":a.answers[i]||"",
                  "Submitted":new Date(a.submitted_at).toLocaleDateString("en-IN",{day:"numeric",month:"short",year:"numeric"}),
                }));
              });
              const ws=XLSX.utils.json_to_sheet(rows);
              const wb2=XLSX.utils.book_new();
              XLSX.utils.book_append_sheet(wb2,ws,"IR Answers");
              XLSX.writeFile(wb2,"Wint_IR_Answers.xlsx");
            }}>Download Excel</Btn>
          </div>
          {allAnswers.length===0?(
            <Card><CardBody><div style={{textAlign:"center",padding:40,color:G.grayText}}>No answers submitted yet.</div></CardBody></Card>
          ):db.schedule.map(day=>{
            const dayAnswers=allAnswers.filter(a=>a.day===day.day);
            if (!dayAnswers.length) return null;
            const qs=db.dayContent[day.day]?.questions||[];
            return (
              <Card key={day.day}>
                <CardHeader title={`Day ${day.day} - ${day.topics||"No topics"}`} right={<span style={{fontSize:12,color:G.grayFaint}}>{dayAnswers.length} submissions</span>}/>
                <div style={{overflowX:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse"}}>
                    <thead><tr>
                      <th style={{textAlign:"left",fontSize:11,fontWeight:700,color:G.grayFaint,padding:"10px 16px",borderBottom:`1.5px solid ${G.grayMid}`,background:G.gray,textTransform:"uppercase",letterSpacing:".5px"}}>IR</th>
                      {qs.map((q,i)=><th key={i} style={{textAlign:"left",fontSize:11,fontWeight:700,color:G.grayFaint,padding:"10px 16px",borderBottom:`1.5px solid ${G.grayMid}`,background:G.gray,textTransform:"uppercase",letterSpacing:".5px",maxWidth:200}}>Q{i+1}: {q.slice(0,40)}{q.length>40?"...":""}</th>)}
                      <th style={{textAlign:"left",fontSize:11,fontWeight:700,color:G.grayFaint,padding:"10px 16px",borderBottom:`1.5px solid ${G.grayMid}`,background:G.gray,textTransform:"uppercase",letterSpacing:".5px"}}>Submitted</th>
                    </tr></thead>
                    <tbody>
                      {dayAnswers.map((a,i)=>(
                        <tr key={i} style={{borderBottom:`1px solid ${G.grayMid}`}}>
                          <td style={{padding:"10px 16px",fontSize:13,fontWeight:600}}>{a.ir_email}</td>
                          {qs.map((_,qi)=><td key={qi} style={{padding:"10px 16px",fontSize:13,color:G.grayText,maxWidth:240,verticalAlign:"top"}}>{a.answers[qi]||"-"}</td>)}
                          <td style={{padding:"10px 16px",fontSize:12,color:G.grayFaint,whiteSpace:"nowrap"}}>{new Date(a.submitted_at).toLocaleDateString("en-IN",{day:"numeric",month:"short",year:"numeric"})}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </Card>
            );
          })}
        </>
      )}

      {adminPage==="irs"&&(
        <>
          <div style={{marginBottom:28,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
            <div><div style={{fontSize:22,fontWeight:700}}>IR List</div><div style={{fontSize:13,color:G.grayText}}>All registered Investor Relations trainees.</div></div>
            <Btn variant="green" size="sm" onClick={()=>setShowAddModal(true)}>+ Add IR</Btn>
          </div>
          <Card>
            <CardBody>
              {db.irs.map((ir,i)=>(
                <div key={ir.email} style={{display:"flex",alignItems:"center",gap:12,padding:"10px 0",borderBottom:i<db.irs.length-1?`1px solid ${G.grayMid}`:"none"}}>
                  <div style={{width:32,height:32,borderRadius:"50%",background:"rgba(108,254,95,.15)",color:G.greenDark,fontSize:12,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",border:`1.5px solid ${G.green}`,flexShrink:0}}>{inits(ir.name)}</div>
                  <div style={{flex:1}}><div style={{fontSize:13,fontWeight:600}}>{ir.name}</div><div style={{fontSize:11,color:G.grayFaint}}>{ir.email}</div></div>
                  <Btn variant="danger" size="sm" onClick={async()=>{await supabase.from("irs").delete().eq("email",ir.email);setDb(p=>({...p,irs:p.irs.filter((_,j)=>j!==i)}));}}>Remove</Btn>
                </div>
              ))}
            </CardBody>
          </Card>
          {showAddModal&&(
            <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.4)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:200}}>
              <div style={{background:G.white,borderRadius:16,padding:28,width:"100%",maxWidth:400}}>
                <h3 style={{fontSize:17,fontWeight:700,marginBottom:18}}>Add IR</h3>
                {[["Full Name","text",newName,setNewName,"Priya Sharma"],["Email","email",newEmail,setNewEmail,"priya@wintwealth.com"]].map(([lbl,type,val,set,ph])=>(
                  <div key={lbl} style={{marginBottom:14}}>
                    <label style={{display:"block",fontSize:11,fontWeight:700,color:G.grayText,marginBottom:5,textTransform:"uppercase",letterSpacing:".5px"}}>{lbl}</label>
                    <input type={type} value={val} onChange={e=>set(e.target.value)} placeholder={ph} style={inputStyle}/>
                  </div>
                ))}
                <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:20}}>
                  <Btn variant="outline" size="sm" onClick={()=>setShowAddModal(false)}>Cancel</Btn>
                  <Btn variant="green" size="sm" onClick={async()=>{if(!newName||!newEmail)return;const ir={name:newName,email:newEmail.toLowerCase()};await supabase.from("irs").upsert(ir,{onConflict:"email"});setDb(p=>({...p,irs:[...p.irs,ir]}));setNewName("");setNewEmail("");setShowAddModal(false);}}>Add IR</Btn>
                </div>
              </div>
            </div>
          )}
        </>
      )}

      {adminPage==="batch"&&(
        <>
          <div style={{marginBottom:28}}><div style={{fontSize:22,fontWeight:700}}>Batch Overview</div><div style={{fontSize:13,color:G.grayText}}>Performance across all IRs and quizzes.</div></div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:24}}>
            <Metric accent label="Batch Avg" value={`${batchAvg}%`}/>
            <Metric label="Pass Rate" value={`${passRate}%`} note={`${PASS_MARK}% to pass`}/>
            <Metric label="Total IRs" value={db.irs.length}/>
            <Metric label="Quizzes" value={db.quizzes.length}/>
          </div>
          <Card>
            <CardHeader title="IR Scores"/>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead><tr>{["IR","Avg Score","Marks","Progress","Status"].map(h=><th key={h} style={{textAlign:"left",fontSize:11,fontWeight:700,letterSpacing:".5px",textTransform:"uppercase",color:G.grayFaint,padding:"10px 16px",borderBottom:`1.5px solid ${G.grayMid}`,background:G.gray}}>{h}</th>)}</tr></thead>
                <tbody>
                  {db.irs.map(ir=>{
                    const scores=db.quizzes.map(q=>irScore(q,ir.email)).filter(Boolean);
                    if (!scores.length) return <tr key={ir.email}><td colSpan={5} style={{padding:"12px 16px",fontSize:12,color:G.grayText}}>{ir.name} - no results yet</td></tr>;
                    const avg=Math.round(scores.reduce((s,x)=>s+x.pct,0)/scores.length);
                    const earned=scores.reduce((s,x)=>s+x.earned,0);
                    const max=scores.reduce((s,x)=>s+x.max,0);
                    return (
                      <tr key={ir.email} style={{borderBottom:`1px solid ${G.grayMid}`}}>
                        <td style={{padding:"11px 16px"}}><div style={{display:"flex",alignItems:"center",gap:10}}><div style={{width:30,height:30,borderRadius:"50%",background:"rgba(108,254,95,.15)",color:G.greenDark,fontSize:11,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",border:`1.5px solid ${G.green}`}}>{inits(ir.name)}</div><div><div style={{fontWeight:600,fontSize:13}}>{ir.name}</div><div style={{fontSize:11,color:G.grayFaint}}>{ir.email}</div></div></div></td>
                        <td style={{padding:"11px 16px"}}><Pill pct={avg} label={`${avg}%`}/></td>
                        <td style={{padding:"11px 16px",fontWeight:600,fontSize:13}}>{earned}/{max}</td>
                        <td style={{padding:"11px 16px",minWidth:120}}><ProgBar pct={avg}/></td>
                        <td style={{padding:"11px 16px"}}><span style={{fontSize:11,fontWeight:700,padding:"3px 10px",borderRadius:100,background:avg>=PASS_MARK?"#dcfce7":G.redLight,color:avg>=PASS_MARK?"#16a34a":G.red}}>{avg>=PASS_MARK?"Pass":"Needs work"}</span></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </Card>
        </>
      )}

      {adminPage==="quizzes"&&(
        <>
          <div style={{marginBottom:28}}><div style={{fontSize:22,fontWeight:700}}>Manage Quizzes</div><div style={{fontSize:13,color:G.grayText}}>Delete quizzes and their results permanently.</div></div>
          <Card>
            <CardBody>
              {db.quizzes.length===0?<div style={{textAlign:"center",padding:40,color:G.grayText}}>No quizzes uploaded yet.</div>
              :db.quizzes.map((q,i)=>(
                <div key={q.id} style={{display:"flex",alignItems:"center",gap:12,padding:"12px 0",borderBottom:i<db.quizzes.length-1?`1px solid ${G.grayMid}`:"none"}}>
                  <div style={{flex:1}}>
                    <div style={{fontSize:13,fontWeight:600}}>{q.title}</div>
                    <div style={{fontSize:11,color:G.grayFaint}}>{q.date} - {Object.keys(q.results).length} IRs</div>
                  </div>
                  <Btn variant="danger" size="sm" onClick={async()=>{
                    if (!window.confirm("Delete this quiz and all its results permanently?")) return;
                    await supabase.from("results").delete().eq("quiz_id",q.id);
                    await supabase.from("quizzes").delete().eq("id",q.id);
                    setDb(p=>({...p,quizzes:p.quizzes.filter((_,j)=>j!==i)}));
                  }}>Delete</Btn>
                </div>
              ))}
            </CardBody>
          </Card>
        </>
      )}

      {adminPage==="sessions"&&(
        <>
          <div style={{marginBottom:28,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
            <div><div style={{fontSize:22,fontWeight:700}}>Session Requests</div><div style={{fontSize:13,color:G.grayText}}>All IR session requests.</div></div>
            <Btn variant="green" size="sm" onClick={()=>{
              const ws=XLSX.utils.json_to_sheet(db.sessions.map(s=({"Name":s.ir_name,"Email":s.ir_email,"Topic":s.topic,"Date Requested":new Date(s.created_at).toLocaleDateString("en-IN",{day:"numeric",month:"short",year:"numeric"}),"Status":s.status})));
              const wb2=XLSX.utils.book_new();
              XLSX.utils.book_append_sheet(wb2,ws,"Session Requests");
              XLSX.writeFile(wb2,"Wint_Session_Requests.xlsx");
            }}>Download Excel</Btn>
          </div>
          <Card>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead><tr>{["Name","Email","Topic","Date","Status","Action"].map(h=><th key={h} style={{textAlign:"left",fontSize:11,fontWeight:700,letterSpacing:".5px",textTransform:"uppercase",color:G.grayFaint,padding:"10px 16px",borderBottom:`1.5px solid ${G.grayMid}`,background:G.gray}}>{h}</th>)}</tr></thead>
                <tbody>
                  {db.sessions.length===0?<tr><td colSpan={6} style={{padding:"40px 16px",textAlign:"center",color:G.grayText,fontSize:13}}>No session requests yet.</td></tr>
                  :db.sessions.map((s,i)=>(
                    <tr key={i} style={{borderBottom:`1px solid ${G.grayMid}`}}>
                      <td style={{padding:"11px 16px",fontWeight:600,fontSize:13}}>{s.ir_name}</td>
                      <td style={{padding:"11px 16px",fontSize:12,color:G.grayText}}>{s.ir_email}</td>
                      <td style={{padding:"11px 16px",fontSize:13,maxWidth:200}}>{s.topic}</td>
                      <td style={{padding:"11px 16px",fontSize:12,color:G.grayFaint,whiteSpace:"nowrap"}}>{new Date(s.created_at).toLocaleDateString("en-IN",{day:"numeric",month:"short",year:"numeric"})}</td>
                      <td style={{padding:"11px 16px"}}><StatusBadge status={s.status}/></td>
                      <td style={{padding:"11px 16px"}}>
                        <div style={{display:"flex",gap:6}}>
                          {["Pending","Confirmed","Rejected"].filter(st=>st!==s.status).map(st=>(
                            <Btn key={st} variant="outline" size="sm" onClick={async()=>{
                              await supabase.from("session_requests").update({status:st}).eq("id",s.id);
                              setDb(p=>({...p,sessions:p.sessions.map((r,j)=>j===i?{...r,status:st}:r)}));
                            }}>{st}</Btn>
                          ))}
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        </>
      )}
    </Layout>
  );
}

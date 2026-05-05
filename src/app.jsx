import { useState, useEffect, useRef } from "react"
import { createClient } from "@supabase/supabase-js"
import * as XLSX from "xlsx"
import {
  BarChart, Bar, AreaChart, Area, PieChart, Pie, Cell,
  XAxis, YAxis, Tooltip, ResponsiveContainer, Legend
} from "recharts"

// ── Supabase client ───────────────────────────────────────────────────────
const supabase = createClient(
  "https://knzhdfykzerwwylwslts.supabase.co",
  "YOUR_ANON_KEY_HERE"   // ← paste your anon public key here (Settings → API)
)

// ── Colour tokens ─────────────────────────────────────────────────────────
const C = {
  navy:"#0C2340", navyMid:"#153458", navyLight:"#1A4068",
  blue:"#185FA5", blueLight:"#378ADD",
  green:"#1D9E75", red:"#E24B4A", amber:"#EF9F27",
  text:"#F0F4F8", textMuted:"#8FA8C8", textDim:"#5A7A9A",
  border:"rgba(255,255,255,0.08)", borderMid:"rgba(255,255,255,0.14)",
  card:"rgba(255,255,255,0.04)",
}

// ── Formatters ────────────────────────────────────────────────────────────
const fmt  = (n,d=0) => new Intl.NumberFormat("en-US",{minimumFractionDigits:d,maximumFractionDigits:d}).format(n)
const fmtM = n => { const a=Math.abs(n); const s=a>=1e6?`$${fmt(a/1e6,2)}M`:a>=1e3?`$${fmt(a/1e3,0)}k`:`$${fmt(a,0)}`; return n<0?`-${s}`:s }
const fmtK = v => v>=1e6?`$${(v/1e6).toFixed(1)}M`:v>=1e3?`$${(v/1e3).toFixed(0)}k`:`$${v}`

// ── QB row-label map ──────────────────────────────────────────────────────
const QB_MAP = {
  gross_income: ["total income","total revenue","gross revenue","gross income","total operating income"],
  expenses:     ["total expenses","total operating expenses","operating expenses"],
  noi:          ["net operating income","net income","net profit","operating income"],
  interest:     ["interest expense","loan interest","interest paid","mortgage interest"],
  principal:    ["loan principal","principal repayment","principal payments","debt principal"],
}
const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
const norm = s => (s||"").toLowerCase().trim().replace(/[^a-z0-9 ]/g,"")
const matchKey = label => { const n=norm(label); for(const [k,a] of Object.entries(QB_MAP)){if(a.some(x=>n.includes(x)||x.includes(n)))return k} return null }

// ── Parse QuickBooks monthly Excel ───────────────────────────────────────
const parseQB = file => new Promise((res,rej) => {
  const r = new FileReader()
  r.onload = e => {
    try {
      const wb = XLSX.read(e.target.result,{type:"array"})
      const ws = wb.Sheets[wb.SheetNames[0]]
      const rows = XLSX.utils.sheet_to_json(ws,{header:1,defval:""})
      let hIdx=-1, mCols={}
      for(let i=0;i<Math.min(rows.length,12);i++){
        const mc={}
        rows[i].forEach((c,j)=>{ const mi=MONTHS.findIndex(m=>norm(String(c)).startsWith(m.toLowerCase())); if(mi>=0) mc[mi]=j })
        if(Object.keys(mc).length>=2){hIdx=i;mCols=mc;break}
      }
      if(hIdx<0){rej(new Error("No monthly columns found. Use Reports → P&L → Columns: Months in QuickBooks."));return}
      const ext={}; Object.keys(mCols).forEach(mi=>{ext[mi]={}})
      for(let r2=hIdx+1;r2<rows.length;r2++){
        const row=rows[r2]; const key=matchKey(String(row[0]||""))
        if(!key) continue
        for(const [mi,ci] of Object.entries(mCols)){
          const raw=row[ci]; const val=typeof raw==="number"?raw:parseFloat(String(raw).replace(/[$,()]/g,""))||0
          ext[mi][key]=val
        }
      }
      const plRows=Object.entries(ext).map(([mi,v])=>({
        month:MONTHS[+mi], month_index:+mi, year:new Date().getFullYear(),
        gross_income:v.gross_income||0, expenses:v.expenses||0,
        noi:v.noi||(v.gross_income-v.expenses)||0,
        interest:v.interest||0, principal:v.principal||0,
        free_cf:(v.noi||0)-(v.interest||0)-(v.principal||0)
      })).filter(r=>r.gross_income>0)
      res(plRows)
    } catch(err){rej(err)}
  }
  r.onerror=()=>rej(new Error("File read failed"))
  r.readAsArrayBuffer(file)
})

// ══════════════════════════════════════════════════════════════════════════
// COMPONENTS
// ══════════════════════════════════════════════════════════════════════════

function KPI({label,value,sub,subColor}){
  return(
    <div style={{background:C.card,border:`0.5px solid ${C.border}`,borderRadius:10,padding:"12px 14px"}}>
      <div style={{fontSize:11,color:C.textMuted,marginBottom:4}}>{label}</div>
      <div style={{fontSize:19,fontWeight:500,color:C.text}}>{value}</div>
      {sub&&<div style={{fontSize:11,marginTop:3,color:subColor||C.green}}>{sub}</div>}
    </div>
  )
}

function BVA({label,budget,actual,isDscr}){
  const bal=actual-budget
  const pos=isDscr?bal>=0:bal>=0
  const fv = isDscr ? v=>`${v.toFixed(2)}x` : v=>`$${fmt(Math.abs(v))}`
  return(
    <div style={{background:C.card,border:`0.5px solid ${C.border}`,borderRadius:10,padding:"10px 14px"}}>
      <div style={{fontSize:11,fontWeight:500,color:C.blueLight,marginBottom:8,paddingBottom:6,borderBottom:`0.5px solid ${C.border}`}}>{label}</div>
      {[["Budget",budget],["Actual",actual]].map(([l,v])=>(
        <div key={l} style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:4}}>
          <span style={{color:C.textMuted}}>{l}</span>
          <span style={{color:C.text}}>{fv(v)}</span>
        </div>
      ))}
      <div style={{display:"flex",justifyContent:"space-between",fontSize:12,fontWeight:500,borderTop:`0.5px solid ${C.border}`,paddingTop:4,marginTop:4}}>
        <span style={{color:C.textMuted}}>Balance</span>
        <span style={{color:pos?C.green:C.red}}>{isDscr?`${bal>=0?"+":""}${bal.toFixed(2)}x`:`${bal>=0?"+":"-"}$${fmt(Math.abs(bal))}`}</span>
      </div>
    </div>
  )
}

const CHART_COLORS=["#378ADD","#888780","#D85A30","#534AB7","#1D9E75"]

// ── Login ─────────────────────────────────────────────────────────────────
function Login({onDone}){
  const [email,setEmail]=useState("")
  const [pw,setPw]=useState("")
  const [err,setErr]=useState("")
  const [loading,setLoading]=useState(false)
  const submit=async()=>{
    setLoading(true);setErr("")
    const {error}=await supabase.auth.signInWithPassword({email:email.trim(),password:pw})
    if(error){setErr(error.message);setLoading(false)}
    else onDone()
  }
  return(
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",background:C.navy}}>
      <div style={{background:C.navyMid,border:`0.5px solid ${C.border}`,borderRadius:14,padding:"2rem",width:340}}>
        <div style={{textAlign:"center",marginBottom:"1.5rem"}}>
          <div style={{width:44,height:44,borderRadius:10,background:C.blue,display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 10px",fontSize:22,color:"#fff",fontWeight:700}}>W</div>
          <div style={{fontSize:17,fontWeight:500,color:C.text}}>Weigel Family</div>
          <div style={{fontSize:12,color:C.textMuted,marginTop:2}}>Investment Portfolio</div>
        </div>
        {[["Email","email",email,setEmail],["Password","password",pw,setPw]].map(([l,t,v,s])=>(
          <div key={l} style={{marginBottom:12}}>
            <div style={{fontSize:11,color:C.textMuted,marginBottom:4}}>{l}</div>
            <input type={t} value={v} onChange={e=>s(e.target.value)} onKeyDown={e=>e.key==="Enter"&&submit()}
              style={{width:"100%",padding:"8px 10px",background:C.navy,border:`0.5px solid ${C.borderMid}`,borderRadius:7,color:C.text,fontSize:13,outline:"none",boxSizing:"border-box"}}/>
          </div>
        ))}
        {err&&<div style={{fontSize:11,color:C.red,marginBottom:8,textAlign:"center"}}>{err}</div>}
        <button onClick={submit} disabled={loading}
          style={{width:"100%",padding:9,background:C.blue,color:"#fff",border:"none",borderRadius:7,fontSize:14,fontWeight:500,cursor:"pointer",opacity:loading?.6:1}}>
          {loading?"Signing in…":"Sign in"}
        </button>
      </div>
    </div>
  )
}

// ── Dashboard ─────────────────────────────────────────────────────────────
function Dashboard({pl,portfolio,properties}){
  if(!portfolio) return <div style={{color:C.textMuted,padding:40,textAlign:"center"}}>Loading portfolio data…</div>
  const occ=portfolio.occupancy||0
  const prev=portfolio.prev_occupancy||0
  const d=occ-prev
  return(
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
        <div style={{fontSize:13,fontWeight:500,color:C.text}}>Portfolio overview <span style={{color:C.textMuted,fontWeight:400}}>— {portfolio.year}</span></div>
        <div style={{fontSize:11,color:C.textDim}}>Through {MONTHS[portfolio.report_month]||""} {portfolio.year}</div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(5,minmax(0,1fr))",gap:10,marginBottom:12}}>
        <KPI label="Assets under management" value={fmtM(portfolio.aum)} sub="37 assets · 82 tenants"/>
        <KPI label="Total equity"            value={fmtM(portfolio.equity)} sub={`LTV ${fmt((portfolio.commercial_loans/portfolio.aum)*100,2)}%`}/>
        <KPI label="Commercial loans"        value={fmtM(portfolio.commercial_loans)} sub={`Estate: ${fmtM(portfolio.estate_loans)}`} subColor={C.textMuted}/>
        <KPI label="All-cash position"       value={fmtM(portfolio.cash_position)} sub={`Managed: ${fmtM(portfolio.managed_cash)}`}/>
        <KPI label="Portfolio occupancy"     value={`${fmt(occ,1)}%`} sub={`${d>=0?"↑":"↓"} ${fmt(Math.abs(d),1)}% from prev month`} subColor={d>=0?C.green:C.red}/>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(0,1fr))",gap:10,marginBottom:12}}>
        <BVA label="Gross income"         budget={portfolio.budget_gross}    actual={portfolio.actual_gross}/>
        <BVA label="Expenses"             budget={portfolio.budget_expenses}  actual={portfolio.actual_expenses}/>
        <BVA label="Net operating income" budget={portfolio.budget_noi}       actual={portfolio.actual_noi}/>
        <BVA label="DSCR"                 budget={portfolio.budget_dscr||1.51} actual={portfolio.actual_dscr||1.01} isDscr/>
      </div>

      {pl.length>0&&(
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
          <div style={{background:C.card,border:`0.5px solid ${C.border}`,borderRadius:10,padding:"12px 14px"}}>
            <div style={{fontSize:12,color:C.textMuted,marginBottom:8}}>P&amp;L by month</div>
            <ResponsiveContainer width="100%" height={150}>
              <BarChart data={pl} margin={{top:4,right:4,bottom:0,left:0}}>
                <XAxis dataKey="month" tick={{fill:C.textMuted,fontSize:10}} axisLine={false} tickLine={false}/>
                <YAxis tickFormatter={fmtK} tick={{fill:C.textMuted,fontSize:9}} axisLine={false} tickLine={false} width={45}/>
                <Tooltip formatter={v=>`$${fmt(v)}`} contentStyle={{background:C.navyMid,border:`0.5px solid ${C.border}`,borderRadius:6,color:C.text,fontSize:11}}/>
                <Bar dataKey="gross_income" name="Income"    fill="#378ADD" radius={[2,2,0,0]}/>
                <Bar dataKey="expenses"     name="Expenses"  fill="#888780" radius={[2,2,0,0]}/>
                <Bar dataKey="noi"          name="NOI"       fill="#1D9E75" radius={[2,2,0,0]}/>
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div style={{background:C.card,border:`0.5px solid ${C.border}`,borderRadius:10,padding:"12px 14px"}}>
            <div style={{fontSize:12,color:C.textMuted,marginBottom:8}}>Free cash flow trend</div>
            <ResponsiveContainer width="100%" height={150}>
              <AreaChart data={pl} margin={{top:4,right:4,bottom:0,left:0}}>
                <defs>
                  <linearGradient id="cfGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%"  stopColor="#378ADD" stopOpacity={0.3}/>
                    <stop offset="95%" stopColor="#378ADD" stopOpacity={0}/>
                  </linearGradient>
                </defs>
                <XAxis dataKey="month" tick={{fill:C.textMuted,fontSize:10}} axisLine={false} tickLine={false}/>
                <YAxis tickFormatter={fmtK} tick={{fill:C.textMuted,fontSize:9}} axisLine={false} tickLine={false} width={45}/>
                <Tooltip formatter={v=>`$${fmt(v)}`} contentStyle={{background:C.navyMid,border:`0.5px solid ${C.border}`,borderRadius:6,color:C.text,fontSize:11}}/>
                <Area dataKey="free_cf" name="Free CF" stroke="#378ADD" fill="url(#cfGrad)" strokeWidth={1.5}/>
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      {properties.length>0&&(
        <div style={{background:C.card,border:`0.5px solid ${C.border}`,borderRadius:10,padding:"12px 14px"}}>
          <div style={{fontSize:12,color:C.textMuted,marginBottom:10}}>Asset register</div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,tableLayout:"fixed"}}>
            <thead><tr>{["Property","Type","FMV","Debt","NOI","DSCR","Occ %","Last report"].map(h=>(
              <th key={h} style={{textAlign:"left",padding:"4px 8px",borderBottom:`0.5px solid ${C.border}`,color:C.textMuted,fontWeight:500,fontSize:11}}>{h}</th>
            ))}</tr></thead>
            <tbody>{properties.map((p,i)=>(
              <tr key={i}>
                <td style={{padding:"5px 8px",borderBottom:`0.5px solid ${C.border}`,color:C.text}}>{p.name}</td>
                <td style={{padding:"5px 8px",borderBottom:`0.5px solid ${C.border}`}}>
                  <span style={{fontSize:10,padding:"2px 7px",borderRadius:10,background:p.type==="Retail"?"rgba(24,95,165,0.2)":"rgba(29,158,117,0.2)",color:p.type==="Retail"?"#85B7EB":"#5DCAA5"}}>{p.type}</span>
                </td>
                {[fmtM(p.fmv),fmtM(p.debt),fmtM(p.noi)].map((v,j)=>(
                  <td key={j} style={{padding:"5px 8px",borderBottom:`0.5px solid ${C.border}`,color:C.text}}>{v}</td>
                ))}
                <td style={{padding:"5px 8px",borderBottom:`0.5px solid ${C.border}`,color:p.dscr<1?C.red:C.green}}>{(p.dscr||0).toFixed(2)}x</td>
                <td style={{padding:"5px 8px",borderBottom:`0.5px solid ${C.border}`,color:(p.occupancy||0)<90?C.amber:C.text}}>{p.occupancy||0}%</td>
                <td style={{padding:"5px 8px",borderBottom:`0.5px solid ${C.border}`,color:C.textMuted,fontSize:11}}>{p.last_report}</td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      )}
    </div>
  )
}

// ── Upload ────────────────────────────────────────────────────────────────
function Upload({userId,onUploaded}){
  const fileRef=useRef()
  const [type,setType]=useState(null)
  const [status,setStatus]=useState(null) // null | parsing | preview | error
  const [preview,setPreview]=useState([])
  const [errMsg,setErrMsg]=useState("")
  const [log,setLog]=useState([])

  useEffect(()=>{
    supabase.from("upload_log").select("*").order("upload_date",{ascending:false}).limit(20)
      .then(({data})=>{ if(data) setLog(data) })
  },[])

  const handleFile=async e=>{
    const file=e.target.files[0]; if(!file) return
    setStatus("parsing"); setErrMsg("")
    try{
      if(type==="QB"){
        const rows=await parseQB(file)
        setPreview(rows); setStatus("preview")
      } else {
        setStatus("preview"); setPreview([])
      }
      const {error}=await supabase.from("upload_log").insert({
        file_name:file.name, file_type:type, uploaded_by:userId, status:"imported"
      })
      if(!error){ const {data}=await supabase.from("upload_log").select("*").order("upload_date",{ascending:false}).limit(20); if(data) setLog(data) }
    } catch(err){ setStatus("error"); setErrMsg(err.message) }
    fileRef.current.value=""
  }

  const confirmImport=async()=>{
    if(!preview.length) return
    const {error}=await supabase.from("monthly_pl").upsert(
      preview.map(r=>({...r,year:r.year||new Date().getFullYear()})),
      {onConflict:"year,month_index"}
    )
    if(error){ setErrMsg(error.message); return }
    setStatus(null); setPreview([]); onUploaded()
  }

  const boxes=[
    {id:"QB", icon:"📊",title:"QuickBooks export",     desc:"P&L with monthly columns .xlsx/.csv"},
    {id:"PDF",icon:"📄",title:"Property manager PDF",  desc:"Monthly PDF report"},
    {id:"XLS",icon:"📋",title:"Property manager Excel",desc:"Excel financial statement"},
  ]
  return(
    <div>
      <div style={{background:C.card,border:`0.5px solid ${C.border}`,borderRadius:10,padding:"14px 16px",marginBottom:14}}>
        <div style={{fontSize:14,fontWeight:500,color:C.text,marginBottom:12}}>Upload monthly reports</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
          {boxes.map(b=>(
            <div key={b.id} onClick={()=>{setType(b.id);fileRef.current.click()}}
              style={{border:`0.5px dashed ${C.borderMid}`,borderRadius:8,padding:"1.25rem",textAlign:"center",cursor:"pointer"}}
              onMouseEnter={e=>e.currentTarget.style.background=C.card}
              onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
              <div style={{fontSize:24,marginBottom:6}}>{b.icon}</div>
              <div style={{fontSize:13,fontWeight:500,color:C.text}}>{b.title}</div>
              <div style={{fontSize:11,color:C.textMuted,marginTop:3}}>{b.desc}</div>
            </div>
          ))}
        </div>
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,.pdf" style={{display:"none"}} onChange={handleFile}/>
      </div>

      {status==="parsing"&&<div style={{textAlign:"center",color:C.textMuted,padding:20,fontSize:13}}>Parsing file…</div>}

      {status==="error"&&(
        <div style={{background:"rgba(226,75,74,0.1)",border:`0.5px solid ${C.red}`,borderRadius:10,padding:"12px 16px",marginBottom:14,color:C.red,fontSize:13}}>
          {errMsg}
        </div>
      )}

      {status==="preview"&&preview.length>0&&(
        <div style={{background:"rgba(29,158,117,0.08)",border:`0.5px solid ${C.green}`,borderRadius:10,padding:"14px 16px",marginBottom:14}}>
          <div style={{fontSize:13,fontWeight:500,color:C.green,marginBottom:10}}>✓ Parsed {preview.length} months — review before importing</div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr>{["Month","Gross income","Expenses","NOI","Interest","Principal","Free CF"].map(h=>(
              <th key={h} style={{textAlign:"left",padding:"3px 8px",color:C.textMuted,fontWeight:500}}>{h}</th>
            ))}</tr></thead>
            <tbody>{preview.map((r,i)=>(
              <tr key={i}>
                <td style={{padding:"3px 8px",color:C.text}}>{r.month}</td>
                {["gross_income","expenses","noi","interest","principal","free_cf"].map(k=>(
                  <td key={k} style={{padding:"3px 8px",color:k==="free_cf"?(r[k]>=0?C.green:C.red):C.text}}>${fmt(r[k])}</td>
                ))}
              </tr>
            ))}</tbody>
          </table>
          <button onClick={confirmImport}
            style={{marginTop:12,padding:"7px 18px",background:C.green,color:"#fff",border:"none",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:500}}>
            Confirm & save to database
          </button>
        </div>
      )}

      <div style={{background:C.card,border:`0.5px solid ${C.border}`,borderRadius:10,padding:"14px 16px"}}>
        <div style={{fontSize:13,fontWeight:500,color:C.text,marginBottom:10}}>Upload log</div>
        {log.length===0&&<div style={{fontSize:12,color:C.textDim}}>No uploads yet.</div>}
        {log.map((item,i)=>(
          <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:i<log.length-1?`0.5px solid ${C.border}`:"none",fontSize:12}}>
            <div>
              <div style={{color:C.text}}>{item.file_name}</div>
              <div style={{fontSize:10,color:C.textDim,marginTop:1}}>{new Date(item.upload_date).toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"})}</div>
            </div>
            <span style={{fontSize:11,color:item.status==="imported"?C.green:C.amber}}>
              {item.status==="imported"?"✓ Imported":"⏳ Pending"}
            </span>
          </div>
        ))}
      </div>
    </div>
  )
}

// ── Admin ─────────────────────────────────────────────────────────────────
function Admin(){
  const [users,setUsers]=useState([])
  const [newEmail,setNewEmail]=useState("")
  const [newName,setNewName]=useState("")
  const [msg,setMsg]=useState("")

  useEffect(()=>{
    supabase.from("profiles").select("*").then(({data})=>{ if(data) setUsers(data) })
  },[])

  const inviteUser=async()=>{
    if(!newEmail||!newName){setMsg("Please enter name and email.");return}
    const {error}=await supabase.auth.admin.inviteUserByEmail(newEmail)
    if(error){setMsg(error.message);return}
    setMsg(`Invitation sent to ${newEmail}`); setNewEmail(""); setNewName("")
  }

  const settings=[
    ["Reporting period","Monthly"],["Base currency","USD"],
    ["2FA requirement","All users"],["Data retention","7 years"],
    ["DSCR alert threshold","1.10x"],
  ]
  return(
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
      <div>
        <div style={{background:C.card,border:`0.5px solid ${C.border}`,borderRadius:10,padding:"14px 16px",marginBottom:14}}>
          <div style={{fontSize:14,fontWeight:500,color:C.text,marginBottom:12}}>User accounts</div>
          {users.map((u,i)=>(
            <div key={i} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"7px 0",borderBottom:i<users.length-1?`0.5px solid ${C.border}`:"none"}}>
              <div style={{display:"flex",alignItems:"center",gap:10}}>
                <div style={{width:30,height:30,borderRadius:"50%",background:u.role==="admin"?"rgba(24,95,165,0.25)":"rgba(29,158,117,0.2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:500,color:u.role==="admin"?"#85B7EB":"#5DCAA5"}}>
                  {(u.initials||u.name?.slice(0,2)||"??")}
                </div>
                <div>
                  <div style={{fontSize:13,fontWeight:500,color:C.text}}>{u.name||"User"}</div>
                  <div style={{fontSize:11,color:C.textDim}}>{u.role||"viewer"}</div>
                </div>
              </div>
              <span style={{fontSize:10,padding:"2px 9px",borderRadius:10,background:u.role==="admin"?"rgba(24,95,165,0.2)":"rgba(255,255,255,0.06)",color:u.role==="admin"?"#85B7EB":C.textMuted}}>
                {u.role==="admin"?"Admin":"Viewer"}
              </span>
            </div>
          ))}
          {users.length===0&&<div style={{fontSize:12,color:C.textDim}}>No profiles yet. Add users below.</div>}
        </div>
        <div style={{background:C.card,border:`0.5px solid ${C.border}`,borderRadius:10,padding:"14px 16px"}}>
          <div style={{fontSize:14,fontWeight:500,color:C.text,marginBottom:12}}>Invite new user</div>
          {[["Full name",newName,setNewName,"text"],["Email address",newEmail,setNewEmail,"email"]].map(([l,v,s,t])=>(
            <div key={l} style={{marginBottom:10}}>
              <div style={{fontSize:11,color:C.textMuted,marginBottom:3}}>{l}</div>
              <input type={t} value={v} onChange={e=>s(e.target.value)}
                style={{width:"100%",padding:"7px 10px",background:C.navy,border:`0.5px solid ${C.borderMid}`,borderRadius:7,color:C.text,fontSize:13,outline:"none",boxSizing:"border-box"}}/>
            </div>
          ))}
          {msg&&<div style={{fontSize:11,color:msg.includes("sent")?C.green:C.red,marginBottom:8}}>{msg}</div>}
          <button onClick={inviteUser}
            style={{padding:"7px 16px",background:C.blue,color:"#fff",border:"none",borderRadius:7,fontSize:13,cursor:"pointer"}}>
            Send invite
          </button>
        </div>
      </div>
      <div>
        <div style={{background:C.card,border:`0.5px solid ${C.border}`,borderRadius:10,padding:"14px 16px",marginBottom:14}}>
          <div style={{fontSize:14,fontWeight:500,color:C.text,marginBottom:12}}>System settings</div>
          {settings.map(([l,v],i)=>(
            <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:i<settings.length-1?`0.5px solid ${C.border}`:"none",fontSize:13}}>
              <span style={{color:C.textMuted}}>{l}</span>
              <span style={{color:C.blueLight,cursor:"pointer"}}>{v} ▾</span>
            </div>
          ))}
        </div>
        <div style={{background:C.card,border:`0.5px solid ${C.border}`,borderRadius:10,padding:"14px 16px"}}>
          <div style={{fontSize:14,fontWeight:500,color:C.text,marginBottom:6}}>AI-assisted updates</div>
          <div style={{fontSize:12,color:C.textMuted,marginBottom:12}}>Use Claude to modify the dashboard, add KPIs, or adjust reporting logic.</div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {["Add KPI card","Add new chart","Modify layout","Add property"].map(l=>(
              <button key={l} onClick={()=>sendPrompt(`Update the Weigel dashboard: ${l}`)}
                style={{padding:"6px 12px",background:C.blue,color:"#fff",border:"none",borderRadius:7,fontSize:12,cursor:"pointer"}}>
                {l} ↗
              </button>
            ))}
          </div>
        </div>
      </div>
    </div>
  )
}

// ══════════════════════════════════════════════════════════════════════════
// ROOT
// ══════════════════════════════════════════════════════════════════════════
export default function App(){
  const [session,setSession]=useState(undefined) // undefined=loading
  const [profile,setProfile]=useState(null)
  const [tab,setTab]=useState("dashboard")
  const [pl,setPL]=useState([])
  const [portfolio,setPortfolio]=useState(null)
  const [properties,setProperties]=useState([])

  useEffect(()=>{
    supabase.auth.getSession().then(({data:{session:s}})=>{
      setSession(s)
      if(s) loadProfile(s.user.id)
    })
    const {data:{subscription}}=supabase.auth.onAuthStateChange((_,s)=>{
      setSession(s)
      if(s) loadProfile(s.user.id); else setProfile(null)
    })
    return ()=>subscription.unsubscribe()
  },[])

  const loadProfile=async id=>{
    const {data}=await supabase.from("profiles").select("*").eq("id",id).single()
    setProfile(data)
  }

  const loadData=async()=>{
    const yr=new Date().getFullYear()
    const [{data:plData},{data:portData},{data:propData}]=await Promise.all([
      supabase.from("monthly_pl").select("*").eq("year",yr).order("month_index"),
      supabase.from("portfolio_data").select("*").eq("year",yr).order("created_at",{ascending:false}).limit(1).single(),
      supabase.from("properties").select("*").order("fmv",{ascending:false})
    ])
    if(plData)  setPL(plData)
    if(portData) setPortfolio(portData)
    if(propData) setProperties(propData)
  }

  useEffect(()=>{ if(session) loadData() },[session])

  if(session===undefined) return(
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",background:C.navy,color:C.textMuted,fontSize:14}}>Loading…</div>
  )

  if(!session) return <Login onDone={()=>{}}/>

  const tabs=["dashboard","upload",...(profile?.role==="admin"?["admin"]:[]) ]
  const labels={dashboard:"Dashboard",upload:"Data upload",admin:"Admin"}

  return(
    <div style={{background:C.navy,minHeight:"100vh",display:"flex",flexDirection:"column",fontFamily:"system-ui,sans-serif"}}>
      <div style={{background:C.navyMid,borderBottom:`0.5px solid ${C.border}`,display:"flex",alignItems:"center",justifyContent:"space-between",padding:"0 16px",height:46,flexShrink:0}}>
        <div style={{fontSize:14,fontWeight:500,color:C.text}}>Weigel Family Investment Portfolio</div>
        <div style={{display:"flex",gap:4}}>
          {tabs.map(t=>(
            <button key={t} onClick={()=>setTab(t)}
              style={{background:tab===t?"rgba(255,255,255,0.16)":"transparent",border:"none",color:tab===t?C.text:C.textMuted,fontSize:12,padding:"5px 12px",borderRadius:6,cursor:"pointer"}}>
              {labels[t]}
            </button>
          ))}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:12,color:C.textMuted}}>{profile?.name||session.user.email}</span>
          <button onClick={()=>supabase.auth.signOut()}
            style={{background:"transparent",border:`0.5px solid ${C.border}`,color:C.textMuted,fontSize:11,padding:"4px 10px",borderRadius:6,cursor:"pointer"}}>
            Sign out
          </button>
        </div>
      </div>
      <div style={{flex:1,padding:14,overflowY:"auto"}}>
        {tab==="dashboard" && <Dashboard pl={pl} portfolio={portfolio} properties={properties}/>}
        {tab==="upload"    && <Upload userId={session.user.id} onUploaded={loadData}/>}
        {tab==="admin"     && profile?.role==="admin" && <Admin/>}
      </div>
    </div>
  )
}

import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from "recharts";

import { useState, useEffect } from "react";
import * as XLSX from "xlsx";

// ── SUPABASE ──────────────────────────────────────────────────────────────────
const SUPA_URL = "https://rgsgjqnlvyossuykemjl.supabase.co";
const SUPA_KEY = "sb_publishable_uC6SF60PfuFnZfDDLRo2Wg_m7l2FoBW";

async function sbGet(table, params = "") {
  const res = await fetch(`${SUPA_URL}/rest/v1/${table}?${params}`, {
    headers: { "apikey": SUPA_KEY, "Authorization": `Bearer ${SUPA_KEY}` },
  });
  if (!res.ok) throw new Error(`GET ${table}: ${res.status}`);
  return res.json();
}

async function sbPost(table, body) {
  const res = await fetch(`${SUPA_URL}/rest/v1/${table}`, {
    method: "POST",
    headers: { "apikey": SUPA_KEY, "Authorization": `Bearer ${SUPA_KEY}`, "Content-Type": "application/json", "Prefer": "return=representation" },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`POST ${table}: ${res.status}`);
  return res.json();
}

async function sbPatch(table, id, body) {
  const res = await fetch(`${SUPA_URL}/rest/v1/${table}?id=eq.${id}`, {
    method: "PATCH",
    headers: { "apikey": SUPA_KEY, "Authorization": `Bearer ${SUPA_KEY}`, "Content-Type": "application/json", "Prefer": "return=representation" },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`PATCH ${table}: ${res.status}`);
  return res.json();
}

async function sbDelete(table, id) {
  const res = await fetch(`${SUPA_URL}/rest/v1/${table}?id=eq.${id}`, {
    method: "DELETE",
    headers: { "apikey": SUPA_KEY, "Authorization": `Bearer ${SUPA_KEY}` },
  });
  if (!res.ok) throw new Error(`DELETE ${table}: ${res.status}`);
  return true;
}

function Loading() {
  return <div style={{color:"#6b8f6b",fontSize:13,padding:"20px 0",textAlign:"center"}}>Carregando...</div>;
}

function ErrMsg({ msg }) {
  return <div style={{color:"#f87171",fontSize:12,padding:"10px 0"}}>Erro: {msg}</div>;
}

// ── EXPORTAÇÃO EXCEL ──────────────────────────────────────────────────────────
function exportXLSX(sheets, filename) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, data }) => {
    if (!data || data.length === 0) return;
    const ws = XLSX.utils.json_to_sheet(data);
    const cols = Object.keys(data[0]).map(k => ({ wch: Math.max(k.length + 2, 16) }));
    ws["!cols"] = cols;
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  XLSX.writeFile(wb, filename);
}

function BtnExport({ onClick }) {
  return (
    <button onClick={onClick} title="Exportar para Excel" style={{
      display:"flex", alignItems:"center", gap:5,
      background:"#0a1a0a", color:"#4ade80",
      border:"1px solid #4ade80", borderRadius:7,
      padding:"6px 13px", cursor:"pointer", fontSize:11, fontWeight:700, whiteSpace:"nowrap",
    }}>⬇ Excel</button>
  );
}


// ── TEMA ──────────────────────────────────────────────────────────────────────
const t = {
  bg:"#f8f8f4",    card:"#ffffff",   bord:"#e0e8d8",
  green:"#4a7c3f", greenLight:"#6aab5e", greenDark:"#2d5a24",
  gold:"#e8b84b",  goldDark:"#c99520",
  text:"#1a1a1a",  muted:"#5a6a5a",  dim:"#f0f4ec",
  red:"#dc2626",   orange:"#ea580c", yellow:"#d97706",
  blue:"#1d4ed8",  cyan:"#0891b2",   purple:"#7c3aed",
  lime:"#4d7c0f",  pink:"#be185d",
};

const ACCENT = "#0891b2";
const ACCENT_LIGHT = "#cffafe";
const ACCENT_DARK = "#164e63";

function AppHeader({ children }) {
  return (
    <div style={{ borderBottom:"1px solid #e0e8d8", background:"#ffffff", boxShadow:"0 1px 4px rgba(0,0,0,0.06)" }}>
      <div style={{ padding:"10px 20px", display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8 }}>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <div style={{ width:42, height:42, borderRadius:10, background:"linear-gradient(135deg,#4a7c3f,#6aab5e)", display:"flex", alignItems:"center", justifyContent:"center", boxShadow:"0 2px 8px rgba(74,124,63,0.3)" }}>
            <span style={{ color:"#fff", fontSize:20, fontFamily:"Georgia,serif", fontWeight:700 }}>D</span>
          </div>
          <div>
            <div style={{ fontSize:11, color:"#5a6a5a", letterSpacing:"0.08em", textTransform:"uppercase", fontWeight:600 }}>Dona Luzia</div>
            <div style={{ fontSize:15, color:ACCENT_DARK, fontWeight:700, fontFamily:"Georgia,serif", lineHeight:1.2 }}>Equipe</div>
          </div>
        </div>
        {children}
      </div>
      <div style={{ height:3, background:`linear-gradient(90deg,${ACCENT_DARK},${ACCENT},transparent)` }} />
    </div>
  );
}



const FONT = {
  title:  { fontFamily:"'Georgia',serif", fontWeight:700 },
  label:  { fontFamily:"system-ui,sans-serif", fontWeight:600, letterSpacing:"0.05em", textTransform:"uppercase", fontSize:11 },
  body:   { fontFamily:"system-ui,sans-serif", fontWeight:400 },
  mono:   { fontFamily:"'Courier New',monospace", fontWeight:700 },
  number: { fontFamily:"'Georgia',serif", fontWeight:700 },
};



const fmt  = v => `R$ ${Number(v).toLocaleString("pt-BR",{minimumFractionDigits:2})}`;
const pct  = (v,tot) => tot>0?((v/tot)*100).toFixed(1)+"%":"0%";
const inp  = {background:t.bg,border:`1px solid ${t.bord}`,borderRadius:8,padding:"9px 13px",color:t.text,fontSize:13,fontFamily:"inherit",outline:"none",width:"100%",boxSizing:"border-box"};
const lbl  = {color:t.muted,fontSize:11,letterSpacing:1,textTransform:"uppercase",display:"block",marginBottom:6};
const TT   = {contentStyle:{background:t.bg,border:`1px solid ${t.bord}`,borderRadius:8,color:t.text}};

// ── DADOS ─────────────────────────────────────────────────────────────────────
const COLABS_INIT = {
  loja: [
    {id:1,nome:"Camila",  cargo:"Atendente", salario:1412,vt:220,vr:198,plano:0,   admissao:"03/2023",cor:t.blue},
    {id:2,nome:"Rebeca",  cargo:"Atendente", salario:1412,vt:180,vr:198,plano:0,   admissao:"08/2024",cor:t.cyan},
    {id:3,nome:"Daniela", cargo:"Caixa",     salario:1600,vt:200,vr:198,plano:120, admissao:"01/2023",cor:t.pink},
  ],
  matriz: [
    {id:1,nome:"Fátima",  cargo:"Cozinheira",salario:1800,vt:240,vr:198,plano:0,   admissao:"05/2022",cor:t.lime},
    {id:2,nome:"Andressa",cargo:"Cozinheira",salario:1650,vt:200,vr:198,plano:0,   admissao:"11/2023",cor:t.orange},
  ],
};

const CONSUMOS_INIT = {
  loja: [
    {data:"21/04",colabId:1,produto:"Smoothie banana",custo:8.50},
    {data:"21/04",colabId:2,produto:"Bowl frutas",    custo:12.00},
    {data:"22/04",colabId:1,produto:"Suco laranja",   custo:6.00},
    {data:"22/04",colabId:3,produto:"Quiche",         custo:9.00},
    {data:"23/04",colabId:2,produto:"Smoothie verde", custo:9.50},
    {data:"24/04",colabId:1,produto:"Bowl frutas",    custo:12.00},
    {data:"25/04",colabId:3,produto:"Suco detox",     custo:7.50},
  ],
  matriz: [
    {data:"21/04",colabId:1,produto:"Marmita frango", custo:18.00},
    {data:"21/04",colabId:2,produto:"Marmita vegana", custo:15.00},
    {data:"22/04",colabId:1,produto:"Marmita carne",  custo:20.00},
    {data:"22/04",colabId:2,produto:"Marmita frango", custo:18.00},
    {data:"23/04",colabId:1,produto:"Marmita vegana", custo:15.00},
    {data:"24/04",colabId:2,produto:"Marmita carne",  custo:20.00},
    {data:"25/04",colabId:1,produto:"Marmita frango", custo:18.00},
  ],
};

// ── BASE COMPONENTS ───────────────────────────────────────────────────────────
function Kpi({label,value,sub,color=t.blue}) {
  return (
    <div style={{background:t.card,border:`1px solid ${t.bord}`,borderRadius:12,padding:"18px 20px"}}>
      <div style={{color:t.muted,fontSize:11,letterSpacing:1,textTransform:"uppercase",marginBottom:6}}>{label}</div>
      <div style={{color,fontSize:24,...FONT.number,lineHeight:1}}>{value}</div>
      {sub&&<div style={{color:t.muted,fontSize:12,marginTop:4}}>{sub}</div>}
    </div>
  );
}

function TabBtn({label,active,onClick}) {
  return (
    <button onClick={onClick} style={{
      background:active?ACCENT:"transparent",color:active?"#fff":t.muted,
      border:`1px solid ${active?ACCENT:t.bord}`,borderRadius:8,
      padding:"8px 20px",cursor:"pointer",fontSize:12,fontWeight:active?700:400,transition:"all 0.15s",
    }}>{label}</button>
  );
}

function Card({children,style={}}) {
  return <div style={{background:t.card,border:`1px solid ${t.bord}`,borderRadius:12,padding:20,...style}}>{children}</div>;
}

function ST({children,color=t.muted}) {
  return <div style={{color,fontSize:11,letterSpacing:1.5,textTransform:"uppercase",marginBottom:14,fontWeight:700,fontFamily:"system-ui,sans-serif"}}>{children}</div>;
}

function OpBtn({op,setOp}) {
  return (
    <div style={{display:"flex",gap:8}}>
      {[["loja","🏪 Loja"],["matriz","🏭 Matriz"]].map(([id,label])=>(
        <button key={id} onClick={()=>setOp(id)} style={{
          background:op===id?t.blue:t.card,color:op===id?t.bg:t.muted,
          border:`1px solid ${op===id?t.blue:t.bord}`,borderRadius:8,
          padding:"8px 20px",cursor:"pointer",fontSize:12,fontWeight:op===id?700:400,
        }}>{label}</button>
      ))}
    </div>
  );
}

// ── EQUIPE & FOLHA ────────────────────────────────────────────────────────────
function Equipe() {
  const [op,     setOp]     = useState("loja");
  const [colabs,setColabs]=useState(COLABS_INIT);
  useEffect(()=>{(async()=>{try{const rows=await sbGet("colaboradores","order=created_at.asc");if(rows.length>0){const grouped={loja:rows.filter(r=>r.operacao==="loja"),matriz:rows.filter(r=>r.operacao==="matriz")};setColabs(grouped);}}catch(e){console.error(e);}})();},[]);
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({nome:"",cargo:"",salario:"",vt:"",vr:"",plano:"",admissao:""});

  const lista = colabs[op];
  const custoTotal  = c => c.salario + c.vt + c.vr + c.plano;
  const totalFolha  = lista.reduce((s,c)=>s+custoTotal(c),0);
  const totalSal    = lista.reduce((s,c)=>s+c.salario,0);
  const totalBenef  = lista.reduce((s,c)=>s+c.vt+c.vr+c.plano,0);

  const graficoFolha = lista.map(c=>({
    nome: c.nome,
    Salário: c.salario,
    Benefícios: c.vt+c.vr+c.plano,
  }));

  const salvar=async()=>{if(!form.nome||!form.salario)return;try{const[novo]=await sbPost("colaboradores",{operacao:op,nome:form.nome,cargo:form.cargo,salario:parseFloat(form.salario)||0,vt:parseFloat(form.vt)||0,vr:parseFloat(form.vr)||0,plano:parseFloat(form.plano)||0,admissao:form.admissao,cor:t.cyan});setColabs(p=>({...p,[op]:[...p[op],novo]}));setForm({nome:"",cargo:"",salario:"",vt:"",vr:"",plano:"",admissao:""});setShowForm(false);}catch(e){alert("Erro: "+e.message);}};

  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
        <OpBtn op={op} setOp={setOp}/>
        <div style={{display:"flex",gap:8}}>
          <BtnExport onClick={()=>exportXLSX([{name:"Colaboradores",data:lista.map(c=>({Nome:c.nome,Cargo:c.cargo,Operação:op,"Salário (R$)":c.salario,"VT (R$)":c.vt,"VR (R$)":c.vr,"Plano (R$)":c.plano,"Total (R$)":c.salario+c.vt+c.vr+c.plano,Admissão:c.admissao}))}],"colaboradores.xlsx")} />
          <button onClick={()=>setShowForm(!showForm)} style={{background:showForm?t.card:t.green,color:showForm?t.muted:t.bg,border:`1px solid ${showForm?t.bord:t.green}`,borderRadius:8,padding:"8px 18px",cursor:"pointer",fontSize:12,fontWeight:700}}>
            {showForm?"✕ Cancelar":"+ Novo Colaborador"}
          </button>
        </div>
      </div>

      {/* KPIs */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12}}>
        <Kpi label="Colaboradores"  value={lista.length}      sub={op==="loja"?"Loja":"Matriz"}  color={t.blue}/>
        <Kpi label="Folha total"    value={fmt(totalFolha)}   sub="Salários + benefícios"        color={t.yellow}/>
        <Kpi label="Só salários"    value={fmt(totalSal)}     sub="Sem benefícios"               color={t.text}/>
        <Kpi label="Benefícios"     value={fmt(totalBenef)}   sub="VT + VR + Plano"             color={t.cyan}/>
      </div>

      {/* Form novo colaborador */}
      {showForm&&(
        <Card style={{border:`1px solid ${t.green}`}}>
          <ST color={t.green}>+ Novo Colaborador — {op==="loja"?"🏪 Loja":"🏭 Matriz"}</ST>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:12}}>
            {[["Nome","nome","text"],["Cargo","cargo","text"],["Salário (R$)","salario","number"],["VT mensal","vt","number"],["VR mensal","vr","number"],["Plano saúde","plano","number"]].map(([label,key,type])=>(
              <div key={key}>
                <label style={lbl}>{label}</label>
                <input type={type} value={form[key]} onChange={e=>setForm(f=>({...f,[key]:e.target.value}))} style={inp}/>
              </div>
            ))}
          </div>
          <div style={{marginBottom:14,maxWidth:200}}>
            <label style={lbl}>Admissão (mm/aaaa)</label>
            <input value={form.admissao} onChange={e=>setForm(f=>({...f,admissao:e.target.value}))} style={inp} placeholder="04/2025"/>
          </div>
          {form.salario&&(
            <div style={{marginBottom:14,background:t.bg,borderRadius:8,padding:"9px 12px",color:t.muted,fontSize:12}}>
              Custo total/mês: <span style={{color:t.yellow,fontFamily:"monospace"}}>{fmt((parseFloat(form.salario)||0)+(parseFloat(form.vt)||0)+(parseFloat(form.vr)||0)+(parseFloat(form.plano)||0))}</span>
            </div>
          )}
          <div style={{display:"flex",gap:10}}>
            <button onClick={salvar} style={{background:t.green,color:t.bg,border:"none",borderRadius:8,padding:"10px 22px",cursor:"pointer",fontSize:12,fontWeight:700}}>Salvar</button>
            <button onClick={()=>setShowForm(false)} style={{background:"transparent",color:t.muted,border:`1px solid ${t.bord}`,borderRadius:8,padding:"10px 18px",cursor:"pointer",fontSize:12}}>Cancelar</button>
          </div>
        </Card>
      )}

      {/* Cards colaboradores */}
      <div style={{display:"flex",flexDirection:"column",gap:14}}>
        {lista.map(c=>{
          const total = custoTotal(c);
          return (
            <Card key={c.id}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16}}>
                <div style={{display:"flex",gap:14,alignItems:"center"}}>
                  <div style={{width:46,height:46,borderRadius:"50%",background:t.bord,display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,border:`2px solid ${c.cor}`}}>👤</div>
                  <div>
                    <div style={{color:t.text,fontSize:15,fontWeight:600}}>{c.nome}</div>
                    <div style={{color:t.muted,fontSize:12,marginTop:2}}>{c.cargo} · Desde {c.admissao}</div>
                  </div>
                </div>
                <div style={{textAlign:"right"}}>
                  <div style={{color:c.cor,fontFamily:"monospace",fontSize:16,fontWeight:700}}>{fmt(total)}</div>
                  <div style={{color:t.muted,fontSize:11}}>custo total/mês</div>
                </div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10}}>
                {[["Salário",c.salario,t.text],["VT",c.vt,t.muted],["VR",c.vr,t.muted],["Plano",c.plano,t.muted]].map(([k,v,cor])=>(
                  <div key={k} style={{background:t.bg,borderRadius:8,padding:"10px 12px",textAlign:"center"}}>
                    <div style={{color:t.muted,fontSize:11,marginBottom:4}}>{k}</div>
                    <div style={{color:cor,fontFamily:"monospace",fontSize:13}}>{fmt(v)}</div>
                  </div>
                ))}
              </div>
            </Card>
          );
        })}
      </div>

      {/* Gráfico comparativo folha */}
      <Card>
        <ST>Custo por Colaborador — {op==="loja"?"Loja":"Matriz"}</ST>
        <ResponsiveContainer width="100%" height={200}>
          <BarChart data={graficoFolha} barCategoryGap="30%">
            <XAxis dataKey="nome" stroke={t.muted} tick={{fill:t.muted,fontSize:11}}/>
            <YAxis stroke={t.muted} tick={{fill:t.muted,fontSize:11}} tickFormatter={v=>`R$${(v/1000).toFixed(1)}k`}/>
            <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle}/>
            <Bar dataKey="Salário"    fill={t.blue}  radius={[4,4,0,0]}/>
            <Bar dataKey="Benefícios" fill={t.cyan}  radius={[4,4,0,0]}/>
          </BarChart>
        </ResponsiveContainer>
        <div style={{display:"flex",gap:16,marginTop:8}}>
          {[["Salário",t.blue],["Benefícios",t.cyan]].map(([l,c])=>(
            <div key={l} style={{display:"flex",alignItems:"center",gap:6}}>
              <div style={{width:10,height:10,borderRadius:2,background:c}}/>
              <span style={{color:t.muted,fontSize:12}}>{l}</span>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );
}

// ── CONSUMO ───────────────────────────────────────────────────────────────────
function Consumo() {
  const [op,       setOp]       = useState("loja");
  const [colabs,]               = useState(COLABS_INIT);
  const [consumos,setConsumos]=useState(CONSUMOS_INIT);
  useEffect(()=>{(async()=>{try{const rows=await sbGet("consumo_colaboradores","order=data.desc");if(rows.length>0){const grouped={loja:rows.filter(r=>r.operacao==="loja").map(r=>({...r,colabId:r.colaborador_id,data:r.data?new Date(r.data).toLocaleDateString("pt-BR"):""})),matriz:rows.filter(r=>r.operacao==="matriz").map(r=>({...r,colabId:r.colaborador_id,data:r.data?new Date(r.data).toLocaleDateString("pt-BR"):""}))}; setConsumos(grouped);}}catch(e){console.error(e);}})();},[]);
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({data:"27/04",colabId:"",produto:"",custo:""});

  const lista      = consumos[op];
  const listaColab = colabs[op];
  const totalGeral = lista.reduce((s,x)=>s+x.custo,0);

  const consumoPorColab = id => lista.filter(x=>x.colabId===id).reduce((s,x)=>s+x.custo,0);

  const grafico = listaColab.map(c=>({
    nome: c.nome,
    consumo: consumoPorColab(c.id),
  }));

  const salvar=async()=>{if(!form.colabId||!form.produto||!form.custo)return;try{const[novo]=await sbPost("consumo_colaboradores",{operacao:op,colaborador_id:form.colabId,data:new Date().toISOString().split("T")[0],produto:form.produto,custo:parseFloat(form.custo)});setConsumos(p=>({...p,[op]:[...p[op],{...novo,colabId:novo.colaborador_id,data:new Date().toLocaleDateString("pt-BR")}]}));setForm(f=>({...f,produto:"",custo:""}));setShowForm(false);}catch(e){alert("Erro: "+e.message);}};

  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
        <OpBtn op={op} setOp={setOp}/>
        <button onClick={()=>setShowForm(!showForm)} style={{background:showForm?t.card:t.yellow,color:showForm?t.muted:t.bg,border:`1px solid ${showForm?t.bord:t.yellow}`,borderRadius:8,padding:"8px 18px",cursor:"pointer",fontSize:12,fontWeight:700}}>
          {showForm?"✕ Cancelar":"+ Registrar Consumo"}
        </button>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
        <Kpi label="Total consumido" value={fmt(totalGeral)}  sub="Custo do mês"        color={t.orange}/>
        <Kpi label="Registros"       value={lista.length}     sub="Lançamentos"         color={t.blue}/>
        <Kpi label="Média/colaborador" value={fmt(listaColab.length>0?totalGeral/listaColab.length:0)} sub="Por pessoa/mês" color={t.cyan}/>
      </div>

      {showForm&&(
        <Card style={{border:`1px solid ${t.yellow}`}}>
          <ST color={t.yellow}>🥤 Registrar Consumo — {op==="loja"?"🏪 Loja":"🏭 Matriz"}</ST>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:10}}>
            <div>
              <label style={lbl}>Data</label>
              <input value={form.data} onChange={e=>setForm(f=>({...f,data:e.target.value}))} style={inp}/>
            </div>
            <div>
              <label style={lbl}>Colaborador</label>
              <select value={form.colabId} onChange={e=>setForm(f=>({...f,colabId:e.target.value}))} style={inp}>
                <option value="">Selecionar...</option>
                {listaColab.map(c=><option key={c.id} value={c.id}>{c.nome}</option>)}
              </select>
            </div>
            <div>
              <label style={lbl}>Produto</label>
              <input value={form.produto} onChange={e=>setForm(f=>({...f,produto:e.target.value}))} style={inp} placeholder="Ex: Marmita frango"/>
            </div>
            <div>
              <label style={lbl}>Custo (R$)</label>
              <input type="number" value={form.custo} onChange={e=>setForm(f=>({...f,custo:e.target.value}))} style={inp} placeholder="18.00"/>
            </div>
          </div>
          <div style={{display:"flex",gap:10,marginTop:14}}>
            <button onClick={salvar} style={{background:t.yellow,color:t.bg,border:"none",borderRadius:8,padding:"10px 22px",cursor:"pointer",fontSize:12,fontWeight:700}}>Confirmar</button>
            <button onClick={()=>setShowForm(false)} style={{background:"transparent",color:t.muted,border:`1px solid ${t.bord}`,borderRadius:8,padding:"10px 18px",cursor:"pointer",fontSize:12}}>Cancelar</button>
          </div>
        </Card>
      )}

      {/* Gráfico consumo por colaborador */}
      <Card>
        <ST>Consumo por Colaborador</ST>
        <ResponsiveContainer width="100%" height={180}>
          <BarChart data={grafico} barCategoryGap="40%">
            <XAxis dataKey="nome" stroke={t.muted} tick={{fill:t.muted,fontSize:11}}/>
            <YAxis stroke={t.muted} tick={{fill:t.muted,fontSize:11}} tickFormatter={v=>`R$${v.toFixed(0)}`}/>
            <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle}/>
            <Bar dataKey="consumo" fill={t.orange} radius={[4,4,0,0]} name="Consumo"/>
          </BarChart>
        </ResponsiveContainer>
      </Card>

      {/* Cards por colaborador */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        {listaColab.map(c=>{
          const cons   = lista.filter(x=>x.colabId===c.id);
          const total  = cons.reduce((s,x)=>s+x.custo,0);
          return (
            <Card key={c.id}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}>
                <div style={{display:"flex",gap:10,alignItems:"center"}}>
                  <div style={{width:36,height:36,borderRadius:"50%",background:t.bord,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,border:`2px solid ${c.cor}`}}>👤</div>
                  <div>
                    <div style={{color:t.text,fontSize:14,fontWeight:600}}>{c.nome}</div>
                    <div style={{color:t.muted,fontSize:11}}>{c.cargo}</div>
                  </div>
                </div>
                <div style={{textAlign:"right"}}>
                  <div style={{color:t.orange,fontFamily:"monospace",fontSize:15,fontWeight:700}}>{fmt(total)}</div>
                  <div style={{color:t.muted,fontSize:11}}>consumo/mês</div>
                </div>
              </div>
              {cons.length===0
                ? <div style={{color:t.muted,fontSize:12}}>Nenhum consumo registrado.</div>
                : cons.slice(-4).reverse().map((x,i)=>(
                  <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:`1px solid ${t.bord}`}}>
                    <span style={{color:t.muted,fontSize:12}}>{x.data} · {x.produto}</span>
                    <span style={{color:t.text,fontFamily:"monospace",fontSize:12}}>{fmt(x.custo)}</span>
                  </div>
                ))
              }
            </Card>
          );
        })}
      </div>

      {/* Histórico completo */}
      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <ST>Histórico Completo — {op==="loja"?"Loja":"Matriz"}</ST>
          <BtnExport onClick={()=>exportXLSX([{name:"Consumo",data:[...lista].reverse().map(x=>{const c=listaColab.find(c=>c.id===x.colabId);return{Data:x.data,Colaborador:c?.nome||"—",Operação:op,Produto:x.produto,"Custo (R$)":x.custo};})}],"consumo-colaboradores.xlsx")} />
        </div>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead>
            <tr style={{background:t.bg}}>
              {["Data","Colaborador","Produto","Custo"].map(h=>(
                <th key={h} style={{padding:"9px 12px",textAlign:"left",color:t.muted,fontSize:11,letterSpacing:1,textTransform:"uppercase"}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {[...lista].reverse().map((x,i)=>{
              const c = listaColab.find(c=>c.id===x.colabId);
              return (
                <tr key={i} style={{borderTop:`1px solid ${t.bord}`}}>
                  <td style={{padding:"9px 12px",color:t.muted,fontFamily:"monospace",fontSize:12}}>{x.data}</td>
                  <td style={{padding:"9px 12px"}}>
                    <span style={{color:c?.cor||t.text,fontSize:13}}>{c?.nome||"—"}</span>
                  </td>
                  <td style={{padding:"9px 12px",color:t.text,fontSize:13}}>{x.produto}</td>
                  <td style={{padding:"9px 12px",color:t.orange,fontFamily:"monospace",fontSize:13,fontWeight:700}}>{fmt(x.custo)}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

// ── BANCO DE HORAS ────────────────────────────────────────────────────────────
const BANCO_INIT = {
  loja: [
    // colabId, data, tipo: "extra"|"folga"|"atraso", horas, obs
    {colabId:1,data:"21/04",tipo:"extra",  horas:1.5, obs:"Fechamento de caixa"},
    {colabId:2,data:"21/04",tipo:"extra",  horas:1.0, obs:"Movimento alto"},
    {colabId:1,data:"22/04",tipo:"atraso", horas:-0.5,obs:"Trânsito"},
    {colabId:3,data:"22/04",tipo:"extra",  horas:2.0, obs:"Cobertura colega"},
    {colabId:2,data:"23/04",tipo:"folga",  horas:-8.0,obs:"Folga compensatória"},
    {colabId:1,data:"24/04",tipo:"extra",  horas:1.0, obs:"Evento especial"},
    {colabId:3,data:"25/04",tipo:"atraso", horas:-0.5,obs:"Consulta médica"},
    {colabId:1,data:"25/04",tipo:"extra",  horas:2.0, obs:"Sábado reforço"},
  ],
  matriz: [
    {colabId:1,data:"21/04",tipo:"extra",  horas:2.0, obs:"Produção emergencial"},
    {colabId:2,data:"21/04",tipo:"extra",  horas:1.5, obs:"Lote extra marmitas"},
    {colabId:1,data:"22/04",tipo:"extra",  horas:1.0, obs:"Entrega atrasada"},
    {colabId:2,data:"23/04",tipo:"atraso", horas:-1.0,obs:"Problema pessoal"},
    {colabId:1,data:"24/04",tipo:"folga",  horas:-8.0,obs:"Folga compensatória"},
    {colabId:2,data:"25/04",tipo:"extra",  horas:3.0, obs:"Sábado produção"},
  ],
};

function BancoHoras() {
  const [op,    setOp]    = useState("loja");
  const [banco, setBanco] = useState(BANCO_INIT);
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({colabId:"",data:"27/04",tipo:"extra",horas:"",obs:""});
  const [filtroColab, setFiltroColab] = useState("Todos");

  const listaColab = COLABS_INIT[op];
  const registros  = banco[op];

  const saldoColab = id => registros.filter(r=>r.colabId===id).reduce((s,r)=>s+r.horas,0);

  const filtrados = filtroColab==="Todos"
    ? registros
    : registros.filter(r=>r.colabId===parseInt(filtroColab));

  const totalExtra  = registros.filter(r=>r.tipo==="extra").reduce((s,r)=>s+r.horas,0);
  const totalFolga  = Math.abs(registros.filter(r=>r.tipo==="folga").reduce((s,r)=>s+r.horas,0));
  const totalAtraso = Math.abs(registros.filter(r=>r.tipo==="atraso").reduce((s,r)=>s+r.horas,0));

  const fmtH = h => {
    const abs = Math.abs(h);
    const int = Math.floor(abs);
    const min = Math.round((abs-int)*60);
    const str = min>0?`${int}h${String(min).padStart(2,"0")}m`:`${int}h`;
    return h<0?`-${str}`:str;
  };

  const corTipo = tipo =>
    tipo==="extra"  ? [t.green,  "#0f2a0f"] :
    tipo==="folga"  ? [t.yellow, "#3b2f00"] :
                     [t.red,    "#3b0f0f"];

  const salvar=async()=>{if(!form.colabId||!form.horas)return;const horas=form.tipo==="extra"?Math.abs(parseFloat(form.horas)):-Math.abs(parseFloat(form.horas));try{const[novo]=await sbPost("banco_horas",{operacao:op,colaborador_id:form.colabId,data:new Date().toISOString().split("T")[0],tipo:form.tipo,horas,obs:form.obs});setBanco(p=>({...p,[op]:[...p[op],{...novo,colabId:novo.colaborador_id,data:new Date().toLocaleDateString("pt-BR"),horas}]}));setForm(f=>({...f,horas:"",obs:""}));setShowForm(false);}catch(e){alert("Erro: "+e.message);}};

  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
        <OpBtn op={op} setOp={id=>{setOp(id);setFiltroColab("Todos");}}/>
        <button onClick={()=>setShowForm(!showForm)} style={{background:showForm?t.card:t.cyan,color:showForm?t.muted:t.bg,border:`1px solid ${showForm?t.bord:t.cyan}`,borderRadius:8,padding:"8px 18px",cursor:"pointer",fontSize:12,fontWeight:700}}>
          {showForm?"✕ Cancelar":"+ Lançar Horas"}
        </button>
      </div>

      {/* KPIs */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12}}>
        <Kpi label="Horas extras"  value={fmtH(totalExtra)}  sub="Acumuladas no mês"   color={t.green}/>
        <Kpi label="Folgas usadas" value={fmtH(totalFolga)}  sub="Compensadas"         color={t.yellow}/>
        <Kpi label="Atrasos"       value={fmtH(totalAtraso)} sub="Total registrado"    color={t.red}/>
        <Kpi label="Colaboradores" value={listaColab.length} sub={op==="loja"?"Loja":"Matriz"} color={t.blue}/>
      </div>

      {/* Form lançamento */}
      {showForm&&(
        <Card style={{border:`1px solid ${t.cyan}`}}>
          <ST color={t.cyan}>+ Lançar Horas — {op==="loja"?"🏪 Loja":"🏭 Matriz"}</ST>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr 2fr",gap:10}}>
            <div>
              <label style={lbl}>Colaborador</label>
              <select value={form.colabId} onChange={e=>setForm(f=>({...f,colabId:e.target.value}))} style={inp}>
                <option value="">Selecionar...</option>
                {listaColab.map(c=><option key={c.id} value={c.id}>{c.nome}</option>)}
              </select>
            </div>
            <div>
              <label style={lbl}>Data</label>
              <input value={form.data} onChange={e=>setForm(f=>({...f,data:e.target.value}))} style={inp}/>
            </div>
            <div>
              <label style={lbl}>Tipo</label>
              <select value={form.tipo} onChange={e=>setForm(f=>({...f,tipo:e.target.value}))} style={inp}>
                <option value="extra">Hora extra</option>
                <option value="folga">Folga compensatória</option>
                <option value="atraso">Atraso</option>
              </select>
            </div>
            <div>
              <label style={lbl}>Horas</label>
              <input type="number" step="0.5" value={form.horas} onChange={e=>setForm(f=>({...f,horas:e.target.value}))} style={inp} placeholder="Ex: 1.5"/>
            </div>
            <div>
              <label style={lbl}>Observação</label>
              <input value={form.obs} onChange={e=>setForm(f=>({...f,obs:e.target.value}))} style={inp} placeholder="Ex: Cobertura colega"/>
            </div>
          </div>
          <div style={{display:"flex",gap:10,marginTop:14}}>
            <button onClick={salvar} style={{background:t.cyan,color:t.bg,border:"none",borderRadius:8,padding:"10px 22px",cursor:"pointer",fontSize:12,fontWeight:700}}>Confirmar</button>
            <button onClick={()=>setShowForm(false)} style={{background:"transparent",color:t.muted,border:`1px solid ${t.bord}`,borderRadius:8,padding:"10px 18px",cursor:"pointer",fontSize:12}}>Cancelar</button>
          </div>
        </Card>
      )}

      {/* Saldo por colaborador */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        {listaColab.map(c=>{
          const saldo   = saldoColab(c.id);
          const extras  = registros.filter(r=>r.colabId===c.id&&r.tipo==="extra").reduce((s,r)=>s+r.horas,0);
          const folgas  = Math.abs(registros.filter(r=>r.colabId===c.id&&r.tipo==="folga").reduce((s,r)=>s+r.horas,0));
          const atrasos = Math.abs(registros.filter(r=>r.colabId===c.id&&r.tipo==="atraso").reduce((s,r)=>s+r.horas,0));
          return (
            <Card key={c.id} style={{border:`1px solid ${saldo>=0?t.bord:t.red}`}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
                <div style={{display:"flex",gap:10,alignItems:"center"}}>
                  <div style={{width:38,height:38,borderRadius:"50%",background:t.bord,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,border:`2px solid ${c.cor}`}}>👤</div>
                  <div>
                    <div style={{color:t.text,fontSize:14,fontWeight:600}}>{c.nome}</div>
                    <div style={{color:t.muted,fontSize:11}}>{c.cargo}</div>
                  </div>
                </div>
                <div style={{textAlign:"right"}}>
                  <div style={{color:saldo>=0?t.green:t.red,fontFamily:"monospace",fontSize:18,fontWeight:700}}>{fmtH(saldo)}</div>
                  <div style={{color:t.muted,fontSize:11}}>saldo atual</div>
                </div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
                {[["Extras",fmtH(extras),t.green],["Folgas",fmtH(folgas),t.yellow],["Atrasos",fmtH(atrasos),t.red]].map(([k,v,cor])=>(
                  <div key={k} style={{background:t.bg,borderRadius:8,padding:"8px 10px",textAlign:"center"}}>
                    <div style={{color:t.muted,fontSize:11,marginBottom:3}}>{k}</div>
                    <div style={{color:cor,fontFamily:"monospace",fontSize:13,fontWeight:600}}>{v}</div>
                  </div>
                ))}
              </div>
              {/* Barra de saldo */}
              <div style={{marginTop:12}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                  <span style={{color:t.muted,fontSize:11}}>Extras usadas em folgas</span>
                  <span style={{color:t.muted,fontSize:11}}>{extras>0?((folgas/extras)*100).toFixed(0):0}%</span>
                </div>
                <div style={{background:t.bg,borderRadius:4,height:5,overflow:"hidden"}}>
                  <div style={{width:extras>0?`${Math.min((folgas/extras)*100,100)}%`:"0%",background:t.yellow,height:"100%",borderRadius:4}}/>
                </div>
              </div>
            </Card>
          );
        })}
      </div>

      {/* Filtro + tabela histórico */}
      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
          <ST>Histórico de Lançamentos</ST>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            <BtnExport onClick={()=>exportXLSX([{name:"Banco de Horas",data:[...filtrados].reverse().map(r=>{const c=listaColab.find(x=>x.id===r.colabId);return{Data:r.data,Colaborador:c?.nome||"—",Operação:op,Tipo:r.tipo==="extra"?"Hora extra":r.tipo==="folga"?"Folga":"Atraso",Horas:r.horas,Obs:r.obs||""};})}],"banco-horas.xlsx")} />
            <select value={filtroColab} onChange={e=>setFiltroColab(e.target.value)} style={{...inp,width:"auto",fontSize:12,padding:"6px 12px"}}>
            <option value="Todos">Todos</option>
            {listaColab.map(c=><option key={c.id} value={c.id}>{c.nome}</option>)}
          </select>
          </div>
        </div>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead>
            <tr style={{background:t.bg}}>
              {["Data","Colaborador","Tipo","Horas","Observação"].map(h=>(
                <th key={h} style={{padding:"9px 12px",textAlign:"left",color:t.muted,fontSize:11,letterSpacing:1,textTransform:"uppercase"}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {[...filtrados].reverse().map((r,i)=>{
              const c = listaColab.find(x=>x.id===r.colabId);
              const [cor,bg] = corTipo(r.tipo);
              return (
                <tr key={i} style={{borderTop:`1px solid ${t.bord}`}}>
                  <td style={{padding:"10px 12px",color:t.muted,fontFamily:"monospace",fontSize:12}}>{r.data}</td>
                  <td style={{padding:"10px 12px"}}>
                    <span style={{color:c?.cor||t.text,fontSize:13}}>{c?.nome||"—"}</span>
                  </td>
                  <td style={{padding:"10px 12px"}}>
                    <span style={{background:bg,color:cor,borderRadius:6,padding:"3px 10px",fontSize:11,fontFamily:"monospace"}}>
                      {r.tipo==="extra"?"Hora extra":r.tipo==="folga"?"Folga":"Atraso"}
                    </span>
                  </td>
                  <td style={{padding:"10px 12px",color:cor,fontFamily:"monospace",fontSize:13,fontWeight:700}}>{fmtH(r.horas)}</td>
                  <td style={{padding:"10px 12px",color:t.muted,fontSize:12}}>{r.obs||"—"}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

// ── RESUMO GERAL ──────────────────────────────────────────────────────────────
function ResumoGeral() {
  const totalFolhaLoja   = COLABS_INIT.loja.reduce((s,c)=>s+c.salario+c.vt+c.vr+c.plano,0);
  const totalFolhaMatriz = COLABS_INIT.matriz.reduce((s,c)=>s+c.salario+c.vt+c.vr+c.plano,0);
  const totalConsLoja    = CONSUMOS_INIT.loja.reduce((s,x)=>s+x.custo,0);
  const totalConsMatriz  = CONSUMOS_INIT.matriz.reduce((s,x)=>s+x.custo,0);
  const totalColabs      = COLABS_INIT.loja.length+COLABS_INIT.matriz.length;
  const totalFolha       = totalFolhaLoja+totalFolhaMatriz;
  const totalConsumo     = totalConsLoja+totalConsMatriz;

  const pieData = [
    {name:"Folha Loja",   value:totalFolhaLoja,   cor:t.blue},
    {name:"Folha Matriz", value:totalFolhaMatriz, cor:t.cyan},
    {name:"Consumo Loja", value:totalConsLoja,    cor:t.orange},
    {name:"Cons. Matriz", value:totalConsMatriz,  cor:t.yellow},
  ];

  const barData = [
    {op:"🏪 Loja",   folha:totalFolhaLoja,   consumo:totalConsLoja},
    {op:"🏭 Matriz", folha:totalFolhaMatriz, consumo:totalConsMatriz},
  ];

  const allColabs = [
    ...COLABS_INIT.loja.map(c=>({...c,op:"Loja"})),
    ...COLABS_INIT.matriz.map(c=>({...c,op:"Matriz"})),
  ];

  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12}}>
        <Kpi label="Total colaboradores" value={totalColabs}       sub="Loja + Matriz"     color={t.blue}/>
        <Kpi label="Folha total"         value={fmt(totalFolha)}   sub="Ambas operações"   color={t.yellow}/>
        <Kpi label="Consumo total"       value={fmt(totalConsumo)} sub="Produtos/mês"      color={t.orange}/>
        <Kpi label="Custo/colaborador"   value={fmt(totalFolha/totalColabs)} sub="Média geral" color={t.cyan}/>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <Card>
          <ST>Distribuição de Custos</ST>
          <ResponsiveContainer width="100%" height={170}>
            <PieChart>
              <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={70} innerRadius={36}>
                {pieData.map((d,i)=><Cell key={i} fill={d.cor}/>)}
              </Pie>
              <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle}/>
            </PieChart>
          </ResponsiveContainer>
          {pieData.map(d=>(
            <div key={d.name} style={{display:"flex",justifyContent:"space-between",padding:"4px 0"}}>
              <div style={{display:"flex",alignItems:"center",gap:7}}>
                <div style={{width:7,height:7,borderRadius:"50%",background:d.cor}}/>
                <span style={{color:t.muted,fontSize:12}}>{d.name}</span>
              </div>
              <span style={{color:t.text,fontFamily:"monospace",fontSize:12}}>{fmt(d.value)}</span>
            </div>
          ))}
        </Card>

        <Card>
          <ST>Loja vs Matriz</ST>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={barData} barCategoryGap="35%">
              <XAxis dataKey="op" stroke={t.muted} tick={{fill:t.muted,fontSize:11}}/>
              <YAxis stroke={t.muted} tick={{fill:t.muted,fontSize:11}} tickFormatter={v=>`R$${(v/1000).toFixed(1)}k`}/>
              <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle}/>
              <Bar dataKey="folha"   fill={t.blue}   radius={[4,4,0,0]} name="Folha"/>
              <Bar dataKey="consumo" fill={t.orange} radius={[4,4,0,0]} name="Consumo"/>
            </BarChart>
          </ResponsiveContainer>
          <div style={{display:"flex",gap:16,marginTop:8}}>
            {[["Folha",t.blue],["Consumo",t.orange]].map(([l,c])=>(
              <div key={l} style={{display:"flex",alignItems:"center",gap:6}}>
                <div style={{width:10,height:10,borderRadius:2,background:c}}/>
                <span style={{color:t.muted,fontSize:12}}>{l}</span>
              </div>
            ))}
          </div>
        </Card>
      </div>

      {/* Tabela geral todos os colaboradores */}
      <div style={{display:"flex",justifyContent:"flex-end"}}>
        <BtnExport onClick={()=>exportXLSX([{name:"Todos Colaboradores",data:allColabs.map(c=>({Nome:c.nome,Cargo:c.cargo,Operação:c.op,"Salário (R$)":c.salario,"VT (R$)":c.vt,"VR (R$)":c.vr,"Plano (R$)":c.plano,"Total (R$)":c.salario+c.vt+c.vr+c.plano,Admissão:c.admissao}))},{name:"Resumo Folha",data:[{Operação:"Loja","Folha (R$)":totalFolhaLoja,"Consumo (R$)":totalConsLoja},{Operação:"Matriz","Folha (R$)":totalFolhaMatriz,"Consumo (R$)":totalConsMatriz}]}],"resumo-equipe.xlsx")} />
      </div>
      <Card>
        <ST>Todos os Colaboradores</ST>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead>
            <tr style={{background:t.bg}}>
              {["Nome","Cargo","Operação","Salário","VT","VR","Plano","Total","Admissão"].map(h=>(
                <th key={h} style={{padding:"9px 12px",textAlign:"left",color:t.muted,fontSize:11,letterSpacing:1,textTransform:"uppercase",whiteSpace:"nowrap"}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {allColabs.map((c,i)=>{
              const total = c.salario+c.vt+c.vr+c.plano;
              return (
                <tr key={i} style={{borderTop:`1px solid ${t.bord}`}}>
                  <td style={{padding:"10px 12px"}}>
                    <span style={{color:c.cor,fontSize:13,fontWeight:600}}>{c.nome}</span>
                  </td>
                  <td style={{padding:"10px 12px",color:t.muted,fontSize:12}}>{c.cargo}</td>
                  <td style={{padding:"10px 12px"}}>
                    <span style={{background:c.op==="Loja"?"#0f1a30":"#0f2a1a",color:c.op==="Loja"?t.blue:t.green,borderRadius:6,padding:"3px 9px",fontSize:11,fontFamily:"monospace"}}>{c.op}</span>
                  </td>
                  <td style={{padding:"10px 12px",color:t.text,fontFamily:"monospace",fontSize:13}}>{fmt(c.salario)}</td>
                  <td style={{padding:"10px 12px",color:t.muted,fontFamily:"monospace",fontSize:12}}>{fmt(c.vt)}</td>
                  <td style={{padding:"10px 12px",color:t.muted,fontFamily:"monospace",fontSize:12}}>{fmt(c.vr)}</td>
                  <td style={{padding:"10px 12px",color:t.muted,fontFamily:"monospace",fontSize:12}}>{fmt(c.plano)}</td>
                  <td style={{padding:"10px 12px",color:t.yellow,fontFamily:"monospace",fontSize:13,fontWeight:700}}>{fmt(total)}</td>
                  <td style={{padding:"10px 12px",color:t.muted,fontFamily:"monospace",fontSize:12}}>{c.admissao}</td>
                </tr>
              );
            })}
            <tr style={{borderTop:`2px solid ${t.bord}`,background:t.bg}}>
              <td colSpan={7} style={{padding:"10px 12px",color:t.muted,fontSize:12,fontFamily:"monospace"}}>TOTAL GERAL</td>
              <td style={{padding:"10px 12px",color:t.yellow,fontFamily:"monospace",fontSize:14,fontWeight:700}}>{fmt(totalFolha)}</td>
              <td></td>
            </tr>
          </tbody>
        </table>
      </Card>
    </div>
  );
}

// ── APP ───────────────────────────────────────────────────────────────────────
const TABS = [
  {id:"resumo",  label:"📋 Resumo Geral"},
  {id:"equipe",  label:"👩‍🍳 Equipe & Folha"},
  {id:"consumo", label:"🥤 Consumo Diário"},
  {id:"banco",   label:"⏱ Banco de Horas"},
];

export default function App() {
  const [tab, setTab] = useState("resumo");
  return (
    <div style={{minHeight:"100vh",background:t.bg,color:t.text,fontFamily:"system-ui,sans-serif"}}>
      <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600&family=DM+Sans:wght@400;500;600&display=swap" rel="stylesheet"/>

      <AppHeader>
        <div style={{ display:"flex", alignItems:"center", gap:6, background:ACCENT_LIGHT, border:`1px solid ${ACCENT}44`, borderRadius:8, padding:"5px 10px" }}>
          <div style={{ width:6, height:6, borderRadius:"50%", background:ACCENT }} />
          <span style={{ color:ACCENT_DARK, fontSize:11, fontWeight:600 }}>Dados de exemplo</span>
        </div>
      </AppHeader>

            <div style={{padding:"16px 28px 0",overflowX:"auto"}}>
        <div style={{display:"flex",gap:8,minWidth:"max-content"}}>
          {TABS.map(tb=><TabBtn key={tb.id} label={tb.label} active={tab===tb.id} onClick={()=>setTab(tb.id)}/>)}
        </div>
      </div>

      <div style={{padding:"22px 28px"}}>
        {tab==="resumo"  && <ResumoGeral/>}
        {tab==="equipe"  && <Equipe/>}
        {tab==="consumo" && <Consumo/>}
        {tab==="banco"   && <BancoHoras/>}
      </div>
    </div>
  );
}

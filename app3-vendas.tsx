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

import { BarChart, Bar, LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from "recharts";

const t = {
  bg:"#f8f8f4",    card:"#ffffff",   bord:"#e0e8d8",
  green:"#4a7c3f", greenLight:"#6aab5e", greenDark:"#2d5a24",
  gold:"#e8b84b",  goldDark:"#c99520",
  text:"#1a1a1a",  muted:"#5a6a5a",  dim:"#f0f4ec",
  red:"#dc2626",   orange:"#ea580c", yellow:"#d97706",
  blue:"#1d4ed8",  cyan:"#0891b2",   purple:"#7c3aed",
  lime:"#4d7c0f",  pink:"#be185d",
};

const ACCENT = "#7c3aed";
const ACCENT_LIGHT = "#ede9fe";
const ACCENT_DARK = "#4c1d95";

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
            <div style={{ fontSize:15, color:ACCENT_DARK, fontWeight:700, fontFamily:"Georgia,serif", lineHeight:1.2 }}>Vendas</div>
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
const fmtK = v => `R$ ${Number(v).toLocaleString("pt-BR",{minimumFractionDigits:0})}`;
const pct  = (v,tot) => tot>0?((v/tot)*100).toFixed(1)+"%":"0%";
const inp  = {background:t.bg,border:`1px solid ${t.bord}`,borderRadius:8,padding:"9px 13px",color:t.text,fontSize:13,fontFamily:"inherit",outline:"none",width:"100%",boxSizing:"border-box"};
const lbl  = {color:t.muted,fontSize:11,letterSpacing:1,textTransform:"uppercase",display:"block",marginBottom:6};

// ── DADOS ─────────────────────────────────────────────────────────────────────
const MENSAL = [
  {mes:"Jan",loja:15000,matriz:10000,total:25000},
  {mes:"Fev",loja:16200,matriz:10800,total:27000},
  {mes:"Mar",loja:17500,matriz:11200,total:28700},
  {mes:"Abr",loja:18400,matriz:12600,total:31000},
];

const LOJA_CANAIS = [
  {canal:"Balcão",  valor:8200,cor:t.purple},
  {canal:"iFood",   valor:6400,cor:t.pink},
  {canal:"Delivery",valor:3800,cor:t.cyan},
];

const LOJA_PRODUTOS = [
  {nome:"Salgados", qtd:410,valor:4920,cor:t.purple},
  {nome:"Smoothies",qtd:310,valor:4650,cor:t.cyan},
  {nome:"Doces",    qtd:280,valor:3360,cor:t.pink},
  {nome:"Bowls",    qtd:120,valor:3000,cor:t.lime},
  {nome:"Sucos",    qtd:190,valor:2280,cor:t.yellow},
];

const LOJA_SEMANAL = [
  {dia:"Seg",balcao:1100,ifood:900, delivery:500},
  {dia:"Ter",balcao:1200,ifood:850, delivery:480},
  {dia:"Qua",balcao:1350,ifood:980, delivery:520},
  {dia:"Qui",balcao:1050,ifood:820, delivery:450},
  {dia:"Sex",balcao:1500,ifood:1100,delivery:600},
  {dia:"Sáb",balcao:1800,ifood:1400,delivery:720},
  {dia:"Dom",balcao:1200,ifood:950, delivery:530},
];

const MATRIZ_CATS = [
  {nome:"Pacote 20 marmitas",       qtd:40, receita:756, cor:t.purple},
  {nome:"Pacote 10 marmitas",       qtd:10, receita:200, cor:t.cyan},
  {nome:"Marmita avulsa",           qtd:5,  receita:110, cor:t.yellow},
  {nome:"Marmita congelada",        qtd:20, receita:390, cor:t.orange},
  {nome:"Encomendas (bolos/tortas)",qtd:2,  receita:205, cor:t.pink},
];

const MATRIZ_MENSAL = [
  {mes:"Jan",marmitas:980, congeladas:280,encomendas:340},
  {mes:"Fev",marmitas:1120,congeladas:390,encomendas:420},
  {mes:"Mar",marmitas:1350,congeladas:450,encomendas:510},
  {mes:"Abr",marmitas:1461,congeladas:624,encomendas:576},
];

const PEDIDOS_INIT = [
  {id:"#001",cliente:"Ana Lima",      tipo:"Pacote 20 marmitas",       qtd:20,unit:18.90,total:378.00,data:"22/04",status:"Entregue"},
  {id:"#002",cliente:"Carlos Souza",  tipo:"Marmita avulsa",            qtd:5, unit:22.00,total:110.00,data:"22/04",status:"Entregue"},
  {id:"#003",cliente:"Maria Fernanda",tipo:"Pacote 10 marmitas",        qtd:10,unit:20.00,total:200.00,data:"23/04",status:"Entregue"},
  {id:"#004",cliente:"João Pedro",    tipo:"Marmita congelada",         qtd:8, unit:19.50,total:156.00,data:"23/04",status:"Entregue"},
  {id:"#005",cliente:"Patrícia Ramos",tipo:"Encomenda — Bolo integral", qtd:1, unit:120.00,total:120.00,data:"24/04",status:"Entregue"},
  {id:"#006",cliente:"Roberto Alves", tipo:"Pacote 20 marmitas",        qtd:20,unit:18.90,total:378.00,data:"24/04",status:"A entregar"},
  {id:"#007",cliente:"Fernanda Costa",tipo:"Encomenda — Torta frango",  qtd:1, unit:85.00,total:85.00, data:"25/04",status:"A entregar"},
  {id:"#008",cliente:"Marcos Vieira", tipo:"Marmita congelada",         qtd:12,unit:19.50,total:234.00,data:"25/04",status:"A entregar"},
];

// ── BASE COMPONENTS ───────────────────────────────────────────────────────────
function Kpi({label,value,sub,color=t.purple}) {
  return (
    <div style={{background:t.card,border:`1px solid ${t.bord}`,borderRadius:12,padding:"18px 20px"}}>
      <div style={{color:t.muted,fontSize:11,letterSpacing:1,textTransform:"uppercase",marginBottom:6}}>{label}</div>
      <div style={{color,fontSize:24,fontFamily:"'Cormorant Garamond',serif",lineHeight:1}}>{value}</div>
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

const TT = {contentStyle:{background:t.bg,border:`1px solid ${t.bord}`,borderRadius:8,color:t.text}};

// ── PAINEL GERAL ──────────────────────────────────────────────────────────────
function PainelGeral() {
  const totalLoja   = LOJA_CANAIS.reduce((s,c)=>s+c.valor,0);
  const totalMatriz = MATRIZ_CATS.reduce((s,c)=>s+c.receita,0);
  const totalAbr    = totalLoja+totalMatriz;

  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12}}>
        <Kpi label="Receita Total Abr" value={fmtK(totalAbr)}    sub="Loja + Matriz"                       color={t.purple}/>
        <Kpi label="Loja — Abril"      value={fmtK(totalLoja)}   sub={pct(totalLoja,totalAbr)+" do total"}  color={t.pink}/>
        <Kpi label="Matriz — Abril"    value={fmtK(totalMatriz)} sub={pct(totalMatriz,totalAbr)+" do total"} color={t.cyan}/>
        <Kpi label="Crescimento"       value="+8%"               sub="vs março"                             color={t.lime}/>
      </div>

      <Card>
        <ST>Receita Mensal — Loja vs Matriz</ST>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={MENSAL} barCategoryGap="30%">
            <XAxis dataKey="mes" stroke={t.muted} tick={{fill:t.muted,fontSize:11}}/>
            <YAxis stroke={t.muted} tick={{fill:t.muted,fontSize:11}} tickFormatter={v=>`R$${(v/1000).toFixed(0)}k`}/>
            <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle}/>
            <Bar dataKey="loja"   fill={t.purple} radius={[4,4,0,0]} name="Loja"/>
            <Bar dataKey="matriz" fill={t.cyan}   radius={[4,4,0,0]} name="Matriz"/>
          </BarChart>
        </ResponsiveContainer>
        <div style={{display:"flex",gap:20,marginTop:8}}>
          {[["Loja",t.purple],["Matriz",t.cyan]].map(([l,c])=>(
            <div key={l} style={{display:"flex",alignItems:"center",gap:6}}>
              <div style={{width:10,height:10,borderRadius:2,background:c}}/>
              <span style={{color:t.muted,fontSize:12}}>{l}</span>
            </div>
          ))}
        </div>
      </Card>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <Card>
          <ST color={t.purple}>🏪 Loja — Abril</ST>
          {LOJA_CANAIS.map(c=>(
            <div key={c.canal} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${t.bord}`}}>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <div style={{width:8,height:8,borderRadius:"50%",background:c.cor}}/>
                <span style={{color:t.muted,fontSize:13}}>{c.canal}</span>
              </div>
              <span style={{color:c.cor,fontFamily:"monospace",fontSize:13}}>{fmt(c.valor)} <span style={{color:t.muted,fontSize:11}}>({pct(c.valor,totalLoja)})</span></span>
            </div>
          ))}
          <div style={{marginTop:10,display:"flex",justifyContent:"space-between"}}>
            <span style={{color:t.muted,fontSize:12}}>Total</span>
            <span style={{color:t.purple,fontFamily:"monospace",fontSize:14,fontWeight:700}}>{fmt(totalLoja)}</span>
          </div>
        </Card>

        <Card>
          <ST color={t.cyan}>🏭 Matriz — Abril</ST>
          {MATRIZ_CATS.map(c=>(
            <div key={c.nome} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${t.bord}`}}>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <div style={{width:8,height:8,borderRadius:"50%",background:c.cor}}/>
                <span style={{color:t.muted,fontSize:12}}>{c.nome}</span>
              </div>
              <span style={{color:c.cor,fontFamily:"monospace",fontSize:13}}>{fmt(c.receita)}</span>
            </div>
          ))}
          <div style={{marginTop:10,display:"flex",justifyContent:"space-between"}}>
            <span style={{color:t.muted,fontSize:12}}>Total</span>
            <span style={{color:t.cyan,fontFamily:"monospace",fontSize:14,fontWeight:700}}>{fmt(totalMatriz)}</span>
          </div>
        </Card>
      </div>

      <div style={{display:"flex",justifyContent:"flex-end"}}>
        <BtnExport onClick={()=>exportXLSX([{name:"Receita Mensal",data:MENSAL.map(m=>({Mês:m.mes,"Loja (R$)":m.loja,"Matriz (R$)":m.matriz,"Total (R$)":m.total}))},{name:"Canais Loja",data:LOJA_CANAIS.map(c=>({Canal:c.canal,"Valor (R$)":c.valor}))},{name:"Produtos Loja",data:LOJA_PRODUTOS.map(p=>({Produto:p.nome,Qtd:p.qtd,"Valor (R$)":p.valor}))},{name:"Categorias Matriz",data:MATRIZ_CATS.map(c=>({Categoria:c.nome,Qtd:c.qtd,"Receita (R$)":c.receita}))}],"painel-geral.xlsx")} />
      </div>
      <Card>
        <ST>Evolução Total — Últimos 4 meses</ST>
        <ResponsiveContainer width="100%" height={180}>
          <LineChart data={MENSAL}>
            <XAxis dataKey="mes" stroke={t.muted} tick={{fill:t.muted,fontSize:11}}/>
            <YAxis stroke={t.muted} tick={{fill:t.muted,fontSize:11}} tickFormatter={v=>`R$${(v/1000).toFixed(0)}k`}/>
            <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle}/>
            <Line type="monotone" dataKey="total" stroke={t.purple} strokeWidth={2.5} dot={{fill:t.purple,r:5}} name="Total"/>
          </LineChart>
        </ResponsiveContainer>
      </Card>
    </div>
  );
}

// ── VENDAS LOJA ───────────────────────────────────────────────────────────────
function VendasLoja() {
  const [vendas,   setVendas]   = useState([]);
  const [form,     setForm]     = useState({data:"27/04",produto:"",qtd:"",valor:"",canal:"Balcão"});
  const [showForm, setShowForm] = useState(false);

  const totalCanais = LOJA_CANAIS.reduce((s,c)=>s+c.valor,0);
  const maxVal      = Math.max(...LOJA_PRODUTOS.map(p=>p.valor));

  const salvar=async()=>{if(!form.produto||!form.valor)return;try{await sbPost("vendas_loja",{data:new Date().toISOString().split("T")[0],produto:form.produto,qtd:parseInt(form.qtd)||1,valor:parseFloat(form.valor),canal:form.canal});setVendas(p=>[...p,{...form,qtd:parseInt(form.qtd)||1,valor:parseFloat(form.valor),id:"#"+String(p.length+1).padStart(3,"0")}]);setForm(f=>({...f,produto:"",qtd:"",valor:""}));setShowForm(false);}catch(e){alert("Erro: "+e.message);}};

  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
        <Kpi label="Receita Loja Abr" value={fmt(totalCanais)} sub="Abril 2025"         color={t.purple}/>
        <Kpi label="Ticket médio"     value="R$ 32,40"         sub="Por transação"      color={t.pink}/>
        <Kpi label="Canal top"        value="Balcão"           sub={fmt(8200)+" (44%)"} color={t.cyan}/>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <Card>
          <ST>Vendas por Canal</ST>
          <ResponsiveContainer width="100%" height={160}>
            <PieChart>
              <Pie data={LOJA_CANAIS} dataKey="valor" nameKey="canal" cx="50%" cy="50%" outerRadius={70} innerRadius={36}>
                {LOJA_CANAIS.map((c,i)=><Cell key={i} fill={c.cor}/>)}
              </Pie>
              <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle}/>
            </PieChart>
          </ResponsiveContainer>
          {LOJA_CANAIS.map(c=>(
            <div key={c.canal} style={{display:"flex",justifyContent:"space-between",padding:"5px 0"}}>
              <div style={{display:"flex",alignItems:"center",gap:7}}>
                <div style={{width:7,height:7,borderRadius:"50%",background:c.cor}}/>
                <span style={{color:t.muted,fontSize:12}}>{c.canal}</span>
              </div>
              <span style={{color:t.text,fontFamily:"monospace",fontSize:12}}>{fmt(c.valor)} · {pct(c.valor,totalCanais)}</span>
            </div>
          ))}
        </Card>

        <Card>
          <ST>Top Produtos</ST>
          {[...LOJA_PRODUTOS].sort((a,b)=>b.valor-a.valor).map(p=>(
            <div key={p.nome} style={{marginBottom:12}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                <span style={{color:t.text,fontSize:13}}>{p.nome}</span>
                <div>
                  <span style={{color:p.cor,fontFamily:"monospace",fontSize:13}}>{fmt(p.valor)}</span>
                  <span style={{color:t.muted,fontSize:11,marginLeft:8}}>{p.qtd} un</span>
                </div>
              </div>
              <div style={{background:t.bg,borderRadius:4,height:5,overflow:"hidden"}}>
                <div style={{width:pct(p.valor,maxVal),background:p.cor,height:"100%",borderRadius:4}}/>
              </div>
            </div>
          ))}
        </Card>
      </div>

      <Card>
        <ST>Vendas por Dia — Semana Atual</ST>
        <ResponsiveContainer width="100%" height={210}>
          <BarChart data={LOJA_SEMANAL} barCategoryGap="20%">
            <XAxis dataKey="dia" stroke={t.muted} tick={{fill:t.muted,fontSize:11}}/>
            <YAxis stroke={t.muted} tick={{fill:t.muted,fontSize:11}} tickFormatter={v=>`R$${(v/1000).toFixed(1)}k`}/>
            <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle}/>
            <Bar dataKey="balcao"   fill={t.purple} radius={[3,3,0,0]} name="Balcão"   stackId="a"/>
            <Bar dataKey="ifood"    fill={t.pink}                      name="iFood"    stackId="a"/>
            <Bar dataKey="delivery" fill={t.cyan}   radius={[0,0,0,0]} name="Delivery" stackId="a"/>
          </BarChart>
        </ResponsiveContainer>
        <div style={{display:"flex",gap:16,marginTop:8}}>
          {[["Balcão",t.purple],["iFood",t.pink],["Delivery",t.cyan]].map(([l,c])=>(
            <div key={l} style={{display:"flex",alignItems:"center",gap:6}}>
              <div style={{width:10,height:10,borderRadius:2,background:c}}/>
              <span style={{color:t.muted,fontSize:12}}>{l}</span>
            </div>
          ))}
        </div>
      </Card>

      <div style={{display:"flex",justifyContent:"flex-end"}}>
        <button onClick={()=>setShowForm(!showForm)} style={{background:showForm?t.card:t.purple,color:showForm?t.muted:t.bg,border:`1px solid ${showForm?t.bord:t.purple}`,borderRadius:8,padding:"9px 20px",cursor:"pointer",fontSize:12,fontWeight:700}}>
          {showForm?"✕ Cancelar":"+ Lançar Venda"}
        </button>
      </div>

      {showForm&&(
        <Card style={{border:`1px solid ${t.purple}`}}>
          <ST color={t.purple}>+ Nova Venda — Loja</ST>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr 1fr",gap:10}}>
            <div><label style={lbl}>Data</label><input value={form.data} onChange={e=>setForm(f=>({...f,data:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Produto</label><input value={form.produto} onChange={e=>setForm(f=>({...f,produto:e.target.value}))} style={inp} placeholder="Ex: Smoothie"/></div>
            <div><label style={lbl}>Qtd</label><input type="number" value={form.qtd} onChange={e=>setForm(f=>({...f,qtd:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Valor (R$)</label><input type="number" value={form.valor} onChange={e=>setForm(f=>({...f,valor:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Canal</label>
              <select value={form.canal} onChange={e=>setForm(f=>({...f,canal:e.target.value}))} style={inp}>
                <option>Balcão</option><option>iFood</option><option>Delivery</option>
              </select>
            </div>
          </div>
          <div style={{display:"flex",gap:10,marginTop:14}}>
            <button onClick={salvar} style={{background:t.purple,color:t.bg,border:"none",borderRadius:8,padding:"10px 22px",cursor:"pointer",fontSize:12,fontWeight:700}}>Salvar</button>
            <button onClick={()=>setShowForm(false)} style={{background:"transparent",color:t.muted,border:`1px solid ${t.bord}`,borderRadius:8,padding:"10px 18px",cursor:"pointer",fontSize:12}}>Cancelar</button>
          </div>
        </Card>
      )}

      {vendas.length>0&&(
        <Card>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <ST>Vendas Lançadas</ST>
            <BtnExport onClick={()=>exportXLSX([{name:"Vendas Loja",data:vendas.map(v=>({ID:v.id,Data:v.data,Produto:v.produto,Qtd:v.qtd,"Valor (R$)":v.valor,Canal:v.canal}))}],"vendas-loja.xlsx")} />
          </div>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr style={{background:t.bg}}>{["Data","Produto","Qtd","Valor","Canal"].map(h=><th key={h} style={{padding:"9px 12px",textAlign:"left",color:t.muted,fontSize:11,letterSpacing:1,textTransform:"uppercase"}}>{h}</th>)}</tr></thead>
            <tbody>
              {[...vendas].reverse().map((v,i)=>(
                <tr key={i} style={{borderTop:`1px solid ${t.bord}`}}>
                  <td style={{padding:"9px 12px",color:t.muted,fontFamily:"monospace",fontSize:12}}>{v.data}</td>
                  <td style={{padding:"9px 12px",color:t.text,fontSize:13}}>{v.produto}</td>
                  <td style={{padding:"9px 12px",color:t.text,fontFamily:"monospace"}}>{v.qtd}</td>
                  <td style={{padding:"9px 12px",color:t.purple,fontFamily:"monospace",fontSize:13,fontWeight:700}}>{fmt(v.valor)}</td>
                  <td style={{padding:"9px 12px",color:t.muted,fontSize:12}}>{v.canal}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </Card>
      )}
    </div>
  );
}

// ── VENDAS MATRIZ ─────────────────────────────────────────────────────────────
function VendasMatriz() {
  const [pedidos,setPedidos]=useState(PEDIDOS_INIT);
  useEffect(()=>{(async()=>{try{const rows=await sbGet("pedidos_matriz","order=created_at.desc");if(rows.length>0)setPedidos(rows.map(p=>({...p,data:p.data?new Date(p.data).toLocaleDateString("pt-BR"):"",unit:p.valor_unit})));}catch(e){console.error(e);}})();},[]);
  const [filtro,   setFiltro]   = useState("Todos");
  const [form,     setForm]     = useState({cliente:"",tipo:"Pacote 20 marmitas",qtd:"",unit:"",data:"27/04"});
  const [showForm, setShowForm] = useState(false);

  const totalReceita = MATRIZ_CATS.reduce((s,c)=>s+c.receita,0);
  const totalUnid    = MATRIZ_CATS.reduce((s,c)=>s+c.qtd,0);
  const aEntregar    = pedidos.filter(p=>p.status==="A entregar").length;
  const filtrados    = filtro==="Todos"?pedidos:pedidos.filter(p=>p.status===filtro);

  const salvar=async()=>{if(!form.cliente||!form.qtd||!form.unit)return;const total=parseFloat(form.unit)*parseInt(form.qtd);try{const[novo]=await sbPost("pedidos_matriz",{cliente:form.cliente,tipo:form.tipo,qtd:parseInt(form.qtd),valor_unit:parseFloat(form.unit),total,data:new Date().toISOString().split("T")[0],status:"A entregar"});setPedidos(p=>[...p,{...novo,unit:novo.valor_unit,data:new Date().toLocaleDateString("pt-BR")}]);setForm({cliente:"",tipo:"Pacote 20 marmitas",qtd:"",unit:"",data:""});setShowForm(false);}catch(e){alert("Erro: "+e.message);}};

  const marcarEntregue=async(id)=>{try{await sbPatch("pedidos_matriz",id,{status:"Entregue"});setPedidos(p=>p.map(x=>x.id!==id?x:{...x,status:"Entregue"}));}catch(e){alert("Erro: "+e.message);}};

  const pieData = MATRIZ_CATS.map(c=>({name:c.nome,value:c.receita,cor:c.cor}));

  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12}}>
        <Kpi label="Receita direta" value={fmt(totalReceita)} sub="Abril 2025"  color={t.cyan}/>
        <Kpi label="Pedidos"        value={pedidos.length}    sub="Mês atual"   color={t.purple}/>
        <Kpi label="Unidades"       value={totalUnid}         sub="Vendidas"    color={t.lime}/>
        <Kpi label="A entregar"     value={aEntregar}         sub="Pendentes"   color={aEntregar>0?t.yellow:t.green}/>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <Card>
          <ST>Mix de Produtos</ST>
          <ResponsiveContainer width="100%" height={160}>
            <PieChart>
              <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={68} innerRadius={34}>
                {pieData.map((d,i)=><Cell key={i} fill={d.cor}/>)}
              </Pie>
              <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle}/>
            </PieChart>
          </ResponsiveContainer>
          {MATRIZ_CATS.map(c=>(
            <div key={c.nome} style={{display:"flex",justifyContent:"space-between",padding:"4px 0"}}>
              <div style={{display:"flex",alignItems:"center",gap:7}}>
                <div style={{width:7,height:7,borderRadius:"50%",background:c.cor}}/>
                <span style={{color:t.muted,fontSize:11}}>{c.nome}</span>
              </div>
              <span style={{color:t.text,fontFamily:"monospace",fontSize:11}}>{pct(c.receita,totalReceita)}</span>
            </div>
          ))}
        </Card>

        <Card>
          <ST>Receita Mensal por Categoria</ST>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={MATRIZ_MENSAL} barCategoryGap="20%">
              <XAxis dataKey="mes" stroke={t.muted} tick={{fill:t.muted,fontSize:11}}/>
              <YAxis stroke={t.muted} tick={{fill:t.muted,fontSize:11}} tickFormatter={v=>`R$${(v/1000).toFixed(1)}k`}/>
              <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle}/>
              <Bar dataKey="marmitas"   fill={t.purple} radius={[3,3,0,0]} name="Marmitas"   stackId="a"/>
              <Bar dataKey="congeladas" fill={t.orange}                    name="Congeladas" stackId="a"/>
              <Bar dataKey="encomendas" fill={t.cyan}   radius={[3,3,0,0]} name="Encomendas" stackId="a"/>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>

      <div style={{display:"flex",justifyContent:"flex-end"}}>
        <button onClick={()=>setShowForm(!showForm)} style={{background:showForm?t.card:t.cyan,color:showForm?t.muted:t.bg,border:`1px solid ${showForm?t.bord:t.cyan}`,borderRadius:8,padding:"9px 20px",cursor:"pointer",fontSize:12,fontWeight:700}}>
          {showForm?"✕ Cancelar":"+ Novo Pedido"}
        </button>
      </div>

      {showForm&&(
        <Card style={{border:`1px solid ${t.cyan}`}}>
          <ST color={t.cyan}>+ Novo Pedido — Matriz</ST>
          <div style={{display:"grid",gridTemplateColumns:"2fr 2fr 1fr 1fr 1fr",gap:10}}>
            <div><label style={lbl}>Cliente</label><input value={form.cliente} onChange={e=>setForm(f=>({...f,cliente:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Tipo</label>
              <select value={form.tipo} onChange={e=>setForm(f=>({...f,tipo:e.target.value}))} style={inp}>
                {["Pacote 20 marmitas","Pacote 10 marmitas","Marmita avulsa","Marmita congelada","Encomenda — Bolo","Encomenda — Torta"].map(o=><option key={o}>{o}</option>)}
              </select>
            </div>
            <div><label style={lbl}>Qtd</label><input type="number" value={form.qtd} onChange={e=>setForm(f=>({...f,qtd:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Vlr/un</label><input type="number" value={form.unit} onChange={e=>setForm(f=>({...f,unit:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Data</label><input value={form.data} onChange={e=>setForm(f=>({...f,data:e.target.value}))} style={inp}/></div>
          </div>
          {form.qtd&&form.unit&&(
            <div style={{marginTop:10,background:t.bg,borderRadius:8,padding:"9px 12px",color:t.muted,fontSize:12}}>
              Total: <span style={{color:t.cyan,fontFamily:"monospace"}}>{fmt(parseFloat(form.unit||0)*parseInt(form.qtd||0))}</span>
            </div>
          )}
          <div style={{display:"flex",gap:10,marginTop:14}}>
            <button onClick={salvar} style={{background:t.cyan,color:t.bg,border:"none",borderRadius:8,padding:"10px 22px",cursor:"pointer",fontSize:12,fontWeight:700}}>Salvar</button>
            <button onClick={()=>setShowForm(false)} style={{background:"transparent",color:t.muted,border:`1px solid ${t.bord}`,borderRadius:8,padding:"10px 18px",cursor:"pointer",fontSize:12}}>Cancelar</button>
          </div>
        </Card>
      )}

      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
          <ST>Pedidos da Matriz</ST>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            <BtnExport onClick={()=>exportXLSX([{name:"Pedidos Matriz",data:pedidos.map(p=>({ID:p.id,Cliente:p.cliente,Tipo:p.tipo,Qtd:p.qtd,"Vlr Unit (R$)":p.unit,"Total (R$)":p.total,Data:p.data,Status:p.status}))}],"pedidos-matriz.xlsx")} />
          </div>
          <div style={{display:"flex",gap:6}}>
            {["Todos","Entregue","A entregar"].map(s=>(
              <button key={s} onClick={()=>setFiltro(s)} style={{background:filtro===s?t.cyan:"transparent",color:filtro===s?t.bg:t.muted,border:`1px solid ${filtro===s?t.cyan:t.bord}`,borderRadius:6,padding:"4px 12px",cursor:"pointer",fontSize:11}}>{s}</button>
            ))}
          </div>
        </div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead>
              <tr style={{background:t.bg}}>
                {["Pedido","Cliente","Tipo","Qtd","Unit.","Total","Data","Status",""].map(h=>(
                  <th key={h} style={{padding:"9px 12px",textAlign:"left",color:t.muted,fontSize:11,letterSpacing:1,textTransform:"uppercase",whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtrados.map((p,i)=>(
                <tr key={i} style={{borderTop:`1px solid ${t.bord}`}}>
                  <td style={{padding:"10px 12px",color:t.muted,fontFamily:"monospace",fontSize:12}}>{p.id}</td>
                  <td style={{padding:"10px 12px",color:t.text,fontSize:13}}>{p.cliente}</td>
                  <td style={{padding:"10px 12px",color:t.purple,fontSize:12}}>{p.tipo}</td>
                  <td style={{padding:"10px 12px",color:t.text,fontFamily:"monospace",fontSize:13}}>{p.qtd}</td>
                  <td style={{padding:"10px 12px",color:t.muted,fontFamily:"monospace",fontSize:12}}>{fmt(p.unit)}</td>
                  <td style={{padding:"10px 12px",color:t.cyan,fontFamily:"monospace",fontSize:13,fontWeight:700}}>{fmt(p.total)}</td>
                  <td style={{padding:"10px 12px",color:t.muted,fontFamily:"monospace",fontSize:12}}>{p.data}</td>
                  <td style={{padding:"10px 12px"}}>
                    <span style={{background:p.status==="Entregue"?"#0f2a0f":"#3b2f00",color:p.status==="Entregue"?t.green:t.yellow,borderRadius:6,padding:"3px 10px",fontSize:11,fontFamily:"monospace"}}>{p.status}</span>
                  </td>
                  <td style={{padding:"10px 12px"}}>
                    {p.status==="A entregar"&&(
                      <button onClick={()=>marcarEntregue(p.id)} style={{background:"#0f2a0f",color:t.green,border:`1px solid ${t.green}`,borderRadius:6,padding:"4px 10px",cursor:"pointer",fontSize:11}}>✓ Entregar</button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ── APP ───────────────────────────────────────────────────────────────────────
const TABS = [
  {id:"geral", label:"📊 Painel Geral"},
  {id:"loja",  label:"🏪 Vendas Loja"},
  {id:"matriz",label:"🏭 Vendas Matriz"},
];

export default function App() {
  const [tab, setTab] = useState("geral");
  return (
    <div style={{minHeight:"100vh",background:t.bg,color:t.text,fontFamily:"system-ui,sans-serif"}}>
      <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;600&family=DM+Sans:wght@400;500;600&display=swap" rel="stylesheet"/>

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
        {tab==="geral"  && <PainelGeral/>}
        {tab==="loja"   && <VendasLoja/>}
        {tab==="matriz" && <VendasMatriz/>}
      </div>
    </div>
  );
}

import { useState } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer } from "recharts";

const ACCENT       = "#1d4ed8";
const ACCENT_LIGHT = "#dbeafe";
const ACCENT_DARK  = "#1e3a8a";

const t = {
  bg:"#f8f8f4", card:"#ffffff", bord:"#e0e8d8",
  green:"#4a7c3f", text:"#1a1a1a", muted:"#5a6a5a",
  red:"#dc2626", orange:"#ea580c", yellow:"#d97706",
};

const fmt = v => "R$ " + Number(v||0).toLocaleString("pt-BR", {minimumFractionDigits:2});
const inp = {background:t.bg,border:"1px solid "+t.bord,borderRadius:8,padding:"9px 13px",color:t.text,fontSize:13,fontFamily:"inherit",outline:"none",width:"100%",boxSizing:"border-box"};
const lbl = {color:t.muted,fontSize:11,letterSpacing:1,textTransform:"uppercase",display:"block",marginBottom:6,fontWeight:600};
const TT  = {contentStyle:{background:t.card,border:"1px solid "+t.bord,borderRadius:8,color:t.text}};

function exportXLSX(sheets, filename) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({name,data}) => {
    if (!data||!data.length) return;
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = Object.keys(data[0]).map(k=>({wch:Math.max(k.length+2,16)}));
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0,31));
  });
  XLSX.writeFile(wb, filename);
}

function AppHeader({children}) {
  return (
    <div style={{borderBottom:"1px solid "+t.bord,background:t.card,boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
      <div style={{padding:"10px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:42,height:42,borderRadius:10,background:"linear-gradient(135deg,#4a7c3f,#6aab5e)",display:"flex",alignItems:"center",justifyContent:"center"}}>
            <span style={{color:"#fff",fontSize:20,fontFamily:"Georgia,serif",fontWeight:700}}>D</span>
          </div>
          <div>
            <div style={{fontSize:11,color:t.muted,letterSpacing:"0.08em",textTransform:"uppercase",fontWeight:600}}>Dona Luzia</div>
            <div style={{fontSize:15,color:ACCENT_DARK,fontWeight:700,fontFamily:"Georgia,serif"}}>Operações</div>
          </div>
        </div>
        {children}
      </div>
      <div style={{height:3,background:ACCENT_DARK}}/>
    </div>
  );
}

function Kpi({label,value,sub,color=ACCENT,icon}) {
  return (
    <div style={{background:t.card,border:"1px solid "+t.bord,borderRadius:12,padding:"16px 18px",position:"relative",overflow:"hidden"}}>
      {icon&&<div style={{position:"absolute",top:12,right:14,fontSize:20,opacity:0.15}}>{icon}</div>}
      <div style={{color:t.muted,fontSize:11,letterSpacing:1,textTransform:"uppercase",marginBottom:6,fontWeight:600}}>{label}</div>
      <div style={{color,fontSize:22,fontFamily:"Georgia,serif",fontWeight:700,lineHeight:1}}>{value}</div>
      {sub&&<div style={{color:t.muted,fontSize:12,marginTop:4}}>{sub}</div>}
    </div>
  );
}

function Card({children,style={}}) {
  return <div style={{background:t.card,border:"1px solid "+t.bord,borderRadius:12,padding:20,...style}}>{children}</div>;
}

function ST({children,color=t.muted}) {
  return <div style={{color,fontSize:11,letterSpacing:1.5,textTransform:"uppercase",marginBottom:14,fontWeight:700}}>{children}</div>;
}

function TabBtn({label,active,onClick}) {
  return (
    <button onClick={onClick} style={{
      background:active?ACCENT:"transparent",color:active?"#fff":t.muted,
      border:"1px solid "+(active?ACCENT:t.bord),borderRadius:8,
      padding:"8px 18px",cursor:"pointer",fontSize:12,fontWeight:active?700:400,
    }}>{label}</button>
  );
}

function OpBtn({op,setOp}) {
  return (
    <div style={{display:"flex",gap:8}}>
      {[["loja","Loja"],["matriz","Matriz"]].map(([id,label])=>(
        <button key={id} onClick={()=>setOp(id)} style={{
          background:op===id?ACCENT:t.card,color:op===id?"#fff":t.muted,
          border:"1px solid "+(op===id?ACCENT:t.bord),borderRadius:8,
          padding:"8px 18px",cursor:"pointer",fontSize:12,fontWeight:op===id?700:400,
        }}>{id==="loja"?"🏪 ":"🏭 "}{label}</button>
      ))}
    </div>
  );
}

const INSUMOS_INIT = {
  loja:[
    {id:1,nome:"Banana",     cat:"Frutas",     un:"kg", qtd:8.5, min:5,  vunit:6.90},
    {id:2,nome:"Morango",    cat:"Frutas",     un:"kg", qtd:2.0, min:3,  vunit:18.00},
    {id:3,nome:"Leite",      cat:"Laticínios", un:"L",  qtd:12,  min:8,  vunit:4.50},
    {id:4,nome:"Whey",       cat:"Suplementos",un:"kg", qtd:0.8, min:1,  vunit:90.00},
    {id:5,nome:"Granola",    cat:"Ingredientes",un:"kg",qtd:3.0, min:2,  vunit:22.00},
  ],
  matriz:[
    {id:1,nome:"Arroz integral",cat:"Grãos",    un:"kg",qtd:18, min:10, vunit:5.80},
    {id:2,nome:"Frango",        cat:"Proteínas",un:"kg",qtd:6.5,min:5,  vunit:14.90},
    {id:3,nome:"Carne moída",   cat:"Proteínas",un:"kg",qtd:3.0,min:4,  vunit:32.00},
    {id:4,nome:"Brócolis",      cat:"Vegetais", un:"kg",qtd:1.5,min:2,  vunit:8.00},
    {id:5,nome:"Azeite",        cat:"Condimentos",un:"L",qtd:2, min:1,  vunit:28.00},
  ],
};

const PROD_INIT = {
  loja:[
    {id:1,data:"25/04",produto:"Bowl frutas",      qtd:18,canal:"Balcão",  status:"Produzido"},
    {id:2,data:"25/04",produto:"Smoothie banana",  qtd:25,canal:"iFood",   status:"Produzido"},
    {id:3,data:"24/04",produto:"Suco detox",       qtd:12,canal:"Balcão",  status:"Produzido"},
  ],
  matriz:[
    {id:1,data:"25/04",produto:"Marmita frango",   qtd:40,canal:"Delivery",status:"Produzido"},
    {id:2,data:"25/04",produto:"Marmita vegana",   qtd:20,canal:"Delivery",status:"Produzindo"},
    {id:3,data:"24/04",produto:"Marmita carne",    qtd:35,canal:"Delivery",status:"Produzido"},
  ],
};

const DESP_INIT = {
  loja:[
    {id:1,data:"23/04",produto:"Morango",    qtd:0.3,un:"kg",custo:5.40, motivo:"Vencido"},
    {id:2,data:"22/04",produto:"Smoothie",   qtd:2,  un:"un",custo:17.00,motivo:"Erro preparo"},
  ],
  matriz:[
    {id:1,data:"23/04",produto:"Frango",     qtd:0.5,un:"kg",custo:7.45, motivo:"Erro preparo"},
    {id:2,data:"22/04",produto:"Arroz",      qtd:0.8,un:"kg",custo:4.64, motivo:"Queimado"},
    {id:3,data:"21/04",produto:"Marmitas",   qtd:6,  un:"un",custo:54.00,motivo:"Erro preparo"},
  ],
};

function Estoque() {
  const [op,setOp]=useState("loja");
  const [insumos,setInsumos]=useState(INSUMOS_INIT);
  const [showForm,setShowForm]=useState(false);
  const [form,setForm]=useState({nome:"",cat:"",un:"kg",qtd:"",min:"",vunit:""});

  const lista=insumos[op];
  const criticos=lista.filter(i=>i.qtd<=i.min);
  const valorTotal=lista.reduce((s,i)=>s+i.qtd*i.vunit,0);

  const salvar=()=>{
    if(!form.nome) return;
    setInsumos(p=>({...p,[op]:[...p[op],{...form,id:Date.now(),qtd:parseFloat(form.qtd)||0,min:parseFloat(form.min)||0,vunit:parseFloat(form.vunit)||0}]}));
    setForm({nome:"",cat:"",un:"kg",qtd:"",min:"",vunit:""});
    setShowForm(false);
  };

  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
        <OpBtn op={op} setOp={setOp}/>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>exportXLSX([{name:"Estoque",data:lista.map(i=>({Nome:i.nome,Cat:i.cat,Un:i.un,Qtd:i.qtd,Min:i.min,"Vlr Unit":i.vunit,"Total":+(i.qtd*i.vunit).toFixed(2)}))}],"estoque.xlsx")} style={{background:ACCENT_LIGHT,color:ACCENT_DARK,border:"1px solid "+ACCENT,borderRadius:7,padding:"6px 12px",cursor:"pointer",fontSize:11,fontWeight:700}}>⬇ Excel</button>
          <button onClick={()=>setShowForm(!showForm)} style={{background:showForm?t.card:t.green,color:showForm?t.muted:"#fff",border:"1px solid "+(showForm?t.bord:t.green),borderRadius:8,padding:"8px 16px",cursor:"pointer",fontSize:12,fontWeight:700}}>{showForm?"✕ Cancelar":"+ Insumo"}</button>
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
        <Kpi label="Itens" value={lista.length} color={ACCENT} icon="📦"/>
        <Kpi label="Críticos" value={criticos.length} color={criticos.length>0?t.red:t.green} icon="🔴"/>
        <Kpi label="Valor total" value={fmt(valorTotal)} color={t.green} icon="💰"/>
      </div>

      {criticos.length>0&&<div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:10,padding:"10px 14px",color:t.red,fontSize:13}}>🔴 Repor urgente: {criticos.map(i=>i.nome).join(", ")}</div>}

      {showForm&&(
        <Card style={{border:"1px solid "+t.green}}>
          <ST color={t.green}>+ Novo Insumo</ST>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:12}}>
            {[["Nome","nome","text"],["Categoria","cat","text"],["Unidade","un","text"],["Qtd atual","qtd","number"],["Mínimo","min","number"],["Vlr unit","vunit","number"]].map(([l,k,tp])=>(
              <div key={k}><label style={lbl}>{l}</label><input type={tp} value={form[k]} onChange={e=>setForm(f=>({...f,[k]:e.target.value}))} style={inp}/></div>
            ))}
          </div>
          <button onClick={salvar} style={{background:t.green,color:"#fff",border:"none",borderRadius:8,padding:"10px 22px",cursor:"pointer",fontSize:13,fontWeight:700}}>Salvar</button>
        </Card>
      )}

      <Card>
        <ST>Insumos — {op==="loja"?"Loja":"Matriz"}</ST>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",minWidth:480}}>
            <thead>
              <tr style={{background:t.bg}}>
                {["Insumo","Cat","Qtd","Mín","Vlr Unit","Total","Status"].map(h=>(
                  <th key={h} style={{padding:"9px 10px",textAlign:"left",color:t.muted,fontSize:11,letterSpacing:1,textTransform:"uppercase",whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {lista.map(i=>{
                const crit=i.qtd<=i.min;
                return (
                  <tr key={i.id} style={{borderTop:"1px solid "+t.bord,background:crit?"#fff8f8":"transparent"}}>
                    <td style={{padding:"9px 10px",color:t.text,fontSize:13,fontWeight:600}}>{i.nome}</td>
                    <td style={{padding:"9px 10px",color:t.muted,fontSize:12}}>{i.cat}</td>
                    <td style={{padding:"9px 10px",color:crit?t.red:t.text,fontFamily:"monospace",fontSize:13,fontWeight:700}}>{i.qtd} {i.un}</td>
                    <td style={{padding:"9px 10px",color:t.muted,fontFamily:"monospace",fontSize:12}}>{i.min}</td>
                    <td style={{padding:"9px 10px",color:t.muted,fontFamily:"monospace",fontSize:12}}>{fmt(i.vunit)}</td>
                    <td style={{padding:"9px 10px",color:ACCENT_DARK,fontFamily:"monospace",fontSize:13,fontWeight:700}}>{fmt(i.qtd*i.vunit)}</td>
                    <td style={{padding:"9px 10px"}}>
                      <span style={{background:crit?"#fef2f2":ACCENT_LIGHT,color:crit?t.red:ACCENT_DARK,borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:700}}>{crit?"⚠️ Crítico":"✅ OK"}</span>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

function Producao() {
  const [op,setOp]=useState("loja");
  const [lista,setLista]=useState(PROD_INIT);
  const [showForm,setShowForm]=useState(false);
  const [form,setForm]=useState({produto:"",qtd:"",canal:"Balcão",status:"Produzindo"});

  const itens=lista[op];

  const salvar=()=>{
    if(!form.produto||!form.qtd) return;
    setLista(p=>({...p,[op]:[{...form,id:Date.now(),data:new Date().toLocaleDateString("pt-BR"),qtd:parseInt(form.qtd)},...p[op]]}));
    setForm({produto:"",qtd:"",canal:"Balcão",status:"Produzindo"});
    setShowForm(false);
  };

  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
        <OpBtn op={op} setOp={setOp}/>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>exportXLSX([{name:"Producao",data:itens.map(p=>({Data:p.data,Produto:p.produto,Qtd:p.qtd,Canal:p.canal,Status:p.status}))}],"producao.xlsx")} style={{background:ACCENT_LIGHT,color:ACCENT_DARK,border:"1px solid "+ACCENT,borderRadius:7,padding:"6px 12px",cursor:"pointer",fontSize:11,fontWeight:700}}>⬇ Excel</button>
          <button onClick={()=>setShowForm(!showForm)} style={{background:showForm?t.card:ACCENT,color:showForm?t.muted:"#fff",border:"1px solid "+(showForm?t.bord:ACCENT),borderRadius:8,padding:"8px 16px",cursor:"pointer",fontSize:12,fontWeight:700}}>{showForm?"✕ Cancelar":"+ Registrar"}</button>
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
        <Kpi label="Registros" value={itens.length} color={ACCENT} icon="📋"/>
        <Kpi label="Total produzido" value={itens.reduce((s,p)=>s+p.qtd,0)+" un"} color={t.green} icon="🍱"/>
        <Kpi label="Em produção" value={itens.filter(p=>p.status==="Produzindo").length} color={t.yellow} icon="⚙️"/>
      </div>

      {showForm&&(
        <Card style={{border:"1px solid "+ACCENT}}>
          <ST color={ACCENT}>+ Registrar Produção</ST>
          <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:10,marginBottom:12}}>
            <div><label style={lbl}>Produto</label><input value={form.produto} onChange={e=>setForm(f=>({...f,produto:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Qtd</label><input type="number" value={form.qtd} onChange={e=>setForm(f=>({...f,qtd:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Canal</label>
              <select value={form.canal} onChange={e=>setForm(f=>({...f,canal:e.target.value}))} style={inp}>
                {["Balcão","iFood","Delivery","Interno"].map(c=><option key={c}>{c}</option>)}
              </select>
            </div>
            <div><label style={lbl}>Status</label>
              <select value={form.status} onChange={e=>setForm(f=>({...f,status:e.target.value}))} style={inp}>
                {["Produzindo","Produzido","Cancelado"].map(s=><option key={s}>{s}</option>)}
              </select>
            </div>
          </div>
          <button onClick={salvar} style={{background:ACCENT,color:"#fff",border:"none",borderRadius:8,padding:"10px 22px",cursor:"pointer",fontSize:13,fontWeight:700}}>Salvar</button>
        </Card>
      )}

      <Card>
        <ST>Registros — {op==="loja"?"Loja":"Matriz"}</ST>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead>
            <tr style={{background:t.bg}}>
              {["Data","Produto","Qtd","Canal","Status"].map(h=>(
                <th key={h} style={{padding:"9px 12px",textAlign:"left",color:t.muted,fontSize:11,letterSpacing:1,textTransform:"uppercase"}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {itens.map(p=>(
              <tr key={p.id} style={{borderTop:"1px solid "+t.bord}}>
                <td style={{padding:"9px 12px",color:t.muted,fontFamily:"monospace",fontSize:12}}>{p.data}</td>
                <td style={{padding:"9px 12px",color:t.text,fontSize:13,fontWeight:600}}>{p.produto}</td>
                <td style={{padding:"9px 12px",color:ACCENT_DARK,fontFamily:"monospace",fontSize:13,fontWeight:700}}>{p.qtd}</td>
                <td style={{padding:"9px 12px",color:t.muted,fontSize:12}}>{p.canal}</td>
                <td style={{padding:"9px 12px"}}>
                  <span style={{background:p.status==="Produzido"?"#f0fdf4":p.status==="Produzindo"?"#fffbeb":"#f9fafb",color:p.status==="Produzido"?t.green:p.status==="Produzindo"?t.yellow:t.muted,borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:600}}>{p.status}</span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

function Desperdicio() {
  const [op,setOp]=useState("loja");
  const [lista,setLista]=useState(DESP_INIT);
  const [showForm,setShowForm]=useState(false);
  const [form,setForm]=useState({produto:"",qtd:"",un:"kg",custo:"",motivo:"Vencido"});

  const itens=lista[op];
  const totalCusto=itens.reduce((s,d)=>s+d.custo,0);
  const motivoData=Object.entries(itens.reduce((acc,d)=>{acc[d.motivo]=(acc[d.motivo]||0)+d.custo;return acc;},{})).map(([motivo,custo])=>({motivo,custo}));

  const salvar=()=>{
    if(!form.produto||!form.custo) return;
    setLista(p=>({...p,[op]:[{...form,id:Date.now(),data:new Date().toLocaleDateString("pt-BR"),qtd:parseFloat(form.qtd)||0,custo:parseFloat(form.custo)},...p[op]]}));
    setForm({produto:"",qtd:"",un:"kg",custo:"",motivo:"Vencido"});
    setShowForm(false);
  };

  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
        <OpBtn op={op} setOp={setOp}/>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>exportXLSX([{name:"Desperdicio",data:itens.map(d=>({Data:d.data,Produto:d.produto,Qtd:d.qtd,Un:d.un,Motivo:d.motivo,"Custo":d.custo}))}],"desperdicio.xlsx")} style={{background:ACCENT_LIGHT,color:ACCENT_DARK,border:"1px solid "+ACCENT,borderRadius:7,padding:"6px 12px",cursor:"pointer",fontSize:11,fontWeight:700}}>⬇ Excel</button>
          <button onClick={()=>setShowForm(!showForm)} style={{background:showForm?t.card:t.red,color:showForm?t.muted:"#fff",border:"1px solid "+(showForm?t.bord:t.red),borderRadius:8,padding:"8px 16px",cursor:"pointer",fontSize:12,fontWeight:700}}>{showForm?"✕ Cancelar":"+ Registrar"}</button>
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
        <Kpi label="Registros" value={itens.length} color={t.red} icon="🗑"/>
        <Kpi label="Custo total" value={fmt(totalCusto)} color={t.orange} icon="💸"/>
        <Kpi label="Motivos" value={motivoData.length} color={t.yellow} icon="📊"/>
      </div>

      {showForm&&(
        <Card style={{border:"1px solid "+t.red}}>
          <ST color={t.red}>+ Registrar Desperdício</ST>
          <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr 2fr",gap:10,marginBottom:12}}>
            <div><label style={lbl}>Produto</label><input value={form.produto} onChange={e=>setForm(f=>({...f,produto:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Qtd</label><input type="number" value={form.qtd} onChange={e=>setForm(f=>({...f,qtd:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Unidade</label><input value={form.un} onChange={e=>setForm(f=>({...f,un:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Custo R$</label><input type="number" value={form.custo} onChange={e=>setForm(f=>({...f,custo:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Motivo</label>
              <select value={form.motivo} onChange={e=>setForm(f=>({...f,motivo:e.target.value}))} style={inp}>
                {["Vencido","Erro preparo","Queda","Queimado","Outro"].map(m=><option key={m}>{m}</option>)}
              </select>
            </div>
          </div>
          <button onClick={salvar} style={{background:t.red,color:"#fff",border:"none",borderRadius:8,padding:"10px 22px",cursor:"pointer",fontSize:13,fontWeight:700}}>Salvar</button>
        </Card>
      )}

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <Card>
          <ST color={t.red}>Por Motivo (R$)</ST>
          <ResponsiveContainer width="100%" height={180}>
            <BarChart data={motivoData} layout="vertical">
              <XAxis type="number" stroke={t.muted} tick={{fill:t.muted,fontSize:10}} tickFormatter={v=>"R$"+v.toFixed(0)}/>
              <YAxis type="category" dataKey="motivo" stroke={t.muted} tick={{fill:t.muted,fontSize:11}} width={80}/>
              <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle}/>
              <Bar dataKey="custo" fill={t.red} radius={[0,4,4,0]} name="Custo"/>
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <ST>Histórico</ST>
          {itens.map(d=>(
            <div key={d.id} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:"1px solid "+t.bord}}>
              <div>
                <div style={{color:t.text,fontSize:13,fontWeight:600}}>{d.produto}</div>
                <div style={{color:t.muted,fontSize:11}}>{d.data} · {d.motivo}</div>
              </div>
              <div style={{color:t.red,fontFamily:"monospace",fontSize:13,fontWeight:700}}>{fmt(d.custo)}</div>
            </div>
          ))}
        </Card>
      </div>
    </div>
  );
}

const TABS = [
  {id:"estoque",    label:"📦 Estoque"},
  {id:"producao",   label:"⚙️ Produção"},
  {id:"desperdicio",label:"🗑 Desperdício"},
];

export default function App() {
  const [tab,setTab]=useState("estoque");
  return (
    <div style={{minHeight:"100vh",background:t.bg,color:t.text,fontFamily:"system-ui,sans-serif"}}>
      <AppHeader>
        <div style={{display:"flex",alignItems:"center",gap:6,background:ACCENT_LIGHT,border:"1px solid "+ACCENT,borderRadius:8,padding:"5px 10px"}}>
          <div style={{width:6,height:6,borderRadius:"50%",background:ACCENT}}/>
          <span style={{color:ACCENT_DARK,fontSize:11,fontWeight:600}}>Dados de exemplo</span>
        </div>
      </AppHeader>
      <div style={{padding:"14px 20px 0",overflowX:"auto"}}>
        <div style={{display:"flex",gap:8,minWidth:"max-content"}}>
          {TABS.map(tb=><TabBtn key={tb.id} label={tb.label} active={tab===tb.id} onClick={()=>setTab(tb.id)}/>)}
        </div>
      </div>
      <div style={{padding:"18px 20px"}}>
        {tab==="estoque"     && <Estoque/>}
        {tab==="producao"    && <Producao/>}
        {tab==="desperdicio" && <Desperdicio/>}
      </div>
    </div>
  );
}

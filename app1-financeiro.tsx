import { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from "recharts";

const SUPA_URL = "https://rgsgjqnlvyossuykemjl.supabase.co";
const SUPA_KEY = "sb_publishable_uC6SF60PfuFnZfDDLRo2Wg_m7l2FoBW";

async function sbGet(table, params) {
  const url = SUPA_URL + "/rest/v1/" + table + "?" + (params||"");
  const res = await fetch(url, {headers:{"apikey":SUPA_KEY,"Authorization":"Bearer "+SUPA_KEY}});
  if (!res.ok) throw new Error(table+": "+res.status);
  return res.json();
}
async function sbPost(table, body) {
  const res = await fetch(SUPA_URL+"/rest/v1/"+table, {method:"POST",headers:{"apikey":SUPA_KEY,"Authorization":"Bearer "+SUPA_KEY,"Content-Type":"application/json","Prefer":"return=representation"},body:JSON.stringify(body)});
  if (!res.ok) throw new Error("POST "+table+": "+res.status);
  return res.json();
}
async function sbPatch(table, id, body) {
  const res = await fetch(SUPA_URL+"/rest/v1/"+table+"?id=eq."+id, {method:"PATCH",headers:{"apikey":SUPA_KEY,"Authorization":"Bearer "+SUPA_KEY,"Content-Type":"application/json","Prefer":"return=representation"},body:JSON.stringify(body)});
  if (!res.ok) throw new Error("PATCH "+table+": "+res.status);
  return res.json();
}

const ACCENT       = "#15803d";
const ACCENT_LIGHT = "#dcfce7";
const ACCENT_DARK  = "#14532d";

const t = {
  bg:"#f8f8f4", card:"#ffffff", bord:"#e0e8d8",
  green:"#4a7c3f", greenLight:"#6aab5e",
  text:"#1a1a1a", muted:"#5a6a5a",
  red:"#dc2626", orange:"#ea580c", yellow:"#d97706",
  blue:"#1d4ed8", cyan:"#0891b2", purple:"#7c3aed",
};

const fmt  = v => "R$ " + Number(v||0).toLocaleString("pt-BR",{minimumFractionDigits:2});
const fmtK = v => { const n=Number(v||0); return n>=1000?"R$"+(n/1000).toFixed(1)+"k":"R$"+n.toFixed(0); };
const inp  = {background:t.bg,border:"1px solid "+t.bord,borderRadius:8,padding:"9px 13px",color:t.text,fontSize:13,fontFamily:"inherit",outline:"none",width:"100%",boxSizing:"border-box"};
const lbl  = {color:t.muted,fontSize:11,letterSpacing:1,textTransform:"uppercase",display:"block",marginBottom:6,fontWeight:600};
const TT   = {contentStyle:{background:t.card,border:"1px solid "+t.bord,borderRadius:8,color:t.text}};

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
            <div style={{fontSize:15,color:ACCENT_DARK,fontWeight:700,fontFamily:"Georgia,serif"}}>Financeiro</div>
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
      {[["loja","🏪 Loja"],["matriz","🏭 Matriz"]].map(([id,label])=>(
        <button key={id} onClick={()=>setOp(id)} style={{
          background:op===id?ACCENT:t.card,color:op===id?"#fff":t.muted,
          border:"1px solid "+(op===id?ACCENT:t.bord),borderRadius:8,
          padding:"8px 16px",cursor:"pointer",fontSize:12,fontWeight:op===id?700:400,
        }}>{label}</button>
      ))}
    </div>
  );
}

// ── CAIXA ──────────────────────────────────────────────────────────────────────
const CAIXA_INIT = [
  {dia:"Seg 22/04",dinheiro:420,pix:680,debito:350,credito:290,anotai:180,ifood:520},
  {dia:"Ter 23/04",dinheiro:380,pix:720,debito:410,credito:320,anotai:210,ifood:490},
  {dia:"Qua 24/04",dinheiro:510,pix:840,debito:390,credito:450,anotai:160,ifood:580},
  {dia:"Qui 25/04",dinheiro:460,pix:790,debito:420,credito:380,anotai:195,ifood:510},
  {dia:"Sex 26/04",dinheiro:620,pix:910,debito:480,credito:520,anotai:230,ifood:640},
  {dia:"Sáb 27/04",dinheiro:780,pix:1050,debito:560,credito:610,anotai:280,ifood:720},
];

function Caixa() {
  const [idx,setIdx]=useState(CAIXA_INIT.length-1);
  const dia = CAIXA_INIT[idx];
  const totalGerado = dia.dinheiro+dia.pix+dia.debito+dia.credito+dia.anotai+dia.ifood;
  const totalLiq = dia.dinheiro+dia.pix; // simplificado

  const rows = [
    {label:"💵 Dinheiro",     val:dia.dinheiro, color:ACCENT,   liq:"Mesmo dia"},
    {label:"📱 Pix",          val:dia.pix,      color:ACCENT,   liq:"Mesmo dia"},
    {label:"💳 Débito",       val:dia.debito,   color:t.yellow, liq:"Prox. dia útil"},
    {label:"💳 Crédito",      val:dia.credito,  color:t.yellow, liq:"Prox. dia útil"},
    {label:"📝 Anotai",       val:dia.anotai,   color:t.orange, liq:"Prox. dia útil"},
    {label:"🟠 iFood",        val:dia.ifood,    color:t.orange, liq:"Toda quarta"},
  ];

  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div style={{display:"flex",gap:8,overflowX:"auto",paddingBottom:4}}>
        {CAIXA_INIT.map((d,i)=>(
          <button key={i} onClick={()=>setIdx(i)} style={{background:i===idx?ACCENT:t.card,color:i===idx?"#fff":t.muted,border:"1px solid "+(i===idx?ACCENT:t.bord),borderRadius:8,padding:"6px 14px",cursor:"pointer",fontSize:12,fontWeight:i===idx?700:400,whiteSpace:"nowrap"}}>{d.dia}</button>
        ))}
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:12}}>
        <Kpi label="Venda gerada" value={fmtK(totalGerado)} sub="Total do dia" color={ACCENT} icon="💰"/>
        <Kpi label="Liquidado hoje" value={fmtK(totalLiq)} sub="Dinheiro + Pix" color={t.green} icon="✅"/>
      </div>

      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
          <ST>Formas de Pagamento</ST>
          <button onClick={()=>exportXLSX([{name:"Caixa",data:rows.map(r=>({Forma:r.label,"Valor (R$)":r.val,Liquidacao:r.liq}))}],"caixa.xlsx")} style={{background:ACCENT_LIGHT,color:ACCENT_DARK,border:"1px solid "+ACCENT,borderRadius:7,padding:"5px 12px",cursor:"pointer",fontSize:11,fontWeight:700}}>⬇ Excel</button>
        </div>
        {rows.map(r=>(
          <div key={r.label} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0",borderBottom:"1px solid "+t.bord}}>
            <div>
              <div style={{color:t.text,fontSize:13,fontWeight:600}}>{r.label}</div>
              <div style={{color:t.muted,fontSize:11}}>{r.liq}</div>
            </div>
            <div style={{color:r.color,fontFamily:"monospace",fontSize:15,fontWeight:700}}>{fmt(r.val)}</div>
          </div>
        ))}
        <div style={{display:"flex",justifyContent:"space-between",padding:"12px 0 0",fontWeight:700}}>
          <span style={{color:t.text,fontSize:14}}>Total</span>
          <span style={{color:ACCENT_DARK,fontFamily:"monospace",fontSize:16}}>{fmt(totalGerado)}</span>
        </div>
      </Card>

      <Card>
        <ST>Evolução Semanal</ST>
        <ResponsiveContainer width="100%" height={180}>
          <BarChart data={CAIXA_INIT.map(d=>({dia:d.dia.split(" ")[0],total:d.dinheiro+d.pix+d.debito+d.credito+d.anotai+d.ifood}))}>
            <XAxis dataKey="dia" stroke={t.muted} tick={{fill:t.muted,fontSize:11}}/>
            <YAxis stroke={t.muted} tick={{fill:t.muted,fontSize:11}} tickFormatter={v=>"R$"+Math.round(v/1000)+"k"}/>
            <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle}/>
            <Bar dataKey="total" fill={ACCENT} radius={[4,4,0,0]} name="Total"/>
          </BarChart>
        </ResponsiveContainer>
      </Card>
    </div>
  );
}

// ── CUSTOS ─────────────────────────────────────────────────────────────────────
const CUSTOS_DATA = {
  loja:   [{cat:"Aluguel",val:2800},{cat:"Energia",val:420},{cat:"Internet",val:120},{cat:"Embalagens",val:380},{cat:"Taxa iFood",val:640},{cat:"Pró-labore",val:1200}],
  matriz: [{cat:"Aluguel",val:1800},{cat:"Energia",val:680},{cat:"Gás",val:290},{cat:"Embalagens",val:450},{cat:"Manutenção",val:180},{cat:"Pró-labore",val:1500}],
};

function Custos() {
  const [op,setOp]=useState("loja");
  const lista = CUSTOS_DATA[op];
  const total = lista.reduce((s,c)=>s+c.val,0);
  const COLORS = [ACCENT,"#0891b2","#7c3aed","#d97706","#dc2626","#ea580c"];

  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
        <OpBtn op={op} setOp={setOp}/>
        <button onClick={()=>exportXLSX([{name:"Custos",data:lista.map(c=>({Categoria:c.cat,Operação:op,"Valor (R$)":c.val,"% Total":(c.val/total*100).toFixed(1)+"%"}))}],"custos.xlsx")} style={{background:ACCENT_LIGHT,color:ACCENT_DARK,border:"1px solid "+ACCENT,borderRadius:7,padding:"6px 12px",cursor:"pointer",fontSize:11,fontWeight:700}}>⬇ Excel</button>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:12}}>
        <Kpi label="Total custos" value={fmtK(total)} sub={op==="loja"?"Loja":"Matriz"} color={ACCENT} icon="📊"/>
        <Kpi label="Itens" value={lista.length} sub="Categorias" color={t.muted} icon="📋"/>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <Card>
          <ST>Por Categoria</ST>
          <ResponsiveContainer width="100%" height={180}>
            <PieChart>
              <Pie data={lista} dataKey="val" nameKey="cat" cx="50%" cy="50%" outerRadius={70} innerRadius={36}>
                {lista.map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]}/>)}
              </Pie>
              <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle}/>
            </PieChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <ST>Detalhamento</ST>
          {lista.map((c,i)=>(
            <div key={c.cat} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:"1px solid "+t.bord}}>
              <div style={{display:"flex",alignItems:"center",gap:6}}>
                <div style={{width:8,height:8,borderRadius:"50%",background:COLORS[i%COLORS.length]}}/>
                <span style={{color:t.text,fontSize:13}}>{c.cat}</span>
              </div>
              <span style={{color:ACCENT_DARK,fontFamily:"monospace",fontSize:13,fontWeight:700}}>{fmt(c.val)}</span>
            </div>
          ))}
          <div style={{display:"flex",justifyContent:"space-between",paddingTop:10,fontWeight:700}}>
            <span style={{color:t.text}}>Total</span>
            <span style={{color:ACCENT_DARK,fontFamily:"monospace",fontSize:15}}>{fmt(total)}</span>
          </div>
        </Card>
      </div>
    </div>
  );
}

// ── EMPRÉSTIMOS ─────────────────────────────────────────────────────────────────
const EMP_INIT = [
  {id:"e1",op:"Loja",  credor:"Maria (sócia)",  valor:3000,data:"03/03/2025",motivo:"Fechamento negativo",
   parcelas:[{num:1,valor:600,venc:"03/04",pago:true,dataPago:"02/04"},{num:2,valor:600,venc:"03/05",pago:true,dataPago:"01/05"},{num:3,valor:600,venc:"03/06",pago:false},{num:4,valor:600,venc:"03/07",pago:false},{num:5,valor:600,venc:"03/08",pago:false}]},
  {id:"e2",op:"Matriz",credor:"Banco Família",   valor:5000,data:"10/01/2025",motivo:"Reforço de caixa",
   parcelas:[{num:1,valor:1000,venc:"10/02",pago:true,dataPago:"09/02"},{num:2,valor:1000,venc:"10/03",pago:true,dataPago:"08/03"},{num:3,valor:1000,venc:"10/04",pago:true,dataPago:"09/04"},{num:4,valor:1000,venc:"10/05",pago:false},{num:5,valor:1000,venc:"10/06",pago:false}]},
];

function Emprestimos() {
  const [emps,setEmps]=useState(EMP_INIT);
  const [exp,setExp]=useState(null);
  const [showForm,setShowForm]=useState(false);
  const [form,setForm]=useState({op:"Loja",credor:"",valor:"",parcelas:"",motivo:""});

  const totalAberto = emps.reduce((s,e)=>s+e.parcelas.filter(p=>!p.pago).reduce((a,p)=>a+p.valor,0),0);
  const totalPago   = emps.reduce((s,e)=>s+e.parcelas.filter(p=>p.pago).reduce((a,p)=>a+p.valor,0),0);

  const status = e => {
    const pg=e.parcelas.filter(p=>p.pago).length;
    if(pg===e.parcelas.length) return ["Quitado",ACCENT,"#f0fdf4"];
    if(pg===0) return ["Em aberto",t.red,"#fef2f2"];
    return [pg+"/"+e.parcelas.length,t.yellow,"#fffbeb"];
  };

  const marcarPago=(eId,pNum)=>setEmps(prev=>prev.map(e=>e.id!==eId?e:{...e,parcelas:e.parcelas.map(p=>p.num!==pNum?p:{...p,pago:true,dataPago:new Date().toLocaleDateString("pt-BR")})}));

  const adicionar=()=>{
    if(!form.credor||!form.valor||!form.parcelas) return;
    const val=parseFloat(form.valor),np=parseInt(form.parcelas),vp=+(val/np).toFixed(2);
    const hoje=new Date();
    const parc=Array.from({length:np},(_,i)=>{const d=new Date(hoje);d.setMonth(d.getMonth()+i+1);return{num:i+1,valor:vp,venc:d.toLocaleDateString("pt-BR"),pago:false,dataPago:null};});
    setEmps(prev=>[...prev,{id:"e"+Date.now(),op:form.op,credor:form.credor,valor:val,data:hoje.toLocaleDateString("pt-BR"),motivo:form.motivo||"Registrado manualmente",parcelas:parc}]);
    setForm({op:"Loja",credor:"",valor:"",parcelas:"",motivo:""});
    setShowForm(false);
  };

  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
        <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:12,flex:1}}>
          <Kpi label="Total em aberto" value={fmt(totalAberto)} color={t.red} icon="🏦"/>
          <Kpi label="Já devolvido"    value={fmt(totalPago)}   color={ACCENT} icon="✅"/>
        </div>
      </div>

      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <button onClick={()=>exportXLSX([{name:"Emprestimos",data:emps.map(e=>({Credor:e.credor,Operacao:e.op,"Valor Total":e.valor,Data:e.data,"Saldo Aberto":e.parcelas.filter(p=>!p.pago).reduce((a,p)=>a+p.valor,0)}))}],"emprestimos.xlsx")} style={{background:ACCENT_LIGHT,color:ACCENT_DARK,border:"1px solid "+ACCENT,borderRadius:7,padding:"6px 12px",cursor:"pointer",fontSize:11,fontWeight:700}}>⬇ Excel</button>
        <button onClick={()=>setShowForm(!showForm)} style={{background:showForm?t.card:t.red,color:showForm?t.muted:"#fff",border:"1px solid "+(showForm?t.bord:t.red),borderRadius:8,padding:"8px 16px",cursor:"pointer",fontSize:12,fontWeight:700}}>{showForm?"✕ Cancelar":"+ Novo Empréstimo"}</button>
      </div>

      {showForm&&(
        <Card style={{border:"1px solid "+t.red}}>
          <ST color={t.red}>+ Novo Empréstimo</ST>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
            <div><label style={lbl}>Operação</label>
              <select value={form.op} onChange={e=>setForm(f=>({...f,op:e.target.value}))} style={inp}>
                <option>Loja</option><option>Matriz</option>
              </select>
            </div>
            <div><label style={lbl}>Credor</label><input value={form.credor} onChange={e=>setForm(f=>({...f,credor:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Valor total (R$)</label><input type="number" value={form.valor} onChange={e=>setForm(f=>({...f,valor:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Nº de parcelas</label><input type="number" value={form.parcelas} onChange={e=>setForm(f=>({...f,parcelas:e.target.value}))} style={inp}/></div>
            <div style={{gridColumn:"1/-1"}}><label style={lbl}>Motivo</label><input value={form.motivo} onChange={e=>setForm(f=>({...f,motivo:e.target.value}))} style={inp}/></div>
          </div>
          <button onClick={adicionar} style={{background:t.red,color:"#fff",border:"none",borderRadius:8,padding:"10px 22px",cursor:"pointer",fontSize:13,fontWeight:700}}>Registrar</button>
        </Card>
      )}

      {emps.map(e=>{
        const [stLabel,stColor,stBg]=status(e);
        const aberto=e.parcelas.filter(p=>!p.pago).reduce((a,p)=>a+p.valor,0);
        return (
          <Card key={e.id}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}>
              <div>
                <div style={{color:t.text,fontSize:15,fontWeight:700}}>{e.credor}</div>
                <div style={{color:t.muted,fontSize:12}}>{e.op} · {e.data} · {e.motivo}</div>
              </div>
              <div style={{textAlign:"right"}}>
                <div style={{background:stBg,color:stColor,borderRadius:8,padding:"4px 10px",fontSize:12,fontWeight:700,marginBottom:4}}>{stLabel}</div>
                <div style={{color:t.red,fontFamily:"monospace",fontSize:14,fontWeight:700}}>{fmt(aberto)} aberto</div>
              </div>
            </div>
            <button onClick={()=>setExp(exp===e.id?null:e.id)} style={{background:t.bg,color:t.muted,border:"1px solid "+t.bord,borderRadius:7,padding:"5px 12px",cursor:"pointer",fontSize:12,marginBottom:exp===e.id?10:0}}>{exp===e.id?"▲ Fechar":"▼ Ver parcelas"}</button>
            {exp===e.id&&(
              <table style={{width:"100%",borderCollapse:"collapse",marginTop:8}}>
                <thead><tr style={{background:t.bg}}>{["Parc","Valor","Venc","Status","Ação"].map(h=><th key={h} style={{padding:"7px 10px",textAlign:"left",color:t.muted,fontSize:11,letterSpacing:1,textTransform:"uppercase"}}>{h}</th>)}</tr></thead>
                <tbody>
                  {e.parcelas.map(p=>(
                    <tr key={p.num} style={{borderTop:"1px solid "+t.bord}}>
                      <td style={{padding:"8px 10px",color:t.muted,fontFamily:"monospace",fontSize:12}}>{p.num}</td>
                      <td style={{padding:"8px 10px",color:ACCENT_DARK,fontFamily:"monospace",fontSize:13,fontWeight:700}}>{fmt(p.valor)}</td>
                      <td style={{padding:"8px 10px",color:t.muted,fontSize:12}}>{p.venc}</td>
                      <td style={{padding:"8px 10px"}}>
                        <span style={{background:p.pago?"#f0fdf4":"#fef2f2",color:p.pago?ACCENT:t.red,borderRadius:5,padding:"2px 8px",fontSize:11,fontWeight:600}}>{p.pago?"Pago":"Pendente"}</span>
                      </td>
                      <td style={{padding:"8px 10px"}}>
                        {!p.pago&&<button onClick={()=>marcarPago(e.id,p.num)} style={{background:ACCENT,color:"#fff",border:"none",borderRadius:6,padding:"4px 10px",cursor:"pointer",fontSize:11,fontWeight:600}}>Marcar pago</button>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </Card>
        );
      })}
    </div>
  );
}

// ── RESERVAS ────────────────────────────────────────────────────────────────────
const RES_DEST = [
  {id:"d1",label:"Cartão Tiago",  icon:"💳",meta:1800,cor:ACCENT},
  {id:"d2",label:"Empréstimos",   icon:"🏦",meta:1200,cor:t.red},
  {id:"d3",label:"Aluguel",       icon:"🏠",meta:2100,cor:t.yellow},
  {id:"d4",label:"Outras Contas", icon:"📋",meta:600, cor:t.orange},
];
const RES_LANC = {
  loja:  [{id:1,dest:"d1",data:"20/04",valor:650,obs:"Pagamento mensal"},{id:2,dest:"d2",data:"22/04",valor:350,obs:"Parcela Maria"}],
  matriz:[{id:1,dest:"d2",data:"21/04",valor:200,obs:"Aporte família"},{id:2,dest:"d3",data:"23/04",valor:550,obs:"Reserva aluguel"}],
};

function Reservas() {
  const [op,setOp]=useState("loja");
  const [dests]=useState(RES_DEST);
  const [lancs,setLancs]=useState(RES_LANC);
  const [showForm,setShowForm]=useState(false);
  const [form,setForm]=useState({dest:"d1",valor:"",obs:""});

  const lista=lancs[op];
  const totalRes=lista.reduce((s,l)=>s+l.valor,0);

  const adicionar=()=>{
    if(!form.valor) return;
    setLancs(p=>({...p,[op]:[...p[op],{id:Date.now(),dest:form.dest,data:new Date().toLocaleDateString("pt-BR"),valor:parseFloat(form.valor),obs:form.obs}]}));
    setForm(f=>({...f,valor:"",obs:""}));
    setShowForm(false);
  };

  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
        <OpBtn op={op} setOp={setOp}/>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>exportXLSX([{name:"Reservas",data:lista.map(l=>{const d=dests.find(x=>x.id===l.dest);return{Data:l.data,Destino:d?.label||l.dest,Operação:op,"Valor (R$)":l.valor,Obs:l.obs||""};})}],"reservas.xlsx")} style={{background:ACCENT_LIGHT,color:ACCENT_DARK,border:"1px solid "+ACCENT,borderRadius:7,padding:"6px 12px",cursor:"pointer",fontSize:11,fontWeight:700}}>⬇ Excel</button>
          <button onClick={()=>setShowForm(!showForm)} style={{background:showForm?t.card:ACCENT,color:showForm?t.muted:"#fff",border:"1px solid "+(showForm?t.bord:ACCENT),borderRadius:8,padding:"8px 16px",cursor:"pointer",fontSize:12,fontWeight:700}}>{showForm?"✕ Cancelar":"+ Lançar"}</button>
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:12}}>
        <Kpi label="Total reservado" value={fmt(totalRes)} color={ACCENT} icon="🎯"/>
        <Kpi label="Meta total" value={fmt(dests.reduce((s,d)=>s+d.meta,0))} color={t.muted} icon="📊"/>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        {dests.map(d=>{
          const totalD=lista.filter(l=>l.dest===d.id).reduce((s,l)=>s+l.valor,0);
          const pct=Math.min((totalD/d.meta)*100,100);
          return (
            <Card key={d.id}>
              <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:10}}>
                <span style={{fontSize:20}}>{d.icon}</span>
                <div>
                  <div style={{color:t.text,fontSize:13,fontWeight:700}}>{d.label}</div>
                  <div style={{color:t.muted,fontSize:11}}>Meta: {fmt(d.meta)}</div>
                </div>
              </div>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
                <span style={{color:d.cor,fontFamily:"monospace",fontSize:15,fontWeight:700}}>{fmt(totalD)}</span>
                <span style={{color:t.muted,fontSize:12}}>{pct.toFixed(0)}%</span>
              </div>
              <div style={{background:t.bord,borderRadius:4,height:6,overflow:"hidden"}}>
                <div style={{width:pct+"%",background:d.cor,height:"100%",borderRadius:4}}/>
              </div>
            </Card>
          );
        })}
      </div>

      {showForm&&(
        <Card style={{border:"1px solid "+ACCENT}}>
          <ST color={ACCENT}>+ Lançar Reserva</ST>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 2fr",gap:10,marginBottom:12}}>
            <div><label style={lbl}>Destino</label>
              <select value={form.dest} onChange={e=>setForm(f=>({...f,dest:e.target.value}))} style={inp}>
                {dests.map(d=><option key={d.id} value={d.id}>{d.icon} {d.label}</option>)}
              </select>
            </div>
            <div><label style={lbl}>Valor (R$)</label><input type="number" value={form.valor} onChange={e=>setForm(f=>({...f,valor:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Observação</label><input value={form.obs} onChange={e=>setForm(f=>({...f,obs:e.target.value}))} style={inp}/></div>
          </div>
          <button onClick={adicionar} style={{background:ACCENT,color:"#fff",border:"none",borderRadius:8,padding:"10px 22px",cursor:"pointer",fontSize:13,fontWeight:700}}>Confirmar</button>
        </Card>
      )}

      <Card>
        <ST>Histórico — {op==="loja"?"Loja":"Matriz"}</ST>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr style={{background:t.bg}}>{["Data","Destino","Valor","Obs"].map(h=><th key={h} style={{padding:"9px 12px",textAlign:"left",color:t.muted,fontSize:11,letterSpacing:1,textTransform:"uppercase"}}>{h}</th>)}</tr></thead>
          <tbody>
            {[...lista].reverse().map((l,i)=>{
              const d=dests.find(x=>x.id===l.dest);
              return (
                <tr key={i} style={{borderTop:"1px solid "+t.bord}}>
                  <td style={{padding:"9px 12px",color:t.muted,fontFamily:"monospace",fontSize:12}}>{l.data}</td>
                  <td style={{padding:"9px 12px",color:t.text,fontSize:13}}>{d?.icon} {d?.label}</td>
                  <td style={{padding:"9px 12px",color:ACCENT_DARK,fontFamily:"monospace",fontSize:13,fontWeight:700}}>{fmt(l.valor)}</td>
                  <td style={{padding:"9px 12px",color:t.muted,fontSize:12}}>{l.obs||"—"}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

const TABS = [
  {id:"caixa",      label:"💵 Caixa Diário"},
  {id:"custos",     label:"📊 Custos"},
  {id:"emprestimos",label:"🏦 Empréstimos"},
  {id:"reservas",   label:"🎯 Reservas"},
];

export default function App() {
  const [tab,setTab]=useState("caixa");
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
        {tab==="caixa"       && <Caixa/>}
        {tab==="custos"      && <Custos/>}
        {tab==="emprestimos" && <Emprestimos/>}
        {tab==="reservas"    && <Reservas/>}
      </div>
    </div>
  );
}

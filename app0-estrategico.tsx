import { useState, useEffect } from "react";
import { BarChart, Bar, LineChart, Line, AreaChart, Area, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, RadarChart, Radar, PolarGrid, PolarAngleAxis } from "recharts";
import * as XLSX from "xlsx";

// ── SUPABASE ──────────────────────────────────────────────────────────────────
const SUPA_URL = "https://rgsgjqnlvyossuykemjl.supabase.co";
const SUPA_KEY = "sb_publishable_uC6SF60PfuFnZfDDLRo2Wg_m7l2FoBW";

async function sbGet(table, params = "") {
  const res = await fetch(`${SUPA_URL}/rest/v1/${table}?${params}`, {
    headers: { "apikey": SUPA_KEY, "Authorization": `Bearer ${SUPA_KEY}` },
  });
  if (!res.ok) throw new Error(`${table}: ${res.status}`);
  return res.json();
}

function exportXLSX(sheets, filename) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, data }) => {
    if (!data || !data.length) return;
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = Object.keys(data[0]).map(k => ({ wch: Math.max(k.length + 2, 16) }));
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  XLSX.writeFile(wb, filename);
}

// ── TEMA ──────────────────────────────────────────────────────────────────────
const t = {
  bg:"#f8f7f2",    card:"#ffffff",   card2:"#f0ede4",
  bord:"#e2ddd4",  gold:"#b45309",   goldDim:"#fef3c7",
  blue:"#1d4ed8",  cyan:"#0891b2",   green:"#15803d",
  lime:"#4d7c0f",  purple:"#7c3aed", pink:"#be185d",
  yellow:"#b45309",red:"#dc2626",    orange:"#ea580c",
  text:"#1a1512",  muted:"#6b5f4e",
};

const fmt  = v => `R$ ${Number(v||0).toLocaleString("pt-BR",{minimumFractionDigits:2})}`;
const fmtK = v => { const n=Number(v||0); return n>=1000?`R$${(n/1000).toFixed(1)}k`:`R$${n.toFixed(0)}`; };
const pct  = (v,t) => t>0?((v/t)*100).toFixed(1)+"%":"0%";
const TT   = { contentStyle:{background:"#ffffff",border:"1px solid #e2ddd4",borderRadius:8,color:t.text} };

// ── DADOS MENSAIS FIXOS (histórico) ──────────────────────────────────────────
const HISTORICO = [
  { mes:"Jan", receita_loja:15000, receita_matriz:10000, custos_loja:8200, custos_matriz:8400, desperdicio:420 },
  { mes:"Fev", receita_loja:16200, receita_matriz:10800, custos_loja:8600, custos_matriz:8900, desperdicio:380 },
  { mes:"Mar", receita_loja:17500, receita_matriz:11200, custos_loja:9000, custos_matriz:9200, desperdicio:510 },
  { mes:"Abr", receita_loja:18400, receita_matriz:12600, custos_loja:9660, custos_matriz:9640, desperdicio:290 },
];

// ── BASE COMPONENTS ───────────────────────────────────────────────────────────
function Kpi({ label, value, sub, color = t.gold, icon, trend }) {
  return (
    <div style={{ background:t.card, border:`1px solid ${t.bord}`, borderRadius:14, padding:"20px 22px", position:"relative", overflow:"hidden" }}>
      <div style={{ position:"absolute", top:14, right:16, fontSize:20, opacity:0.25 }}>{icon}</div>
      <div style={{ color:t.muted, fontSize:11, letterSpacing:1.5, textTransform:"uppercase", marginBottom:8 }}>{label}</div>
      <div style={{ color, fontSize:28, fontFamily:"'Cormorant Garamond',serif", lineHeight:1, marginBottom:4 }}>{value}</div>
      {sub && <div style={{ color:t.muted, fontSize:12 }}>{sub}</div>}
      {trend && (
        <div style={{ marginTop:8, display:"flex", alignItems:"center", gap:4 }}>
          <span style={{ color:trend>0?t.green:t.red, fontSize:12 }}>{trend>0?"▲":"▼"} {Math.abs(trend)}%</span>
          <span style={{ color:t.muted, fontSize:11 }}>vs mês anterior</span>
        </div>
      )}
    </div>
  );
}

function Card({ children, style={} }) {
  return <div style={{ background:t.card, border:`1px solid ${t.bord}`, borderRadius:14, padding:20, ...style }}>{children}</div>;
}

function ST({ children, color=t.muted }) {
  return <div style={{ color, fontSize:11, letterSpacing:1.5, textTransform:"uppercase", marginBottom:14, fontWeight:700 }}>{children}</div>;
}

function TabBtn({ label, active, onClick }) {
  return (
    <button onClick={onClick} style={{
      background:active?t.gold:"transparent", color:active?t.bg:t.muted,
      border:`1px solid ${active?t.gold:t.bord}`, borderRadius:8,
      padding:"8px 20px", cursor:"pointer", fontSize:12, fontWeight:active?700:400, transition:"all 0.15s",
    }}>{label}</button>
  );
}

function Loading() {
  return (
    <div style={{ display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", padding:60, gap:16 }}>
      <div style={{ width:40, height:40, border:`3px solid #e2ddd4`, borderTop:`3px solid ${t.gold}`, borderRadius:"50%", animation:"spin 1s linear infinite" }} />
      <div style={{ color:t.muted, fontSize:13 }}>Carregando dados do Supabase...</div>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}

function Alerta({ tipo, msg }) {
  const cores = { info:[t.blue,"#eff6ff"], warn:[t.yellow,"#fffbeb"], crit:[t.red,"#fef2f2"], ok:[t.green,"#f0fdf4"] };
  const [cor, bg] = cores[tipo] || cores.info;
  return (
    <div style={{ background:bg, border:`1px solid ${cor}`, borderRadius:10, padding:"12px 16px", display:"flex", alignItems:"center", gap:10 }}>
      <span style={{ fontSize:16 }}>{tipo==="crit"?"🔴":tipo==="warn"?"🟡":tipo==="ok"?"✅":"ℹ️"}</span>
      <span style={{ color:cor, fontSize:13 }}>{msg}</span>
    </div>
  );
}

// ── VISÃO GERAL ───────────────────────────────────────────────────────────────
function Geral({ dados }) {
  const { vendas_loja, pedidos_matriz, colaboradores, consumo, banco, desperdicio, emprestimos, parcelas, reservas_lanc, reservas_dest } = dados;

  const receitaLoja   = vendas_loja.reduce((s,v)=>s+parseFloat(v.valor||0),0);
  const receitaMatriz = pedidos_matriz.reduce((s,p)=>s+parseFloat(p.total||0),0);
  const receitaTotal  = receitaLoja + receitaMatriz;

  const custosLoja   = HISTORICO[HISTORICO.length-1].custos_loja;
  const custosMatriz = HISTORICO[HISTORICO.length-1].custos_matriz;
  const custosTotal  = custosLoja + custosMatriz;

  const lucroLoja   = receitaLoja - custosLoja;
  const lucroMatriz = receitaMatriz - custosMatriz;
  const lucroTotal  = lucroLoja + lucroMatriz;

  const folhaTotal   = colaboradores.reduce((s,c)=>s+parseFloat(c.salario||0)+parseFloat(c.vt||0)+parseFloat(c.vr||0)+parseFloat(c.plano||0),0);
  const consumoTotal = consumo.reduce((s,x)=>s+parseFloat(x.custo||0),0);
  const despTotal    = desperdicio.reduce((s,d)=>s+parseFloat(d.custo||0),0);
  const empAberto    = emprestimos.reduce((s,e)=>s+parcelas.filter(p=>p.emprestimo_id===e.id&&!p.pago).reduce((a,p)=>a+parseFloat(p.valor||0),0),0);
  const reservasMeta = reservas_dest.reduce((s,d)=>s+parseFloat(d.meta||0),0);
  const reservasFeito= reservas_lanc.reduce((s,l)=>s+parseFloat(l.valor||0),0);

  // Alertas automáticos
  const alertas = [];
  if (lucroTotal < 0) alertas.push({ tipo:"crit", msg:`Resultado negativo este mês: ${fmt(lucroTotal)}. Considere registrar um empréstimo.` });
  if (despTotal > 500) alertas.push({ tipo:"warn", msg:`Desperdício acima de R$500 no período: ${fmt(despTotal)}. Verifique os motivos no App 2.` });
  if (empAberto > 5000) alertas.push({ tipo:"warn", msg:`Saldo de empréstimos em aberto: ${fmt(empAberto)}. Confira as parcelas no App 1.` });
  if (reservasFeito < reservasMeta * 0.5) alertas.push({ tipo:"warn", msg:`Reservas abaixo de 50% da meta: ${fmt(reservasFeito)} de ${fmt(reservasMeta)}.` });
  if (lucroTotal > 0 && despTotal < 300) alertas.push({ tipo:"ok", msg:`Operação saudável. Lucro de ${fmt(lucroTotal)} com desperdício controlado.` });

  const pizzaOp = [
    { name:"Loja", value:receitaLoja, cor:t.purple },
    { name:"Matriz", value:receitaMatriz, cor:t.cyan },
  ];

  const barMensal = HISTORICO.map(h => ({
    mes:h.mes,
    "Receita":h.receita_loja+h.receita_matriz,
    "Custos":h.custos_loja+h.custos_matriz,
    "Lucro":(h.receita_loja+h.receita_matriz)-(h.custos_loja+h.custos_matriz),
  }));

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:20 }}>
      {/* Alertas */}
      {alertas.length>0 && (
        <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
          {alertas.map((a,i) => <Alerta key={i} tipo={a.tipo} msg={a.msg} />)}
        </div>
      )}

      {/* KPIs principais */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(2,1fr)", gap:12 }}>
        <Kpi label="Receita Total" value={fmtK(receitaTotal)} sub="Loja + Matriz" color={t.gold} icon="💰" trend={8} />
        <Kpi label="Lucro Total"   value={fmtK(lucroTotal)}   sub={pct(lucroTotal,receitaTotal)+" de margem"} color={lucroTotal>0?t.green:t.red} icon="📈" />
        <Kpi label="Custos Totais" value={fmtK(custosTotal)}  sub={pct(custosTotal,receitaTotal)+" da receita"} color={t.yellow} icon="📊" />
        <Kpi label="Folha de Pagamento" value={fmtK(folhaTotal)} sub={`${colaboradores.length} colaboradores`} color={t.blue} icon="👥" />
      </div>

      <div style={{ display:"grid", gridTemplateColumns:"repeat(2,1fr)", gap:12 }}>
        <Kpi label="Lucro Loja"     value={fmtK(lucroLoja)}   sub={pct(lucroLoja,receitaLoja)+" margem"}   color={t.purple} icon="🏪" />
        <Kpi label="Lucro Matriz"   value={fmtK(lucroMatriz)} sub={pct(lucroMatriz,receitaMatriz)+" margem"} color={t.cyan}   icon="🏭" />
        <Kpi label="Desperdício"    value={fmtK(despTotal)}   sub={`${desperdicio.length} ocorrências`}    color={t.red}    icon="🗑" />
        <Kpi label="Empréstimos"    value={fmtK(empAberto)}   sub="Saldo em aberto"                        color={t.orange} icon="🏦" />
      </div>

      {/* Gráfico evolução + pizza */}
      <div style={{ display:"grid", gridTemplateColumns:"2fr 1fr", gap:16 }}>
        <Card>
          <ST>Receita · Custos · Lucro — Mensal</ST>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={barMensal} barCategoryGap="25%">
              <XAxis dataKey="mes" stroke={t.muted} tick={{ fill:t.muted, fontSize:11 }} />
              <YAxis stroke={t.muted} tick={{ fill:t.muted, fontSize:11 }} tickFormatter={v=>`R$${(v/1000).toFixed(0)}k`} />
              <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle} />
              <Bar dataKey="Receita" fill={t.gold}   radius={[4,4,0,0]} />
              <Bar dataKey="Custos"  fill={t.red}    radius={[4,4,0,0]} />
              <Bar dataKey="Lucro"   fill={t.green}  radius={[4,4,0,0]} />
            </BarChart>
          </ResponsiveContainer>
          <div style={{ display:"flex", gap:16, marginTop:8 }}>
            {[["Receita",t.gold],["Custos",t.red],["Lucro",t.green]].map(([l,c])=>(
              <div key={l} style={{ display:"flex", alignItems:"center", gap:6 }}>
                <div style={{ width:10,height:10,borderRadius:2,background:c }} />
                <span style={{ color:t.muted, fontSize:12 }}>{l}</span>
              </div>
            ))}
          </div>
        </Card>

        <Card>
          <ST>Receita por Operação</ST>
          <ResponsiveContainer width="100%" height={160}>
            <PieChart>
              <Pie data={pizzaOp} dataKey="value" cx="50%" cy="50%" outerRadius={70} innerRadius={38}>
                {pizzaOp.map((d,i)=><Cell key={i} fill={d.cor} />)}
              </Pie>
              <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle} />
            </PieChart>
          </ResponsiveContainer>
          {pizzaOp.map(d=>(
            <div key={d.name} style={{ display:"flex", justifyContent:"space-between", padding:"5px 0" }}>
              <div style={{ display:"flex", alignItems:"center", gap:7 }}>
                <div style={{ width:8,height:8,borderRadius:"50%",background:d.cor }} />
                <span style={{ color:t.muted, fontSize:13 }}>{d.name}</span>
              </div>
              <span style={{ color:d.cor, fontFamily:"monospace", fontSize:13 }}>{fmt(d.value)}</span>
            </div>
          ))}
          <div style={{ marginTop:10, borderTop:`1px solid ${t.bord}`, paddingTop:8, display:"flex", justifyContent:"space-between" }}>
            <span style={{ color:t.muted, fontSize:12 }}>Total</span>
            <span style={{ color:t.gold, fontFamily:"monospace", fontSize:14, fontWeight:700 }}>{fmt(receitaTotal)}</span>
          </div>
        </Card>
      </div>

      {/* Linha lucro acumulado */}
      <Card>
        <ST color={t.gold}>Evolução do Lucro — Últimos 4 Meses</ST>
        <ResponsiveContainer width="100%" height={180}>
          <AreaChart data={HISTORICO.map(h=>({ mes:h.mes, lucro:(h.receita_loja+h.receita_matriz)-(h.custos_loja+h.custos_matriz), loja:h.receita_loja-h.custos_loja, matriz:h.receita_matriz-h.custos_matriz }))}>
            <defs>
              <linearGradient id="gradGold" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%"  stopColor="#b45309" stopOpacity={0.15} />
                <stop offset="95%" stopColor="#b45309" stopOpacity={0}   />
              </linearGradient>
            </defs>
            <XAxis dataKey="mes" stroke={t.muted} tick={{ fill:t.muted, fontSize:11 }} />
            <YAxis stroke={t.muted} tick={{ fill:t.muted, fontSize:11 }} tickFormatter={v=>`R$${(v/1000).toFixed(1)}k`} />
            <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle} />
            <Area type="monotone" dataKey="lucro" stroke={t.gold} fill="url(#gradGold)" strokeWidth={2.5} name="Lucro Total" />
            <Line type="monotone" dataKey="loja"   stroke={t.purple} strokeWidth={1.5} dot={false} name="Loja" />
            <Line type="monotone" dataKey="matriz" stroke={t.cyan}   strokeWidth={1.5} dot={false} name="Matriz" />
          </AreaChart>
        </ResponsiveContainer>
      </Card>
    </div>
  );
}

// ── CRUZAMENTO FINANCEIRO ─────────────────────────────────────────────────────
function Financeiro({ dados }) {
  const { vendas_loja, pedidos_matriz, desperdicio, emprestimos, parcelas, consumo, colaboradores } = dados;

  const receitaLoja   = vendas_loja.reduce((s,v)=>s+parseFloat(v.valor||0),0);
  const receitaMatriz = pedidos_matriz.reduce((s,p)=>s+parseFloat(p.total||0),0);
  const folhaLoja     = colaboradores.filter(c=>c.operacao==="loja").reduce((s,c)=>s+parseFloat(c.salario||0)+parseFloat(c.vt||0)+parseFloat(c.vr||0)+parseFloat(c.plano||0),0);
  const folhaMatriz   = colaboradores.filter(c=>c.operacao==="matriz").reduce((s,c)=>s+parseFloat(c.salario||0)+parseFloat(c.vt||0)+parseFloat(c.vr||0)+parseFloat(c.plano||0),0);
  const despLoja      = desperdicio.filter(d=>d.operacao==="loja").reduce((s,d)=>s+parseFloat(d.custo||0),0);
  const despMatriz    = desperdicio.filter(d=>d.operacao==="matriz").reduce((s,d)=>s+parseFloat(d.custo||0),0);
  const consLoja      = consumo.filter(c=>c.operacao==="loja").reduce((s,c)=>s+parseFloat(c.custo||0),0);
  const consMatriz    = consumo.filter(c=>c.operacao==="matriz").reduce((s,c)=>s+parseFloat(c.custo||0),0);
  const custosFixosLoja   = HISTORICO[HISTORICO.length-1].custos_loja;
  const custosFixosMatriz = HISTORICO[HISTORICO.length-1].custos_matriz;
  const lucroLoja   = receitaLoja - custosFixosLoja - despLoja - consLoja;
  const lucroMatriz = receitaMatriz - custosFixosMatriz - despMatriz - consMatriz;

  const dre = [
    { item:"Receita",          loja:receitaLoja,        matriz:receitaMatriz,        tot:receitaLoja+receitaMatriz,        tipo:"receita" },
    { item:"(−) Custos Fixos", loja:-custosFixosLoja,   matriz:-custosFixosMatriz,   tot:-(custosFixosLoja+custosFixosMatriz), tipo:"custo" },
    { item:"(−) Folha",        loja:-folhaLoja,         matriz:-folhaMatriz,         tot:-(folhaLoja+folhaMatriz),         tipo:"custo" },
    { item:"(−) Desperdício",  loja:-despLoja,          matriz:-despMatriz,          tot:-(despLoja+despMatriz),           tipo:"custo" },
    { item:"(−) Consumo Colab",loja:-consLoja,          matriz:-consMatriz,          tot:-(consLoja+consMatriz),           tipo:"custo" },
    { item:"= Lucro Líquido",  loja:lucroLoja,          matriz:lucroMatriz,          tot:lucroLoja+lucroMatriz,            tipo:"lucro" },
  ];

  const empAberto = emprestimos.reduce((s,e)=>s+parcelas.filter(p=>p.emprestimo_id===e.id&&!p.pago).reduce((a,p)=>a+parseFloat(p.valor||0),0),0);

  const radarData = [
    { cat:"Receita",   loja:Math.round((receitaLoja/20000)*100),   matriz:Math.round((receitaMatriz/15000)*100) },
    { cat:"Margem",    loja:Math.round((lucroLoja/receitaLoja)*100||0), matriz:Math.round((lucroMatriz/receitaMatriz)*100||0) },
    { cat:"Eficiência",loja:Math.round(((receitaLoja-custosFixosLoja)/receitaLoja)*100||0), matriz:Math.round(((receitaMatriz-custosFixosMatriz)/receitaMatriz)*100||0) },
    { cat:"Desp. Zero",loja:Math.max(0,100-Math.round((despLoja/receitaLoja)*1000||0)), matriz:Math.max(0,100-Math.round((despMatriz/receitaMatriz)*1000||0)) },
    { cat:"Folha OK",  loja:Math.max(0,100-Math.round((folhaLoja/receitaLoja)*100||0)), matriz:Math.max(0,100-Math.round((folhaMatriz/receitaMatriz)*100||0)) },
  ];

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:20 }}>
      {/* DRE */}
      <Card>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14 }}>
          <ST color={t.gold}>DRE Simplificado — Loja vs Matriz</ST>
          <button onClick={()=>exportXLSX([{name:"DRE",data:dre.map(r=>({Item:r.item,"Loja (R$)":r.loja,"Matriz (R$)":r.matriz,"Total (R$)":r.tot}))}],"dre-dona-luzia.xlsx")} style={{ background:"#fefce8",color:t.gold,border:`1px solid ${t.gold}`,borderRadius:7,padding:"5px 12px",cursor:"pointer",fontSize:11,fontWeight:700 }}>⬇ Excel</button>
        </div>
        <table style={{ width:"100%", borderCollapse:"collapse" }}>
          <thead>
            <tr style={{ background:"#f0ede4" }}>
              {["Item","🏪 Loja","🏭 Matriz","Total"].map(h=>(
                <th key={h} style={{ padding:"9px 14px", textAlign:h==="Item"?"left":"right", color:t.muted, fontSize:11, letterSpacing:1, textTransform:"uppercase" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {dre.map((row,i)=>{
              const isLucro = row.tipo==="lucro";
              const isReceita = row.tipo==="receita";
              return (
                <tr key={i} style={{ borderTop:`1px solid ${isLucro?"#fde68a":t.bord}`, background:isLucro?"#fefce8":"transparent" }}>
                  <td style={{ padding:"11px 14px", color:isLucro?t.gold:isReceita?t.text:t.muted, fontSize:13, fontWeight:isLucro?700:400 }}>{row.item}</td>
                  {[row.loja,row.matriz,row.tot].map((v,j)=>(
                    <td key={j} style={{ padding:"11px 14px", textAlign:"right", color:isLucro?t.gold:v<0?t.red:t.green, fontFamily:"monospace", fontSize:13, fontWeight:isLucro?700:400 }}>{fmt(v)}</td>
                  ))}
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>

      {/* Radar + Empréstimos */}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16 }}>
        <Card>
          <ST>Saúde das Operações — Radar</ST>
          <ResponsiveContainer width="100%" height={240}>
            <RadarChart data={radarData}>
              <PolarGrid stroke="#e2ddd4" />
              <PolarAngleAxis dataKey="cat" tick={{ fill:t.muted, fontSize:11 }} />
              <Radar name="Loja"   dataKey="loja"   stroke={t.purple} fill={t.purple} fillOpacity={0.2} />
              <Radar name="Matriz" dataKey="matriz" stroke={t.cyan}   fill={t.cyan}   fillOpacity={0.2} />
              <Tooltip contentStyle={TT.contentStyle} formatter={v=>`${v}pts`} />
            </RadarChart>
          </ResponsiveContainer>
          <div style={{ display:"flex", gap:16, justifyContent:"center" }}>
            {[["🏪 Loja",t.purple],["🏭 Matriz",t.cyan]].map(([l,c])=>(
              <div key={l} style={{ display:"flex", alignItems:"center", gap:6 }}>
                <div style={{ width:10,height:10,borderRadius:2,background:c }} />
                <span style={{ color:t.muted, fontSize:12 }}>{l}</span>
              </div>
            ))}
          </div>
        </Card>

        <Card>
          <ST color={t.orange}>Empréstimos em Aberto</ST>
          <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
            {emprestimos.length===0
              ? <div style={{ color:t.muted, fontSize:13, padding:"20px 0", textAlign:"center" }}>Nenhum empréstimo registrado.</div>
              : emprestimos.map(e=>{
                const aberto = parcelas.filter(p=>p.emprestimo_id===e.id&&!p.pago).reduce((a,p)=>a+parseFloat(p.valor||0),0);
                const total  = parseFloat(e.valor||0);
                const prog   = total>0?Math.min(((total-aberto)/total)*100,100):0;
                return (
                  <div key={e.id} style={{ background:"#f8f7f2", borderRadius:10, padding:"12px 14px" }}>
                    <div style={{ display:"flex", justifyContent:"space-between", marginBottom:8 }}>
                      <span style={{ color:t.text, fontSize:13 }}>{e.credor} <span style={{ color:t.muted, fontSize:11 }}>· {e.operacao}</span></span>
                      <span style={{ color:t.orange, fontFamily:"monospace", fontSize:13, fontWeight:700 }}>{fmt(aberto)}</span>
                    </div>
                    <div style={{ background:"#e2ddd4", borderRadius:4, height:5, overflow:"hidden" }}>
                      <div style={{ width:`${prog}%`, background:t.green, height:"100%", borderRadius:4 }} />
                    </div>
                    <div style={{ display:"flex", justifyContent:"space-between", marginTop:4 }}>
                      <span style={{ color:t.muted, fontSize:11 }}>Pago: {prog.toFixed(0)}%</span>
                      <span style={{ color:t.muted, fontSize:11 }}>Total: {fmt(total)}</span>
                    </div>
                  </div>
                );
              })
            }
            {empAberto>0&&<div style={{ borderTop:`1px solid ${t.bord}`, paddingTop:10, display:"flex", justifyContent:"space-between" }}>
              <span style={{ color:t.muted, fontSize:12 }}>Total em aberto</span>
              <span style={{ color:t.orange, fontFamily:"monospace", fontSize:14, fontWeight:700 }}>{fmt(empAberto)}</span>
            </div>}
          </div>
        </Card>
      </div>
    </div>
  );
}

// ── CRUZAMENTO OPERAÇÕES ──────────────────────────────────────────────────────
function Operacoes({ dados }) {
  const { desperdicio, insumos } = dados;
  const despTotal   = desperdicio.reduce((s,d)=>s+parseFloat(d.custo||0),0);
  const despLoja    = desperdicio.filter(d=>d.operacao==="loja").reduce((s,d)=>s+parseFloat(d.custo||0),0);
  const despMatriz  = desperdicio.filter(d=>d.operacao==="matriz").reduce((s,d)=>s+parseFloat(d.custo||0),0);

  const motivosPorCusto = Object.entries(
    desperdicio.reduce((acc,d)=>{ acc[d.motivo]=(acc[d.motivo]||0)+parseFloat(d.custo||0); return acc; },{})
  ).map(([motivo,custo])=>({ motivo, custo })).sort((a,b)=>b.custo-a.custo);

  const criticos = insumos.filter(i=>{
    const r = (i.qtd||0)/(i.minimo||1);
    return i.alerta && r<=1;
  });
  const atencao = insumos.filter(i=>{
    const r = (i.qtd||0)/(i.minimo||1);
    return i.alerta && r>1 && r<=1.5;
  });

  const insumosCores = ["#a78bfa","#22d3ee","#4ade80","#fbbf24","#fb923c","#f472b6","#a3e635"];
  const topInsumos = [...insumos].sort((a,b)=>(b.qtd*b.vunit)-(a.qtd*a.vunit)).slice(0,8).map((i,idx)=>({
    nome:i.nome.length>14?i.nome.slice(0,14)+"…":i.nome,
    valor:+(i.qtd*(i.vunit||0)).toFixed(2),
    cor:insumosCores[idx%insumosCores.length],
  }));

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:20 }}>
      {/* Alertas insumos */}
      {criticos.length>0&&<Alerta tipo="crit" msg={`${criticos.length} insumo(s) crítico(s) ou zerado(s): ${criticos.map(i=>i.nome).join(", ")}`} />}
      {atencao.length>0&&<Alerta tipo="warn" msg={`${atencao.length} insumo(s) em atenção: ${atencao.map(i=>i.nome).join(", ")}`} />}

      <div style={{ display:"grid", gridTemplateColumns:"repeat(2,1fr)", gap:12 }}>
        <Kpi label="Total insumos"  value={insumos.length}       sub="Cadastrados"          color={t.blue}   icon="📦" />
        <Kpi label="Críticos"       value={criticos.length}      sub="Repor urgente"        color={criticos.length>0?t.red:t.green} icon="🔴" />
        <Kpi label="Em atenção"     value={atencao.length}       sub="Estoque baixo"        color={atencao.length>0?t.yellow:t.green} icon="🟡" />
        <Kpi label="Custo desperdício" value={fmt(despTotal)}    sub="Período atual"        color={t.red}    icon="🗑" />
      </div>

      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16 }}>
        {/* Desperdício por motivo */}
        <Card>
          <ST color={t.red}>Desperdício por Motivo (R$)</ST>
          {motivosPorCusto.length===0
            ? <div style={{ color:t.muted, fontSize:13, padding:"20px 0", textAlign:"center" }}>Nenhum desperdício registrado.</div>
            : motivosPorCusto.map(m=>(
              <div key={m.motivo} style={{ marginBottom:12 }}>
                <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}>
                  <span style={{ color:t.text, fontSize:13 }}>{m.motivo}</span>
                  <span style={{ color:t.red, fontFamily:"monospace", fontSize:13 }}>{fmt(m.custo)}</span>
                </div>
                <div style={{ background:"#f8f7f2", borderRadius:4, height:5, overflow:"hidden" }}>
                  <div style={{ width:pct(m.custo,despTotal), background:t.red, height:"100%", borderRadius:4 }} />
                </div>
              </div>
            ))
          }
          {despTotal>0&&(
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginTop:12 }}>
              {[["🏪 Loja",despLoja,t.purple],["🏭 Matriz",despMatriz,t.cyan]].map(([l,v,c])=>(
                <div key={l} style={{ background:"#f8f7f2", borderRadius:8, padding:"9px 12px", textAlign:"center" }}>
                  <div style={{ color:t.muted, fontSize:11 }}>{l}</div>
                  <div style={{ color:c, fontFamily:"monospace", fontSize:14, fontWeight:700, marginTop:4 }}>{fmt(v)}</div>
                </div>
              ))}
            </div>
          )}
        </Card>

        {/* Valor em estoque */}
        <Card>
          <ST>Valor em Estoque — Top Insumos</ST>
          {topInsumos.length===0
            ? <div style={{ color:t.muted, fontSize:13, padding:"20px 0", textAlign:"center" }}>Nenhum insumo cadastrado.</div>
            : <ResponsiveContainer width="100%" height={220}>
                <BarChart data={topInsumos} layout="vertical">
                  <XAxis type="number" stroke={t.muted} tick={{ fill:t.muted, fontSize:10 }} tickFormatter={v=>`R$${v.toFixed(0)}`} />
                  <YAxis type="category" dataKey="nome" stroke={t.muted} tick={{ fill:t.muted, fontSize:11 }} width={90} />
                  <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle} />
                  <Bar dataKey="valor" radius={[0,4,4,0]} name="Valor (R$)">
                    {topInsumos.map((d,i)=><Cell key={i} fill={d.cor} />)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
          }
        </Card>
      </div>
    </div>
  );
}

// ── CRUZAMENTO EQUIPE ─────────────────────────────────────────────────────────
function Equipe({ dados }) {
  const { colaboradores, consumo, banco } = dados;

  const totalColabs  = colaboradores.length;
  const folhaTotal   = colaboradores.reduce((s,c)=>s+parseFloat(c.salario||0)+parseFloat(c.vt||0)+parseFloat(c.vr||0)+parseFloat(c.plano||0),0);
  const consumoTotal = consumo.reduce((s,x)=>s+parseFloat(x.custo||0),0);
  const extrasTot    = banco.filter(b=>b.tipo==="extra").reduce((s,b)=>s+parseFloat(b.horas||0),0);
  const atrasosTot   = banco.filter(b=>b.tipo==="atraso").reduce((s,b)=>s+Math.abs(parseFloat(b.horas||0)),0);

  const porColab = colaboradores.map(c=>({
    nome:c.nome, operacao:c.operacao,
    custo:parseFloat(c.salario||0)+parseFloat(c.vt||0)+parseFloat(c.vr||0)+parseFloat(c.plano||0),
    consumo:consumo.filter(x=>x.colaborador_id===c.id).reduce((s,x)=>s+parseFloat(x.custo||0),0),
    extras:banco.filter(b=>b.colaborador_id===c.id&&b.tipo==="extra").reduce((s,b)=>s+parseFloat(b.horas||0),0),
    atrasos:banco.filter(b=>b.colaborador_id===c.id&&b.tipo==="atraso").reduce((s,b)=>s+Math.abs(parseFloat(b.horas||0)),0),
  }));

  const receitaTotal = 18400 + 12600;
  const custoPorReceita = folhaTotal / receitaTotal * 100;

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:20 }}>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(2,1fr)", gap:12 }}>
        <Kpi label="Colaboradores"   value={totalColabs}        sub="Ativos"                  color={t.blue}   icon="👥" />
        <Kpi label="Folha total"     value={fmtK(folhaTotal)}   sub={pct(folhaTotal,receitaTotal)+" da receita"} color={t.yellow} icon="💼" />
        <Kpi label="Horas extras"    value={`${extrasTot.toFixed(1)}h`}  sub="Acumuladas"    color={t.green}  icon="⏱" />
        <Kpi label="Atrasos"         value={`${atrasosTot.toFixed(1)}h`} sub="Registrados"   color={atrasosTot>5?t.red:t.muted} icon="⚠️" />
      </div>

      {custoPorReceita>35 && <Alerta tipo="warn" msg={`Folha representa ${custoPorReceita.toFixed(1)}% da receita. Acima de 35% pode comprimir a margem.`} />}

      <Card>
        <ST>Custo + Consumo por Colaborador</ST>
        {porColab.length===0
          ? <div style={{ color:t.muted, fontSize:13, padding:"20px 0", textAlign:"center" }}>Nenhum colaborador cadastrado.</div>
          : <ResponsiveContainer width="100%" height={200}>
              <BarChart data={porColab} barCategoryGap="30%">
                <XAxis dataKey="nome" stroke={t.muted} tick={{ fill:t.muted, fontSize:11 }} />
                <YAxis stroke={t.muted} tick={{ fill:t.muted, fontSize:11 }} tickFormatter={v=>`R$${(v/1000).toFixed(1)}k`} />
                <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle} />
                <Bar dataKey="custo"   fill={t.blue}   radius={[4,4,0,0]} name="Custo (Folha)" stackId="a" />
                <Bar dataKey="consumo" fill={t.orange} radius={[0,0,0,0]} name="Consumo"        stackId="a" />
              </BarChart>
            </ResponsiveContainer>
        }
        <div style={{ display:"flex", gap:16, marginTop:8 }}>
          {[["Custo Folha",t.blue],["Consumo",t.orange]].map(([l,c])=>(
            <div key={l} style={{ display:"flex", alignItems:"center", gap:6 }}>
              <div style={{ width:10,height:10,borderRadius:2,background:c }} />
              <span style={{ color:t.muted, fontSize:12 }}>{l}</span>
            </div>
          ))}
        </div>
      </Card>

      <Card>
        <ST>Banco de Horas — Visão Geral</ST>
        <table style={{ width:"100%", borderCollapse:"collapse" }}>
          <thead>
            <tr style={{ background:"#f0ede4" }}>
              {["Colaborador","Operação","Extras","Atrasos","Saldo","Consumo Mês"].map(h=>(
                <th key={h} style={{ padding:"9px 12px", textAlign:h==="Colaborador"||h==="Operação"?"left":"right", color:t.muted, fontSize:11, letterSpacing:1, textTransform:"uppercase" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {porColab.length===0
              ? <tr><td colSpan={6} style={{ padding:"20px",color:t.muted,textAlign:"center",fontSize:13 }}>Nenhum dado disponível.</td></tr>
              : porColab.map((c,i)=>{
                  const saldo = c.extras - c.atrasos;
                  return (
                    <tr key={i} style={{ borderTop:`1px solid ${t.bord}` }}>
                      <td style={{ padding:"10px 12px", color:t.text, fontSize:13, fontWeight:600 }}>{c.nome}</td>
                      <td style={{ padding:"10px 12px", color:t.muted, fontSize:12 }}>{c.operacao==="loja"?"🏪 Loja":"🏭 Matriz"}</td>
                      <td style={{ padding:"10px 12px", textAlign:"right", color:t.green, fontFamily:"monospace", fontSize:13 }}>{c.extras.toFixed(1)}h</td>
                      <td style={{ padding:"10px 12px", textAlign:"right", color:t.red, fontFamily:"monospace", fontSize:13 }}>{c.atrasos.toFixed(1)}h</td>
                      <td style={{ padding:"10px 12px", textAlign:"right", color:saldo>=0?t.green:t.red, fontFamily:"monospace", fontSize:13, fontWeight:700 }}>{saldo>=0?"+":""}{saldo.toFixed(1)}h</td>
                      <td style={{ padding:"10px 12px", textAlign:"right", color:t.orange, fontFamily:"monospace", fontSize:13 }}>{fmt(c.consumo)}</td>
                    </tr>
                  );
                })
            }
          </tbody>
        </table>
      </Card>
    </div>
  );
}

// ── CMV & DRE ─────────────────────────────────────────────────────────────────
function CmvDre({ dados }) {
  const { vendas_loja, pedidos_matriz, insumos_mov, colaboradores, desperdicio, consumo } = dados;
  const [cmvPctLoja,   setCmvPctLoja]   = useState(35);
  const [cmvPctMatriz, setCmvPctMatriz] = useState(40);
  const [op, setOp] = useState("loja");

  // Receitas por mês (histórico fixo + atual do banco)
  const receitas = {
    loja:   { Jan:15000, Fev:16200, Mar:17500, Abr: vendas_loja.reduce((s,v)=>s+parseFloat(v.valor||0),0)||18400 },
    matriz: { Jan:10000, Fev:10800, Mar:11200, Abr: pedidos_matriz.reduce((s,p)=>s+parseFloat(p.total||0),0)||12600 },
  };

  // Custos fixos por mês
  const custosFixos = {
    loja:   { Jan:2800, Fev:2900, Mar:3000, Abr:3100 }, // aluguel+contas
    matriz: { Jan:2200, Fev:2300, Mar:2400, Abr:2500 },
  };

  // CMV real = custo dos insumos baixados do estoque
  const cmvRealLoja   = (insumos_mov||[]).filter(m=>m.operacao==="loja"  &&m.tipo!=="entrada").reduce((s,m)=>{
    const ins = (dados.insumos||[]).find(i=>i.id===m.insumo_id);
    return s + parseFloat(m.qtd||0)*parseFloat(ins?.vunit||0);
  },0);
  const cmvRealMatriz = (insumos_mov||[]).filter(m=>m.operacao==="matriz"&&m.tipo!=="entrada").reduce((s,m)=>{
    const ins = (dados.insumos||[]).find(i=>i.id===m.insumo_id);
    return s + parseFloat(m.qtd||0)*parseFloat(ins?.vunit||0);
  },0);

  const meses = ["Jan","Fev","Mar","Abr"];

  // Gera DRE completo por operação por mês
  const gerarDRE = (oper) => {
    const rec    = receitas[oper];
    const cfix   = custosFixos[oper];
    const folha  = colaboradores.filter(c=>c.operacao===oper).reduce((s,c)=>s+parseFloat(c.salario||0)+parseFloat(c.vt||0)+parseFloat(c.vr||0)+parseFloat(c.plano||0),0);
    const desp   = desperdicio.filter(d=>d.operacao===oper).reduce((s,d)=>s+parseFloat(d.custo||0),0);
    const cons   = consumo.filter(c=>c.operacao===oper).reduce((s,c)=>s+parseFloat(c.custo||0),0);
    const taxas  = oper==="loja" ? 740 : 380;
    const transp = oper==="loja" ? 320 : 560;
    const cmvPct = oper==="loja" ? cmvPctLoja : cmvPctMatriz;
    const cmvReal= oper==="loja" ? cmvRealLoja : cmvRealMatriz;

    return meses.map((mes,i) => {
      const receita     = rec[mes];
      // CMV: usa o real se disponível (mês atual), senão usa percentual
      const cmvPercent  = receita * (cmvPct/100);
      const cmvUsado    = mes==="Abr" && cmvReal>0 ? cmvReal : cmvPercent;
      const cmvDesvio   = mes==="Abr" && cmvReal>0 ? cmvReal - cmvPercent : 0;
      const margBruta   = receita - cmvUsado;
      const despTotal   = (i===3?folha:folha*0.95) + (i===3?cfix[mes]:cfix[mes]*0.97) + (i===3?desp:desp*0.6) + (i===3?cons:cons*0.7) + taxas + transp;
      const lucroLiq    = margBruta - despTotal;
      return {
        mes, receita, cmvUsado, cmvPercent, cmvReal:mes==="Abr"?cmvReal:0,
        cmvDesvio, margBruta, despTotal, folha:i===3?folha:folha*0.95,
        custosFixos:cfix[mes], taxas, transporte:transp,
        desperdicio:i===3?desp:desp*0.6, consumoColab:i===3?cons:cons*0.7,
        lucroLiq,
        margemBruta: receita>0?((margBruta/receita)*100).toFixed(1):0,
        margemLiq:   receita>0?((lucroLiq/receita)*100).toFixed(1):0,
        cmvPct:      receita>0?((cmvUsado/receita)*100).toFixed(1):0,
      };
    });
  };

  const dre = gerarDRE(op);
  const dreOutra = gerarDRE(op==="loja"?"matriz":"loja");
  const recAtual = dre[3].receita;
  const cmvAtual = dre[3].cmvUsado;
  const margAtual= dre[3].margBruta;
  const lucAtual = dre[3].lucroLiq;
  const cmvReal  = op==="loja" ? cmvRealLoja : cmvRealMatriz;
  const cmvPct   = op==="loja" ? cmvPctLoja : cmvPctMatriz;

  const linhasDRE = [
    { label:"(+) Receita Bruta",      key:"receita",      cor:t.gold,   bold:true  },
    { label:"(−) CMV",                key:"cmvUsado",     cor:t.red,    bold:false, neg:true },
    { label:"(=) Margem Bruta",       key:"margBruta",    cor:t.green,  bold:true, sep:true },
    { label:"(−) Folha de Pagamento", key:"folha",        cor:t.muted,  bold:false, neg:true },
    { label:"(−) Aluguel / Contas",   key:"custosFixos",  cor:t.muted,  bold:false, neg:true },
    { label:"(−) Taxas iFood/Anotai", key:"taxas",        cor:t.muted,  bold:false, neg:true },
    { label:"(−) Transporte",         key:"transporte",   cor:t.muted,  bold:false, neg:true },
    { label:"(−) Desperdício",        key:"desperdicio",  cor:t.muted,  bold:false, neg:true },
    { label:"(−) Consumo Colabor.",   key:"consumoColab", cor:t.muted,  bold:false, neg:true },
    { label:"(=) Lucro Líquido",      key:"lucroLiq",     cor:t.gold,   bold:true, sep:true  },
  ];

  const barDRE = dre.map(d=>({ mes:d.mes, Receita:d.receita, CMV:d.cmvUsado, "Marg.Bruta":d.margBruta, "Lucro Líq.":d.lucroLiq>0?d.lucroLiq:0 }));

  const inp2 = { background:"#f8f7f2", border:`1px solid ${t.bord}`, borderRadius:7, padding:"7px 10px", color:t.text, fontSize:13, fontFamily:"inherit", outline:"none", width:70, textAlign:"center" };

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:20 }}>

      {/* Seletor operação */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:10 }}>
        <div style={{ display:"flex", gap:8 }}>
          {[["loja","🏪 Loja"],["matriz","🏭 Matriz"]].map(([id,label])=>(
            <button key={id} onClick={()=>setOp(id)} style={{ background:op===id?t.gold:t.card, color:op===id?t.bg:t.muted, border:`1px solid ${op===id?t.gold:t.bord}`, borderRadius:8, padding:"8px 20px", cursor:"pointer", fontSize:12, fontWeight:op===id?700:400 }}>{label}</button>
          ))}
        </div>
        <button onClick={()=>exportXLSX([
          {name:`DRE ${op==="loja"?"Loja":"Matriz"}`, data:dre.map(d=>linhasDRE.map(l=>({Mês:d.mes,Item:l.label,"Valor (R$)":l.neg?-d[l.key]:d[l.key]})).flat()).flat()},
          {name:"CMV Comparativo", data:dre.map(d=>({Mês:d.mes,"CMV Real (R$)":d.cmvReal,"CMV %Receita":d.cmvPct+"%","Desvio (R$)":d.cmvDesvio,"Margem Bruta %":d.margemBruta+"%","Margem Líq %":d.margemLiq+"%"}))},
        ],"dre-cmv.xlsx")} style={{ background:"#fefce8", color:t.gold, border:`1px solid ${t.gold}`, borderRadius:7, padding:"7px 16px", cursor:"pointer", fontSize:11, fontWeight:700 }}>⬇ Exportar DRE</button>
      </div>

      {/* KPIs CMV */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(2,1fr)", gap:12 }}>
        <Kpi label="Receita"       value={fmt(recAtual)}              sub="Mês atual"              color={t.gold}                              icon="💰" />
        <Kpi label="CMV"           value={fmt(cmvAtual)}              sub={`${dre[3].cmvPct}% da receita`} color={t.red}                     icon="📦" />
        <Kpi label="Margem Bruta"  value={fmt(margAtual)}             sub={`${dre[3].margemBruta}% da receita`} color={t.green}              icon="📊" />
        <Kpi label="Lucro Líquido" value={fmt(lucAtual)}              sub={`${dre[3].margemLiq}% de margem`} color={lucAtual>0?t.lime:t.red} icon="📈" />
      </div>

      {/* CMV: real vs percentual */}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16 }}>
        <Card style={{ border:`1px solid ${t.gold}44` }}>
          <ST color={t.gold}>CMV — Real vs Percentual ({op==="loja"?"Loja":"Matriz"})</ST>

          {/* Configurar % esperado */}
          <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:16, background:"#f8f7f2", borderRadius:8, padding:"10px 14px" }}>
            <span style={{ color:t.muted, fontSize:13 }}>% CMV esperado:</span>
            <input type="number" value={op==="loja"?cmvPctLoja:cmvPctMatriz}
              onChange={e=>op==="loja"?setCmvPctLoja(parseFloat(e.target.value)||0):setCmvPctMatriz(parseFloat(e.target.value)||0)}
              style={inp2} min={0} max={100} />
            <span style={{ color:t.muted, fontSize:13 }}>% da receita</span>
          </div>

          <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
            {[
              ["CMV Real (insumos baixados)", cmvReal>0?cmvReal:null, t.cyan,   "Calculado pelas saídas do estoque"],
              ["CMV % Esperado",              recAtual*(cmvPct/100),  t.yellow, `${cmvPct}% de ${fmt(recAtual)}`],
              ["Desvio (real − esperado)",    cmvReal>0?cmvReal-recAtual*(cmvPct/100):null, cmvReal>recAtual*(cmvPct/100)?t.red:t.green, cmvReal===0?"Lance dados de estoque para calcular":""],
            ].map(([label,val,cor,obs])=>(
              <div key={label} style={{ background:"#f8f7f2", borderRadius:9, padding:"12px 14px" }}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                  <span style={{ color:t.muted, fontSize:12 }}>{label}</span>
                  <span style={{ color:cor, fontFamily:"monospace", fontSize:14, fontWeight:700 }}>
                    {val===null?"—":fmt(val)}
                  </span>
                </div>
                {obs&&<div style={{ color:t.muted, fontSize:11, marginTop:3 }}>{obs}</div>}
              </div>
            ))}
          </div>

          {/* Gauge margem bruta */}
          <div style={{ marginTop:16 }}>
            <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
              <span style={{ color:t.muted, fontSize:12 }}>Margem Bruta</span>
              <span style={{ color:parseFloat(dre[3].margemBruta)>50?t.green:parseFloat(dre[3].margemBruta)>35?t.yellow:t.red, fontFamily:"monospace", fontSize:14, fontWeight:700 }}>{dre[3].margemBruta}%</span>
            </div>
            <div style={{ background:"#e2ddd4", borderRadius:6, height:10, overflow:"hidden" }}>
              <div style={{ width:`${dre[3].margemBruta}%`, background:parseFloat(dre[3].margemBruta)>50?t.green:parseFloat(dre[3].margemBruta)>35?t.yellow:t.red, height:"100%", borderRadius:6, transition:"width 0.5s" }} />
            </div>
            <div style={{ display:"flex", justifyContent:"space-between", marginTop:4 }}>
              <span style={{ color:t.red, fontSize:10 }}>Crítico &lt;35%</span>
              <span style={{ color:t.yellow, fontSize:10 }}>Atenção 35–50%</span>
              <span style={{ color:t.green, fontSize:10 }}>Saudável &gt;50%</span>
            </div>
          </div>
        </Card>

        {/* Evolução CMV % nos 4 meses */}
        <Card>
          <ST>Evolução CMV % e Margem Bruta % — 4 meses</ST>
          <ResponsiveContainer width="100%" height={200}>
            <LineChart data={dre}>
              <XAxis dataKey="mes" stroke={t.muted} tick={{ fill:t.muted, fontSize:11 }} />
              <YAxis stroke={t.muted} tick={{ fill:t.muted, fontSize:11 }} tickFormatter={v=>`${v}%`} domain={[0,100]} />
              <Tooltip formatter={v=>`${v}%`} contentStyle={TT.contentStyle} />
              <Line type="monotone" dataKey="cmvPct"      stroke={t.red}   strokeWidth={2} dot={{ fill:t.red,   r:4 }} name="CMV %" />
              <Line type="monotone" dataKey="margemBruta" stroke={t.green} strokeWidth={2} dot={{ fill:t.green, r:4 }} name="Margem Bruta %" />
              <Line type="monotone" dataKey="margemLiq"   stroke={t.gold}  strokeWidth={2} dot={{ fill:t.gold,  r:4 }} name="Margem Líq %" strokeDasharray="5 3" />
            </LineChart>
          </ResponsiveContainer>
          <div style={{ display:"flex", gap:16, marginTop:8, flexWrap:"wrap" }}>
            {[["CMV %",t.red],["Margem Bruta %",t.green],["Margem Líq %",t.gold]].map(([l,c])=>(
              <div key={l} style={{ display:"flex", alignItems:"center", gap:6 }}>
                <div style={{ width:10,height:10,borderRadius:2,background:c }} />
                <span style={{ color:t.muted, fontSize:12 }}>{l}</span>
              </div>
            ))}
          </div>
        </Card>
      </div>

      {/* DRE — 4 meses lado a lado */}
      <Card>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
          <ST color={op==="loja"?t.purple:t.cyan}>DRE — {op==="loja"?"🏪 Loja":"🏭 Matriz"} — Últimos 4 Meses</ST>
        </div>
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", minWidth:560 }}>
            <thead>
              <tr style={{ background:"#f0ede4" }}>
                <th style={{ padding:"10px 14px", textAlign:"left", color:t.muted, fontSize:11, letterSpacing:1, textTransform:"uppercase", width:200 }}>Item</th>
                {meses.map(m=>(
                  <th key={m} style={{ padding:"10px 14px", textAlign:"right", color:m==="Abr"?t.gold:t.muted, fontSize:11, letterSpacing:1, textTransform:"uppercase" }}>
                    {m}{m==="Abr"?" ★":""}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {linhasDRE.map((linha,i)=>(
                <tr key={i} style={{ borderTop:`1px solid ${linha.sep?"#fde68a":t.bord}`, background:linha.bold&&linha.sep?"#fefce8":"transparent" }}>
                  <td style={{ padding:"10px 14px", color:linha.bold?linha.cor:t.muted, fontSize:13, fontWeight:linha.bold?700:400 }}>{linha.label}</td>
                  {dre.map((d,j)=>{
                    const val = d[linha.key];
                    const display = linha.neg ? -val : val;
                    const isNeg = display < 0;
                    return (
                      <td key={j} style={{ padding:"10px 14px", textAlign:"right", fontFamily:"monospace", fontSize:13, fontWeight:linha.bold?700:400,
                        color:linha.bold?(val<0&&linha.key==="lucroLiq"?t.red:linha.cor):(linha.neg?t.muted:t.text)
                      }}>
                        {fmt(display)}
                      </td>
                    );
                  })}
                </tr>
              ))}
              {/* Linha margem % */}
              <tr style={{ borderTop:`1px solid ${t.bord}`, background:"#f0ede4" }}>
                <td style={{ padding:"8px 14px", color:t.muted, fontSize:11, letterSpacing:1, textTransform:"uppercase" }}>Margem Líquida %</td>
                {dre.map((d,j)=>(
                  <td key={j} style={{ padding:"8px 14px", textAlign:"right", fontFamily:"monospace", fontSize:12, color:parseFloat(d.margemLiq)>0?t.green:t.red }}>
                    {d.margemLiq}%
                  </td>
                ))}
              </tr>
              <tr style={{ background:"#f0ede4" }}>
                <td style={{ padding:"8px 14px", color:t.muted, fontSize:11, letterSpacing:1, textTransform:"uppercase" }}>CMV %</td>
                {dre.map((d,j)=>(
                  <td key={j} style={{ padding:"8px 14px", textAlign:"right", fontFamily:"monospace", fontSize:12, color:parseFloat(d.cmvPct)<cmvPct?t.green:t.red }}>
                    {d.cmvPct}%
                  </td>
                ))}
              </tr>
            </tbody>
          </table>
        </div>
      </Card>

      {/* Gráfico barras DRE */}
      <Card>
        <ST>Receita · CMV · Margem Bruta · Lucro — {op==="loja"?"Loja":"Matriz"}</ST>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={barDRE} barCategoryGap="25%">
            <XAxis dataKey="mes" stroke={t.muted} tick={{ fill:t.muted, fontSize:11 }} />
            <YAxis stroke={t.muted} tick={{ fill:t.muted, fontSize:11 }} tickFormatter={v=>`R$${(v/1000).toFixed(0)}k`} />
            <Tooltip formatter={v=>fmt(v)} contentStyle={TT.contentStyle} />
            <Bar dataKey="Receita"     fill={t.gold}   radius={[4,4,0,0]} />
            <Bar dataKey="CMV"         fill={t.red}    radius={[4,4,0,0]} />
            <Bar dataKey="Marg.Bruta"  fill={t.green}  radius={[4,4,0,0]} />
            <Bar dataKey="Lucro Líq."  fill={t.purple} radius={[4,4,0,0]} />
          </BarChart>
        </ResponsiveContainer>
        <div style={{ display:"flex", gap:16, marginTop:8, flexWrap:"wrap" }}>
          {[["Receita",t.gold],["CMV",t.red],["Margem Bruta",t.green],["Lucro Líq.",t.purple]].map(([l,c])=>(
            <div key={l} style={{ display:"flex", alignItems:"center", gap:6 }}>
              <div style={{ width:10,height:10,borderRadius:2,background:c }} />
              <span style={{ color:t.muted, fontSize:12 }}>{l}</span>
            </div>
          ))}
        </div>
      </Card>

      {/* Comparativo Loja vs Matriz */}
      <Card>
        <ST color={t.gold}>Comparativo Loja vs Matriz — Mês Atual</ST>
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse" }}>
            <thead>
              <tr style={{ background:"#f0ede4" }}>
                {["Item","🏪 Loja","🏭 Matriz","Diferença"].map(h=>(
                  <th key={h} style={{ padding:"10px 14px", textAlign:h==="Item"?"left":"right", color:t.muted, fontSize:11, letterSpacing:1, textTransform:"uppercase" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {linhasDRE.map((linha,i)=>{
                const vLoja   = gerarDRE("loja")[3][linha.key];
                const vMatriz = gerarDRE("matriz")[3][linha.key];
                const diff    = vLoja - vMatriz;
                return (
                  <tr key={i} style={{ borderTop:`1px solid ${linha.sep?"#fde68a":t.bord}`, background:linha.bold&&linha.sep?"#fefce8":"transparent" }}>
                    <td style={{ padding:"10px 14px", color:linha.bold?linha.cor:t.muted, fontSize:13, fontWeight:linha.bold?700:400 }}>{linha.label}</td>
                    <td style={{ padding:"10px 14px", textAlign:"right", color:t.purple, fontFamily:"monospace", fontSize:13, fontWeight:linha.bold?700:400 }}>{fmt(linha.neg?-vLoja:vLoja)}</td>
                    <td style={{ padding:"10px 14px", textAlign:"right", color:t.cyan,   fontFamily:"monospace", fontSize:13, fontWeight:linha.bold?700:400 }}>{fmt(linha.neg?-vMatriz:vMatriz)}</td>
                    <td style={{ padding:"10px 14px", textAlign:"right", color:diff>0?t.green:t.red, fontFamily:"monospace", fontSize:13 }}>{diff>=0?"+":""}{fmt(diff)}</td>
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

// ── APP ───────────────────────────────────────────────────────────────────────
const TABS = [
  { id:"geral",       label:"🏠 Visão Geral"     },
  { id:"financeiro",  label:"💰 Financeiro"       },
  { id:"cmv",         label:"📉 CMV & DRE"        },
  { id:"operacoes",   label:"📦 Operações"        },
  { id:"equipe",      label:"👥 Equipe"           },
];

export default function App() {
  const [tab,     setTab]     = useState("geral");
  const [dados,   setDados]   = useState(null);
  const [loading, setLoading] = useState(true);
  const [error,   setError]   = useState(null);
  const [lastUp,  setLastUp]  = useState(null);
  const [usoExemplo, setUsoExemplo] = useState(false);

  // Dados de exemplo para quando o banco estiver vazio ou inacessível
  const DADOS_EXEMPLO = {
    vendas_loja: [
      {valor:4920,canal:"Balcão"},{valor:4650,canal:"iFood"},{valor:3360,canal:"Balcão"},
      {valor:3000,canal:"Delivery"},{valor:2280,canal:"iFood"},
    ],
    pedidos_matriz: [
      {total:378,status:"Entregue"},{total:200,status:"Entregue"},{total:156,status:"Entregue"},
      {total:120,status:"Entregue"},{total:378,status:"A entregar"},{total:85,status:"A entregar"},
    ],
    colaboradores: [
      {id:"c1",nome:"Camila",  cargo:"Atendente", operacao:"loja",   salario:1412,vt:220,vr:198,plano:0},
      {id:"c2",nome:"Rebeca",  cargo:"Atendente", operacao:"loja",   salario:1412,vt:180,vr:198,plano:0},
      {id:"c3",nome:"Daniela", cargo:"Caixa",     operacao:"loja",   salario:1600,vt:200,vr:198,plano:120},
      {id:"c4",nome:"Fátima",  cargo:"Cozinheira",operacao:"matriz", salario:1800,vt:240,vr:198,plano:0},
      {id:"c5",nome:"Andressa",cargo:"Cozinheira",operacao:"matriz", salario:1650,vt:200,vr:198,plano:0},
    ],
    consumo: [
      {operacao:"loja",  colaborador_id:"c1",custo:8.50},{operacao:"loja",  colaborador_id:"c2",custo:12.00},
      {operacao:"loja",  colaborador_id:"c3",custo:9.00},{operacao:"matriz",colaborador_id:"c4",custo:18.00},
      {operacao:"matriz",colaborador_id:"c5",custo:15.00},{operacao:"matriz",colaborador_id:"c4",custo:20.00},
    ],
    banco: [
      {operacao:"loja",  colaborador_id:"c1",tipo:"extra", horas:1.5},
      {operacao:"loja",  colaborador_id:"c2",tipo:"extra", horas:1.0},
      {operacao:"loja",  colaborador_id:"c1",tipo:"atraso",horas:-0.5},
      {operacao:"matriz",colaborador_id:"c4",tipo:"extra", horas:2.0},
      {operacao:"matriz",colaborador_id:"c5",tipo:"extra", horas:1.5},
    ],
    desperdicio: [
      {operacao:"loja",  produto:"Morango",  custo:5.40, motivo:"Vencido"},
      {operacao:"loja",  produto:"Smoothie", custo:17.00,motivo:"Erro de preparo"},
      {operacao:"matriz",produto:"Frango",   custo:7.45, motivo:"Erro de preparo"},
      {operacao:"matriz",produto:"Arroz",    custo:4.64, motivo:"Queimado"},
      {operacao:"matriz",produto:"Marmita",  custo:54.00,motivo:"Erro de preparo"},
    ],
    emprestimos: [
      {id:"e1",operacao:"Loja",  credor:"Maria (sócia)",valor:3000,data:"2025-03-03",motivo:"Fechamento negativo"},
      {id:"e2",operacao:"Matriz",credor:"Banco Família", valor:5000,data:"2025-01-10",motivo:"Reforço de caixa"},
    ],
    parcelas: [
      {emprestimo_id:"e1",num:1,valor:600,pago:true}, {emprestimo_id:"e1",num:2,valor:600,pago:true},
      {emprestimo_id:"e1",num:3,valor:600,pago:false},{emprestimo_id:"e1",num:4,valor:600,pago:false},{emprestimo_id:"e1",num:5,valor:600,pago:false},
      {emprestimo_id:"e2",num:1,valor:1000,pago:true},{emprestimo_id:"e2",num:2,valor:1000,pago:true},{emprestimo_id:"e2",num:3,valor:1000,pago:true},
      {emprestimo_id:"e2",num:4,valor:1000,pago:false},{emprestimo_id:"e2",num:5,valor:1000,pago:false},
    ],
    reservas_lanc: [
      {operacao:"loja",destino_id:"d1",valor:650},{operacao:"loja",destino_id:"d2",valor:350},
      {operacao:"matriz",destino_id:"d2",valor:200},{operacao:"matriz",destino_id:"d3",valor:550},
    ],
    reservas_dest: [
      {id:"d1",label:"Cartão Tiago",meta:1800},{id:"d2",label:"Empréstimos",meta:1200},
      {id:"d3",label:"Aluguel",meta:2100},{id:"d4",label:"Outras Contas",meta:600},
    ],
    insumos: [
      {id:"i1",nome:"Banana",       operacao:"loja",  qtd:8.5, minimo:5,  alerta:true, vunit:6.90},
      {id:"i2",nome:"Morango",      operacao:"loja",  qtd:2.0, minimo:3,  alerta:true, vunit:18.00},
      {id:"i3",nome:"Leite",        operacao:"loja",  qtd:12,  minimo:8,  alerta:true, vunit:4.50},
      {id:"i4",nome:"Whey",         operacao:"loja",  qtd:0.8, minimo:1,  alerta:true, vunit:90.00},
      {id:"i5",nome:"Arroz integral",operacao:"matriz",qtd:18, minimo:10, alerta:true, vunit:5.80},
      {id:"i6",nome:"Frango",       operacao:"matriz",qtd:6.5, minimo:5,  alerta:true, vunit:14.90},
      {id:"i7",nome:"Carne moída",  operacao:"matriz",qtd:3.0, minimo:4,  alerta:true, vunit:32.00},
      {id:"i8",nome:"Brócolis",     operacao:"matriz",qtd:1.5, minimo:2,  alerta:true, vunit:8.00},
    ],
    insumos_mov: [
      {insumo_id:"i5",operacao:"matriz",tipo:"saida_lote",qtd:4.5},
      {insumo_id:"i6",operacao:"matriz",tipo:"saida_receita",qtd:6.0},
      {insumo_id:"i7",operacao:"matriz",tipo:"saida_lote",qtd:1.8},
      {insumo_id:"i1",operacao:"loja",  tipo:"saida_lote",qtd:3.0},
      {insumo_id:"i3",operacao:"loja",  tipo:"saida_lote",qtd:5.0},
    ],
  };

  const carregar = async () => {
    setLoading(true); setError(null);
    try {
      const [vendas_loja, pedidos_matriz, colaboradores, consumo, banco, desperdicio, emprestimos, parcelas, reservas_lanc, reservas_dest, insumos, insumos_mov] = await Promise.all([
        sbGet("vendas_loja",           "order=created_at.desc"),
        sbGet("pedidos_matriz",        "order=created_at.desc"),
        sbGet("colaboradores",         "order=created_at.asc"),
        sbGet("consumo_colaboradores", "order=data.desc"),
        sbGet("banco_horas",           "order=data.desc"),
        sbGet("desperdicio",           "order=data.desc"),
        sbGet("emprestimos",           "order=created_at.desc"),
        sbGet("emprestimo_parcelas",   "order=num.asc"),
        sbGet("reservas_lancamentos",  "order=data.desc"),
        sbGet("reservas_destinos",     "order=created_at.asc"),
        sbGet("insumos",               "order=created_at.asc"),
        sbGet("insumos_movimentos",    "order=data.desc"),
      ]);
      const d = { vendas_loja, pedidos_matriz, colaboradores, consumo, banco, desperdicio, emprestimos, parcelas, reservas_lanc, reservas_dest, insumos, insumos_mov };
      // Se banco vazio, usa exemplo
      const temDados = colaboradores.length>0 || vendas_loja.length>0 || pedidos_matriz.length>0;
      setDados(temDados ? d : { ...DADOS_EXEMPLO });
      setUsoExemplo(!temDados);
      setLastUp(new Date().toLocaleTimeString("pt-BR"));
    } catch(e) {
      // Fallback para dados de exemplo em caso de erro de rede
      setDados({ ...DADOS_EXEMPLO });
      setUsoExemplo(true);
      setLastUp(new Date().toLocaleTimeString("pt-BR"));
    }
    finally { setLoading(false); }
  };

  useEffect(() => { carregar(); }, []);

  return (
    <div style={{ minHeight:"100vh", background:"#f8f7f2", color:t.text, fontFamily:"'DM Sans', sans-serif" }}>
      <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;600&family=DM+Sans:wght@400;500;600&display=swap" rel="stylesheet" />

      {/* Header */}
      <div style={{ borderBottom:`1px solid ${t.bord}`, padding:"16px 20px" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:10 }}>
          <div>
            <div style={{ fontFamily:"'Cormorant Garamond',serif", fontSize:22, color:t.gold, letterSpacing:0.5 }}>✦ Dona Luzia</div>
            <div style={{ color:t.muted, fontSize:11, letterSpacing:1, marginTop:1 }}>PAINEL ESTRATÉGICO</div>
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" }}>
            <div style={{ display:"flex", alignItems:"center", gap:5, background:t.card, border:`1px solid ${usoExemplo?t.yellow:t.green}`, borderRadius:8, padding:"5px 10px" }}>
              <div style={{ width:6, height:6, borderRadius:"50%", background:usoExemplo?t.yellow:t.green, boxShadow:`0 0 5px ${usoExemplo?t.yellow:t.green}` }} />
              <span style={{ color:t.muted, fontSize:11 }}>{usoExemplo?"Exemplo":"Supabase ✓"}</span>
            </div>
            <button onClick={carregar} style={{ background:t.gold, color:t.bg, border:"none", borderRadius:8, padding:"7px 14px", cursor:"pointer", fontSize:12, fontWeight:700 }}>↺ Atualizar</button>
          </div>
        </div>
        {lastUp && <div style={{ color:t.muted, fontSize:10, marginTop:6 }}>Última atualização: {lastUp}</div>}
        {usoExemplo && <div style={{ marginTop:8, background:"#fefce8", border:`1px solid #b45309`, borderRadius:8, padding:"8px 12px", color:"#92400e", fontSize:11 }}>📊 Exibindo dados de exemplo. Cadastre dados nos Apps 1–4 para ver seus números reais.</div>}
      </div>

      {/* Tabs */}
      <div style={{ padding:"14px 16px 0", overflowX:"auto" }}>
        <div style={{ display:"flex", gap:8, minWidth:"max-content" }}>
          {TABS.map(tb=><TabBtn key={tb.id} label={tb.label} active={tab===tb.id} onClick={()=>setTab(tb.id)} />)}
        </div>
      </div>

      {/* Conteúdo */}
      <div style={{ padding:"16px" }}>
        {loading && <Loading />}
        {!loading && dados && (
          <>
            {tab==="geral"      && <Geral      dados={dados} />}
            {tab==="financeiro" && <Financeiro dados={dados} />}
            {tab==="cmv"        && <CmvDre     dados={dados} />}
            {tab==="operacoes"  && <Operacoes  dados={dados} />}
            {tab==="equipe"     && <Equipe     dados={dados} />}
          </>
        )}
      </div>
    </div>
  );
}

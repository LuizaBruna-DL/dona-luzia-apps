import { useState, useEffect } from "react";
import { createClient } from "@supabase/supabase-js";
import AppEstrategico from "./app0-estrategico";
import AppFinanceiro  from "./app1-financeiro";
import AppOperacoes   from "./app2-operacoes";
import AppVendas      from "./app3-vendas";
import AppEquipe      from "./app4-equipe";

const SUPA_URL = "https://rgsgjqnlvyossuykemjl.supabase.co";
const SUPA_KEY = "sb_publishable_uC6SF60PfuFnZfDDLRo2Wg_m7l2FoBW";
const supabase = createClient(SUPA_URL, SUPA_KEY);

const MODULES = [
  { id:"estrategico", label:"Estrategico", icon:"[E]", desc:"Dashboard geral, relatorios e exportacao", color:"#b45309", bg:"#fef3c7", component:AppEstrategico },
  { id:"financeiro",  label:"Financeiro",  icon:"[F]", desc:"CMV, DRE e analises financeiras",          color:"#15803d", bg:"#dcfce7", component:AppFinanceiro  },
  { id:"operacoes",   label:"Operacoes",   icon:"[O]", desc:"Caixa, custos e emprestimos",               color:"#1d4ed8", bg:"#dbeafe", component:AppOperacoes   },
  { id:"vendas",      label:"Vendas",      icon:"[V]", desc:"Estoque, producao e desperdicio",           color:"#16a34a", bg:"#f0fdf4", component:AppVendas      },
  { id:"equipe",      label:"Equipe",      icon:"[Q]", desc:"Gestao de colaboradores",                   color:"#7c3aed", bg:"#ede9fe", component:AppEquipe      },
];

function LoginScreen() {
  const [loading, setLoading] = useState(false);
  const handleGoogle = async () => {
    setLoading(true);
    await supabase.auth.signInWithOAuth({ provider: "google", options: { redirectTo: window.location.origin } });
    setLoading(false);
  };
  return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#fef3c7 0%,#f8f7f2 60%,#dcfce7 100%)"}}>
      <div style={{background:"#fff",borderRadius:20,padding:"52px 44px",maxWidth:420,width:"100%",textAlign:"center",boxShadow:"0 8px 40px rgba(0,0,0,0.10)"}}>
        <div style={{fontSize:52,marginBottom:8}}>*</div>
        <h1 style={{fontFamily:"Georgia,serif",fontSize:28,color:"#1a1512",margin:"0 0 4px"}}>Dona Luzia</h1>
        <p style={{color:"#6b5f4e",fontSize:14,marginBottom:36}}>Culinaria Saudavel - Painel de Gestao</p>
        <button onClick={handleGoogle} disabled={loading} style={{display:"flex",alignItems:"center",justifyContent:"center",gap:12,width:"100%",padding:"14px 20px",borderRadius:12,border:"1.5px solid #e2ddd4",background:"#fff",fontSize:15,fontWeight:600,color:"#1a1512",cursor:"pointer"}}>
          {loading ? "Entrando..." : "Entrar com Google"}
        </button>
      </div>
    </div>
  );
}

function MenuScreen({ user, onSelect, onLogout }: { user: any, onSelect: (id: string) => void, onLogout: () => void }) {
  const firstName = (user?.user_metadata?.full_name || user?.email || "").split(" ")[0];
  return (
    <div>
      <div style={{background:"#fff",borderBottom:"1.5px solid #e2ddd4",padding:"0 28px",height:62,display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100}}>
        <span style={{fontFamily:"Georgia,serif",fontSize:18,color:"#b45309",fontWeight:700}}>Dona Luzia</span>
        <div style={{display:"flex",alignItems:"center",gap:14}}>
          <span style={{fontSize:13,color:"#6b5f4e"}}>{user?.user_metadata?.full_name || user?.email}</span>
          <button onClick={onLogout} style={{padding:"7px 16px",borderRadius:8,border:"1.5px solid #e2ddd4",background:"#fff",fontSize:13,color:"#6b5f4e",cursor:"pointer",fontWeight:600}}>Sair</button>
        </div>
      </div>
      <div style={{maxWidth:960,margin:"0 auto",padding:"40px 24px"}}>
        <h2 style={{fontFamily:"Georgia,serif",fontSize:26,color:"#1a1512",marginBottom:6}}>Ola, {firstName}!</h2>
        <p style={{color:"#6b5f4e",fontSize:14,marginBottom:36}}>Selecione um modulo para comecar</p>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill, minmax(260px, 1fr))",gap:18}}>
          {MODULES.map(m => (
            <div key={m.id} onClick={() => onSelect(m.id)} style={{background:"#fff",border:"1.5px solid #e2ddd4",borderRadius:16,padding:"24px 22px",cursor:"pointer",display:"flex",flexDirection:"column",gap:10}}>
              <div style={{width:46,height:46,borderRadius:12,background:m.bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,fontWeight:700,color:m.color}}>{m.icon}</div>
              <div style={{fontFamily:"Georgia,serif",fontSize:19,color:m.color,fontWeight:700}}>{m.label}</div>
              <div style={{fontSize:13,color:"#6b5f4e",lineHeight:1.5}}>{m.desc}</div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [session, setSession] = useState<any>(null);
  const [loading, setLoading] = useState(true);
  const [activeModule, setActiveModule] = useState<string | null>(null);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => { setSession(data.session); setLoading(false); });
    const { data: listener } = supabase.auth.onAuthStateChange((_event, sess) => { setSession(sess); setLoading(false); });
    return () => listener.subscription.unsubscribe();
  }, []);

  const handleLogout = async () => { await supabase.auth.signOut(); setActiveModule(null); };

  if (loading) return <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center"}}><span>Carregando...</span></div>;
  if (!session) return <LoginScreen />;

  if (activeModule) {
    const mod = MODULES.find(m => m.id === activeModule)!;
    const ModuleComponent = mod.component;
    return (
      <div style={{minHeight:"100vh",background:"#f8f7f2"}}>
        <div style={{background:"#fff",borderBottom:"1.5px solid #e2ddd4",padding:"0 24px",height:60,display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <button onClick={() => setActiveModule(null)} style={{padding:"7px 14px",borderRadius:8,border:"1.5px solid #b45309",background:"#fef3c7",fontSize:13,color:"#b45309",cursor:"pointer",fontWeight:600}}>Menu</button>
            <span style={{fontFamily:"Georgia,serif",fontSize:17,color:mod.color,fontWeight:700}}>{mod.label}</span>
          </div>
          <button onClick={handleLogout} style={{padding:"7px 14px",borderRadius:8,border:"1.5px solid #e2ddd4",background:"#fff",fontSize:13,color:"#6b5f4e",cursor:"pointer",fontWeight:600}}>Sair</button>
        </div>
        <ModuleComponent />
      </div>
    );
  }

  return <div style={{minHeight:"100vh",background:"#f8f7f2"}}><MenuScreen user={session.user} onSelect={setActiveModule} onLogout={handleLogout} /></div>;
}

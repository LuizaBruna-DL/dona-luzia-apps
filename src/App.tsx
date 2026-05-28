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
  { id:"estrategico", label:"Estratégico", icon:"📊", desc:"Dashboard geral, relatórios e exportação", color:"#b45309", bg:"#fef3c7", component:AppEstrategico },
  { id:"financeiro",  label:"Financeiro",  icon:"💰", desc:"CMV, DRE e análises financeiras",          color:"#15803d", bg:"#dcfce7", component:AppFinanceiro  },
  { id:"operacoes",   label:"Operações",   icon:"⚙️", desc:"Caixa, custos e empréstimos",               color:"#1d4ed8", bg:"#dbeafe", component:AppOperacoes   },
  { id:"vendas",      label:"Vendas",      icon:"🥗", desc:"Estoque, produção e desperdício",           color:"#16a34a", bg:"#f0fdf4", component:AppVendas      },
  { id:"equipe",      label:"Equipe",      icon:"👥", desc:"Gestão de colaboradores",                   color:"#7c3aed", bg:"#ede9fe", component:AppEquipe      },
];
function LoginScreen() {
  const [loading, setLoading] = useState(false);
  const handleGoogle = async () => {
    setLoading(true);
    await supabase.auth.signInWithOAuth({
      provider: "google",
      options: { redirectTo: window.location.origin },
    });
    setLoading(false);
  };
  return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#fef3c7 0%,#f8f7f2 60%,#dcfce7 100%)"}}>
      <div style={{background:"#fff",borderRadius:20,padding:"52px 44px",maxWidth:420,width:"100%",textAlign:"center",boxShadow:"0 8px 40px rgba(0,0,0,0.10)"}}>
        <div style={{fontSize:52,marginBottom:8}}>🥗</div>
        <h1 style={{fontFamily:"Georgia,serif",fontSize:28,color:"#1a1512",margin:"0 0 4px"}}>Dona Luzia</h1>
        <p style={{color:"#6b5f4e",fontSize:14,marginBottom:36}}>Culinária Saudável · Painel de Gestão</p>
        <button onClick={handleGoogle} disabled={loading} style={{display:"flex",alignItems:"center",justifyContent:"center",gap:12,width:"100%",padding:"14px 20px",borderRadius:12,border:"1.5px solid #e2ddd4",background:"#fff",fontSize:15,fontWeight:600,color:"#1a1512",cursor:"pointer"}}>
          <svg width="20" height="20" viewBox="0 0 48 48"><path fill="#EA4335" d="M24 9.5c3.5 0 6.6 1.2 9 3.2l6.7-6.7C35.7 2.5 30.2 0 24 0 14.8 0 6.9 5.4 3 13.3l7.8 6C12.8 13.1 17.9 9.5 24 9.5z"/><path fill="#4285F4" d="M46.5 24.5c0-1.6-.1-3.1-.4-4.5H24v8.5h12.7c-.6 3-2.3 5.5-4.8 7.2l7.5 5.8c4.4-4 7.1-10 7.1-17z"/><path fill="#FBBC05" d="M10.8 28.7A14.8 14.8 0 0 1 9.5 24c0-1.7.3-3.3.8-4.8L2.5 13.3A24 24 0 0 0 0 24c0 3.8.9 7.4 2.5 10.6l8.3
            function MenuScreen({ user, onSelect, onLogout }) {
  const firstName = (user?.user_metadata?.full_name || user?.email || "").split(" ")[0];
  return (
    <div>
      <div style={{background:"#fff",borderBottom:"1.5px solid #e2ddd4",padding:"0 28px",height:62,display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:22}}>🥗</span>
          <span
            export default function App() {
  const [session, setSession] = useState(null);
  const [loading, setLoading] = useState(true);
  const [activeModule, setActiveModule] = useState(null);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session);
      setLoading(false);
    });
    const { data: listener } = supabase.auth.onAuthStateChange((_event, sess) => {
      setSession(sess);
      setLoading(false);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  const handleLogout = async () => {
    await supabase.auth.signOut();
    setActiveModule(null);
  };

  if (loading) return <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center"}}><span>Carregando…</span></div>;
  if (!session) return <LoginScreen />;

  if (activeModule) {
    const mod = MODULES.find(m => m.id === activeModule);
    const ModuleComponent = mod.component;
    return (
      <div style={{minHeight:"100vh",background:"#f8f7f2"}}>
        <div style={{background:"#fff",borderBottom:"1.5px solid #e2ddd4",padding:"0 24px",height:60,display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <button onClick={() => setActiveModule(null)} style={{padding:"7px 14px",borderRadius:8,border:"1.5px solid #b45309",background:"#fef3c7",fontSize:13,color:"#b45309",cursor:"pointer",fontWeight:600}}>← Menu</button>
            <span style={{fontSize:18}}>{mod.icon}</span>
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

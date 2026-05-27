import React, { useState, useEffect } from 'react';
import { supabase } from './lib/supabase';

function App() {
  const [user, setUser] = useState<any>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setUser(session?.user ?? null);
      setLoading(false);
    });

    supabase.auth.onAuthStateChange((_event, session) => {
      setUser(session?.user ?? null);
    });
  }, []);

  const handleGoogleLogin = async () => {
    await supabase.auth.signInWithOAuth({
      provider: 'google',
      options: {
        redirectTo: window.location.origin
      }
    });
  };

  const handleLogout = async () => {
    await supabase.auth.signOut();
  };

  if (loading) return <div>Carregando...</div>;

  return (
    <div style={{ textAlign: 'center', padding: '50px' }}>
      <h1>🌿 Dona Luzia</h1>
      <h2>Culinária Saudável</h2>
      {user ? (
        <div>
          <p>Bem-vinda, {user.email}! 👋</p>
          <button onClick={handleLogout}>Sair</button>
        </div>
      ) : (
        <button onClick={handleGoogleLogin}>
          Entrar com Google
        </button>
      )}
    </div>
  );
}

export default App;

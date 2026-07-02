"use client";

import { createContext, useContext, useState, useEffect } from 'react';
import { supabase } from '@/lib/supabase';
import { useRouter, usePathname } from 'next/navigation';

const AuthContext = createContext({});

export function AuthProvider({ children }) {
  const [user, setUser] = useState(null);
  const [loading, setLoading] = useState(true);
  const router = useRouter();
  const pathname = usePathname();

  useEffect(() => {
    // Load user from localStorage
    const savedUser = localStorage.getItem('bp_user');
    if (savedUser) {
      try {
        setUser(jsonParseSecure(savedUser));
      } catch (e) {
        localStorage.removeItem('bp_user');
      }
    }
    setLoading(false);
  }, []);

  function jsonParseSecure(str) {
    try {
      return JSON.parse(str);
    } catch (e) {
      return null;
    }
  }

  // Redirect on auth change or route change
  useEffect(() => {
    if (loading) return;
    
    const isLoginPage = pathname === '/login';
    if (!user && !isLoginPage) {
      router.push('/login');
    } else if (user && isLoginPage) {
      router.push('/');
    }
  }, [user, pathname, loading, router]);

  const login = async (email, password) => {
    try {
      const { data, error } = await supabase
        .from('users')
        .select('*')
        .eq('email', email.trim().toLowerCase())
        .eq('password', password)
        .single();

      if (error || !data) {
        return { success: false, message: 'Credenciales inválidas' };
      }

      const userData = {
        email: data.email,
        name: data.name,
        role: data.role
      };

      setUser(userData);
      localStorage.setItem('bp_user', JSON.stringify(userData));
      router.push('/');
      return { success: true };
    } catch (e) {
      return { success: false, message: 'Error de servidor. Intenta de nuevo.' };
    }
  };

  const logout = () => {
    setUser(null);
    localStorage.removeItem('bp_user');
    router.push('/login');
  };

  return (
    <AuthContext.Provider value={{ user, role: user?.role, login, logout, loading }}>
      {!loading && children}
    </AuthContext.Provider>
  );
}

export const useAuth = () => useContext(AuthContext);

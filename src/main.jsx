import React from 'react'
import ReactDOM from 'react-dom/client'
import { createClient } from '@supabase/supabase-js'
import { SUPABASE_URL, SUPABASE_ANON_KEY } from './supabase.js'
import App from './App.jsx'

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY)

// ── Almacenamiento (la app usa window.storage; lo redirigimos a Supabase) ──
window.storage = {
  async get(key) {
    const { data, error } = await supabase
      .from('plaat_data')
      .select('value')
      .eq('key', key)
      .maybeSingle()
    if (error) throw error
    return data ? { value: data.value } : null
  },
  async set(key, value) {
    const { error } = await supabase
      .from('plaat_data')
      .upsert({ key, value })
    if (error) throw error
    return { value }
  },
}

// ── Autenticación (login con correo y contraseña) ──
window.auth = {
  async signIn(email, password) {
    const { data, error } = await supabase.auth.signInWithPassword({ email, password })
    if (error) throw error
    return data
  },
  async signOut() {
    await supabase.auth.signOut()
  },
  async getUser() {
    const { data } = await supabase.auth.getUser()
    return data?.user || null
  },
  onChange(cb) {
    supabase.auth.onAuthStateChange((_event, session) => cb(session?.user || null))
  },
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)

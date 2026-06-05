import React from 'react'
import ReactDOM from 'react-dom/client'
import { createClient } from '@supabase/supabase-js'
import { SUPABASE_URL, SUPABASE_ANON_KEY } from './supabase.js'
import App from './App.jsx'

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY)

// La app usa window.storage (que solo existe dentro de Claude).
// Fuera de Claude lo redirigimos a Supabase con la MISMA interfaz,
// así no hay que tocar nada de App.jsx.
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

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)

import React from 'react'
import ReactDOM from 'react-dom/client'
import { createClient } from '@supabase/supabase-js'
import { SUPABASE_URL, SUPABASE_ANON_KEY } from './supabase.js'
import App from './App.jsx'

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  realtime: { params: { eventsPerSecond: 10 } },
})

// ── Storage legacy (migración) ───────────────────────────────────────────────
window.storage = {
  async get(key) {
    const { data, error } = await supabase
      .from('plaat_data').select('value').eq('key', key).maybeSingle()
    if (error) throw error
    return data ? { value: data.value } : null
  },
  async set(key, value) {
    const { error } = await supabase.from('plaat_data').upsert({ key, value })
    if (error) throw error
    return { value }
  },
  async delete(key) {
    const { error } = await supabase.from('plaat_data').delete().eq('key', key)
    if (error) throw error
    return { deleted: true }
  },
}

// ── DB ───────────────────────────────────────────────────────────────────────
window.db = {

  async crearObraConOwner(obraId, obraData) {
    const { error } = await supabase.rpc('crear_obra_con_owner', {
      p_obra_id: obraId,
      p_obra_data: obraData,
    });
    if (error) throw error;
  },

  // ─ Obras (RLS via obra_usuarios) ─────────────────────────────────────────
  async getObras() {
    // RLS filtra automáticamente las obras a las que el usuario tiene acceso
    const { data, error } = await supabase
      .from('obras').select('*').order('created_at', { ascending: false })
    if (error) throw error
    return data || []
  },

  async upsertObra(row) {
    const { error } = await supabase.from('obras').upsert(row)
    if (error) throw error
  },

  async deleteObra(id) {
    // Solo funciona si el usuario es owner (RLS lo bloquea si no)
    const { error } = await supabase.from('obras').delete().eq('id', id)
    if (error) throw error
  },

  // ─ Permisos: obra_usuarios ───────────────────────────────────────────────
  async getUsuariosObra(obraId) {
    const { data, error } = await supabase
      .from('obra_usuarios')
      .select('user_id, rol, created_at, invitado_por')
      .eq('obra_id', obraId)
    if (error) throw error
    return data || []
  },

  async addOwner(obraId, userId) {
    const { error } = await supabase.from('obra_usuarios').upsert({
      obra_id: obraId, user_id: userId, rol: 'owner', invitado_por: userId,
    })
    if (error) throw error
  },

  async invitarUsuario(obraId, emailInvitado, invitadoPorId) {
    // Busca el user_id por email via función RPC
    const { data, error } = await supabase.rpc('get_user_id_by_email', {
      email_input: emailInvitado.trim().toLowerCase()
    })
    if (error) throw new Error('No se encontró ningún usuario con ese email')
    if (!data) throw new Error('No se encontró ningún usuario con ese email')
    const { error: e2 } = await supabase.from('obra_usuarios').upsert({
      obra_id: obraId, user_id: data, rol: 'editor', invitado_por: invitadoPorId,
    })
    if (e2) throw e2
    return data
  },

  async quitarAcceso(obraId, userId) {
    const { error } = await supabase.from('obra_usuarios')
      .delete().eq('obra_id', obraId).eq('user_id', userId)
    if (error) throw error
  },

  async getRolUsuario(obraId, userId) {
    const { data } = await supabase.from('obra_usuarios')
      .select('rol').eq('obra_id', obraId).eq('user_id', userId).maybeSingle()
    return data?.rol || null
  },

  // ─ Fotos en Storage ──────────────────────────────────────────────────────
  async subirFoto(obraId, fotoId, base64) {
    // Extraer tipo y datos del base64
    const [meta, data] = base64.split(',');
    const mime = meta.match(/:(.*?);/)[1];
    const ext  = mime.split('/')[1] || 'jpg';
    const path = `${obraId}/${fotoId}.${ext}`;
    // Convertir base64 a Blob
    const bytes = atob(data);
    const arr   = new Uint8Array(bytes.length);
    for (let i = 0; i < bytes.length; i++) arr[i] = bytes.charCodeAt(i);
    const blob  = new Blob([arr], { type: mime });
    const { error } = await supabase.storage.from('plaat-fotos').upload(path, blob, { upsert: true, contentType: mime });
    if (error) throw error;
    // URL pública — no caduca mai (el bucket plaat-fotos ha d'estar en mode públic)
    const { data: urlData } = supabase.storage.from('plaat-fotos').getPublicUrl(path);
    return { path, url: urlData?.publicUrl };
  },

  async getFotoUrl(path) {
    // URL pública — no caduca mai
    const { data } = supabase.storage.from('plaat-fotos').getPublicUrl(path);
    return data?.publicUrl;
  },

  async eliminarFoto(path) {
    const { error } = await supabase.storage.from('plaat-fotos').remove([path]);
    if (error) console.error('Error eliminando foto:', error);
  },

  // ─ Perfiles de usuario ───────────────────────────────────────────────────
  async getPerfiles() {
    const { data, error } = await supabase.from('perfiles').select('user_id, nombre');
    if (error) throw error;
    return data || [];
  },
  async upsertPerfil(userId, nombre) {
    const { error } = await supabase.from('perfiles').upsert({ user_id: userId, nombre, updated_at: new Date().toISOString() });
    if (error) throw error;
  },

  // ─ Módulos independientes ────────────────────────────────────────────────
  async getModulo(tabla, obraId) {
    const { data, error } = await supabase
      .from(tabla).select('*').eq('obra_id', obraId)
    if (error) throw error
    return data || []
  },

  async upsertModulo(tabla, row) {
    const { error } = await supabase.from(tabla).upsert(row)
    if (error) throw error
  },

  async deleteModulo(tabla, id) {
    const { error } = await supabase.from(tabla).delete().eq('id', id)
    if (error) throw error
  },

  // ─ Seguimiento global ────────────────────────────────────────────────────
  async getSeguimiento() {
    const { data, error } = await supabase
      .from('seguimiento').select('*').order('updated_at', { ascending: true });
    if (error) throw error;
    return (data || []).map(r => r.data);
  },
  async upsertPuntoSeg(punto) {
    const { error } = await supabase.from('seguimiento').upsert({
      id: punto.id, data: punto, updated_at: new Date().toISOString()
    });
    if (error) throw error;
  },
  async deletePuntoSeg(id) {
    const { error } = await supabase.from('seguimiento').delete().eq('id', id);
    if (error) throw error;
  },
  subscribeSeguimiento(onCambio) {
    const canal = supabase.channel('seguimiento-global');
    canal.on('postgres_changes', { event: '*', schema: 'public', table: 'seguimiento' },
      () => onCambio());
    canal.subscribe();
    return () => supabase.removeChannel(canal);
  },
  subscribeObra(obraId, onCambio) {
    const canal = supabase.channel(`obra-${obraId}`)
    const tablas = ['obras', 'incidencias', 'actas_vo', 'actas_insp', 'notas', 'calidad']
    tablas.forEach(tabla => {
      canal.on('postgres_changes', {
        event: '*', schema: 'public', table: tabla,
        filter: tabla === 'obras' ? `id=eq.${obraId}` : `obra_id=eq.${obraId}`,
      }, payload => onCambio(tabla, payload))
    })
    canal.subscribe()
    return () => supabase.removeChannel(canal)
  },

  subscribeListaObras(onCambio) {
    const canal = supabase.channel('obras-lista')
    canal.on('postgres_changes', {
      event: '*', schema: 'public', table: 'obra_usuarios',
    }, () => onCambio())
    canal.subscribe()
    return () => supabase.removeChannel(canal)
  },
}

// ── Auth ─────────────────────────────────────────────────────────────────────
window.auth = {
  async signIn(email, password) {
    const { data, error } = await supabase.auth.signInWithPassword({ email, password })
    if (error) throw error
    return data
  },
  async signOut() { await supabase.auth.signOut() },
  async getUser() {
    const { data } = await supabase.auth.getUser()
    return data?.user || null
  },
  onChange(cb) {
    supabase.auth.onAuthStateChange((_e, session) => cb(session?.user || null))
  },
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode><App /></React.StrictMode>,
)

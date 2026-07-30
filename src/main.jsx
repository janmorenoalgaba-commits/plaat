import React from 'react'
import ReactDOM from 'react-dom/client'
import { createClient } from '@supabase/supabase-js'
import { SUPABASE_URL, SUPABASE_ANON_KEY } from './supabase.js'
import App from './App.jsx'

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  realtime: { params: { eventsPerSecond: 10 } },
})

window._supabaseUrl = SUPABASE_URL

// ═══════════════════════════════════════════════════════════════════════════
//  CACHÉ LOCAL (IndexedDB)
//  Dos magatzems: `kv` per a dades, `queue` per a mutacions pendents.
// ═══════════════════════════════════════════════════════════════════════════
const IDB_NAME = 'plaat-offline'
const IDB_VER  = 1

let _dbPromise = null
function idb() {
  if (_dbPromise) return _dbPromise
  _dbPromise = new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, IDB_VER)
    req.onupgradeneeded = () => {
      const db = req.result
      if (!db.objectStoreNames.contains('kv'))    db.createObjectStore('kv')
      if (!db.objectStoreNames.contains('queue')) db.createObjectStore('queue', { keyPath: 'qid' })
    }
    req.onsuccess = () => resolve(req.result)
    req.onerror   = () => reject(req.error)
  })
  return _dbPromise
}

async function idbGet(store, key) {
  try {
    const db = await idb()
    return await new Promise((res, rej) => {
      const tx = db.transaction(store, 'readonly')
      const r  = tx.objectStore(store).get(key)
      r.onsuccess = () => res(r.result)
      r.onerror   = () => rej(r.error)
    })
  } catch { return undefined }
}

async function idbSet(store, key, val) {
  try {
    const db = await idb()
    return await new Promise((res, rej) => {
      const tx = db.transaction(store, 'readwrite')
      const os = tx.objectStore(store)
      const r  = store === 'queue' ? os.put(val) : os.put(val, key)
      r.onsuccess = () => res(true)
      r.onerror   = () => rej(r.error)
    })
  } catch { return false }
}

async function idbDel(store, key) {
  try {
    const db = await idb()
    return await new Promise((res, rej) => {
      const tx = db.transaction(store, 'readwrite')
      const r  = tx.objectStore(store).delete(key)
      r.onsuccess = () => res(true)
      r.onerror   = () => rej(r.error)
    })
  } catch { return false }
}

async function idbAll(store) {
  try {
    const db = await idb()
    return await new Promise((res, rej) => {
      const tx = db.transaction(store, 'readonly')
      const r  = tx.objectStore(store).getAll()
      r.onsuccess = () => res(r.result || [])
      r.onerror   = () => rej(r.error)
    })
  } catch { return [] }
}

// ═══════════════════════════════════════════════════════════════════════════
//  ESTAT DE CONNEXIÓ
// ═══════════════════════════════════════════════════════════════════════════
const net = {
  online: typeof navigator !== 'undefined' ? navigator.onLine !== false : true,
  pendents: 0,
  sincronitzant: false,
  listeners: new Set(),
}

function notificar() {
  const snap = { online: net.online, pendents: net.pendents, sincronitzant: net.sincronitzant }
  net.listeners.forEach(fn => { try { fn(snap) } catch {} })
}

function marcarOffline() {
  if (net.online) { net.online = false; notificar() }
}
function marcarOnline() {
  if (!net.online) { net.online = true; notificar() }
  flushCua()
}

if (typeof window !== 'undefined') {
  window.addEventListener('online',  marcarOnline)
  window.addEventListener('offline', marcarOffline)
}

// Un error és "de xarxa" (recuperable) i no de dades (permanent)?
function esErrorDeXarxa(e) {
  if (!e) return false
  const m = (e.message || String(e)).toLowerCase()
  return m.includes('failed to fetch') || m.includes('networkerror') ||
         m.includes('load failed')     || m.includes('timeout')      ||
         m.includes('network request')  || e.name === 'TypeError' ||
         e.name === 'AbortError'
}

// ═══════════════════════════════════════════════════════════════════════════
//  CUA DE MUTACIONS
//  Cada element: { qid, key, op, tabla, payload, ts }
//  `key` col·lapsa escriptures repetides de la mateixa fila (guanya l'última).
// ═══════════════════════════════════════════════════════════════════════════
async function encuar(op, tabla, payload, key) {
  const qid = `${key}` // la clau és l'identitat: reescriure substitueix
  await idbSet('queue', qid, { qid, key, op, tabla, payload, ts: Date.now() })
  const q = await idbAll('queue')
  net.pendents = q.length
  notificar()
}

let _flushing = false
async function flushCua() {
  if (_flushing || !net.online) return
  const cua = await idbAll('queue')
  if (!cua.length) { net.pendents = 0; notificar(); return }

  _flushing = true
  net.sincronitzant = true
  notificar()

  cua.sort((a, b) => a.ts - b.ts)
  let fallides = 0

  for (const m of cua) {
    try {
      await executar(m)
      await idbDel('queue', m.qid)
    } catch (e) {
      if (esErrorDeXarxa(e)) { marcarOffline(); fallides++; break }
      // Error permanent (dades invàlides, permisos): descartar per no bloquejar la cua
      console.warn('[sync] mutació descartada:', m.op, m.tabla, e?.message)
      await idbDel('queue', m.qid)
    }
  }

  const rest = await idbAll('queue')
  net.pendents = rest.length
  net.sincronitzant = false
  _flushing = false
  notificar()

  if (rest.length && net.online && !fallides) setTimeout(flushCua, 3000)
}

async function executar(m) {
  switch (m.op) {
    case 'upsertObra':    { const { error } = await supabase.from('obras').upsert(m.payload); if (error) throw error; break }
    case 'deleteObra':    { const { error } = await supabase.from('obras').delete().eq('id', m.payload.id); if (error) throw error; break }
    case 'upsertModulo':  { const { error } = await supabase.from(m.tabla).upsert(m.payload); if (error) throw error; break }
    case 'deleteModulo':  { const { error } = await supabase.from(m.tabla).delete().eq('id', m.payload.id); if (error) throw error; break }
    case 'upsertPuntoSeg':{ const { error } = await supabase.from('seguimiento').upsert(m.payload); if (error) throw error; break }
    case 'deletePuntoSeg':{ const { error } = await supabase.from('seguimiento').delete().eq('id', m.payload.id); if (error) throw error; break }
    case 'crearObra':     { const { error } = await supabase.rpc('crear_obra_con_owner', m.payload); if (error) throw error; break }
    case 'upsertPerfil':  { const { error } = await supabase.from('perfiles').upsert(m.payload); if (error) throw error; break }
    case 'subirFoto':     { await pujarFotoReal(m.payload.obraId, m.payload.fotoId, m.payload.base64); break }
    default: console.warn('[sync] operació desconeguda:', m.op)
  }
}

// Escriptura: local sempre, xarxa si es pot, cua si falla.
async function escriure(op, tabla, payload, key, cacheFn) {
  if (cacheFn) { try { await cacheFn() } catch {} }
  if (!net.online) { await encuar(op, tabla, payload, key); return { encuat: true } }
  try {
    await executar({ op, tabla, payload })
    return { encuat: false }
  } catch (e) {
    if (esErrorDeXarxa(e)) {
      marcarOffline()
      await encuar(op, tabla, payload, key)
      return { encuat: true }
    }
    throw e
  }
}

// Lectura: xarxa si es pot (i refresca caché), si no la caché.
async function llegir(cacheKey, fetchFn) {
  if (net.online) {
    try {
      const dades = await fetchFn()
      idbSet('kv', cacheKey, dades)
      return dades
    } catch (e) {
      if (!esErrorDeXarxa(e)) throw e
      marcarOffline()
    }
  }
  const cached = await idbGet('kv', cacheKey)
  return cached !== undefined ? cached : null
}

// ═══════════════════════════════════════════════════════════════════════════
//  FOTOS
// ═══════════════════════════════════════════════════════════════════════════
async function pujarFotoReal(obraId, fotoId, base64) {
  const [meta, data] = base64.split(',')
  const mime = meta.match(/:(.*?);/)[1]
  const ext  = mime.split('/')[1] || 'jpg'
  const path = `${obraId}/${fotoId}.${ext}`
  const bytes = atob(data)
  const arr   = new Uint8Array(bytes.length)
  for (let i = 0; i < bytes.length; i++) arr[i] = bytes.charCodeAt(i)
  const blob  = new Blob([arr], { type: mime })
  const { error } = await supabase.storage.from('plaat-fotos')
    .upload(path, blob, { upsert: true, contentType: mime })
  if (error) throw error
  const { data: urlData } = supabase.storage.from('plaat-fotos').getPublicUrl(path)
  return { path, url: urlData?.publicUrl }
}

// ── Storage legacy ───────────────────────────────────────────────────────────
window.storage = {
  async get(key) {
    const { data, error } = await supabase.from('plaat_data').select('value').eq('key', key).maybeSingle()
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

// ═══════════════════════════════════════════════════════════════════════════
//  API PÚBLICA — mateixa signatura que abans
// ═══════════════════════════════════════════════════════════════════════════
window.db = {

  // ─ Estat de sincronització ───────────────────────────────────────────────
  estatXarxa: () => ({ online: net.online, pendents: net.pendents, sincronitzant: net.sincronitzant }),
  onCanviXarxa(fn) { net.listeners.add(fn); return () => net.listeners.delete(fn) },
  sincronitzarAra: () => flushCua(),

  async crearObraConOwner(obraId, obraData) {
    return escriure('crearObra', null, { p_obra_id: obraId, p_obra_data: obraData }, `crearObra:${obraId}`)
  },

  // ─ Obres ─────────────────────────────────────────────────────────────────
  async getObras() {
    const r = await llegir('obras', async () => {
      const { data, error } = await supabase.from('obras').select('*').order('created_at', { ascending: false })
      if (error) throw error
      return data || []
    })
    return r || []
  },

  async upsertObra(row) {
    // Refresca la còpia local de la llista d'obres
    const cacheFn = async () => {
      const llista = (await idbGet('kv', 'obras')) || []
      const i = llista.findIndex(o => o.id === row.id)
      if (i >= 0) llista[i] = { ...llista[i], ...row }; else llista.unshift(row)
      await idbSet('kv', 'obras', llista)
    }
    return escriure('upsertObra', 'obras', row, `upsertObra:${row.id}`, cacheFn)
  },

  async deleteObra(id) {
    const cacheFn = async () => {
      const llista = (await idbGet('kv', 'obras')) || []
      await idbSet('kv', 'obras', llista.filter(o => o.id !== id))
    }
    return escriure('deleteObra', 'obras', { id }, `deleteObra:${id}`, cacheFn)
  },

  // ─ Permisos ──────────────────────────────────────────────────────────────
  async getUsuariosObra(obraId) {
    const r = await llegir(`usuarios:${obraId}`, async () => {
      const { data, error } = await supabase.from('obra_usuarios')
        .select('user_id, rol, created_at, invitado_por').eq('obra_id', obraId)
      if (error) throw error
      return data || []
    })
    return r || []
  },

  async addOwner(obraId, userId) {
    const { error } = await supabase.from('obra_usuarios')
      .upsert({ obra_id: obraId, user_id: userId, rol: 'owner', invitado_por: userId })
    if (error) throw error
  },

  async invitarUsuario(obraId, emailInvitado, invitadoPorId) {
    const { data, error } = await supabase.rpc('get_user_id_by_email', {
      email_input: emailInvitado.trim().toLowerCase()
    })
    if (error || !data) throw new Error('No s\u2019ha trobat cap usuari amb aquest email')
    const { error: e2 } = await supabase.from('obra_usuarios')
      .upsert({ obra_id: obraId, user_id: data, rol: 'editor', invitado_por: invitadoPorId })
    if (e2) throw e2
    return data
  },

  async quitarAcceso(obraId, userId) {
    const { error } = await supabase.from('obra_usuarios').delete().eq('obra_id', obraId).eq('user_id', userId)
    if (error) throw error
  },

  async getRolUsuario(obraId, userId) {
    const r = await llegir(`rol:${obraId}:${userId}`, async () => {
      const { data } = await supabase.from('obra_usuarios')
        .select('rol').eq('obra_id', obraId).eq('user_id', userId).maybeSingle()
      return data?.rol || null
    })
    return r
  },

  // Tots els rols de l'usuari en una sola consulta
  async getRolesUsuario(userId) {
    const r = await llegir(`roles:${userId}`, async () => {
      const { data, error } = await supabase.from('obra_usuarios')
        .select('obra_id, rol').eq('user_id', userId)
      if (error) throw error
      const mapa = {}
      ;(data || []).forEach(x => { mapa[x.obra_id] = x.rol })
      return mapa
    })
    return r || {}
  },

  // ─ Fotos ─────────────────────────────────────────────────────────────────
  async subirFoto(obraId, fotoId, base64) {
    const ext  = (base64.split(',')[0].match(/:(.*?);/)?.[1] || 'image/jpeg').split('/')[1] || 'jpg'
    const path = `${obraId}/${fotoId}.${ext}`
    const url  = supabase.storage.from('plaat-fotos').getPublicUrl(path).data?.publicUrl

    if (!net.online) {
      // Guarda el base64 localment i encua la pujada; la UI ja pot mostrar la foto
      await idbSet('kv', `foto:${path}`, base64)
      await encuar('subirFoto', null, { obraId, fotoId, base64 }, `subirFoto:${path}`)
      return { path, url, pendent: true }
    }
    try {
      return await pujarFotoReal(obraId, fotoId, base64)
    } catch (e) {
      if (!esErrorDeXarxa(e)) throw e
      marcarOffline()
      await idbSet('kv', `foto:${path}`, base64)
      await encuar('subirFoto', null, { obraId, fotoId, base64 }, `subirFoto:${path}`)
      return { path, url, pendent: true }
    }
  },

  async getFotoUrl(path) {
    return supabase.storage.from('plaat-fotos').getPublicUrl(path).data?.publicUrl
  },

  // base64 local d'una foto encara no pujada (per pintar-la sense connexió)
  async getFotoLocal(path) { return idbGet('kv', `foto:${path}`) },

  async eliminarFoto(path) {
    try { await supabase.storage.from('plaat-fotos').remove([path]) }
    catch (e) { console.error('Error eliminant foto:', e) }
    idbDel('kv', `foto:${path}`)
  },

  // ─ Perfils ───────────────────────────────────────────────────────────────
  async getPerfiles() {
    const r = await llegir('perfiles', async () => {
      const { data, error } = await supabase.from('perfiles').select('user_id, nombre')
      if (error) throw error
      return data || []
    })
    return r || []
  },
  async upsertPerfil(userId, nombre) {
    const row = { user_id: userId, nombre, updated_at: new Date().toISOString() }
    return escriure('upsertPerfil', 'perfiles', row, `upsertPerfil:${userId}`)
  },

  // ─ Mòduls ────────────────────────────────────────────────────────────────
  async getModulo(tabla, obraId) {
    const r = await llegir(`mod:${tabla}:${obraId}`, async () => {
      const { data, error } = await supabase.from(tabla).select('*').eq('obra_id', obraId)
      if (error) throw error
      return data || []
    })
    return r || []
  },

  // ★ Una sola consulta per a TOTES les obres — elimina l'N+1
  async getModulosBatch(tabla, obraIds) {
    if (!obraIds?.length) return {}
    const cacheKey = `modbatch:${tabla}`
    const dades = await llegir(cacheKey, async () => {
      const { data, error } = await supabase.from(tabla).select('*').in('obra_id', obraIds)
      if (error) throw error
      return data || []
    })
    const perObra = {}
    obraIds.forEach(id => { perObra[id] = [] })
    ;(dades || []).forEach(row => {
      if (!perObra[row.obra_id]) perObra[row.obra_id] = []
      perObra[row.obra_id].push(row)
    })
    // Refresca també les caus individuals perquè getModulo funcioni offline
    Object.entries(perObra).forEach(([id, rows]) => idbSet('kv', `mod:${tabla}:${id}`, rows))
    return perObra
  },

  async upsertModulo(tabla, row) {
    const cacheFn = async () => {
      const k    = `mod:${tabla}:${row.obra_id}`
      const rows = (await idbGet('kv', k)) || []
      const i    = rows.findIndex(r => r.id === row.id)
      if (i >= 0) rows[i] = { ...rows[i], ...row }; else rows.push(row)
      await idbSet('kv', k, rows)
    }
    return escriure('upsertModulo', tabla, row, `upsertModulo:${tabla}:${row.id}`, cacheFn)
  },

  async deleteModulo(tabla, id, obraId) {
    const cacheFn = obraId ? async () => {
      const k    = `mod:${tabla}:${obraId}`
      const rows = (await idbGet('kv', k)) || []
      await idbSet('kv', k, rows.filter(r => r.id !== id))
    } : null
    return escriure('deleteModulo', tabla, { id }, `deleteModulo:${tabla}:${id}`, cacheFn)
  },

  // ─ Seguiment global ──────────────────────────────────────────────────────
  async getSeguimiento() {
    const r = await llegir('seguimiento', async () => {
      const { data, error } = await supabase.from('seguimiento').select('*').order('updated_at', { ascending: true })
      if (error) throw error
      return (data || []).map(x => x.data)
    })
    return r || []
  },
  async upsertPuntoSeg(punto) {
    const row = { id: punto.id, data: punto, updated_at: new Date().toISOString() }
    const cacheFn = async () => {
      const llista = (await idbGet('kv', 'seguimiento')) || []
      const i = llista.findIndex(p => p.id === punto.id)
      if (i >= 0) llista[i] = punto; else llista.push(punto)
      await idbSet('kv', 'seguimiento', llista)
    }
    return escriure('upsertPuntoSeg', 'seguimiento', row, `upsertPuntoSeg:${punto.id}`, cacheFn)
  },
  async deletePuntoSeg(id) {
    const cacheFn = async () => {
      const llista = (await idbGet('kv', 'seguimiento')) || []
      await idbSet('kv', 'seguimiento', llista.filter(p => p.id !== id))
    }
    return escriure('deletePuntoSeg', 'seguimiento', { id }, `deletePuntoSeg:${id}`, cacheFn)
  },

  // ─ Realtime (només amb connexió) ─────────────────────────────────────────
  subscribeSeguimiento(onCambio) {
    const canal = supabase.channel('seguimiento-global')
    canal.on('postgres_changes', { event: '*', schema: 'public', table: 'seguimiento' }, () => onCambio())
    canal.subscribe()
    return () => supabase.removeChannel(canal)
  },
  subscribeObra(obraId, onCambio) {
    const canal  = supabase.channel(`obra-${obraId}`)
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
    canal.on('postgres_changes', { event: '*', schema: 'public', table: 'obra_usuarios' }, () => onCambio())
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
    // La sessió persistida és local: permet entrar a l'app sense connexió
    try {
      const { data: s } = await supabase.auth.getSession()
      if (s?.session?.user) return s.session.user
    } catch {}
    if (!net.online) return null
    try {
      const { data } = await supabase.auth.getUser()
      return data?.user || null
    } catch (e) {
      if (esErrorDeXarxa(e)) marcarOffline()
      return null
    }
  },
  onChange(cb) {
    supabase.auth.onAuthStateChange((_e, session) => cb(session?.user || null))
  },
}

// Buida la cua pendent en arrencar
if (typeof window !== 'undefined') {
  idbAll('queue').then(q => { net.pendents = q.length; notificar(); if (q.length) flushCua() })
  setInterval(() => { if (net.online && net.pendents > 0) flushCua() }, 30000)
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode><App /></React.StrictMode>,
)

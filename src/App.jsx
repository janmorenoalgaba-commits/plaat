import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Config ───────────────────────────────────────────────────────────────────

const SK     = 'plaat_deo_v1';         // clave legada (migración)
const SK_IDX = 'plaat_v1_idx';         // índice de IDs de obras
const SK_OBR = id => 'plaat_v1_o_' + id; // datos completos por obra

const RESPONSABLES = [
  'Sergi Castellar','Alex Pla','Ferran Sancho','Adriana de la Barrera',
  'Alex Pallares','Nuria Armengol','Paula Loaiza','Tatiana Acaro',
  'Angela Martin','Agnes Ademà','Roberto Fernandez',
];

const ESTADOS_OBRA = {
  en_curso:   { label: 'En curso',   bg: '#FEF3DB', color: '#7C4A00' },
  pendiente:  { label: 'Pendiente',  bg: '#EEEDE7', color: '#52524E' },
  acabada:    { label: 'Acabada',    bg: '#E8F5E0', color: '#2D5E10' },
  paralizada: { label: 'Paralizada', bg: '#FDECEC', color: '#8A1F1F' },
};

const ESTADOS_INSP = {
  pendiente:    { label: 'Pendiente',    bg: '#EEEDE7', color: '#52524E' },
  inspeccionado:{ label: 'Inspeccionado',bg: '#E8F5E0', color: '#2D5E10' },
  incidencia:   { label: 'Con incidencia',bg:'#FDECEC', color: '#8A1F1F' },
};

const uid      = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const today    = () => new Date().toISOString().slice(0, 10);
const fmtDate  = iso => iso ? new Date(iso).toLocaleDateString('es-ES', { day: '2-digit', month: 'short', year: 'numeric' }) : '—';
const fmtShort = iso => iso ? new Date(iso).toLocaleDateString('es-ES', { day: '2-digit', month: 'short' }) : '—';
const now      = () => new Date().toISOString();

// ─── CSS ──────────────────────────────────────────────────────────────────────

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500;9..40,600;9..40,700&display=swap');
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'DM Sans', -apple-system, sans-serif; background: #F2F1ED; color: #141412; font-size: 14px; }
button, input, select, textarea { font-family: inherit; }
input, select, textarea {
  display: block; width: 100%; padding: 8px 11px;
  border: 1px solid #E0DFD9; border-radius: 8px;
  font-size: 13px; color: #141412; background: #fff;
  outline: none; transition: border-color .15s, box-shadow .15s;
}
input:focus, select:focus, textarea:focus {
  border-color: #141412; box-shadow: 0 0 0 3px rgba(20,20,18,.07);
}
textarea { resize: vertical; min-height: 72px; line-height: 1.5; }
::-webkit-scrollbar { width: 3px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #D0CFC9; border-radius: 2px; }
/* ── Animaciones sutiles ─────────────────────────────── */
@keyframes fi    { from { opacity: 0; transform: translateY(5px); } to { opacity: 1; transform: none; } }
@keyframes pop   { from { opacity: 0; transform: scale(.97);     } to { opacity: 1; transform: none; } }
@keyframes sheet { from { transform: translateY(100%);          } to { transform: none; } }
@keyframes overlay { from { opacity: 0; } to { opacity: 1; } }
@keyframes spin  { to { transform: rotate(360deg); } }

.fade { animation: fi .22s ease both; }

/* Entrada escalonada de listas */
.list-in > * { animation: fi .28s ease both; }
.list-in > *:nth-child(1) { animation-delay: .03s; }
.list-in > *:nth-child(2) { animation-delay: .07s; }
.list-in > *:nth-child(3) { animation-delay: .11s; }
.list-in > *:nth-child(4) { animation-delay: .15s; }
.list-in > *:nth-child(5) { animation-delay: .19s; }
.list-in > *:nth-child(6) { animation-delay: .23s; }
.list-in > *:nth-child(7) { animation-delay: .27s; }
.list-in > *:nth-child(n+8) { animation-delay: .30s; }

.obra-card { transition: box-shadow .2s, transform .14s ease; cursor: pointer; }
.obra-card:hover  { box-shadow: 0 4px 20px rgba(0,0,0,0.1) !important; transform: translateY(-1px); }
.obra-card:active { transform: scale(.985); }

/* Realimentación al tocar */
.tap { transition: transform .14s ease, box-shadow .15s, border-color .15s; }
.tap:active { transform: scale(.98); }

/* Barras de progreso que se rellenan suave */
.bar-fill { transition: width .55s cubic-bezier(.3,.8,.3,1); }

/* Entrada de modales */
.modal-overlay { animation: overlay .2s ease both; }
.modal-in { animation: pop .22s cubic-bezier(.2,.8,.2,1) both; }
.sheet-in { animation: sheet .3s cubic-bezier(.2,.85,.25,1) both; }

@media (prefers-reduced-motion: reduce) {
  .fade, .list-in > *, .modal-in, .sheet-in, .modal-overlay { animation: none !important; }
  .obra-card:active, .tap:active { transform: none !important; }
}

.hov-nav:hover { background: #ECEAE4 !important; }
.hov-row:hover   { background: #F9F8F5 !important; }
.hov-chip:hover  { background: #ECEAE4 !important; }
.no-scrollbar::-webkit-scrollbar { display: none; height: 0; }
.no-scrollbar { -ms-overflow-style: none; scrollbar-width: none; }
`;

// Detecta móvil: pantalla estrecha Y en vertical. En horizontal usa la interfaz de ordenador.
function useIsMobile() {
  const calc = () => {
    if (typeof window === 'undefined') return false;
    const portrait = window.matchMedia
      ? window.matchMedia('(orientation: portrait)').matches
      : window.innerHeight >= window.innerWidth;
    return window.innerWidth < 760 && portrait;
  };
  const [m, setM] = useState(calc);
  useEffect(() => {
    const onR = () => setM(calc());
    window.addEventListener('resize', onR);
    window.addEventListener('orientationchange', onR);
    return () => {
      window.removeEventListener('resize', onR);
      window.removeEventListener('orientationchange', onR);
    };
  }, []);
  return m;
}

// ─── Átomos ───────────────────────────────────────────────────────────────────

function Pill({ label, bg, color }) {
  return <span style={{ fontSize: 11, padding: '2px 9px', borderRadius: 20, fontWeight: 500, background: bg, color, whiteSpace: 'nowrap' }}>{label}</span>;
}

function Btn({ children, onClick, primary, sm, danger, ghost, disabled, full }) {
  const base = {
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 5,
    padding: sm ? '5px 11px' : '7px 14px', borderRadius: 8, border: '1.5px solid',
    cursor: disabled ? 'not-allowed' : 'pointer', fontSize: 13, fontWeight: 500,
    transition: 'opacity .15s', opacity: disabled ? 0.45 : 1, width: full ? '100%' : 'auto',
  };
  const style = primary ? { ...base, background: '#18180F', color: '#fff', borderColor: '#18180F' }
    : danger  ? { ...base, background: '#FDECEC', color: '#8A1F1F', borderColor: '#F9CACA' }
    : ghost   ? { ...base, background: 'transparent', color: '#6B6B66', borderColor: 'transparent' }
    : { ...base, background: 'transparent', color: '#18180F', borderColor: '#E0DFD9' };
  return <button style={style} onClick={disabled ? undefined : onClick}>{children}</button>;
}

function ConfirmMini({ titulo, texto, onSi, onNo }) {
  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.25)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1100, backdropFilter: 'blur(3px)' }}
      onClick={e => { if (e.target === e.currentTarget) onNo(); }}>
      <div className="modal-in fade" style={{ background: '#fff', borderRadius: 14, width: 340, maxWidth: '92vw', padding: 20, border: '1px solid #E0DFD9', boxShadow: '0 16px 48px rgba(0,0,0,.15)' }}>
        <div style={{ fontSize: 14, fontWeight: 600, color: '#141412', marginBottom: 8 }}>{titulo || '¿Eliminar?'}</div>
        <div style={{ fontSize: 13, color: '#52524E', lineHeight: 1.55, marginBottom: 18 }}>{texto || 'Esta acción no se puede deshacer.'}</div>
        <div style={{ display: 'flex', gap: 8 }}>
          <Btn full onClick={onNo}>Cancelar</Btn>
          <Btn full danger onClick={onSi}>Eliminar</Btn>
        </div>
      </div>
    </div>
  );
}

function Field({ label, children, hint }) {
  return (
    <div style={{ marginBottom: 13 }}>
      <label style={{ display: 'block', fontSize: 12, fontWeight: 500, color: '#52524E', marginBottom: 5 }}>{label}</label>
      {children}
      {hint && <p style={{ fontSize: 11, color: '#A5A5A0', marginTop: 4 }}>{hint}</p>}
    </div>
  );
}

function DashedBtn({ children, onClick }) {
  return (
    <button onClick={onClick} style={{ width: '100%', padding: '8px', borderRadius: 8, border: '1.5px dashed #E0DFD9', background: 'transparent', cursor: 'pointer', fontSize: 12, color: '#6B6B66', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 5, marginTop: 8, transition: 'border-color .15s' }}>
      {children}
    </button>
  );
}

// ─── Modal base ───────────────────────────────────────────────────────────────

function Modal({ title, onClose, children, footer, wide }) {
  const [confirmando, setConfirmando] = useState(false);
  const isMobile = useIsMobile();
  const panelStyle = isMobile
    ? { background: '#fff', borderRadius: '16px 16px 0 0', width: '100%', maxWidth: '100%', maxHeight: '94vh', borderTop: '1px solid #E0DFD9', boxShadow: '0 -8px 40px rgba(0,0,0,.18)', display: 'flex', flexDirection: 'column' }
    : { background: '#fff', borderRadius: 14, width: wide ? 540 : 460, maxWidth: '95vw', maxHeight: '90vh', border: '1px solid #E0DFD9', boxShadow: '0 24px 64px rgba(0,0,0,.14)', display: 'flex', flexDirection: 'column' };
  return (
    <div className="modal-overlay"
      style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.2)', display: 'flex', alignItems: isMobile ? 'flex-end' : 'center', justifyContent: 'center', zIndex: 1000, backdropFilter: 'blur(3px)' }}
      onClick={e => { if (e.target === e.currentTarget) setConfirmando(true); }}
    >
      <div className={isMobile ? 'sheet-in' : 'modal-in'} style={panelStyle}>
        <div style={{ padding: '15px 18px', borderBottom: '1px solid #ECEAE4', display: 'flex', alignItems: 'center', flexShrink: 0 }}>
          <div style={{ fontSize: 15, fontWeight: 600, flex: 1 }}>{title}</div>
          <button onClick={() => setConfirmando(true)} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 22, color: '#A5A5A0', lineHeight: 1, padding: '0 4px' }}>×</button>
        </div>
        <div style={{ padding: '16px 18px', overflowY: 'auto', flex: 1 }}>{children}</div>
        {footer && (
          <div style={{ padding: '12px 18px', borderTop: '1px solid #ECEAE4', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexShrink: 0, gap: 10, paddingBottom: isMobile ? 'calc(12px + env(safe-area-inset-bottom))' : 12 }}>
            {footer}
          </div>
        )}
      </div>
      {confirmando && (
        <div style={{ position: 'fixed', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1100 }}>
          <div className="fade" style={{ background: '#fff', borderRadius: 12, width: 340, padding: '20px 22px', border: '1px solid #E0DFD9', boxShadow: '0 12px 40px rgba(0,0,0,.16)' }}>
            <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 8 }}>¿Cerrar el formulario?</div>
            <p style={{ fontSize: 13, color: '#6B6B66', lineHeight: 1.5, marginBottom: 18 }}>Perderás los datos introducidos. Esta acción no se puede deshacer.</p>
            <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
              <Btn onClick={() => setConfirmando(false)}>Continuar editando</Btn>
              <Btn danger onClick={onClose}>Sí, cerrar</Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Sidebar ──────────────────────────────────────────────────────────────────

function Sidebar({ nav, setNav, stats, user }) {
  const navItems = [
    { id: 'alertas',    label: 'Alertas',     badge: stats.alertas, alert: stats.alertas > 0 },
    { id: 'tablero',    label: 'Tablero',     badge: stats.total, alert: false },
  ];
  const email = user?.email || '';
  const iniciales = email ? email.slice(0, 2).toUpperCase() : 'PL';

  function salir() {
    if (window.auth) window.auth.signOut();
  }

  return (
    <div style={{ width: 210, background: '#1C1C1A', display: 'flex', flexDirection: 'column', flexShrink: 0, height: '100%' }}>

      {/* Logo */}
      <div style={{ padding: '24px 20px 20px', borderBottom: '1px solid rgba(255,255,255,0.06)' }}>
        <div style={{ fontSize: 21, fontWeight: 600, letterSpacing: '0.14em', color: '#F2F1ED', display: 'flex', alignItems: 'baseline', gap: 7 }}>
          PLAAT
          <span style={{ fontSize: 12, fontWeight: 400, letterSpacing: '0.18em', color: 'rgba(255,255,255,0.4)' }}>/ DEO</span>
        </div>
      </div>

      {/* Nav */}
      <nav style={{ padding: '16px 10px', flex: 1 }}>
        <div style={{ fontSize: 9.5, color: 'rgba(255,255,255,0.25)', letterSpacing: '0.18em', textTransform: 'uppercase', padding: '4px 10px 10px' }}>Vistas</div>
        {navItems.map(item => (
          <div key={item.id} onClick={() => setNav(item.id)}
            style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '9px 10px', borderRadius: 6, marginBottom: 2, cursor: 'pointer', background: nav === item.id ? 'rgba(255,255,255,0.08)' : 'transparent', fontSize: 13, color: nav === item.id ? '#F2F1ED' : 'rgba(255,255,255,0.45)', fontWeight: nav === item.id ? 500 : 400, letterSpacing: '0.01em', transition: 'all .15s' }}>
            {nav === item.id && <span style={{ color: 'rgba(255,255,255,0.5)' }}>/</span>}
            {item.label}
            {item.badge > 0 && (
              <span style={{ marginLeft: 'auto', fontSize: 11, padding: '1px 7px', borderRadius: 20, background: item.alert ? 'rgba(228,75,74,0.22)' : 'rgba(255,255,255,0.1)', color: item.alert ? '#FF9090' : 'rgba(255,255,255,0.5)', fontWeight: 500 }}>
                {item.badge}
              </span>
            )}
          </div>
        ))}
      </nav>

      {/* Usuario */}
      <div style={{ padding: '12px 14px', borderTop: '1px solid rgba(255,255,255,0.06)' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: email ? 8 : 0 }}>
          <div style={{ width: 32, height: 32, borderRadius: '50%', background: 'rgba(255,255,255,0.1)', color: '#F2F1ED', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 500, flexShrink: 0, letterSpacing: '0.02em' }}>{iniciales}</div>
          <div style={{ minWidth: 0 }}>
            <div style={{ fontSize: 12, fontWeight: 500, color: '#F2F1ED', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{email || 'PLAAT'}</div>
            <div style={{ fontSize: 11, color: 'rgba(255,255,255,0.35)' }}>Arquitectura Técnica</div>
          </div>
        </div>
        {email && (
          <button onClick={salir} style={{ width: '100%', padding: '7px', borderRadius: 6, border: '1px solid rgba(255,255,255,0.12)', background: 'transparent', color: 'rgba(255,255,255,0.5)', fontSize: 12, cursor: 'pointer', letterSpacing: '0.02em' }}>
            Cerrar sesión
          </button>
        )}
      </div>
    </div>
  );
}

// ─── Barra inferior (móvil) ───────────────────────────────────────────────────

function BottomNav({ nav, setNav, stats }) {
  const ICONS = {
    alertas: 'M10 4a3 3 0 0 0-3 3c0 4-1.5 5-2 6h10c-.5-1-2-2-2-6a3 3 0 0 0-3-3Z M8.5 16a1.5 1.5 0 0 0 3 0',
    tablero: 'M3 17V8l5-3 5 3v9 M7 17v-4h2v4',
    salir:   'M7 4H4v12h3 M10 10h7 M14 7l3 3-3 3',
  };
  const items = [
    { id: 'alertas', label: 'Alertas', badge: stats.alertas, alert: true },
    { id: 'tablero', label: 'Obras',   badge: stats.total,   alert: false },
  ];
  function salir() { if (window.auth) window.auth.signOut(); }
  const Icono = ({ d, activo }) => (
    <svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{ opacity: activo ? 1 : 0.65 }}>
      {(d || '').split(' M').map((seg, i) => <path key={i} d={(i === 0 ? seg : 'M' + seg)} />)}
    </svg>
  );
  return (
    <div style={{ display: 'flex', background: '#1C1C1A', borderTop: '1px solid rgba(255,255,255,0.08)', flexShrink: 0, paddingBottom: 'env(safe-area-inset-bottom)' }}>
      {items.map(it => {
        const activo = nav === it.id;
        return (
          <button key={it.id} onClick={() => setNav(it.id)} style={{ flex: 1, background: 'none', border: 'none', cursor: 'pointer', padding: '10px 0 12px', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 4, color: activo ? '#F2F1ED' : 'rgba(255,255,255,0.4)', position: 'relative' }}>
            <Icono d={ICONS[it.id]} activo={activo} />
            <span style={{ fontSize: 11, fontWeight: activo ? 500 : 400, letterSpacing: '0.02em' }}>{it.label}</span>
            {it.badge > 0 && (
              <span style={{ position: 'absolute', top: 6, left: '50%', marginLeft: 7, background: it.alert ? '#E24B4A' : 'rgba(255,255,255,0.2)', color: '#fff', fontSize: 9, minWidth: 16, height: 16, borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '0 4px', fontWeight: 600 }}>{it.badge}</span>
            )}
          </button>
        );
      })}
      <button onClick={salir} style={{ flex: 1, background: 'none', border: 'none', cursor: 'pointer', padding: '10px 0 12px', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 4, color: 'rgba(255,255,255,0.4)' }}>
        <Icono d={ICONS.salir} activo={false} />
        <span style={{ fontSize: 11, letterSpacing: '0.02em' }}>Salir</span>
      </button>
    </div>
  );
}

const STATUS_ACCENT = {
  en_curso:   '#D48A0C',
  pendiente:  '#9B9B97',
  acabada:    '#52A124',
  paralizada: '#E24B4A',
};

function ObraCard({ obra, onClick, onEditar, onEliminar }) {
  const [menu, setMenu] = useState(false);
  const accentColor = STATUS_ACCENT[obra.estado] || STATUS_ACCENT.en_curso;
  const e           = ESTADOS_OBRA[obra.estado]  || ESTADOS_OBRA.en_curso;
  const incPend     = obra.incidencias.filter(i => i.estado !== 'resuelta').length;
  const tareasPend  = (obra.apuntes || []).filter(a => a.tipo === 'tarea' && !a.hecha).length;
  const totalPuntos = obra.disciplinas.reduce((s, d) => s + d.puntos.length, 0);
  const inspDone    = obra.disciplinas.reduce((s, d) => s + d.puntos.filter(p => p.estado === 'inspeccionado').length, 0);
  const inspPct     = totalPuntos > 0 ? Math.round(inspDone / totalPuntos * 100) : 0;
  const tareasVenc  = (obra.apuntes || []).some(a => a.tipo === 'tarea' && !a.hecha && a.fechaLimite && new Date(a.fechaLimite) < new Date(today()));
  const diasV       = obra.diasVisita || [];
  const letras      = ['D', 'L', 'M', 'X', 'J', 'V', 'S'];
  const iniciales   = (obra.nombre || '?').trim().slice(0, 2).toUpperCase();

  return (
    <div className="obra-card" onClick={onClick}
      style={{ background: '#fff', borderRadius: 12, boxShadow: '0 1px 3px rgba(0,0,0,0.05), 0 5px 18px rgba(0,0,0,0.05)', overflow: 'hidden', borderLeft: `3px solid ${accentColor}`, position: 'relative' }}>

      {/* Parte superior */}
      <div style={{ display: 'flex', gap: 13, padding: '15px 16px 13px' }}>
        {/* Iniciales ancla */}
        <div style={{ width: 44, height: 44, borderRadius: 8, flexShrink: 0, background: '#1C1C1A', color: '#F2F1ED', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 14, fontWeight: 500, letterSpacing: '0.04em' }}>{iniciales}</div>

        {/* Texto */}
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ display: 'flex', alignItems: 'flex-start', gap: 8 }}>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontSize: 15, fontWeight: 600, color: '#141412', lineHeight: 1.25, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{obra.nombre}</div>
              <div style={{ fontSize: 12.5, color: '#6B6B66', marginTop: 2, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{obra.cliente}</div>
              {obra.direccion && <div style={{ fontSize: 11.5, color: '#A5A5A0', marginTop: 2, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{obra.direccion}</div>}
            </div>
            {/* Estado + menú */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 2, flexShrink: 0 }}>
              <span style={{ fontSize: 9.5, color: accentColor, fontWeight: 600, whiteSpace: 'nowrap', letterSpacing: '0.1em', textTransform: 'uppercase' }}>{e.label}</span>
              <div style={{ position: 'relative' }}>
                <button onClick={ev => { ev.stopPropagation(); setMenu(m => !m); }}
                  style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#A5A5A0', fontSize: 18, lineHeight: 1, padding: '0 2px' }}>⋮</button>
                {menu && (
                  <>
                    <div onClick={ev => { ev.stopPropagation(); setMenu(false); }} style={{ position: 'fixed', inset: 0, zIndex: 20 }} />
                    <div style={{ position: 'absolute', right: 0, top: '100%', background: '#fff', border: '1px solid #E0DFD9', borderRadius: 9, boxShadow: '0 8px 24px rgba(0,0,0,.12)', padding: 5, zIndex: 21, minWidth: 130 }}>
                      <div onClick={ev => { ev.stopPropagation(); setMenu(false); onEditar(obra); }} style={{ padding: '7px 11px', borderRadius: 6, cursor: 'pointer', fontSize: 13, color: '#141412' }} className="hov-row">Editar</div>
                      <div onClick={ev => { ev.stopPropagation(); setMenu(false); onEliminar(obra); }} style={{ padding: '7px 11px', borderRadius: 6, cursor: 'pointer', fontSize: 13, color: '#8A1F1F' }} className="hov-row">Eliminar</div>
                    </div>
                  </>
                )}
              </div>
            </div>
          </div>

          {/* Barra de inspección */}
          {totalPuntos > 0 && (
            <div style={{ marginTop: 11, display: 'flex', alignItems: 'center', gap: 8 }}>
              <div style={{ flex: 1, height: 4, background: '#F0EFEA', borderRadius: 2, overflow: 'hidden' }}>
                <div className="bar-fill" style={{ width: inspPct + '%', height: '100%', background: inspPct === 100 ? '#52A124' : '#1C1C1A', borderRadius: 2 }} />
              </div>
              <span style={{ fontSize: 11, color: '#6B6B66', fontWeight: 500, whiteSpace: 'nowrap' }}>{inspDone}/{totalPuntos} insp.</span>
            </div>
          )}
        </div>
      </div>

      {/* Footer: chips de estado + días de visita */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 7, padding: '9px 16px', borderTop: '1px solid #F2F1ED', background: '#FBFAF8', flexWrap: 'wrap' }}>
        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 5, fontSize: 11.5, fontWeight: 500, padding: '3px 9px', borderRadius: 4, background: incPend > 0 ? '#FDECEC' : '#F0EFEA', color: incPend > 0 ? '#8A1F1F' : '#9B9B97' }}>
          <span style={{ width: 6, height: 6, borderRadius: '50%', background: incPend > 0 ? '#E24B4A' : '#C5C4BE' }} />
          {incPend > 0 ? `${incPend} incidencia${incPend > 1 ? 's' : ''}` : 'Sin incidencias'}
        </span>
        {tareasPend > 0 && (
          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, fontSize: 11.5, fontWeight: 500, padding: '3px 9px', borderRadius: 4, background: tareasVenc ? '#FEF3DB' : '#F0EFEA', color: tareasVenc ? '#7C4A00' : '#6B6B66' }}>
            {tareasPend} tarea{tareasPend > 1 ? 's' : ''}{tareasVenc ? ' · vencida' : ''}
          </span>
        )}
        {/* Días de visita */}
        {diasV.length > 0 && (
          <span style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 3 }}>
            {[1, 2, 3, 4, 5, 6, 0].map(d => (
              <span key={d} style={{ width: 17, height: 17, borderRadius: 3, fontSize: 9.5, fontWeight: 600, display: 'flex', alignItems: 'center', justifyContent: 'center', background: diasV.includes(d) ? '#1C1C1A' : '#F0EFEA', color: diasV.includes(d) ? '#fff' : '#C5C4BE' }}>{letras[d]}</span>
            ))}
          </span>
        )}
      </div>
    </div>
  );
}

// ─── Selector días de visita ──────────────────────────────────────────────────

const DIAS = [
  { num: 1, corto: 'L', largo: 'Lunes' },
  { num: 2, corto: 'M', largo: 'Martes' },
  { num: 3, corto: 'X', largo: 'Miércoles' },
  { num: 4, corto: 'J', largo: 'Jueves' },
  { num: 5, corto: 'V', largo: 'Viernes' },
  { num: 6, corto: 'S', largo: 'Sábado' },
];

function DiasPicker({ value, onChange }) {
  function toggle(num) {
    onChange(value.includes(num) ? value.filter(d => d !== num) : [...value, num].sort());
  }
  return (
    <div>
      <div style={{ display: 'flex', gap: 6 }}>
        {DIAS.map(d => {
          const on = value.includes(d.num);
          return (
            <button key={d.num} onClick={() => toggle(d.num)} title={d.largo}
              style={{ width: 36, height: 36, borderRadius: 8, border: `1.5px solid ${on ? '#18180F' : '#E0DFD9'}`, background: on ? '#18180F' : 'transparent', color: on ? '#fff' : '#6B6B66', fontSize: 13, fontWeight: on ? 600 : 400, cursor: 'pointer', transition: 'all .15s', flexShrink: 0 }}>
              {d.corto}
            </button>
          );
        })}
      </div>
      {value.length > 0 && (
        <div style={{ fontSize: 11, color: '#A5A5A0', marginTop: 6 }}>
          Visitas: {value.map(n => DIAS.find(d => d.num === n)?.largo).join(', ')}
        </div>
      )}
    </div>
  );
}

// ─── Modal Nueva Obra ─────────────────────────────────────────────────────────

function ModalNuevaObra({ onClose, onCreate, obra }) {
  const editando = !!obra;
  const [form, setForm] = useState(obra
    ? { nombre: obra.nombre || '', cliente: obra.cliente || '', direccion: obra.direccion || '', responsable: obra.responsable || RESPONSABLES[0], diasVisita: obra.diasVisita || [], emplazamiento: obra.emplazamiento || '', propiedad: obra.propiedad || '', proyectista: obra.proyectista || '', direccionObra: obra.direccionObra || '', constructora: obra.constructora || '', deoFirmante: obra.deoFirmante || '' }
    : { nombre: '', cliente: '', direccion: '', responsable: RESPONSABLES[0], diasVisita: [], emplazamiento: '', propiedad: '', proyectista: '', direccionObra: '', constructora: '', deoFirmante: '' });
  const upd = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const footer = (
    <>
      <Btn onClick={onClose}>Cancelar</Btn>
      <Btn primary disabled={!form.nombre.trim() || !form.cliente.trim()} onClick={() => onCreate(form)}>{editando ? 'Guardar cambios' : 'Crear obra'}</Btn>
    </>
  );
  return (
    <Modal title={editando ? 'Editar obra' : 'Nueva obra'} onClose={onClose} footer={footer}>
      <Field label="Nombre de la obra *"><input autoFocus placeholder="Rehabilitación fachada C/ Urgell 88" value={form.nombre} onChange={e => upd('nombre', e.target.value)} /></Field>
      <Field label="Cliente / Promotora *"><input placeholder="Nombre del cliente" value={form.cliente} onChange={e => upd('cliente', e.target.value)} /></Field>
      <Field label="Dirección"><input placeholder="C/ Ejemplo 10, Barcelona" value={form.direccion} onChange={e => upd('direccion', e.target.value)} /></Field>
      <Field label="Responsable DEO">
        <select value={form.responsable} onChange={e => upd('responsable', e.target.value)}>
          {RESPONSABLES.map(r => <option key={r}>{r}</option>)}
        </select>
      </Field>
      <Field label="Días de visita a obra" hint="Selecciona los días que sueles ir a esta obra">
        <DiasPicker value={form.diasVisita} onChange={v => upd('diasVisita', v)} />
      </Field>

      <div style={{ margin: '6px 0 14px', padding: '12px 14px', background: '#F9F8F5', borderRadius: 10, border: '1px solid #ECEAE4' }}>
        <div style={{ fontSize: 11, textTransform: 'uppercase', letterSpacing: '0.1em', color: '#A5A5A0', fontWeight: 600, marginBottom: 12 }}>Datos para el acta de inspección</div>
        <Field label="Emplazamiento"><input placeholder="Carrer Pamplona, 93-99. 08018 Barcelona" value={form.emplazamiento} onChange={e => upd('emplazamiento', e.target.value)} /></Field>
        <Field label="Propiedad"><input placeholder="Conren Tramway Catorce, S.L.U" value={form.propiedad} onChange={e => upd('propiedad', e.target.value)} /></Field>
        <Field label="Proyectista"><input placeholder="BCRA Arquitectes Associats SLP" value={form.proyectista} onChange={e => upd('proyectista', e.target.value)} /></Field>
        <Field label="Dirección de Obra (DO)"><input placeholder="BCRA Arquitectes Associats SLP" value={form.direccionObra} onChange={e => upd('direccionObra', e.target.value)} /></Field>
        <Field label="Constructora"><input placeholder="Certis" value={form.constructora} onChange={e => upd('constructora', e.target.value)} /></Field>
        <Field label="DEO firmante"><input placeholder="Nombre del director de ejecución" value={form.deoFirmante} onChange={e => upd('deoFirmante', e.target.value)} /></Field>
      </div>
    </Modal>
  );
}

// ─── MÓDULO: Inspecciones ─────────────────────────────────────────────────────

// ─── PLANOS DE INSPECCIÓN ─────────────────────────────────────────────────────

async function loadPdfJs() {
  if (window.pdfjsLib) return window.pdfjsLib;
  await new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js';
    s.onload = res; s.onerror = () => rej(new Error('No se pudo cargar PDF.js'));
    document.head.appendChild(s);
  });
  window.pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
  return window.pdfjsLib;
}

async function pdfFileToImgData(file) {
  const lib = await loadPdfJs();
  const ab  = await file.arrayBuffer();
  const pdf  = await lib.getDocument({ data: new Uint8Array(ab) }).promise;
  const page = await pdf.getPage(1);
  const vp   = page.getViewport({ scale: 1.5 });
  const canvas = document.createElement('canvas');
  canvas.width = vp.width; canvas.height = vp.height;
  await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
  return { imgData: canvas.toDataURL('image/jpeg', 0.65), w: canvas.width, h: canvas.height };
}

function comprimirImagen(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = ev => {
      const img = new Image();
      img.onload = () => {
        const MAX = 1200;  // reducido para ahorrar espacio
        const sc  = Math.min(1, MAX / Math.max(img.width, img.height));
        const c   = document.createElement('canvas');
        c.width = Math.round(img.width * sc);
        c.height = Math.round(img.height * sc);
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
        res({ imgData: c.toDataURL('image/jpeg', 0.60), w: c.width, h: c.height });
      };
      img.onerror = rej;
      img.src = ev.target.result;
    };
    r.onerror = rej;
    r.readAsDataURL(file);
  });
}

function PlanoPunto({ punto, onUpdate, obra, nombreDisciplina, numActa, onActaGenerada }) {
  const isMobile    = useIsMobile();
  const imgRef      = useRef(null);
  const [cargando,      setCargando]      = useState(false);
  const [marcando,      setMarcando]      = useState(null);
  const [formMarca,     setFormMarca]     = useState({ titulo: '', desc: '', foto: null });
  const [marcaOpen,     setMarcaOpen]     = useState(null);
  const [planoVisible,  setPlanoVisible]  = useState(true);
  const [generandoActa, setGenerandoActa] = useState(false);
  const [confirmacion,  setConfirmacion]  = useState(null);

  const plano  = punto.plano;
  const marcas = plano?.marcas || [];

  // ── Generar acta PDF — réplica del acta oficial PLAAT ───────────────────────
  async function generarActa() {
    if (!marcas.length) { alert('No hay marcas en el plano para generar el acta.'); return; }
    setGenerandoActa(true);
    try {
      if (!window.jspdf) {
        await new Promise((res, rej) => {
          const s = document.createElement('script');
          s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
          s.onload = res; s.onerror = () => rej(new Error('No se pudo cargar jsPDF'));
          document.head.appendChild(s);
        });
      }
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
      const PW = 210, PH = 297, M = 18, CW = PW - M * 2;
      const num = String(numActa || 1).padStart(2, '0');
      const fechaCorta = new Date().toLocaleDateString('es-ES');
      const fechaLarga = new Date().toLocaleDateString('es-ES', { day: '2-digit', month: 'long', year: 'numeric' });
      const NEGRO = [0, 0, 0], GRIS = [217, 217, 217], LH = 4.6;

      // Cabecera (cada página): Parte/Fecha bold izq + "Plaat." grande dcha
      function cabecera() {
        doc.setTextColor(...NEGRO); doc.setFont('helvetica', 'bold'); doc.setFontSize(8);
        doc.text(`PARTE DE INSPECCIÓN Nº: ${num}`, M, 12);
        doc.text(`FECHA PARTE INSPECCIÓN: ${fechaCorta}`, M, 16.5);
        doc.setFontSize(30);
        doc.text('Plaat.', PW - M, 17, { align: 'right' });
      }
      // Pie (cada página)
      function pie() {
        doc.setTextColor(...NEGRO); doc.setFont('helvetica', 'normal'); doc.setFontSize(8);
        doc.text('Plaat Arquitectura Técnica', M, PH - 12);
        doc.text('Barcelona – Madrid', PW / 2, PH - 12, { align: 'center' });
        doc.text('www.plaat.es', PW - M, PH - 12, { align: 'right' });
      }
      function nuevaPagina() { doc.addPage(); cabecera(); }

      // Tabla clave/valor: etiqueta gris + valor blanco, todo bordeado, negrita
      function fila(yy, label, value, wL) {
        wL = wL || 50;
        const vlines = doc.splitTextToSize(value || '', CW - wL - 5);
        const h = Math.max(8, vlines.length * LH + 4);
        doc.setFillColor(...GRIS); doc.rect(M, yy, wL, h, 'F');
        doc.setDrawColor(...NEGRO); doc.setLineWidth(0.2);
        doc.rect(M, yy, wL, h); doc.rect(M + wL, yy, CW - wL, h);
        doc.setTextColor(...NEGRO); doc.setFont('helvetica', 'bold'); doc.setFontSize(9);
        doc.text(label, M + 2.5, yy + h / 2 + 1.2);
        doc.text(vlines, M + wL + 2.5, yy + h / 2 - (vlines.length - 1) * (LH / 2) + 1.2);
        return yy + h;
      }

      // ═══ PÁGINA 1 ═══
      cabecera();
      let y = 30;
      // Banda título
      doc.setFillColor(...GRIS); doc.rect(M, y, CW, 8, 'F');
      doc.setDrawColor(...NEGRO); doc.setLineWidth(0.2); doc.rect(M, y, CW, 8);
      doc.setTextColor(...NEGRO); doc.setFont('helvetica', 'bold'); doc.setFontSize(11);
      doc.text('ACTA DE INSPECCIÓN DE OBRA', M + 3, y + 5.5);
      doc.text(`NÚM.: ${num}`, PW - M - 3, y + 5.5, { align: 'right' });
      y += 14;

      // OBRA / EMPLAZAMIENTO
      y = fila(y, 'OBRA', (obra?.nombre || '').toUpperCase());
      y = fila(y, 'EMPLAZAMIENTO', (obra?.emplazamiento || obra?.direccion || '').toUpperCase());
      y += 12;

      // DATOS DE LA OBRA
      doc.setTextColor(...NEGRO); doc.setFont('helvetica', 'bold'); doc.setFontSize(10);
      doc.text('DATOS DE LA OBRA:', M, y); y += 7;
      [
        ['PROPIEDAD', (obra?.propiedad || obra?.cliente || '').toUpperCase()],
        ['PROYECTISTA', obra?.proyectista || ''],
        ['DO', obra?.direccionObra || ''],
        ['DEO', 'PLAAT ARQUITECTURA TÉCNICA S.L.'],
        ['CONSTRUCTORA', (obra?.constructora || '').toUpperCase()],
      ].forEach(([k, v]) => { y = fila(y, k, v); });
      y += 16;

      // FECHA FIRMA
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10);
      doc.text('FECHA FIRMA:', M, y); y += 7;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(10);
      doc.text(`Barcelona, ${fechaLarga}.`, M, y);

      // Firmante (parte inferior)
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10);
      doc.text('DIRECTOR DE EJECUCIÓN DE OBRA', M, PH - 40);
      doc.setFont('helvetica', 'normal');
      doc.text(obra?.deoFirmante || obra?.responsable || '', M, PH - 35);

      // ═══ PÁGINA 2 — Aspecto revisado ═══
      nuevaPagina(); y = 30;
      doc.setTextColor(...NEGRO); doc.setFont('helvetica', 'bold'); doc.setFontSize(10);
      doc.text('Aspecto revisado:', M, y); y += 7;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5);
      doc.text('Durante la visita realizada, se han revisado los siguientes aspectos:', M, y); y += 8;

      // Tabla N.º / INCIDENCIAS
      const wN = 18;
      doc.setFillColor(...GRIS); doc.setDrawColor(...NEGRO); doc.setLineWidth(0.2);
      doc.rect(M, y, wN, 8, 'F'); doc.rect(M + wN, y, CW - wN, 8, 'F');
      doc.rect(M, y, wN, 8); doc.rect(M + wN, y, CW - wN, 8);
      doc.setTextColor(...NEGRO); doc.setFont('helvetica', 'bold'); doc.setFontSize(9);
      doc.text('N.º', M + 4, y + 5.3);
      doc.text('INCIDENCIAS', M + wN + 3, y + 5.3);
      y += 8;
      marcas.forEach((m, idx) => {
        const titulo = (m.titulo || m.desc || '').toUpperCase();
        const tl = doc.splitTextToSize(titulo, CW - wN - 6);
        const h = Math.max(10, tl.length * LH + 5);
        doc.setDrawColor(...NEGRO); doc.setLineWidth(0.2);
        doc.rect(M, y, wN, h); doc.rect(M + wN, y, CW - wN, h);
        doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(...NEGRO);
        doc.text(`N${idx + 1}`, M + 4, y + h / 2 + 1);
        doc.text(tl, M + wN + 3, y + h / 2 - (tl.length - 1) * (LH / 2) + 1.2);
        y += h;
      });
      y += 8;

      doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5);
      doc.text('Se adjunta tabla con las zonas revisadas.', M, y); y += 10;

      // Incidencias detectadas (subrayado)
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10);
      doc.text('Incidencias detectadas:', M, y);
      const wTit = doc.getTextWidth('Incidencias detectadas:');
      doc.setLineWidth(0.3); doc.line(M, y + 1, M + wTit, y + 1);
      y += 8;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5);
      [
        'Las incidencias detectadas, deberán subsanarse y dar respuesta a la DEO.',
        'Para ello deberán rellenar los datos solicitados en cada una de las incidencias detectadas y enviar fotografías a la DEO con las rectificaciones.',
        'Se adjuntan tablas con las incidencias detectadas.',
        'Este Acta de Inspección de obra se ha llevado a cabo en base a las inspecciones y muestreos realizados por el DEO en la fecha indicada y en base a Partes de Inspección procedimentados.',
      ].forEach(p => {
        const pl = doc.splitTextToSize(p, CW);
        doc.text(pl, M, y); y += pl.length * LH + 4;
      });

      // ═══ PÁGINA 3 — Zonas revisadas (plano + fichas) ═══
      nuevaPagina(); y = 30;
      doc.setTextColor(...NEGRO); doc.setFont('helvetica', 'bold'); doc.setFontSize(11);
      doc.text('1. ZONAS REVISADAS', M, y); y += 8;

      // Plano con marcas (proporción real)
      if (plano?.imgData) {
        const ratio = plano.h && plano.w ? plano.h / plano.w : 0.6;
        let pw = CW, ph = CW * ratio;
        if (ph > 150) { ph = 150; pw = ph / ratio; }
        const px = M + (CW - pw) / 2;
        if (y + ph > PH - 18) { nuevaPagina(); y = 30; }
        doc.addImage(plano.imgData, 'JPEG', px, y, pw, ph);
        doc.setDrawColor(...NEGRO); doc.setLineWidth(0.2); doc.rect(px, y, pw, ph);
        marcas.forEach((m, idx) => {
          const mx = px + (m.x / 100) * pw, my = y + (m.y / 100) * ph;
          doc.setFillColor(226, 75, 74); doc.circle(mx, my, 1.9, 'F');
          doc.setTextColor(255, 255, 255); doc.setFontSize(5); doc.setFont('helvetica', 'bold');
          doc.text(String(idx + 1), mx, my + 0.7, { align: 'center' });
        });
        y += ph + 10;
      }

      // Fichas por incidencia (estilo tabla: col Nº gris + contenido)
      for (let idx = 0; idx < marcas.length; idx++) {
        const m = marcas[idx];
        // Medir contenido
        doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5);
        const dl = doc.splitTextToSize(m.desc || '', CW - wN - 6);
        let fw = 0, fh = 0;
        if (m.foto) {
          fw = 78; fh = 58;
          try { const pr = doc.getImageProperties(m.foto); const r = pr.height / pr.width; fh = fw * r; if (fh > 72) { fh = 72; fw = fh / r; } } catch (e) {}
        }
        const hTitulo = 8;
        const hCont = 4 + dl.length * LH + (m.foto ? fh + 5 : 0) + 6;
        const hTotal = hTitulo + hCont;

        if (y + hTotal > PH - 18) { nuevaPagina(); y = 30; }

        // Fila título: Nº gris + título gris
        doc.setFillColor(...GRIS); doc.setDrawColor(...NEGRO); doc.setLineWidth(0.2);
        doc.rect(M, y, wN, hTitulo, 'F'); doc.rect(M + wN, y, CW - wN, hTitulo, 'F');
        doc.rect(M, y, wN, hTitulo); doc.rect(M + wN, y, CW - wN, hTitulo);
        doc.setTextColor(...NEGRO); doc.setFont('helvetica', 'bold'); doc.setFontSize(9.5);
        doc.text(`N${idx + 1}`, M + 4, y + 5.3);
        doc.text((m.titulo || `Incidencia ${idx + 1}`).toUpperCase(), M + wN + 3, y + 5.3);
        // Fila contenido: col Nº vacía + contenido
        const yc = y + hTitulo;
        doc.rect(M, yc, wN, hCont); doc.rect(M + wN, yc, CW - wN, hCont);
        let cy = yc + 6;
        doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5); doc.setTextColor(...NEGRO);
        doc.text(dl, M + wN + 3, cy); cy += dl.length * LH + 4;
        if (m.foto) { doc.addImage(m.foto, 'JPEG', M + wN + 3, cy, fw, fh); cy += fh + 3; }
        y += hTotal + 6;
      }

      // Cabecera y pie en todas las páginas + numeración
      const total = doc.getNumberOfPages();
      for (let p = 1; p <= total; p++) {
        doc.setPage(p);
        pie();
      }

      const fname = new Date().toISOString().slice(0, 10);
      doc.save(`Acta_Inspeccion_${num}_${(obra?.nombre || 'obra').replace(/\s+/g, '_')}_${fname}.pdf`);
      if (onActaGenerada) onActaGenerada(numActa);
    } catch (e) {
      alert('Error generando el acta: ' + e.message);
    }
    setGenerandoActa(false);
  }

  async function abrirSelector() {
    const inp = document.createElement('input');
    inp.type = 'file'; inp.accept = '.pdf,image/*';
    inp.style.cssText = 'position:fixed;left:-9999px';
    document.body.appendChild(inp);
    inp.onchange = async () => {
      const file = inp.files[0];
      if (!file) return;
      document.body.removeChild(inp);
      if (file.size > 30 * 1024 * 1024) { alert('El archivo supera 30 MB'); return; }
      setCargando(true);
      try {
        const r = file.type === 'application/pdf' || file.name.endsWith('.pdf')
          ? await pdfFileToImgData(file)
          : await comprimirImagen(file);
        onUpdate({ ...punto, plano: { nombre: file.name, imgData: r.imgData, w: r.w, h: r.h, marcas: plano?.marcas || [] } });
      } catch (e) { alert('Error al procesar el plano: ' + e.message); }
      setCargando(false);
    };
    inp.click();
  }

  function coords(e) {
    if (!imgRef.current) return null;
    const rect = imgRef.current.getBoundingClientRect();
    const cx = e.touches && e.touches[0] ? e.touches[0].clientX : e.clientX;
    const cy = e.touches && e.touches[0] ? e.touches[0].clientY : e.clientY;
    let x = (cx - rect.left) / rect.width * 100;
    let y = (cy - rect.top) / rect.height * 100;
    x = Math.max(0, Math.min(100, x));
    y = Math.max(0, Math.min(100, y));
    return { x: parseFloat(x.toFixed(1)), y: parseFloat(y.toFixed(1)) };
  }

  function handleTapPlano(e) {
    if (marcando) return;
    const c = coords(e);
    if (!c) return;
    setMarcando(c);
    setFormMarca({ titulo: '', desc: '', foto: null });
  }

  // Arrastrar el punto pendiente para ajustarlo con precisión
  function iniciarArrastre(e) {
    e.stopPropagation();
    const mover = ev => { const c = coords(ev); if (c) setMarcando(c); };
    const soltar = () => {
      window.removeEventListener('mousemove', mover);
      window.removeEventListener('mouseup', soltar);
      window.removeEventListener('touchmove', mover);
      window.removeEventListener('touchend', soltar);
    };
    window.addEventListener('mousemove', mover);
    window.addEventListener('mouseup', soltar);
    window.addEventListener('touchmove', mover, { passive: false });
    window.addEventListener('touchend', soltar);
  }

  function guardarMarca() {
    if (!formMarca.desc.trim()) return;
    const nuevas = [...marcas, {
      id: uid(), x: marcando.x, y: marcando.y,
      titulo: formMarca.titulo.trim(), desc: formMarca.desc.trim(), foto: formMarca.foto,
      createdAt: new Date().toISOString(),
    }];
    onUpdate({ ...punto, plano: { ...plano, marcas: nuevas } });
    setMarcando(null);
  }

  function deleteMarca(id) {
    onUpdate({ ...punto, plano: { ...plano, marcas: marcas.filter(m => m.id !== id) } });
    if (marcaOpen === id) setMarcaOpen(null);
  }

  // — Sin plano —
  if (!plano) return (
    <div style={{ margin: '10px 0 4px', padding: '14px 16px', background: '#F9F8F5', borderRadius: 10, border: '1.5px dashed #E0DFD9' }}>
      <div style={{ fontSize: 13, color: '#9B9B97', marginBottom: 10, lineHeight: 1.5 }}>
        Adjunta un plano (PDF o imagen) para marcar incidencias con su ubicación exacta.
      </div>
      {cargando
        ? <div style={{ fontSize: 13, color: '#9B9B97', display: 'flex', alignItems: 'center', gap: 8 }}>
            <span style={{ display: 'inline-block', width: 14, height: 14, border: '2px solid #D0D0CB', borderTopColor: '#141412', borderRadius: '50%', animation: 'spin .7s linear infinite' }} />
            Procesando plano...
          </div>
        : <button onClick={abrirSelector} style={{ padding: '8px 14px', borderRadius: 8, border: '1px solid #D0D0CB', background: '#fff', fontSize: 13, cursor: 'pointer', fontWeight: 500 }}>+ Subir plano (PDF o imagen)</button>
      }
    </div>
  );

  // — Con plano —
  return (
    <div style={{ margin: '10px 0 4px' }}>
      {/* Cabecera plano */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: planoVisible ? 8 : 0 }}>
        <button onClick={() => setPlanoVisible(v => !v)}
          style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 12, color: '#141412', fontWeight: 500, display: 'flex', alignItems: 'center', gap: 5, padding: 0, flex: 1, minWidth: 0, textAlign: 'left' }}>
          <span style={{ fontSize: 10, color: '#A5A5A0', transition: 'transform .2s', display: 'inline-block', transform: planoVisible ? 'rotate(90deg)' : 'rotate(0deg)' }}>▶</span>
          <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{plano.nombre}</span>
          <span style={{ fontSize: 11, color: '#A5A5A0', fontWeight: 400, flexShrink: 0 }}>{marcas.length} marca{marcas.length !== 1 ? 's' : ''}</span>
        </button>
        {marcas.length > 0 && (
          <button onClick={generarActa} disabled={generandoActa}
            style={{ fontSize: 11, padding: '4px 10px', borderRadius: 6, border: '1px solid #1C1C1A', background: generandoActa ? '#ECEAE4' : '#1C1C1A', color: generandoActa ? '#A5A5A0' : '#F2F1ED', cursor: generandoActa ? 'default' : 'pointer', fontWeight: 500, flexShrink: 0 }}>
            {generandoActa ? 'Generando...' : '↓ Acta PDF'}
          </button>
        )}
        {cargando
          ? <span style={{ fontSize: 11, color: '#9B9B97', flexShrink: 0 }}>Procesando...</span>
          : <button onClick={abrirSelector} style={{ fontSize: 11, color: '#9B9B97', background: 'none', border: 'none', cursor: 'pointer', padding: '2px 6px', borderRadius: 6, border: '1px solid #E8E7E1', flexShrink: 0 }}>↺</button>}
      </div>

      {/* Imagen del plano — ocultable */}
      {planoVisible && (<>
      <div style={{ position: 'relative', borderRadius: 10, overflow: 'hidden', border: '1px solid #E8E7E1', cursor: marcando ? 'default' : 'crosshair', userSelect: 'none', touchAction: 'none' }}
        onClick={!marcando ? handleTapPlano : undefined}
        onTouchEnd={!marcando ? e => { e.preventDefault(); handleTapPlano(e); } : undefined}>
        <img ref={imgRef} src={plano.imgData} alt="Plano" style={{ width: '100%', display: 'block', pointerEvents: 'none' }} />
        {/* Marcas existentes */}
        {marcas.map((m, idx) => (
          <button key={m.id} onClick={e => { e.stopPropagation(); setMarcaOpen(marcaOpen === m.id ? null : m.id); }}
            style={{ position: 'absolute', left: m.x + '%', top: m.y + '%', transform: 'translate(-50%,-50%)', width: 19, height: 19, borderRadius: '50%', background: '#E24B4A', color: '#fff', border: '1.5px solid #fff', boxShadow: '0 1px 5px rgba(0,0,0,.35)', fontSize: 9.5, fontWeight: 700, cursor: 'pointer', zIndex: 2, padding: 0, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            {idx + 1}
          </button>
        ))}
        {/* Nueva marca (pendiente, arrastrable) */}
        {marcando && (
          <div
            onMouseDown={iniciarArrastre}
            onTouchStart={iniciarArrastre}
            style={{ position: 'absolute', left: marcando.x + '%', top: marcando.y + '%', transform: 'translate(-50%,-50%)', width: 40, height: 40, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 3, cursor: 'grab', touchAction: 'none' }}>
            <span style={{ position: 'absolute', width: 40, height: 40, borderRadius: '50%', background: 'rgba(28,28,26,0.15)' }} />
            <span style={{ position: 'relative', width: 22, height: 22, borderRadius: '50%', background: '#1C1C1A', color: '#fff', border: '2px solid #fff', boxShadow: '0 1px 6px rgba(0,0,0,.45)', fontSize: 14, display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 300 }}>+</span>
          </div>
        )}
        {/* Hint sobre el plano */}
        {!marcando && marcas.length === 0 && (
          <div style={{ position: 'absolute', bottom: 8, left: '50%', transform: 'translateX(-50%)', background: 'rgba(0,0,0,0.55)', color: '#fff', fontSize: 11, padding: '5px 11px', borderRadius: 20, pointerEvents: 'none', whiteSpace: 'nowrap' }}>
            {isMobile ? 'Toca el plano' : 'Clic en el plano'} para marcar una incidencia
          </div>
        )}
      </div>
      </>)}

      {/* Formulario nueva marca */}
      {marcando && (
        <div className="fade" style={{ marginTop: 8, padding: '12px 14px', background: '#F9F8F5', borderRadius: 10, border: '1px solid #E8E7E1' }}>
          <div style={{ fontSize: 13, fontWeight: 500, marginBottom: 4, color: '#141412' }}>/ Nueva marca — ¿qué has visto?</div>
          <div style={{ fontSize: 11, color: '#9B9B97', marginBottom: 8 }}>Arrastra el punto sobre el plano para ajustar su posición exacta.</div>
          <input placeholder="Título breve (ej. Anclajes pantallas)" value={formMarca.titulo} onChange={e => setFormMarca(f => ({ ...f, titulo: e.target.value }))} style={{ marginBottom: 8 }} />
          <textarea placeholder="Describe la incidencia observada..." value={formMarca.desc} onChange={e => setFormMarca(f => ({ ...f, desc: e.target.value }))} style={{ marginBottom: 8, minHeight: 64 }} />
          {formMarca.foto
            ? <div style={{ position: 'relative', marginBottom: 8, display: 'inline-block' }}>
                <img src={formMarca.foto} style={{ maxWidth: '100%', maxHeight: 140, objectFit: 'contain', borderRadius: 8, display: 'block' }} />
                <button onClick={() => setFormMarca(f => ({ ...f, foto: null }))} style={{ position: 'absolute', top: 6, right: 6, background: 'rgba(0,0,0,0.5)', color: '#fff', border: 'none', borderRadius: '50%', width: 24, height: 24, cursor: 'pointer', fontSize: 14, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>×</button>
              </div>
            : <button onClick={() => pickFiles('image/*', f => setFormMarca(fm => ({ ...fm, foto: f.data })))}
                style={{ width: '100%', padding: '7px', borderRadius: 8, border: '1.5px dashed #E0DFD9', background: 'transparent', cursor: 'pointer', fontSize: 12, color: '#9B9B97', marginBottom: 8 }}>+ Añadir foto</button>}
          <div style={{ display: 'flex', gap: 8 }}>
            <Btn primary full disabled={!formMarca.desc.trim()} onClick={guardarMarca}>Guardar marca</Btn>
            <Btn onClick={() => setMarcando(null)}>✕</Btn>
          </div>
        </div>
      )}

      {/* Lista de marcas */}
      {marcas.length > 0 && (
        <div style={{ marginTop: 8, display: 'flex', flexDirection: 'column', gap: 5 }}>
          {marcas.map((m, idx) => (
            <div key={m.id} style={{ border: `1px solid ${marcaOpen === m.id ? '#1C1C1A' : '#E8E7E1'}`, borderRadius: 9, overflow: 'hidden', transition: 'border-color .15s' }}>
              <div onClick={() => setMarcaOpen(marcaOpen === m.id ? null : m.id)}
                style={{ display: 'flex', alignItems: 'center', gap: 9, padding: '9px 12px', cursor: 'pointer', background: marcaOpen === m.id ? '#F5F4F0' : '#fff' }}>
                <div style={{ width: 22, height: 22, borderRadius: '50%', background: '#E24B4A', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, fontWeight: 700, flexShrink: 0 }}>{idx + 1}</div>
                <div style={{ flex: 1, fontSize: 13, fontWeight: m.titulo ? 500 : 400, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{m.titulo || m.desc}</div>
                <div style={{ fontSize: 11, color: '#A5A5A0', flexShrink: 0 }}>{fmtShort(m.createdAt)}</div>
                <span style={{ fontSize: 12, color: '#C5C4BE' }}>{marcaOpen === m.id ? '▲' : '▼'}</span>
              </div>
              {marcaOpen === m.id && (
                <div className="fade" style={{ padding: '0 12px 12px' }}>
                  {m.foto && <img src={m.foto} style={{ maxWidth: '100%', maxHeight: 180, objectFit: 'contain', borderRadius: 8, marginBottom: 8, marginTop: 6, display: 'block' }} />}
                  {m.titulo && <div style={{ fontSize: 13, fontWeight: 600, color: '#141412', marginBottom: 3 }}>{m.titulo}</div>}
                  <p style={{ fontSize: 13, color: '#18180F', lineHeight: 1.55, marginBottom: 10 }}>{m.desc}</p>
                  <Btn sm danger onClick={() => setConfirmacion({ titulo: 'Eliminar marca', texto: 'Vas a eliminar esta marca y su contenido. Esta acción no se puede deshacer.', onSi: () => { deleteMarca(m.id); setConfirmacion(null); } })}>Eliminar marca</Btn>
                </div>
              )}
            </div>
          ))}
        </div>
      )}
      {confirmacion && <ConfirmMini titulo={confirmacion.titulo} texto={confirmacion.texto} onSi={confirmacion.onSi} onNo={() => setConfirmacion(null)} />}
    </div>
  );
}

function ModuloInspecciones({ obra, onSave }) {
  const isMobile = useIsMobile();
  const [disciplinaActiva,    setDisciplinaActiva]    = useState(null);
  const [showNuevaDisciplina, setShowNuevaDisciplina] = useState(false);
  const [showNuevoPunto,      setShowNuevoPunto]      = useState(false);
  const [nombreDisciplina,    setNombreDisciplina]    = useState('');
  const [nombrePunto,         setNombrePunto]         = useState('');
  const [puntosExpandidos,    setPuntosExpandidos]    = useState(new Set());
  const [confirmacion,        setConfirmacion]         = useState(null); // {titulo,texto,onSi}
  const [menuDisciplina,      setMenuDisciplina]       = useState(false);

  function togglePunto(id) {
    setPuntosExpandidos(prev => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  }

  function updatePunto(updated) {
    const disciplinas = obra.disciplinas.map(d => d.id === disciplinaActiva
      ? { ...d, puntos: d.puntos.map(p => p.id === updated.id ? updated : p) }
      : d);
    onSave({ ...obra, disciplinas });
  }

  const disciplina = obra.disciplinas.find(d => d.id === disciplinaActiva);

  function addDisciplina() {
    if (!nombreDisciplina.trim()) return;
    const nueva = { id: uid(), nombre: nombreDisciplina.trim(), puntos: [] };
    onSave({ ...obra, disciplinas: [...obra.disciplinas, nueva] });
    setNombreDisciplina('');
    setShowNuevaDisciplina(false);
    setDisciplinaActiva(nueva.id);
  }

  function deleteDisciplina(id) {
    onSave({ ...obra, disciplinas: obra.disciplinas.filter(d => d.id !== id) });
    if (disciplinaActiva === id) setDisciplinaActiva(null);
  }

  function addPunto() {
    if (!nombrePunto.trim() || !disciplinaActiva) return;
    const punto = { id: uid(), nombre: nombrePunto.trim(), estado: 'pendiente', notas: '', fecha: '' };
    const disciplinas = obra.disciplinas.map(d => d.id === disciplinaActiva ? { ...d, puntos: [...d.puntos, punto] } : d);
    onSave({ ...obra, disciplinas });
    setNombrePunto('');
    setShowNuevoPunto(false);
  }

  function updatePuntoEstado(puntoId, estado) {
    const disciplinas = obra.disciplinas.map(d => d.id === disciplinaActiva
      ? { ...d, puntos: d.puntos.map(p => p.id === puntoId ? { ...p, estado, fecha: estado !== 'pendiente' ? today() : '' } : p) }
      : d
    );
    onSave({ ...obra, disciplinas });
  }

  function deletePunto(puntoId) {
    const disciplinas = obra.disciplinas.map(d => d.id === disciplinaActiva
      ? { ...d, puntos: d.puntos.filter(p => p.id !== puntoId) }
      : d
    );
    onSave({ ...obra, disciplinas });
  }

  const totalPuntos     = obra.disciplinas.reduce((s, d) => s + d.puntos.length, 0);
  const inspeccionados  = obra.disciplinas.reduce((s, d) => s + d.puntos.filter(p => p.estado === 'inspeccionado').length, 0);
  const conIncidencia   = obra.disciplinas.reduce((s, d) => s + d.puntos.filter(p => p.estado === 'incidencia').length, 0);

  return (
    <div style={{ display: isMobile ? 'flex' : 'grid', flexDirection: isMobile ? 'column' : undefined, gridTemplateColumns: isMobile ? undefined : '220px 1fr', gap: 12, height: isMobile ? 'auto' : '100%' }}>
      {/* Panel izquierdo: disciplinas (en móvil, solo si no hay disciplina abierta) */}
      {(!isMobile || !disciplina) && (
      <div style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 12, padding: '12px', display: 'flex', flexDirection: 'column', gap: 4 }}>
        <div style={{ fontSize: 11, textTransform: 'uppercase', letterSpacing: '0.1em', color: '#A5A5A0', fontWeight: 500, padding: '2px 6px 8px' }}>
          Disciplinas
        </div>

        {obra.disciplinas.length === 0 && (
          <div style={{ fontSize: 12, color: '#A5A5A0', padding: '8px 6px' }}>Sin disciplinas todavía</div>
        )}

        {obra.disciplinas.map(d => {
          const pts   = d.puntos.length;
          const insp  = d.puntos.filter(p => p.estado === 'inspeccionado').length;
          const inc   = d.puntos.filter(p => p.estado === 'incidencia').length;
          const activa = disciplinaActiva === d.id;
          return (
            <div key={d.id} onClick={() => setDisciplinaActiva(d.id)} style={{ padding: '8px 10px', borderRadius: 8, cursor: 'pointer', background: activa ? '#F0EFEA' : 'transparent', border: `1px solid ${activa ? '#C5C4BE' : 'transparent'}` }} className={activa ? '' : 'hov-row'}>
              <div style={{ fontSize: 13, fontWeight: activa ? 500 : 400, color: '#18180F', marginBottom: 3 }}>{d.nombre}</div>
              <div style={{ display: 'flex', gap: 6, fontSize: 11, color: '#A5A5A0' }}>
                <span>{insp}/{pts} insp.</span>
                {inc > 0 && <span style={{ color: '#8A1F1F' }}>⚠ {inc}</span>}
              </div>
            </div>
          );
        })}

        {showNuevaDisciplina ? (
          <div style={{ padding: '6px 4px', display: 'flex', flexDirection: 'column', gap: 6, marginTop: 4 }}>
            <input autoFocus placeholder="Nombre disciplina" value={nombreDisciplina} onChange={e => setNombreDisciplina(e.target.value)} onKeyDown={e => e.key === 'Enter' && addDisciplina()} style={{ fontSize: 12 }} />
            <div style={{ display: 'flex', gap: 5 }}>
              <Btn sm full primary disabled={!nombreDisciplina.trim()} onClick={addDisciplina}>Añadir</Btn>
              <Btn sm onClick={() => { setShowNuevaDisciplina(false); setNombreDisciplina(''); }}>✕</Btn>
            </div>
          </div>
        ) : (
          <DashedBtn onClick={() => setShowNuevaDisciplina(true)}>+ Disciplina</DashedBtn>
        )}

        {/* Resumen global */}
        {totalPuntos > 0 && (
          <div style={{ marginTop: 'auto', padding: '10px 8px 0', borderTop: '1px solid #E8E7E1' }}>
            <div style={{ fontSize: 11, color: '#A5A5A0', marginBottom: 6 }}>Global</div>
            <div style={{ height: 4, background: '#ECEAE4', borderRadius: 2, marginBottom: 5 }}>
              <div style={{ width: (inspeccionados / totalPuntos * 100) + '%', height: 4, borderRadius: 2, background: '#52A124', transition: 'width .3s' }} />
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 11 }}>
              <span style={{ color: '#6B6B66' }}>{inspeccionados}/{totalPuntos}</span>
              {conIncidencia > 0 && <span style={{ color: '#8A1F1F' }}>⚠ {conIncidencia}</span>}
            </div>
          </div>
        )}
      </div>
      )}

      {/* Panel derecho: puntos de inspección (en móvil, solo si hay disciplina abierta) */}
      {(!isMobile || disciplina) && (
      <div style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 12, padding: '14px 16px', overflow: isMobile ? 'visible' : 'auto' }}>
        {isMobile && disciplina && (
          <button onClick={() => setDisciplinaActiva(null)} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 13, color: '#6B6B66', padding: '0 0 12px', display: 'flex', alignItems: 'center', gap: 5 }}>← Disciplinas</button>
        )}
        {!disciplina ? (
          <div style={{ height: '100%', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 10, color: '#A5A5A0' }}>
            <div style={{ fontSize: 32 }}>📋</div>
            <div style={{ fontSize: 13 }}>Selecciona una disciplina para ver sus puntos de control</div>
          </div>
        ) : (
          <>
            <div style={{ display: 'flex', alignItems: 'center', marginBottom: 14 }}>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 14, fontWeight: 600 }}>{disciplina.nombre}</div>
                <div style={{ fontSize: 12, color: '#A5A5A0', marginTop: 1 }}>{disciplina.puntos.length} puntos de control</div>
              </div>
              <div style={{ position: 'relative' }}>
                <button onClick={() => setMenuDisciplina(v => !v)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#A5A5A0', fontSize: 20, lineHeight: 1, padding: '0 4px' }}>⋮</button>
                {menuDisciplina && (
                  <>
                    <div onClick={() => setMenuDisciplina(false)} style={{ position: 'fixed', inset: 0, zIndex: 20 }} />
                    <div style={{ position: 'absolute', right: 0, top: '100%', background: '#fff', border: '1px solid #E0DFD9', borderRadius: 9, boxShadow: '0 8px 24px rgba(0,0,0,.12)', padding: 5, zIndex: 21, minWidth: 160 }}>
                      <div onClick={() => { setMenuDisciplina(false); setConfirmacion({ titulo: 'Eliminar disciplina', texto: `Vas a eliminar "${disciplina.nombre}" y todos sus puntos de control. Esta acción no se puede deshacer.`, onSi: () => { deleteDisciplina(disciplina.id); setConfirmacion(null); } }); }}
                        style={{ padding: '7px 11px', borderRadius: 6, cursor: 'pointer', fontSize: 13, color: '#8A1F1F' }} className="hov-row">Eliminar disciplina</div>
                    </div>
                  </>
                )}
              </div>
            </div>

            {disciplina.puntos.length === 0 && (
              <div style={{ fontSize: 13, color: '#A5A5A0', padding: '12px 0' }}>Sin puntos de control. Añade el primero.</div>
            )}

            <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 8 }}>
              {disciplina.puntos.map(p => {
                const est       = ESTADOS_INSP[p.estado] || ESTADOS_INSP.pendiente;
                const expandido = puntosExpandidos.has(p.id);
                const nMarcas   = p.plano?.marcas?.length || 0;
                return (
                  <div key={p.id} style={{ border: `1px solid ${expandido ? '#C5C4BE' : '#E8E7E1'}`, borderRadius: 9, overflow: 'hidden', transition: 'border-color .15s' }}>
                    {/* Fila principal */}
                    <div style={{ display: 'flex', flexDirection: isMobile ? 'column' : 'row', alignItems: isMobile ? 'stretch' : 'center', gap: isMobile ? 8 : 10, padding: '10px 12px', background: expandido ? '#FAFAF8' : '#fff' }}>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: 13, fontWeight: 400, color: '#18180F' }}>{p.nombre}</div>
                        {p.fecha && <div style={{ fontSize: 11, color: '#A5A5A0', marginTop: 2 }}>{fmtDate(p.fecha)}</div>}
                      </div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8, justifyContent: isMobile ? 'space-between' : 'flex-end' }}>
                        <select value={p.estado} onChange={e => updatePuntoEstado(p.id, e.target.value)}
                          style={{ width: isMobile ? '100%' : 'auto', fontSize: 12, padding: '5px 9px', borderRadius: 20, border: `1px solid ${est.color}30`, background: est.bg, color: est.color, fontWeight: 500, cursor: 'pointer' }}>
                          {Object.entries(ESTADOS_INSP).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
                        </select>
                        {/* Botón plano */}
                        <button onClick={() => togglePunto(p.id)}
                          style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '4px 9px', borderRadius: 6, border: '1px solid #E0DFD9', background: expandido ? '#1C1C1A' : '#fff', color: expandido ? '#F2F1ED' : '#6B6B66', fontSize: 11, fontWeight: 500, cursor: 'pointer', whiteSpace: 'nowrap', flexShrink: 0 }}>
                          / Plano{nMarcas > 0 ? ` (${nMarcas})` : ''}
                        </button>
                        <button onClick={() => setConfirmacion({ titulo: 'Eliminar punto de control', texto: `Vas a eliminar "${p.nombre}" y su plano. Esta acción no se puede deshacer.`, onSi: () => { deletePunto(p.id); setConfirmacion(null); } })} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#C4C3BE', fontSize: 18, padding: '0 2px', lineHeight: 1, flexShrink: 0 }}>×</button>
                      </div>
                    </div>
                    {/* Plano expandido */}
                    {expandido && (
                      <div className="fade" style={{ padding: '0 12px 12px', borderTop: '1px solid #F2F1ED', background: '#FAFAF8' }}>
                        <PlanoPunto punto={p} onUpdate={updatePunto} obra={obra} nombreDisciplina={disciplina.nombre} numActa={(obra.numActaSeq || 0) + 1} onActaGenerada={n => onSave({ ...obra, numActaSeq: n })} />
                      </div>
                    )}
                  </div>
                );
              })}
            </div>

            {showNuevoPunto ? (
              <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                <input autoFocus placeholder="Nombre del punto de control" value={nombrePunto} onChange={e => setNombrePunto(e.target.value)} onKeyDown={e => e.key === 'Enter' && addPunto()} style={{ flex: 1 }} />
                <Btn sm primary disabled={!nombrePunto.trim()} onClick={addPunto}>Añadir</Btn>
                <Btn sm onClick={() => { setShowNuevoPunto(false); setNombrePunto(''); }}>✕</Btn>
              </div>
            ) : (
              <DashedBtn onClick={() => setShowNuevoPunto(true)}>+ Añadir punto de control</DashedBtn>
            )}
          </>
        )}
      </div>
      )}
    </div>
  );
}

// ─── MÓDULO: Incidencias ──────────────────────────────────────────────────────

const ESTADOS_INC = {
  detectada: { label: 'Detectada',        bg: '#EEEDE7', color: '#52524E' },
  pendiente: { label: 'Pte. resolución',  bg: '#FEF3DB', color: '#7C4A00' },
  resuelta:  { label: 'Resuelta',         bg: '#E8F5E0', color: '#2D5E10' },
};

const DIAS_SEMANA = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];

// ─── Acta de Visita de Obra (VO) ──────────────────────────────────────────────
const ROLES_VO = [
  { k: 'promotor',     label: 'Promotor (PR)' },
  { k: 'do',           label: 'Dirección de Obra (DO)' },
  { k: 'deo',          label: 'Dirección de Ejecución (DEO)' },
  { k: 'estructuras',  label: 'Ing. Estructuras (DOE)' },
  { k: 'instalaciones',label: 'Ing. Instalaciones (DOI)' },
  { k: 'css',          label: 'Coordinador Seguridad (CSS)' },
  { k: 'contratista',  label: 'Contratista (EC)' },
];
const SECCIONES_VO = [
  { codigo: '1',   titulo: 'TEMAS TRATADOS' },
  { codigo: '2',   titulo: 'CONTROL DE CALIDAD' },
  { codigo: '3.1', titulo: 'INSTALACIONES' },
  { codigo: '3.2', titulo: 'ACOMETIDAS' },
  { codigo: '4',   titulo: 'SEGURIDAD Y SALUD' },
  { codigo: '5',   titulo: 'SEGUIMIENTO PLANIFICACIÓN' },
  { codigo: '6',   titulo: 'SEGUIMIENTO HITOS CONTRACTUALES' },
  { codigo: '7',   titulo: 'SEGUIMIENTO CONTRATACIÓN' },
  { codigo: '8',   titulo: 'SEGUIMIENTO PERSONAL' },
  { codigo: '9',   titulo: 'SEGUIMIENTO LEED / WELL / WIREDSCORE' },
];
const ESTADOS_VO = {
  P: { label: 'Pendiente',   bg: '#FEF3DB', color: '#7C4A00' },
  R: { label: 'Resuelto',    bg: '#E8F5E0', color: '#2D5E10' },
  I: { label: 'Informativo', bg: '#EEEDE7', color: '#52524E' },
};
const RESP_VO = ['', 'EC', 'DO', 'DEO', 'PR', 'DOE', 'DOI', 'CSS'];
const DIAS_DEFAULT = [1, 3]; // Lunes y Miércoles

function esHoyVisita(obra) {
  const dias = obra.diasVisita || [];
  if (dias.length === 0) return false;
  return dias.includes(new Date().getDay());
}
function revisadaHoy(inc) {
  return (inc.revisiones || []).some(r => r.fecha === today());
}
function diasSinRevisar(inc) {
  const revisiones = inc.revisiones || [];
  const ultima = revisiones.length > 0 ? revisiones[revisiones.length - 1].fecha : (inc.fechaCreacion || inc.ultimaActualizacion);
  return diasDesde(ultima);
}
function diasHasta(iso) {
  if (!iso) return null;
  return Math.ceil((new Date(iso) - new Date(today())) / 86400000);
}
function diasDesde(iso) {
  if (!iso) return 0;
  return Math.floor((Date.now() - new Date(iso)) / 86400000);
}
function pickFiles(accept, cb) {
  const inp = document.createElement('input');
  inp.type = 'file'; inp.multiple = true; inp.accept = accept;
  inp.style.position = 'fixed'; inp.style.left = '-9999px';
  document.body.appendChild(inp);
  inp.onchange = e => {
    Array.from(e.target.files).forEach(f => {
      if (f.size > 15 * 1024 * 1024) { alert(`"${f.name}" supera 15 MB`); return; }
      const r = new FileReader();
      r.onload = ev => {
        if (f.type.startsWith('image/')) {
          const img = new Image();
          img.onload = () => {
            const maxW = 720, ratio = Math.min(1, maxW / img.width);
            const c = document.createElement('canvas');
            c.width = Math.round(img.width * ratio);
            c.height = Math.round(img.height * ratio);
            c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
            cb({ id: uid(), nombre: f.name, tipo: 'imagen', data: c.toDataURL('image/jpeg', 0.55) });
          };
          img.onerror = () => alert('No se pudo procesar la imagen: ' + f.name);
          img.src = ev.target.result;
        } else {
          cb({ id: uid(), nombre: f.name, tipo: 'documento', data: ev.target.result });
        }
      };
      r.onerror = () => alert('No se pudo leer el archivo: ' + f.name);
      r.readAsDataURL(f);
    });
    setTimeout(() => { if (inp.parentNode) document.body.removeChild(inp); }, 100);
  };
  inp.click();
}

// ── Card de incidencia ────────────────────────────────────────────────────────
function IncCard({ inc, esVisitaHoy, onClick, onRevisar }) {
  const isMobile = useIsMobile();
  const est      = ESTADOS_INC[inc.estado] || ESTADOS_INC.detectada;
  const fotos    = (inc.historial || []).flatMap(h => h.adjuntos || []).filter(a => a.tipo === 'imagen');
  const foto     = fotos[0];
  const revisada = revisadaHoy(inc);
  const dias     = diasDesde(inc.ultimaActualizacion || inc.fechaCreacion);
  const sinRevisarAlerta = inc.estado !== 'resuelta' && diasSinRevisar(inc) >= 10;
  const ultNota  = (inc.historial || []).slice().reverse().find(h => h.nota && h.nota.trim());
  const mostrarRevisar = esVisitaHoy && inc.estado !== 'resuelta';

  return (
    <div className="tap" style={{ background: '#fff', border: `1px solid ${esVisitaHoy && !revisada && inc.estado !== 'resuelta' ? '#F5D98B' : '#E8E7E1'}`, borderRadius: 11, overflow: 'hidden', borderLeft: `3px solid ${est.color}` }}>

      <div style={{ display: 'flex', alignItems: 'stretch' }}>
        {/* Foto */}
        <div onClick={onClick} style={{ width: isMobile ? 70 : 80, flexShrink: 0, position: 'relative', background: foto ? 'transparent' : '#F5F4F0', display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer' }}>
          {foto
            ? <img src={foto.data} alt="" style={{ width: '100%', height: '100%', minHeight: isMobile ? 78 : 80, objectFit: 'cover', display: 'block' }} />
            : <span style={{ fontSize: 20, color: '#C5C4BE', fontWeight: 300 }}>—</span>}
          {fotos.length > 1 && (
            <span style={{ position: 'absolute', bottom: 4, right: 4, background: 'rgba(0,0,0,0.7)', color: '#fff', fontSize: 10, fontWeight: 500, padding: '1px 6px', borderRadius: 3 }}>{fotos.length} fotos</span>
          )}
        </div>

        {/* Info */}
        <div onClick={onClick} style={{ flex: 1, minWidth: 0, padding: '10px 13px', cursor: 'pointer', display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
          <div style={{ fontSize: 13.5, fontWeight: 600, color: '#141412', marginBottom: 3, lineHeight: 1.3, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{inc.titulo}</div>
          {ultNota && (
            <div style={{ fontSize: 12, color: '#9B9B97', marginBottom: 6, lineHeight: 1.35, overflow: 'hidden', display: '-webkit-box', WebkitLineClamp: 1, WebkitBoxOrient: 'vertical' }}>{ultNota.nota}</div>
          )}
          <div style={{ display: 'flex', gap: 7, alignItems: 'center', flexWrap: 'wrap' }}>
            <Pill label={est.label} bg={est.bg} color={est.color} />
            <span style={{ fontSize: 11, color: sinRevisarAlerta ? '#8A1F1F' : '#A5A5A0', fontWeight: sinRevisarAlerta ? 600 : 400 }}>
              {sinRevisarAlerta ? `${diasSinRevisar(inc)}d sin revisar` : dias === 0 ? 'Hoy' : `Hace ${dias}d`}
            </span>
          </div>
        </div>

        {/* Revisar a la derecha — solo escritorio */}
        {mostrarRevisar && !isMobile && (
          <div style={{ display: 'flex', alignItems: 'center', padding: '0 12px', borderLeft: '1px solid #E8E7E1', flexShrink: 0 }}>
            {revisada
              ? <span style={{ fontSize: 12, color: '#2D5E10', fontWeight: 500, display: 'flex', alignItems: 'center', gap: 4 }}>✓ Revisada</span>
              : <button onClick={e => { e.stopPropagation(); onRevisar(); }} style={{ padding: '6px 11px', borderRadius: 8, border: '1px solid #D4820A', background: '#FEF3DB', color: '#7C4A00', fontSize: 12, fontWeight: 500, cursor: 'pointer', whiteSpace: 'nowrap' }}>Revisar</button>}
          </div>
        )}
      </div>

      {/* Revisar abajo a lo ancho — solo móvil */}
      {mostrarRevisar && isMobile && (
        <div style={{ borderTop: '1px solid #F2F1ED', padding: '8px 12px' }}>
          {revisada
            ? <span style={{ fontSize: 12.5, color: '#2D5E10', fontWeight: 500, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 5 }}>✓ Revisada hoy</span>
            : <button onClick={e => { e.stopPropagation(); onRevisar(); }} style={{ width: '100%', padding: '9px', borderRadius: 8, border: '1px solid #D4820A', background: '#FEF3DB', color: '#7C4A00', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Revisar en visita</button>}
        </div>
      )}
    </div>
  );
}

// ── Formulario nueva incidencia ───────────────────────────────────────────────
function FormNuevaIncidencia({ onClose, onCrear }) {
  const [fotos,  setFotos]  = useState([]);
  const [titulo, setTitulo] = useState('');
  const [nota,   setNota]   = useState('');
  const [estado, setEstado] = useState('detectada');

  function crear() {
    if (!titulo.trim()) return;
    const ahora = now();
    const inc = {
      id: uid(), titulo: titulo.trim(), estado,
      revisiones: [],
      historial: [{ id: uid(), tipo: 'creacion', estado, nota: nota.trim(), adjuntos: fotos, fecha: ahora }],
      fechaCreacion: ahora, ultimaActualizacion: ahora,
    };
    onCrear(inc);
  }

  return (
    <div style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 12, padding: '16px' }}>
      <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 14 }}>Nueva incidencia</div>

      {/* Fotos */}
      <div style={{ marginBottom: 12 }}>
        <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 6 }}>Fotos</label>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          {fotos.map(f => (
            <div key={f.id} style={{ position: 'relative', flexShrink: 0 }}>
              <img src={f.data} alt="" style={{ width: 70, height: 56, objectFit: 'cover', borderRadius: 7, border: '1px solid #E0DFD9', display: 'block' }} />
              <button onClick={() => setFotos(p => p.filter(x => x.id !== f.id))} style={{ position: 'absolute', top: -5, right: -5, width: 16, height: 16, borderRadius: '50%', background: '#8A1F1F', color: '#fff', border: 'none', cursor: 'pointer', fontSize: 10, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>×</button>
            </div>
          ))}
          <button onClick={() => pickFiles('image/*', f => setFotos(p => [...p, f]))} style={{ width: 70, height: 56, borderRadius: 7, border: '1.5px dashed #E0DFD9', background: 'transparent', cursor: 'pointer', fontSize: 20, color: '#A5A5A0', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>📷</button>
        </div>
      </div>

      {/* Título */}
      <div style={{ marginBottom: 10 }}>
        <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 4 }}>Descripción breve *</label>
        <input autoFocus placeholder="¿Qué se ha detectado?" value={titulo} onChange={e => setTitulo(e.target.value)} />
      </div>

      {/* Nota */}
      <div style={{ marginBottom: 10 }}>
        <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 4 }}>Nota adicional</label>
        <textarea placeholder="Ubicación, contexto, a quién se notificará..." value={nota} onChange={e => setNota(e.target.value)} style={{ minHeight: 56 }} />
      </div>

      {/* Estado */}
      <div style={{ marginBottom: 14 }}>
        <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 6 }}>Estado inicial</label>
        <div style={{ display: 'flex', gap: 6 }}>
          {Object.entries(ESTADOS_INC).filter(([k]) => k !== 'resuelta').map(([k, v]) => (
            <button key={k} onClick={() => setEstado(k)} style={{ padding: '6px 12px', borderRadius: 20, border: `1.5px solid ${estado === k ? v.color : '#E0DFD9'}`, background: estado === k ? v.bg : 'transparent', color: estado === k ? v.color : '#6B6B66', fontSize: 12, cursor: 'pointer', fontWeight: estado === k ? 600 : 400, transition: 'all .15s' }}>
              {v.label}
            </button>
          ))}
        </div>
      </div>

      <div style={{ display: 'flex', gap: 8 }}>
        <Btn onClick={onClose} full>Cancelar</Btn>
        <Btn primary onClick={crear} disabled={!titulo.trim()} full>Registrar</Btn>
      </div>
    </div>
  );
}

// ── Detalle de incidencia ─────────────────────────────────────────────────────
function DetalleIncidencia({ inc, onClose, onActualizar, onEliminar }) {
  const [nota,     setNota]     = useState('');
  const [adjuntos, setAdjuntos] = useState([]);
  const [estado,   setEstado]   = useState(inc.estado);
  const [preview,  setPreview]  = useState(null); // imagen ampliada
  const [menu,     setMenu]     = useState(false);
  const [confirmar, setConfirmar] = useState(false);

  function guardar() {
    if (!nota.trim() && adjuntos.length === 0 && estado === inc.estado) return;
    const estadoCambio = estado !== inc.estado;
    const entrada = {
      id: uid(),
      tipo: estadoCambio ? 'cambio_estado' : 'nota',
      estado,
      nota: nota.trim(),
      adjuntos,
      fecha: now(),
    };
    onActualizar({ ...inc, estado, historial: [...(inc.historial || []), entrada], ultimaActualizacion: now() });
    setNota(''); setAdjuntos([]);
  }

  const est = ESTADOS_INC[inc.estado] || ESTADOS_INC.detectada;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden' }}>
      {/* Header */}
      <div style={{ padding: '12px 16px', borderBottom: '1px solid #E8E7E1', display: 'flex', alignItems: 'center', gap: 10, flexShrink: 0 }}>
        <button onClick={onClose} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 13, color: '#6B6B66', display: 'flex', alignItems: 'center', gap: 4 }}>← Volver</button>
        <span style={{ color: '#D4D3CE' }}>/</span>
        <span style={{ fontSize: 13, fontWeight: 500, flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{inc.titulo}</span>
        <Pill label={est.label} bg={est.bg} color={est.color} />
        <div style={{ position: 'relative', flexShrink: 0 }}>
          <button onClick={() => setMenu(m => !m)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#A5A5A0', fontSize: 18, lineHeight: 1, padding: '0 2px' }}>⋮</button>
          {menu && (
            <>
              <div onClick={() => setMenu(false)} style={{ position: 'fixed', inset: 0, zIndex: 20 }} />
              <div style={{ position: 'absolute', right: 0, top: '100%', background: '#fff', border: '1px solid #E0DFD9', borderRadius: 9, boxShadow: '0 8px 24px rgba(0,0,0,.12)', padding: 5, zIndex: 21, minWidth: 130 }}>
                <div onClick={() => { setMenu(false); setConfirmar(true); }} className="hov-row" style={{ padding: '7px 11px', borderRadius: 6, cursor: 'pointer', fontSize: 13, color: '#8A1F1F' }}>Eliminar</div>
              </div>
            </>
          )}
        </div>
      </div>

      {confirmar && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.2)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000, backdropFilter: 'blur(3px)' }} onClick={e => { if (e.target === e.currentTarget) setConfirmar(false); }}>
          <div className="fade" style={{ background: '#fff', borderRadius: 14, width: 380, maxWidth: '95vw', border: '1px solid #E0DFD9', boxShadow: '0 24px 64px rgba(0,0,0,.14)', padding: 20 }}>
            <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 8 }}>Eliminar incidencia</div>
            <div style={{ fontSize: 13, color: '#52524E', lineHeight: 1.5, marginBottom: 18 }}>Vas a eliminar <strong>{inc.titulo}</strong> y todo su historial. Esta acción no se puede deshacer.</div>
            <div style={{ display: 'flex', gap: 8 }}>
              <Btn onClick={() => setConfirmar(false)} full>Cancelar</Btn>
              <Btn danger full onClick={() => { setConfirmar(false); onEliminar(); }}>Eliminar</Btn>
            </div>
          </div>
        </div>
      )}

      <div style={{ flex: 1, overflow: 'auto', padding: '14px 16px' }}>

        {/* Cambio de estado */}
        <div style={{ background: '#F9F8F5', borderRadius: 10, padding: '12px 14px', marginBottom: 14 }}>
          <div style={{ fontSize: 11, color: '#A5A5A0', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8 }}>Estado</div>
          <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
            {Object.entries(ESTADOS_INC).map(([k, v]) => (
              <button key={k} onClick={() => setEstado(k)} style={{ padding: '6px 12px', borderRadius: 20, border: `1.5px solid ${estado === k ? v.color : '#E0DFD9'}`, background: estado === k ? v.bg : 'transparent', color: estado === k ? v.color : '#6B6B66', fontSize: 12, cursor: 'pointer', fontWeight: estado === k ? 600 : 400, transition: 'all .15s' }}>
                {v.label}
              </button>
            ))}
          </div>
        </div>

        {/* Historial */}
        <div style={{ fontSize: 11, textTransform: 'uppercase', letterSpacing: '0.1em', color: '#A5A5A0', fontWeight: 500, marginBottom: 10 }}>
          Historial
        </div>

            {(inc.historial || []).slice().reverse().map(h => (
              <div key={h.id} style={{
                padding: h.tipo === 'revision' ? '7px 12px' : '10px 12px',
                background: h.tipo === 'revision' ? 'transparent' : '#fff',
                border: `1px solid ${h.tipo === 'revision' ? '#EEECEA' : '#E8E7E1'}`,
                borderRadius: 9, marginBottom: 6,
              }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: (h.nota || h.adjuntos?.length) ? 6 : 0 }}>
                  <span style={{ fontSize: 11, color: '#A5A5A0' }}>{fmtDate(h.fecha)}</span>
                  {h.tipo === 'creacion' && (
                    <span style={{ fontSize: 11, color: '#6B6B66', background: '#F0EFEA', padding: '1px 7px', borderRadius: 20 }}>Registrada</span>
                  )}
                  {h.tipo === 'cambio_estado' && (
                    <Pill label={ESTADOS_INC[h.estado]?.label || h.estado} bg={ESTADOS_INC[h.estado]?.bg || '#eee'} color={ESTADOS_INC[h.estado]?.color || '#333'} />
                  )}
                  {h.tipo === 'nota' && (
                    <span style={{ fontSize: 11, color: '#A5A5A0', fontStyle: 'italic' }}>Actualización</span>
                  )}
                  {h.tipo === 'revision' && (
                    <span style={{ fontSize: 11, color: '#A5A5A0' }}>✓ Revisada en visita · Sin cambios</span>
                  )}
                </div>
                {h.nota && <div style={{ fontSize: 13, lineHeight: 1.5, marginBottom: h.adjuntos?.length ? 8 : 0 }}>{h.nota}</div>}
                {h.adjuntos?.length > 0 && (
                  <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                    {h.adjuntos.map(a => a.tipo === 'imagen'
                      ? <img key={a.id} src={a.data} alt={a.nombre} title="Clic para ampliar" onClick={() => setPreview(a)} style={{ width: 140, height: 110, objectFit: 'cover', borderRadius: 8, border: '1px solid #E0DFD9', display: 'block', cursor: 'zoom-in' }} />
                      : <a key={a.id} href={a.data} download={a.nombre} style={{ display: 'inline-flex', alignItems: 'center', gap: 4, fontSize: 12, padding: '4px 9px', borderRadius: 7, background: '#F5F4F0', border: '1px solid #E0DFD9', color: '#18180F', textDecoration: 'none' }}>📄 {a.nombre}</a>
                    )}
                  </div>
                )}
              </div>
            ))}

        {/* Añadir actualización */}
        <div style={{ background: '#F9F8F5', borderRadius: 10, padding: '12px 14px', marginTop: 4 }}>
          <div style={{ fontSize: 11, color: '#A5A5A0', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8 }}>Añadir actualización</div>
          <textarea placeholder="Nota de seguimiento, notificación enviada, acta adjunta..." value={nota} onChange={e => setNota(e.target.value)} style={{ marginBottom: 8, minHeight: 64 }} />

          {adjuntos.length > 0 && (
            <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginBottom: 8 }}>
              {adjuntos.map(a => (
                <div key={a.id} style={{ position: 'relative' }}>
                  {a.tipo === 'imagen'
                    ? <img src={a.data} alt="" style={{ width: 60, height: 48, objectFit: 'cover', borderRadius: 6, border: '1px solid #E0DFD9' }} />
                    : <div style={{ fontSize: 11, padding: '4px 8px', background: '#fff', border: '1px solid #E0DFD9', borderRadius: 6 }}>📄 {a.nombre}</div>
                  }
                  <button onClick={() => setAdjuntos(p => p.filter(x => x.id !== a.id))} style={{ position: 'absolute', top: -4, right: -4, width: 14, height: 14, borderRadius: '50%', background: '#8A1F1F', color: '#fff', border: 'none', cursor: 'pointer', fontSize: 9, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>×</button>
                </div>
              ))}
            </div>
          )}

          <div style={{ display: 'flex', gap: 8 }}>
            <button onClick={() => pickFiles('image/*,.pdf,.doc,.docx', f => setAdjuntos(p => [...p, f]))} style={{ padding: '7px 12px', borderRadius: 8, border: '1px solid #E0DFD9', background: '#fff', cursor: 'pointer', fontSize: 13, whiteSpace: 'nowrap' }}>📎 Adjuntar</button>
            <Btn primary full onClick={guardar} disabled={!nota.trim() && adjuntos.length === 0 && estado === inc.estado}>Guardar</Btn>
          </div>
        </div>
      </div>

      {/* Lightbox de imagen */}
      {preview && (
        <div onClick={() => setPreview(null)}
          style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.85)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 2000, padding: 30, cursor: 'zoom-out' }}>
          <img src={preview.data} alt={preview.nombre} style={{ maxWidth: '100%', maxHeight: '100%', objectFit: 'contain', borderRadius: 6 }} />
          <button onClick={() => setPreview(null)} style={{ position: 'absolute', top: 20, right: 24, background: 'rgba(255,255,255,0.15)', border: 'none', color: '#fff', fontSize: 24, width: 40, height: 40, borderRadius: '50%', cursor: 'pointer', lineHeight: 1 }}>×</button>
          <a href={preview.data} download={preview.nombre} onClick={e => e.stopPropagation()} style={{ position: 'absolute', bottom: 24, background: 'rgba(255,255,255,0.15)', color: '#fff', fontSize: 13, padding: '8px 16px', borderRadius: 8, textDecoration: 'none' }}>↓ Descargar {preview.nombre}</a>
        </div>
      )}
    </div>
  );
}

// ── Modal de revisión en visita ───────────────────────────────────────────────
function ModalRevision({ inc, onSinCambios, onConCambios, onClose }) {
  const [fase,     setFase]     = useState('pregunta'); // pregunta | cambios
  const [estado,   setEstado]   = useState(inc.estado);
  const [comentario, setComentario] = useState('');
  const [adjuntos, setAdjuntos] = useState([]);

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.25)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000, backdropFilter: 'blur(3px)' }}
      onClick={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="fade" style={{ background: '#fff', borderRadius: 14, width: 440, maxWidth: '95vw', border: '1px solid #E0DFD9', boxShadow: '0 24px 64px rgba(0,0,0,.14)' }}>

        {/* Header */}
        <div style={{ padding: '15px 18px', borderBottom: '1px solid #ECEAE4' }}>
          <div style={{ fontSize: 15, fontWeight: 600 }}>Revisar en visita</div>
          <div style={{ fontSize: 12, color: '#9B9B97', marginTop: 2 }}>{inc.titulo}</div>
        </div>

        <div style={{ padding: '18px' }}>
          {fase === 'pregunta' ? (
            <>
              <p style={{ fontSize: 14, color: '#52524E', marginBottom: 18, textAlign: 'center' }}>
                ¿Hay cambios desde la última visita?
              </p>
              <div style={{ display: 'flex', gap: 10 }}>
                <button onClick={() => onSinCambios(inc.id)} style={{ flex: 1, padding: '14px', borderRadius: 10, border: '1.5px solid #E0DFD9', background: '#fff', cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6, transition: 'all .15s' }}>
                  <span style={{ fontSize: 22 }}>✓</span>
                  <span style={{ fontSize: 13, fontWeight: 500, color: '#52524E' }}>Sin cambios</span>
                  <span style={{ fontSize: 11, color: '#A5A5A0' }}>Sigue igual</span>
                </button>
                <button onClick={() => setFase('cambios')} style={{ flex: 1, padding: '14px', borderRadius: 10, border: '1.5px solid #D48A0C', background: '#FEF3DB', cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6, transition: 'all .15s' }}>
                  <span style={{ fontSize: 22 }}>✎</span>
                  <span style={{ fontSize: 13, fontWeight: 500, color: '#7C4A00' }}>Hay cambios</span>
                  <span style={{ fontSize: 11, color: '#A5780A' }}>Actualizar</span>
                </button>
              </div>
            </>
          ) : (
            <>
              {/* Estado */}
              <div style={{ marginBottom: 14 }}>
                <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 6 }}>Estado</label>
                <div style={{ display: 'flex', gap: 6 }}>
                  {Object.entries(ESTADOS_INC).map(([k, v]) => (
                    <button key={k} onClick={() => setEstado(k)} style={{ padding: '6px 12px', borderRadius: 20, border: `1.5px solid ${estado === k ? v.color : '#E0DFD9'}`, background: estado === k ? v.bg : 'transparent', color: estado === k ? v.color : '#6B6B66', fontSize: 12, cursor: 'pointer', fontWeight: estado === k ? 600 : 400, transition: 'all .15s' }}>
                      {v.label}
                    </button>
                  ))}
                </div>
                {estado === inc.estado && <div style={{ fontSize: 11, color: '#A5A5A0', marginTop: 5 }}>Mismo estado · solo se añadirá el comentario</div>}
              </div>

              {/* Comentario */}
              <div style={{ marginBottom: 12 }}>
                <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 5 }}>Comentario</label>
                <textarea autoFocus placeholder="Qué ha cambiado, qué se ha hecho..." value={comentario} onChange={e => setComentario(e.target.value)} style={{ minHeight: 70 }} />
              </div>

              {/* Adjuntos */}
              {adjuntos.length > 0 && (
                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginBottom: 10 }}>
                  {adjuntos.map(a => (
                    <div key={a.id} style={{ position: 'relative' }}>
                      {a.tipo === 'imagen'
                        ? <img src={a.data} alt="" style={{ width: 60, height: 48, objectFit: 'cover', borderRadius: 6, border: '1px solid #E0DFD9' }} />
                        : <div style={{ fontSize: 11, padding: '4px 8px', background: '#F5F4F0', border: '1px solid #E0DFD9', borderRadius: 6 }}>📄 {a.nombre}</div>}
                      <button onClick={() => setAdjuntos(p => p.filter(x => x.id !== a.id))} style={{ position: 'absolute', top: -4, right: -4, width: 14, height: 14, borderRadius: '50%', background: '#8A1F1F', color: '#fff', border: 'none', cursor: 'pointer', fontSize: 9 }}>×</button>
                    </div>
                  ))}
                </div>
              )}

              <button onClick={() => pickFiles('image/*,.pdf,.doc,.docx', f => setAdjuntos(p => [...p, f]))} style={{ width: '100%', padding: '8px', borderRadius: 8, border: '1.5px dashed #E0DFD9', background: 'transparent', cursor: 'pointer', fontSize: 12, color: '#6B6B66', marginBottom: 14 }}>📎 Adjuntar foto o documento</button>

              <div style={{ display: 'flex', gap: 8 }}>
                <Btn onClick={() => setFase('pregunta')} full>← Atrás</Btn>
                <Btn primary full onClick={() => onConCambios(inc.id, estado, comentario, adjuntos)} disabled={estado === inc.estado && !comentario.trim()}>
                  Guardar
                </Btn>
              </div>
            </>
          )}
        </div>
      </div>
      {confirmacion && <ConfirmMini titulo={confirmacion.titulo} texto={confirmacion.texto} onSi={confirmacion.onSi} onNo={() => setConfirmacion(null)} />}
    </div>
  );
}

// ── Módulo principal ──────────────────────────────────────────────────────────
function ModuloIncidencias({ obra, onSave }) {
  const [vista,     setVista]     = useState('pendientes');
  const [incActiva, setIncActiva] = useState(null);
  const [showNueva, setShowNueva] = useState(false);
  const [revisando, setRevisando] = useState(null); // incidencia que se está revisando

  const esVisita    = esHoyVisita(obra);
  const pendientes  = obra.incidencias.filter(i => i.estado !== 'resuelta');
  const resueltas   = obra.incidencias.filter(i => i.estado === 'resuelta');
  const mostradas   = vista === 'pendientes' ? pendientes : vista === 'resueltas' ? resueltas : obra.incidencias;
  const sinRevisar  = pendientes.filter(i => !revisadaHoy(i));
  const incDetalle  = obra.incidencias.find(i => i.id === incActiva);

  function crearInc(inc) {
    onSave({ ...obra, incidencias: [inc, ...obra.incidencias] });
    setShowNueva(false);
    setIncActiva(inc.id);
  }

  function actualizarInc(updated) {
    onSave({ ...obra, incidencias: obra.incidencias.map(i => i.id === updated.id ? updated : i) });
  }

  function borrarInc(incId) {
    onSave({ ...obra, incidencias: obra.incidencias.filter(i => i.id !== incId) });
    setIncActiva(null);
  }

  // Revisión sin cambios — solo deja constancia
  function revisarSinCambios(incId) {
    const entrada = { id: uid(), tipo: 'revision', estado: obra.incidencias.find(i => i.id === incId)?.estado, nota: '', adjuntos: [], fecha: now() };
    onSave({
      ...obra,
      incidencias: obra.incidencias.map(i =>
        i.id === incId
          ? { ...i, revisiones: [...(i.revisiones || []), { fecha: today() }], historial: [...(i.historial || []), entrada], ultimaActualizacion: now() }
          : i
      ),
    });
    setRevisando(null);
  }

  // Revisión con cambios — cambia estado y/o añade comentario
  function revisarConCambios(incId, nuevoEstado, comentario, adjuntos) {
    const inc = obra.incidencias.find(i => i.id === incId);
    const estadoCambio = nuevoEstado !== inc.estado;
    const entrada = {
      id: uid(),
      tipo: estadoCambio ? 'cambio_estado' : 'nota',
      estado: nuevoEstado,
      nota: comentario.trim(),
      adjuntos: adjuntos || [],
      fecha: now(),
    };
    onSave({
      ...obra,
      incidencias: obra.incidencias.map(i =>
        i.id === incId
          ? { ...i, estado: nuevoEstado, revisiones: [...(i.revisiones || []), { fecha: today() }], historial: [...(i.historial || []), entrada], ultimaActualizacion: now() }
          : i
      ),
    });
    setRevisando(null);
  }

  // Vista detalle
  if (incDetalle) {
    return (
      <div style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 12, height: '100%', overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
        <DetalleIncidencia
          inc={incDetalle}
          onClose={() => setIncActiva(null)}
          onActualizar={updated => { actualizarInc(updated); }}
          onEliminar={() => borrarInc(incDetalle.id)}
        />
      </div>
    );
  }

  const proximaDiaVisita = () => {
    const diasVisita = obra.diasVisita || DIAS_DEFAULT;
    const hoy = new Date().getDay();
    let dias = 1;
    while (dias <= 7) {
      if (diasVisita.includes((hoy + dias) % 7)) return DIAS_SEMANA[(hoy + dias) % 7];
      dias++;
    }
    return '';
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>

      {/* Banner visita */}
      {esVisita && pendientes.length > 0 && (
        <div style={{ background: sinRevisar.length === 0 ? '#E8F5E0' : '#FEF3DB', border: `1px solid ${sinRevisar.length === 0 ? '#B8DFA8' : '#F5D98B'}`, borderRadius: 10, padding: '10px 14px', display: 'flex', alignItems: 'center', gap: 10 }}>
          <span style={{ fontSize: 18 }}>{sinRevisar.length === 0 ? '✅' : '📋'}</span>
          <div>
            <div style={{ fontSize: 13, fontWeight: 500, color: '#18180F' }}>
              {sinRevisar.length === 0
                ? 'Todas las incidencias revisadas en la visita de hoy'
                : `${sinRevisar.length} incidencia${sinRevisar.length > 1 ? 's' : ''} pendiente${sinRevisar.length > 1 ? 's' : ''} de revisar hoy`
              }
            </div>
            {sinRevisar.length > 0 && <div style={{ fontSize: 11, color: '#7C4A00', marginTop: 2 }}>Pulsa "Revisar" en cada una para marcarlas como vistas esta visita</div>}
          </div>
        </div>
      )}

      {/* Sin visita hoy: próxima visita */}
      {!esVisita && pendientes.length > 0 && (
        <div style={{ background: '#F9F8F5', border: '1px solid #E8E7E1', borderRadius: 10, padding: '8px 14px', display: 'flex', alignItems: 'center', gap: 8 }}>
          <span style={{ fontSize: 13 }}>📅</span>
          <span style={{ fontSize: 12, color: '#6B6B66' }}>Próxima visita: <strong>{proximaDiaVisita()}</strong> · {pendientes.length} incidencia{pendientes.length > 1 ? 's' : ''} pendiente{pendientes.length > 1 ? 's' : ''}</span>
        </div>
      )}

      {/* Config días de visita */}
      <div style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 10, padding: '10px 14px', display: 'flex', alignItems: 'center', gap: 12 }}>
        <div style={{ fontSize: 12, color: '#6B6B66', whiteSpace: 'nowrap', flexShrink: 0 }}>Días de visita:</div>
        <DiasPicker
          value={obra.diasVisita || []}
          onChange={dias => onSave({ ...obra, diasVisita: dias })}
        />
        {(!obra.diasVisita || obra.diasVisita.length === 0) && (
          <span style={{ fontSize: 11, color: '#A5A5A0' }}>Sin días configurados — selecciona los días que vas a esta obra</span>
        )}
      </div>

      {/* Filtros + botón nueva */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
        <div style={{ display: 'flex', gap: 4, flex: 1, flexWrap: 'wrap' }}>
          {[
            ['pendientes', 'Pendientes', pendientes.length],
            ['todas',      'Todas',      obra.incidencias.length],
            ['resueltas',  'Resueltas',  resueltas.length],
          ].map(([id, label, count]) => (
            <button key={id} onClick={() => setVista(id)} style={{ padding: '5px 12px', borderRadius: 20, border: `1px solid ${vista === id ? '#18180F' : '#E0DFD9'}`, background: vista === id ? '#18180F' : 'transparent', color: vista === id ? '#fff' : '#6B6B66', fontSize: 12, cursor: 'pointer', fontWeight: vista === id ? 500 : 400, transition: 'all .15s' }}>
              {label}{count > 0 ? ` (${count})` : ''}
            </button>
          ))}
        </div>
        <Btn primary onClick={() => setShowNueva(v => !v)}>
          {showNueva ? '✕ Cancelar' : '+ Nueva incidencia'}
        </Btn>
      </div>

      {/* Form nueva */}
      {showNueva && <FormNuevaIncidencia onClose={() => setShowNueva(false)} onCrear={crearInc} />}

      {/* Lista */}
      {mostradas.length === 0 ? (
        <div style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 12, padding: '40px 20px', textAlign: 'center', color: '#A5A5A0' }}>
          <div style={{ fontSize: 32, marginBottom: 10 }}>{vista === 'resueltas' ? '✅' : '✓'}</div>
          <div style={{ fontSize: 13 }}>
            {vista === 'pendientes' ? 'Sin incidencias pendientes' : vista === 'resueltas' ? 'Aún no hay incidencias resueltas' : 'Sin incidencias registradas'}
          </div>
          {vista === 'pendientes' && !showNueva && <div style={{ marginTop: 12 }}><Btn onClick={() => setShowNueva(true)}>+ Nueva incidencia</Btn></div>}
        </div>
      ) : (
        <div className="list-in" style={{ display: 'flex', flexDirection: 'column', gap: 7 }}>
          {mostradas.map(i => (
            <IncCard
              key={i.id}
              inc={i}
              esVisitaHoy={esVisita}
              onClick={() => setIncActiva(i.id)}
              onRevisar={() => setRevisando(i)}
            />
          ))}
        </div>
      )}

      {/* Modal de revisión */}
      {revisando && (
        <ModalRevision
          inc={revisando}
          onSinCambios={revisarSinCambios}
          onConCambios={revisarConCambios}
          onClose={() => setRevisando(null)}
        />
      )}
    </div>
  );
}

// ─── MÓDULO: Notas y tareas ───────────────────────────────────────────────────

const CATEGORIAS_APT = ['Estructuras','Instalaciones','Obra civil','Acabados','Costes','Documentación','Otros'];

function ModuloApuntes({ obra, onSave }) {
  const apuntes = obra.apuntes || [];
  const [filtro,   setFiltro]   = useState('todo');
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ tipo: 'tarea', texto: '', categoria: 'Otros', fechaLimite: '' });
  const [confirmacion, setConfirmacion] = useState(null);
  const upd = (k, v) => setForm(f => ({ ...f, [k]: v }));

  function guardar() {
    if (!form.texto.trim()) return;
    const item = {
      id: uid(), tipo: form.tipo,
      texto: form.texto.trim(),
      categoria: form.categoria,
      fechaLimite: form.tipo === 'tarea' ? form.fechaLimite : '',
      hecha: false,
      creadaEn: now(),
    };
    onSave({ ...obra, apuntes: [item, ...apuntes] });
    setForm({ tipo: 'tarea', texto: '', categoria: 'Otros', fechaLimite: '' });
    setShowForm(false);
  }

  function toggleHecha(id) {
    onSave({ ...obra, apuntes: apuntes.map(a => a.id === id ? { ...a, hecha: !a.hecha } : a) });
  }

  function eliminar(id) {
    onSave({ ...obra, apuntes: apuntes.filter(a => a.id !== id) });
  }

  const tareasPend  = apuntes.filter(a => a.tipo === 'tarea' && !a.hecha);
  const tareasHecha = apuntes.filter(a => a.tipo === 'tarea' && a.hecha);
  const notas       = apuntes.filter(a => a.tipo === 'nota');

  const mostrados = filtro === 'todo'    ? apuntes
    : filtro === 'pendientes'            ? tareasPend
    : filtro === 'hechas'                ? tareasHecha
    : notas;

  function isVencida(item) {
    if (!item.fechaLimite || item.hecha) return false;
    return new Date(item.fechaLimite) < new Date(today());
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>

      {/* Stats rápidas */}
      {tareasPend.length > 0 && (
        <div style={{ background: apuntes.some(isVencida) ? '#FFF0F0' : '#FEF3DB', border: `1px solid ${apuntes.some(isVencida) ? '#FDCECE' : '#F5D98B'}`, borderRadius: 10, padding: '9px 14px', display: 'flex', alignItems: 'center', gap: 8 }}>
          <span style={{ fontSize: 16 }}>{apuntes.some(isVencida) ? '⚠️' : '📋'}</span>
          <span style={{ fontSize: 13, fontWeight: 500 }}>
            {tareasPend.length} tarea{tareasPend.length > 1 ? 's' : ''} pendiente{tareasPend.length > 1 ? 's' : ''}
            {apuntes.some(isVencida) && <span style={{ color: '#8A1F1F' }}> · {apuntes.filter(isVencida).length} vencida{apuntes.filter(isVencida).length > 1 ? 's' : ''}</span>}
          </span>
        </div>
      )}

      {/* Filtros + botón nuevo */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
        <div style={{ display: 'flex', gap: 4, flex: 1, flexWrap: 'wrap' }}>
          {[
            ['todo',       'Todo',              apuntes.length],
            ['pendientes', 'Tareas pendientes', tareasPend.length],
            ['hechas',     'Tareas hechas',     tareasHecha.length],
            ['notas',      'Notas',             notas.length],
          ].map(([id, label, count]) => (
            <button key={id} onClick={() => setFiltro(id)} style={{ padding: '5px 12px', borderRadius: 20, border: `1px solid ${filtro === id ? '#18180F' : '#E0DFD9'}`, background: filtro === id ? '#18180F' : 'transparent', color: filtro === id ? '#fff' : '#6B6B66', fontSize: 12, cursor: 'pointer', fontWeight: filtro === id ? 500 : 400, transition: 'all .15s' }}>
              {label}{count > 0 ? ` (${count})` : ''}
            </button>
          ))}
        </div>
        <Btn primary onClick={() => setShowForm(v => !v)}>
          {showForm ? '✕ Cancelar' : '+ Nuevo apunte'}
        </Btn>
      </div>

      {/* Formulario nuevo apunte */}
      {showForm && (
        <div style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 12, padding: '14px 16px' }}>
          {/* Tipo toggle */}
          <div style={{ display: 'flex', gap: 6, marginBottom: 12 }}>
            {[['tarea','☑ Tarea'],['nota','📝 Nota']].map(([t, l]) => (
              <button key={t} onClick={() => upd('tipo', t)} style={{ padding: '6px 14px', borderRadius: 20, border: `1.5px solid ${form.tipo === t ? '#18180F' : '#E0DFD9'}`, background: form.tipo === t ? '#18180F' : 'transparent', color: form.tipo === t ? '#fff' : '#6B6B66', fontSize: 12, cursor: 'pointer', fontWeight: form.tipo === t ? 600 : 400, transition: 'all .15s' }}>
                {l}
              </button>
            ))}
          </div>

          <div style={{ marginBottom: 10 }}>
            <textarea autoFocus placeholder={form.tipo === 'tarea' ? 'Descripción de la tarea...' : 'Escribe tu nota...'} value={form.texto} onChange={e => upd('texto', e.target.value)} style={{ minHeight: 70 }} />
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: form.tipo === 'tarea' ? '1fr 1fr' : '1fr', gap: 10, marginBottom: 12 }}>
            <div>
              <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 4 }}>Categoría</label>
              <select value={form.categoria} onChange={e => upd('categoria', e.target.value)}>
                {CATEGORIAS_APT.map(c => <option key={c}>{c}</option>)}
              </select>
            </div>
            {form.tipo === 'tarea' && (
              <div>
                <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 4 }}>Fecha límite</label>
                <input type="date" value={form.fechaLimite} onChange={e => upd('fechaLimite', e.target.value)} />
              </div>
            )}
          </div>

          <div style={{ display: 'flex', gap: 8 }}>
            <Btn onClick={() => setShowForm(false)} full>Cancelar</Btn>
            <Btn primary onClick={guardar} disabled={!form.texto.trim()} full>
              {form.tipo === 'tarea' ? 'Añadir tarea' : 'Añadir nota'}
            </Btn>
          </div>
        </div>
      )}

      {/* Lista */}
      {mostrados.length === 0 ? (
        <div style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 12, padding: '40px 20px', textAlign: 'center', color: '#A5A5A0' }}>
          <div style={{ fontSize: 32, marginBottom: 10 }}>
            {filtro === 'hechas' ? '✅' : filtro === 'notas' ? '📝' : '📋'}
          </div>
          <div style={{ fontSize: 13 }}>
            {filtro === 'todo' ? 'Sin apuntes todavía' : filtro === 'pendientes' ? 'Sin tareas pendientes' : filtro === 'hechas' ? 'Sin tareas completadas' : 'Sin notas'}
          </div>
          {filtro === 'todo' && !showForm && <div style={{ marginTop: 12 }}><Btn onClick={() => setShowForm(true)}>+ Nuevo apunte</Btn></div>}
        </div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          {mostrados.map(item => {
            const vencida = isVencida(item);
            return (
              <div key={item.id} style={{ background: '#fff', border: `1px solid ${vencida ? '#F4ABAB' : '#E8E7E1'}`, borderLeft: vencida ? '3px solid #E24B4A' : '1px solid #E8E7E1', borderRadius: 10, padding: '10px 13px', display: 'flex', alignItems: 'flex-start', gap: 10 }}>

                {/* Checkbox (solo tareas) */}
                {item.tipo === 'tarea' && (
                  <button onClick={() => toggleHecha(item.id)} style={{ width: 20, height: 20, borderRadius: 5, border: `2px solid ${item.hecha ? '#52A124' : '#D4D3CE'}`, background: item.hecha ? '#52A124' : 'transparent', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0, marginTop: 1, fontSize: 11, color: '#fff', transition: 'all .15s' }}>
                    {item.hecha ? '✓' : ''}
                  </button>
                )}

                {/* Icono nota */}
                {item.tipo === 'nota' && (
                  <span style={{ fontSize: 15, flexShrink: 0, marginTop: 1 }}>📝</span>
                )}

                {/* Contenido */}
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 13, color: item.hecha ? '#A5A5A0' : '#18180F', textDecoration: item.hecha ? 'line-through' : 'none', lineHeight: 1.4, marginBottom: 5 }}>
                    {item.texto}
                  </div>
                  <div style={{ display: 'flex', gap: 6, alignItems: 'center', flexWrap: 'wrap' }}>
                    <span style={{ fontSize: 11, padding: '1px 7px', borderRadius: 20, background: '#F0EFEA', color: '#6B6B66' }}>{item.categoria}</span>
                    {item.tipo === 'tarea' && item.fechaLimite && (
                      <span style={{ fontSize: 11, color: vencida ? '#8A1F1F' : '#A5A5A0', fontWeight: vencida ? 500 : 400 }}>
                        {vencida ? '⚠ Vencida · ' : '📅 '}{fmtDate(item.fechaLimite)}
                      </span>
                    )}
                    {item.tipo === 'nota' && (
                      <span style={{ fontSize: 11, color: '#A5A5A0' }}>{fmtShort(item.creadaEn)}</span>
                    )}
                  </div>
                </div>

                {/* Eliminar */}
                <button onClick={() => setConfirmacion({ titulo: 'Eliminar', texto: 'Vas a eliminar este elemento. Esta acción no se puede deshacer.', onSi: () => { eliminar(item.id); setConfirmacion(null); } })} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#D4D3CE', fontSize: 16, padding: '0 2px', lineHeight: 1, flexShrink: 0 }}>×</button>
              </div>
            );
          })}
        </div>
      )}
      {confirmacion && <ConfirmMini titulo={confirmacion.titulo} texto={confirmacion.texto} onSi={confirmacion.onSi} onNo={() => setConfirmacion(null)} />}
    </div>
  );
}

// ─── MÓDULO: Ensayos (importación desde Presto) ───────────────────────────────

// Lee el Excel de Presto y agrupa partidas por capítulo, extrayendo líneas de medición
function parseExcelEnsayos(uint8) {
  const wb = XLSX.read(uint8, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  const capitulos = [];
  let capActual = null;
  let ensActual = null;

  for (const row of rows) {
    const cod     = row[0];   // A · Código
    const nat     = row[1];   // B · Naturaleza
    const ud      = row[2];   // C · Unidad
    const resumen = row[3];   // D · Resumen
    const coment  = row[4];   // E · Comentario (línea de medición)
    const cantLin = row[9];   // J · Cantidad de la línea
    const canPres = row[10];  // K · CanPres (total partida)

    if (nat === 'Capítulo') {
      ensActual = null;
      if (resumen && String(resumen).trim().toUpperCase() !== 'NOTA') {
        capActual = { codigo: String(cod || ''), nombre: String(resumen).trim(), ensayos: [] };
        capitulos.push(capActual);
      } else capActual = null;
    } else if (nat === 'Partida' && capActual) {
      if (resumen && String(resumen).trim().toUpperCase() !== 'NOTA') {
        ensActual = {
          id: uid(),
          codigo: String(cod || ''),
          nombre: String(resumen).trim(),
          unidad: ud ? String(ud).trim() : '',
          cantidad: Number(canPres) || 0,
          _lineas: [],
        };
        capActual.ensayos.push(ensActual);
      } else ensActual = null;
    } else if (ensActual) {
      const esTotal = (resumen && String(resumen).startsWith('Total')) || (row[9] && String(row[9]).startsWith('Total'));
      if (coment && String(coment).trim() && !esTotal) {
        ensActual._lineas.push({ nombre: String(coment).trim(), cantidad: Number(cantLin) || 0 });
      }
    }
  }

  // Post-proceso: decidir unidades con nombre vs criterio único
  for (const cap of capitulos) {
    for (const ens of cap.ensayos) {
      const lineas = ens._lineas || [];
      const conCantidad = lineas.filter(l => l.cantidad >= 1);
      const notas       = lineas.filter(l => l.cantidad < 1);

      if (conCantidad.length >= 2) {
        // Unidades con nombre: se desglosa cada línea según su cantidad
        ens.unidades = [];
        conCantidad.forEach(l => {
          const n = Math.max(1, Math.round(l.cantidad));
          if (n === 1) {
            ens.unidades.push({ id: uid(), nombre: l.nombre, marca: null });
          } else {
            for (let i = 1; i <= n; i++) {
              ens.unidades.push({ id: uid(), nombre: `${l.nombre} (${i}/${n})`, marca: null });
            }
          }
        });
        ens.criterio = notas.map(n => n.nombre).join(' · ');
        ens.registros = [];
      } else {
        // Criterio único: contador de ejecutadas + actas adjuntas
        ens.unidades = [];
        ens.criterio = lineas.map(l => l.nombre).join(' · ');
        ens.ejecutadas = 0;
        ens.registros = [];
      }
      delete ens._lineas;
    }
  }
  return capitulos.filter(c => c.ensayos.length > 0);
}

// Progreso de un ensayo según su modo
function ensayoProgreso(ens) {
  if (ens.unidades && ens.unidades.length > 0) {
    return { modo: 'named', hechas: ens.unidades.filter(u => u.marca).length, total: ens.unidades.length };
  }
  return { modo: 'criterio', hechas: ens.ejecutadas || 0, total: ens.cantidad };
}

// Abre selector y lee el Excel como arrayBuffer
function importarExcel(cb) {
  const inp = document.createElement('input');
  inp.type = 'file'; inp.accept = '.xlsx,.xls';
  inp.style.position = 'fixed'; inp.style.left = '-9999px';
  document.body.appendChild(inp);
  inp.onchange = e => {
    const f = e.target.files[0];
    if (!f) return;
    const r = new FileReader();
    r.onload = ev => {
      try {
        const data = new Uint8Array(ev.target.result);
        const caps = parseExcelEnsayos(data);
        if (caps.length === 0) { alert('No se encontraron ensayos en el Excel. ¿Es un export de presupuesto de Presto?'); }
        else cb(caps, f.name);
      } catch (err) {
        alert('No se pudo leer el Excel: ' + err.message);
      }
      setTimeout(() => { if (inp.parentNode) document.body.removeChild(inp); }, 100);
    };
    r.readAsArrayBuffer(f);
  };
  inp.click();
}

const RESULTADOS = {
  apto:    { label: 'Apto',     bg: '#E8F5E0', color: '#2D5E10' },
  no_apto: { label: 'No apto',  bg: '#FDECEC', color: '#8A1F1F' },
  info:    { label: 'Sin valoración', bg: '#EEEDE7', color: '#52524E' },
};

function ModuloEnsayos({ obra, onSave }) {
  const datos = obra.ensayos || null;
  const [expandido, setExpandido] = useState(null);   // capítulo abierto
  const [ensayoActivo, setEnsayoActivo] = useState(null); // ensayo abierto para registros
  const [preview, setPreview] = useState(null);
  const [pendiente, setPendiente] = useState(null);   // import a confirmar { capitulos, nombre }

  function onImportar(capitulos, nombre) {
    if (datos) {
      // Ya hay datos: pedir confirmación con modal propio (window.confirm no funciona en iframe)
      setPendiente({ capitulos, nombre });
    } else {
      guardarImport(capitulos, nombre);
    }
  }

  function guardarImport(capitulos, nombre) {
    onSave({ ...obra, ensayos: { archivoNombre: nombre, importadoEn: now(), capitulos } });
    setExpandido(null); setEnsayoActivo(null); setPendiente(null);
  }

  function actualizarEnsayo(capIdx, ensId, mut) {
    const capitulos = datos.capitulos.map((c, i) => {
      if (i !== capIdx) return c;
      return { ...c, ensayos: c.ensayos.map(e => e.id === ensId ? mut(e) : e) };
    });
    onSave({ ...obra, ensayos: { ...datos, capitulos } });
  }

  // Estado vacío — importar
  if (!datos) {
    return (
      <div style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 12, padding: '48px 24px', textAlign: 'center', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 14 }}>
        <div style={{ width: 56, height: 56, borderRadius: 16, background: '#1C1C1A', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 26 }}>🧪</div>
        <div>
          <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 6 }}>Importar ensayos desde Presto</div>
          <div style={{ fontSize: 13, color: '#9B9B97', maxWidth: 420, lineHeight: 1.5 }}>
            Carga el Excel exportado del presupuesto de Presto. La app reconocerá los capítulos y creará la lista de ensayos automáticamente.
          </div>
        </div>
        <Btn primary onClick={() => importarExcel(onImportar)}>📂 Cargar Excel de Presto</Btn>
      </div>
    );
  }

  const totalEnsayos = datos.capitulos.reduce((s, c) => s + c.ensayos.length, 0);
  const totalEjecutados = datos.capitulos.reduce((s, c) => s + c.ensayos.filter(e => ensayoProgreso(e).hechas > 0).length, 0);
  const ensayoAbierto = (() => {
    for (let ci = 0; ci < datos.capitulos.length; ci++) {
      const e = datos.capitulos[ci].ensayos.find(x => x.id === ensayoActivo);
      if (e) return { capIdx: ci, ensayo: e };
    }
    return null;
  })();

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>

      {/* Cabecera */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
        <div style={{ flex: 1, minWidth: 200 }}>
          <div style={{ fontSize: 13, fontWeight: 500 }}>{totalEjecutados}/{totalEnsayos} ensayos iniciados</div>
          <div style={{ fontSize: 11, color: '#9B9B97', marginTop: 1 }}>📄 {datos.archivoNombre} · importado {fmtShort(datos.importadoEn)}</div>
        </div>
        <button onClick={() => importarExcel(onImportar)} style={{ fontSize: 12, color: '#6B6B66', background: '#fff', border: '1px solid #E0DFD9', borderRadius: 8, padding: '6px 11px', cursor: 'pointer' }}>Reimportar</button>
      </div>

      {/* Detalle de ensayo abierto */}
      {ensayoAbierto ? (
        <DetalleEnsayo
          ensayo={ensayoAbierto.ensayo}
          onClose={() => setEnsayoActivo(null)}
          onPreview={setPreview}
          onUpdate={(mut) => actualizarEnsayo(ensayoAbierto.capIdx, ensayoAbierto.ensayo.id, mut)}
        />
      ) : (
        /* Lista de capítulos */
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {datos.capitulos.map((cap, capIdx) => {
            const ejec = cap.ensayos.filter(e => ensayoProgreso(e).hechas > 0).length;
            const abierto = expandido === cap.codigo;
            return (
              <div key={cap.codigo} style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 11, overflow: 'hidden' }}>
                {/* Cabecera capítulo */}
                <div onClick={() => setExpandido(abierto ? null : cap.codigo)} style={{ padding: '12px 15px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 10 }}>
                  <span style={{ fontSize: 11, fontWeight: 600, color: '#A5A5A0', minWidth: 24 }}>{cap.codigo}</span>
                  <div style={{ flex: 1, fontSize: 14, fontWeight: 600, color: '#141412' }}>{cap.nombre}</div>
                  <span style={{ fontSize: 11, padding: '2px 9px', borderRadius: 20, background: ejec === cap.ensayos.length ? '#E8F5E0' : '#F0EFEA', color: ejec === cap.ensayos.length ? '#2D5E10' : '#6B6B66', fontWeight: 500 }}>
                    {ejec}/{cap.ensayos.length}
                  </span>
                  <span style={{ fontSize: 11, color: '#A5A5A0' }}>{abierto ? '▲' : '▼'}</span>
                </div>

                {/* Ensayos del capítulo */}
                {abierto && (
                  <div style={{ borderTop: '1px solid #F2F1ED' }}>
                    {cap.ensayos.map(ens => {
                      const p = ensayoProgreso(ens);
                      const completo = p.total > 0 && p.hechas >= p.total;
                      const etiqueta = p.modo === 'named'
                        ? (p.hechas > 0 ? `${p.hechas}/${p.total} ud.` : `${p.total} ud. · Pendiente`)
                        : (p.hechas > 0 ? `${p.hechas}/${p.total}` : 'Pendiente');
                      return (
                        <div key={ens.id} onClick={() => setEnsayoActivo(ens.id)} className="hov-row" style={{ padding: '10px 15px', borderTop: '1px solid #F7F6F2', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 10 }}>
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontSize: 13, color: '#141412', lineHeight: 1.35 }}>{ens.nombre}</div>
                            <div style={{ fontSize: 11, color: '#A5A5A0', marginTop: 2 }}>{ens.codigo} · {ens.cantidad} {ens.unidad}{p.modo === 'named' ? ' · por unidades' : ''}</div>
                          </div>
                          <span style={{ fontSize: 11, padding: '2px 9px', borderRadius: 20, whiteSpace: 'nowrap', background: completo ? '#E8F5E0' : p.hechas > 0 ? '#FEF3DB' : '#F0EFEA', color: completo ? '#2D5E10' : p.hechas > 0 ? '#7C4A00' : '#9B9B97', fontWeight: 500 }}>
                            {etiqueta}
                          </span>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {/* Lightbox */}
      {preview && (
        <div onClick={() => setPreview(null)} style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.85)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 2000, padding: 30, cursor: 'zoom-out' }}>
          <img src={preview.data} alt={preview.nombre} style={{ maxWidth: '100%', maxHeight: '100%', objectFit: 'contain', borderRadius: 6 }} />
          <button onClick={() => setPreview(null)} style={{ position: 'absolute', top: 20, right: 24, background: 'rgba(255,255,255,0.15)', border: 'none', color: '#fff', fontSize: 24, width: 40, height: 40, borderRadius: '50%', cursor: 'pointer', lineHeight: 1 }}>×</button>
        </div>
      )}

      {/* Confirmación de reimportación */}
      {pendiente && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.2)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000, backdropFilter: 'blur(3px)' }} onClick={e => { if (e.target === e.currentTarget) setPendiente(null); }}>
          <div className="fade" style={{ background: '#fff', borderRadius: 14, width: 420, maxWidth: '95vw', border: '1px solid #E0DFD9', boxShadow: '0 24px 64px rgba(0,0,0,.14)', padding: '20px' }}>
            <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 8 }}>Reemplazar ensayos</div>
            <div style={{ fontSize: 13, color: '#52524E', lineHeight: 1.5, marginBottom: 18 }}>
              Ya hay ensayos importados. Al cargar <strong>{pendiente.nombre}</strong> se reemplazará la lista actual y se perderán los registros y unidades marcadas. ¿Continuar?
            </div>
            <div style={{ display: 'flex', gap: 8 }}>
              <Btn onClick={() => setPendiente(null)} full>Cancelar</Btn>
              <Btn danger full onClick={() => guardarImport(pendiente.capitulos, pendiente.nombre)}>Reemplazar</Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
function DetalleEnsayo({ ensayo, onClose, onUpdate, onPreview }) {
  const esNamed = (ensayo.unidades || []).length > 0;
  const p = ensayoProgreso(ensayo);

  return (
    <div style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 12, overflow: 'hidden' }}>
      {/* Header */}
      <div style={{ padding: '13px 16px', borderBottom: '1px solid #ECEAE4', display: 'flex', alignItems: 'flex-start', gap: 10 }}>
        <button onClick={onClose} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 13, color: '#6B6B66', padding: 0, whiteSpace: 'nowrap' }}>← Volver</button>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 14, fontWeight: 600 }}>{ensayo.nombre}</div>
          <div style={{ fontSize: 12, color: '#9B9B97', marginTop: 2 }}>
            {ensayo.codigo} · {ensayo.cantidad} {ensayo.unidad}
            {esNamed ? ` · ${p.hechas}/${p.total} unidades ejecutadas` : ` · ${p.hechas}/${p.total} ejecutadas`}
          </div>
        </div>
      </div>

      <div style={{ padding: '14px 16px' }}>
        {/* Criterio de medición */}
        {ensayo.criterio && (
          <div style={{ background: '#F0F4FA', border: '1px solid #D3E2F5', borderRadius: 9, padding: '9px 12px', marginBottom: 14, display: 'flex', gap: 8 }}>
            <span style={{ fontSize: 13 }}>📐</span>
            <div>
              <div style={{ fontSize: 10, color: '#0C447C', textTransform: 'uppercase', letterSpacing: '0.06em', fontWeight: 600, marginBottom: 2 }}>Criterio de medición</div>
              <div style={{ fontSize: 12, color: '#1A3A5C', lineHeight: 1.45 }}>{ensayo.criterio}</div>
            </div>
          </div>
        )}

        {esNamed
          ? <UnidadesNombradas ensayo={ensayo} onUpdate={onUpdate} onPreview={onPreview} />
          : <RegistrosLibres ensayo={ensayo} onUpdate={onUpdate} onPreview={onPreview} />
        }
      </div>
    </div>
  );
}

// ── Modo A: unidades con nombre (marcar cada una) ─────────────────────────────
function UnidadesNombradas({ ensayo, onUpdate, onPreview }) {
  const [editId,     setEditId]     = useState(null);
  const [resultado,  setResultado]  = useState('apto');
  const [comentario, setComentario] = useState('');
  const [adjuntos,   setAdjuntos]   = useState([]);

  function abrir(u) {
    setEditId(u.id);
    setResultado(u.marca?.resultado || 'apto');
    setComentario(u.marca?.comentario || '');
    setAdjuntos(u.marca?.adjuntos || []);
  }
  function guardar() {
    const marca = { resultado, comentario: comentario.trim(), adjuntos, fecha: now() };
    onUpdate(e => ({ ...e, unidades: e.unidades.map(u => u.id === editId ? { ...u, marca } : u) }));
    setEditId(null);
  }
  function quitar(id) {
    onUpdate(e => ({ ...e, unidades: e.unidades.map(u => u.id === id ? { ...u, marca: null } : u) }));
    setEditId(null);
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 7 }}>
      {ensayo.unidades.map(u => {
        const res = u.marca ? (RESULTADOS[u.marca.resultado] || RESULTADOS.info) : null;
        const editando = editId === u.id;
        return (
          <div key={u.id} style={{ border: `1px solid ${editando ? '#C5C4BE' : '#E8E7E1'}`, borderRadius: 9, overflow: 'hidden' }}>
            {/* Fila unidad */}
            <div onClick={() => editando ? setEditId(null) : abrir(u)} style={{ padding: '10px 12px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 10, background: u.marca ? '#FAFCF9' : '#fff' }}>
              <div style={{ width: 24, height: 24, borderRadius: 6, flexShrink: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700, background: u.marca ? (u.marca.resultado === 'no_apto' ? '#FDECEC' : u.marca.resultado === 'apto' ? '#E8F5E0' : '#EEEDE7') : '#F0EFEA', color: u.marca ? (u.marca.resultado === 'no_apto' ? '#8A1F1F' : u.marca.resultado === 'apto' ? '#2D5E10' : '#52524E') : '#C5C4BE' }}>
                {u.marca ? (u.marca.resultado === 'no_apto' ? '✕' : '✓') : '○'}
              </div>
              <span style={{ flex: 1, fontSize: 13, fontWeight: 500, color: '#141412' }}>{u.nombre}</span>
              {u.marca && <Pill label={res.label} bg={res.bg} color={res.color} />}
              {u.marca?.adjuntos?.length > 0 && <span style={{ fontSize: 11, color: '#A5A5A0' }}>📎 {u.marca.adjuntos.length}</span>}
            </div>

            {/* Editor inline */}
            {editando && (
              <div className="fade" style={{ padding: '12px', borderTop: '1px solid #ECEAE4', background: '#F9F8F5' }}>
                <div style={{ display: 'flex', gap: 6, marginBottom: 10 }}>
                  {Object.entries(RESULTADOS).map(([k, v]) => (
                    <button key={k} onClick={() => setResultado(k)} style={{ padding: '6px 12px', borderRadius: 20, border: `1.5px solid ${resultado === k ? v.color : '#E0DFD9'}`, background: resultado === k ? v.bg : 'transparent', color: resultado === k ? v.color : '#6B6B66', fontSize: 12, cursor: 'pointer', fontWeight: resultado === k ? 600 : 400 }}>{v.label}</button>
                  ))}
                </div>
                <textarea placeholder="Comentario, nº de acta, observaciones..." value={comentario} onChange={e => setComentario(e.target.value)} style={{ marginBottom: 8, minHeight: 48 }} />
                {adjuntos.length > 0 && (
                  <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginBottom: 8 }}>
                    {adjuntos.map(a => (
                      <div key={a.id} style={{ position: 'relative' }}>
                        {a.tipo === 'imagen'
                          ? <img src={a.data} alt="" onClick={() => onPreview(a)} style={{ width: 60, height: 48, objectFit: 'cover', borderRadius: 6, border: '1px solid #E0DFD9', cursor: 'zoom-in' }} />
                          : <div style={{ fontSize: 11, padding: '4px 8px', background: '#fff', border: '1px solid #E0DFD9', borderRadius: 6 }}>📄 {a.nombre}</div>}
                        <button onClick={() => setAdjuntos(p => p.filter(x => x.id !== a.id))} style={{ position: 'absolute', top: -4, right: -4, width: 14, height: 14, borderRadius: '50%', background: '#8A1F1F', color: '#fff', border: 'none', cursor: 'pointer', fontSize: 9 }}>×</button>
                      </div>
                    ))}
                  </div>
                )}
                <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                  <button onClick={() => pickFiles('image/*,.pdf,.doc,.docx', f => setAdjuntos(p => [...p, f]))} style={{ padding: '7px 12px', borderRadius: 8, border: '1px solid #E0DFD9', background: '#fff', cursor: 'pointer', fontSize: 13, whiteSpace: 'nowrap' }}>📎 Adjuntar</button>
                  {u.marca && <Btn danger onClick={() => quitar(u.id)}>Desmarcar</Btn>}
                  <div style={{ flex: 1 }} />
                  <Btn onClick={() => setEditId(null)}>Cancelar</Btn>
                  <Btn primary onClick={guardar}>Guardar</Btn>
                </div>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

// ── Modo B: criterio único (contador + actas, p.ej. probetas) ─────────────────
function RegistrosLibres({ ensayo, onUpdate, onPreview }) {
  const [show, setShow]               = useState(false);
  const [resultado, setResultado]     = useState('apto');
  const [comentario, setComentario]   = useState('');
  const [adjuntos, setAdjuntos]       = useState([]);
  const [confirmacion, setConfirmacion] = useState(null);
  const registros = ensayo.registros || [];
  const ejecutadas = ensayo.ejecutadas || 0;
  const cantidad = ensayo.cantidad || 0;

  function setEjec(v) {
    const n = Math.max(0, Math.min(cantidad || 99999, v));
    onUpdate(e => ({ ...e, ejecutadas: n }));
  }

  function guardar() {
    const reg = { id: uid(), fecha: now(), resultado, comentario: comentario.trim(), adjuntos };
    onUpdate(e => ({ ...e, registros: [...(e.registros || []), reg] }));
    setResultado('apto'); setComentario(''); setAdjuntos([]); setShow(false);
  }
  function borrar(rid) {
    onUpdate(e => ({ ...e, registros: (e.registros || []).filter(r => r.id !== rid) }));
  }

  const pct = cantidad > 0 ? Math.round(ejecutadas / cantidad * 100) : 0;

  return (
    <div>
      {/* Contador de ejecutadas */}
      <div style={{ background: '#F9F8F5', borderRadius: 10, padding: '14px', marginBottom: 14 }}>
        <div style={{ fontSize: 11, color: '#A5A5A0', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 10 }}>Ensayos ejecutados</div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <button onClick={() => setEjec(ejecutadas - 1)} style={{ width: 36, height: 36, borderRadius: 9, border: '1.5px solid #E0DFD9', background: '#fff', cursor: 'pointer', fontSize: 18, color: '#52524E', flexShrink: 0 }}>−</button>
          <div style={{ display: 'flex', alignItems: 'baseline', gap: 5 }}>
            <input type="number" value={ejecutadas} onChange={e => setEjec(parseInt(e.target.value) || 0)} style={{ width: 70, textAlign: 'center', fontSize: 22, fontWeight: 700, padding: '4px', border: '1px solid #E0DFD9' }} />
            <span style={{ fontSize: 15, color: '#9B9B97' }}>/ {cantidad}</span>
          </div>
          <button onClick={() => setEjec(ejecutadas + 1)} style={{ width: 36, height: 36, borderRadius: 9, border: '1.5px solid #E0DFD9', background: '#fff', cursor: 'pointer', fontSize: 18, color: '#52524E', flexShrink: 0 }}>+</button>
          <div style={{ flex: 1, marginLeft: 6 }}>
            <div style={{ height: 6, background: '#ECEAE4', borderRadius: 3 }}>
              <div style={{ width: pct + '%', height: 6, borderRadius: 3, background: pct >= 100 ? '#52A124' : '#D48A0C', transition: 'width .3s' }} />
            </div>
            <div style={{ fontSize: 11, color: '#A5A5A0', marginTop: 4, textAlign: 'right' }}>{pct}%</div>
          </div>
        </div>
      </div>

      {/* Actas adjuntas */}
      <div style={{ fontSize: 11, textTransform: 'uppercase', letterSpacing: '0.08em', color: '#A5A5A0', marginBottom: 10 }}>
        Actas y registros {registros.length > 0 && `(${registros.length})`}
      </div>

      {registros.length > 0 && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 12 }}>
          {registros.slice().reverse().map(r => {
            const res = RESULTADOS[r.resultado] || RESULTADOS.info;
            return (
              <div key={r.id} style={{ background: '#FAFAF8', border: '1px solid #E8E7E1', borderRadius: 9, padding: '10px 12px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: r.comentario || r.adjuntos?.length ? 7 : 0 }}>
                  <span style={{ fontSize: 11, color: '#A5A5A0' }}>{fmtDate(r.fecha)}</span>
                  <Pill label={res.label} bg={res.bg} color={res.color} />
                  <button onClick={() => setConfirmacion({ titulo: 'Eliminar registro', texto: 'Vas a eliminar este registro de ensayo. Esta acción no se puede deshacer.', onSi: () => { borrar(r.id); setConfirmacion(null); } })} style={{ marginLeft: 'auto', background: 'none', border: 'none', cursor: 'pointer', color: '#D4D3CE', fontSize: 15, lineHeight: 1 }}>×</button>
                </div>
                {r.comentario && <div style={{ fontSize: 13, lineHeight: 1.5, marginBottom: r.adjuntos?.length ? 8 : 0 }}>{r.comentario}</div>}
                {r.adjuntos?.length > 0 && (
                  <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                    {r.adjuntos.map(a => a.tipo === 'imagen'
                      ? <img key={a.id} src={a.data} alt={a.nombre} onClick={() => onPreview(a)} style={{ width: 120, height: 92, objectFit: 'cover', borderRadius: 8, border: '1px solid #E0DFD9', cursor: 'zoom-in' }} />
                      : <a key={a.id} href={a.data} download={a.nombre} style={{ display: 'inline-flex', alignItems: 'center', gap: 4, fontSize: 12, padding: '4px 9px', borderRadius: 7, background: '#F5F4F0', border: '1px solid #E0DFD9', color: '#18180F', textDecoration: 'none' }}>📄 {a.nombre}</a>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {show ? (
        <div style={{ background: '#F9F8F5', borderRadius: 10, padding: '12px 14px' }}>
          <div style={{ fontSize: 11, color: '#A5A5A0', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8 }}>Nueva acta / registro</div>
          <div style={{ display: 'flex', gap: 6, marginBottom: 10 }}>
            {Object.entries(RESULTADOS).map(([k, v]) => (
              <button key={k} onClick={() => setResultado(k)} style={{ padding: '6px 12px', borderRadius: 20, border: `1.5px solid ${resultado === k ? v.color : '#E0DFD9'}`, background: resultado === k ? v.bg : 'transparent', color: resultado === k ? v.color : '#6B6B66', fontSize: 12, cursor: 'pointer', fontWeight: resultado === k ? 600 : 400 }}>{v.label}</button>
            ))}
          </div>
          <textarea placeholder="Nº de acta, laboratorio, observaciones..." value={comentario} onChange={e => setComentario(e.target.value)} style={{ marginBottom: 8, minHeight: 56 }} />
          {adjuntos.length > 0 && (
            <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginBottom: 8 }}>
              {adjuntos.map(a => (
                <div key={a.id} style={{ position: 'relative' }}>
                  {a.tipo === 'imagen'
                    ? <img src={a.data} alt="" onClick={() => onPreview(a)} style={{ width: 60, height: 48, objectFit: 'cover', borderRadius: 6, border: '1px solid #E0DFD9', cursor: 'zoom-in' }} />
                    : <div style={{ fontSize: 11, padding: '4px 8px', background: '#fff', border: '1px solid #E0DFD9', borderRadius: 6 }}>📄 {a.nombre}</div>}
                  <button onClick={() => setAdjuntos(p => p.filter(x => x.id !== a.id))} style={{ position: 'absolute', top: -4, right: -4, width: 14, height: 14, borderRadius: '50%', background: '#8A1F1F', color: '#fff', border: 'none', cursor: 'pointer', fontSize: 9 }}>×</button>
                </div>
              ))}
            </div>
          )}
          <div style={{ display: 'flex', gap: 8 }}>
            <button onClick={() => pickFiles('image/*,.pdf,.doc,.docx', f => setAdjuntos(p => [...p, f]))} style={{ padding: '7px 12px', borderRadius: 8, border: '1px solid #E0DFD9', background: '#fff', cursor: 'pointer', fontSize: 13, whiteSpace: 'nowrap' }}>📎 Adjuntar acta</button>
            <Btn onClick={() => setShow(false)} full>Cancelar</Btn>
            <Btn primary full onClick={guardar}>Guardar</Btn>
          </div>
        </div>
      ) : (
        <Btn onClick={() => setShow(true)} full>+ Adjuntar acta / registro</Btn>
      )}
      {confirmacion && <ConfirmMini titulo={confirmacion.titulo} texto={confirmacion.texto} onSi={confirmacion.onSi} onNo={() => setConfirmacion(null)} />}
    </div>
  );
}

// ─── MÓDULO: Control de calidad ───────────────────────────────────────────────// ─── MÓDULO: Control de calidad ───────────────────────────────────────────────// ─── MÓDULO: Control de calidad ───────────────────────────────────────────────

// Tabla de lotificación — Código Estructural (oct 2025), control estadístico Modalidad 1
const TIPOS_ELEMENTO = {
  cim_grande: {
    label: 'Cimentaciones > 200 m³',
    desc: 'Volumen vertido de forma continua',
    volumen: null, superficie: null, tiempo: 1, formulaN: true,
  },
  cim_superficial: {
    label: 'Cimentaciones superficiales < 200 m³',
    desc: 'Zapatas, losas de cimentación',
    volumen: 100, superficie: null, tiempo: 1, formulaN: false,
  },
  flexion: {
    label: 'Vigas, forjados, soleras (flexión)',
    desc: 'Elementos trabajando a flexión',
    volumen: 100, superficie: 1000, tiempo: 2, formulaN: false,
  },
  pilares: {
    label: 'Pilares y muros portantes',
    desc: 'Elementos comprimidos de edificación',
    volumen: 100, superficie: 500, tiempo: 2, formulaN: false,
  },
};

// Calcula la lotificación completa de un elemento según el CE
function calcularLotificacion(tipo, volumen, superficie, conDOR) {
  const t = TIPOS_ELEMENTO[tipo];
  const V = parseFloat(volumen) || 0;
  const S = parseFloat(superficie) || 0;

  // Cimentaciones grandes: 1 lote, N por fórmula sobre volumen
  if (t.formulaN) {
    const N = conDOR ? Math.max(Math.ceil(V / 105), 1) : Math.max(Math.ceil(V / 35), 3);
    return {
      numLotes: 1,
      seriesPorLote: N,
      totalSeries: N,
      motivo: `Volumen continuo · N ${conDOR ? '≥ V/105' : '≥ V/35'}`,
    };
  }

  // Resto: nº de lotes = límite más restrictivo entre volumen y superficie
  const lotesVol = t.volumen ? Math.ceil(V / t.volumen) : 1;
  const lotesSup = t.superficie ? Math.ceil(S / t.superficie) : 1;
  const numLotes = Math.max(lotesVol, lotesSup, 1);
  const N = conDOR ? 1 : 3;

  let motivo;
  if (t.superficie && lotesSup >= lotesVol) motivo = `Manda superficie (${lotesSup} lote${lotesSup > 1 ? 's' : ''} por ${t.superficie} m²)`;
  else motivo = `Manda volumen (${lotesVol} lote${lotesVol > 1 ? 's' : ''} por ${t.volumen} m³)`;

  return { numLotes, seriesPorLote: N, totalSeries: numLotes * N, motivo };
}

function ModuloCalidad({ obra, onSave }) {
  const [sub, setSub] = useState('hormigon'); // hormigon | pcq | ensayos

  const subTabs = [
    { id: 'hormigon', label: 'Control de hormigón' },
    { id: 'pcq',      label: 'PCQ' },
    { id: 'ensayos',  label: 'Ensayos' },
  ];

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
      {/* Sub-navegación de calidad */}
      <div style={{ display: 'flex', gap: 4, background: '#fff', border: '1px solid #E8E7E1', borderRadius: 10, padding: 4, width: 'fit-content' }}>
        {subTabs.map(t => (
          <button key={t.id} onClick={() => setSub(t.id)} style={{ padding: '6px 14px', borderRadius: 7, border: 'none', background: sub === t.id ? '#1C1C1A' : 'transparent', color: sub === t.id ? '#F2F1ED' : '#6B6B66', fontSize: 13, cursor: 'pointer', fontWeight: sub === t.id ? 500 : 400, transition: 'all .15s' }}>
            {t.label}
          </button>
        ))}
      </div>

      {sub === 'hormigon' && <ControlHormigon obra={obra} onSave={onSave} />}
      {sub === 'pcq'      && <ModuloProximamente icono="📋" titulo="Plan de Control de Calidad" descripcion="Generación automática del PCQ a partir de los procesos constructivos de la obra. Se integrará al desplegar la app con la biblioteca de documentos oficiales." />}
      {sub === 'ensayos'  && <ModuloEnsayos obra={obra} onSave={onSave} />}
    </div>
  );
}

// ── Control estadístico de hormigón ───────────────────────────────────────────
function ControlHormigon({ obra, onSave }) {
  const elementosRaw = obra.lotes || [];
  // Filtra elementos con la estructura nueva (tienen nombre y lotes con series)
  const elementos = elementosRaw.filter(e => e && e.nombre && Array.isArray(e.lotes));
  const hayAntiguos = elementosRaw.length > elementos.length;
  const [showNuevo, setShowNuevo] = useState(false);
  const [expandido, setExpandido] = useState(null);
  const [form, setForm] = useState({ nombre: '', tipo: 'flexion', volumen: '', superficie: '', conDOR: false });
  const upd = (k, v) => setForm(f => ({ ...f, [k]: v }));

  const tipoForm = TIPOS_ELEMENTO[form.tipo];
  const preview  = calcularLotificacion(form.tipo, form.volumen, form.superficie, form.conDOR);

  function crearElemento() {
    if (!form.nombre.trim() || !form.volumen) return;
    const calc = calcularLotificacion(form.tipo, form.volumen, form.superficie, form.conDOR);
    const lotes = Array.from({ length: calc.numLotes }, (_, i) => ({
      id: uid(),
      num: i + 1,
      series: Array.from({ length: calc.seriesPorLote }, (_, j) => ({ id: uid(), num: j + 1, acta: null })),
    }));
    const elemento = {
      id: uid(),
      nombre: form.nombre.trim(),
      tipo: form.tipo,
      volumen: form.volumen,
      superficie: form.superficie,
      conDOR: form.conDOR,
      numLotes: calc.numLotes,
      seriesPorLote: calc.seriesPorLote,
      lotes,
      creadoEn: now(),
    };
    onSave({ ...obra, lotes: [elemento, ...elementos] });
    setForm({ nombre: '', tipo: 'flexion', volumen: '', superficie: '', conDOR: false });
    setShowNuevo(false);
    setExpandido(elemento.id);
  }

  function eliminar(id) {
    onSave({ ...obra, lotes: elementos.filter(e => e.id !== id) });
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>

      {/* Aviso de datos antiguos */}
      {hayAntiguos && (
        <div style={{ background: '#FEF3DB', border: '1px solid #F5D98B', borderRadius: 10, padding: '10px 14px', display: 'flex', alignItems: 'center', gap: 10 }}>
          <span style={{ fontSize: 15 }}>⚠️</span>
          <span style={{ fontSize: 12, color: '#7C4A00', flex: 1 }}>Hay lotes creados con una versión anterior que no son compatibles.</span>
          <button onClick={() => onSave({ ...obra, lotes: elementos })} style={{ fontSize: 12, padding: '5px 11px', borderRadius: 8, border: '1px solid #D48A0C', background: '#fff', color: '#7C4A00', cursor: 'pointer', fontWeight: 500, whiteSpace: 'nowrap' }}>
            Limpiar antiguos
          </button>
        </div>
      )}

      {/* Cabecera */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
        <div style={{ flex: 1, fontSize: 13, color: '#6B6B66' }}>
          {elementos.length} elemento{elementos.length !== 1 ? 's' : ''} · Control estadístico Modalidad 1 (CE)
        </div>
        <Btn primary onClick={() => setShowNuevo(v => !v)}>{showNuevo ? '✕ Cancelar' : '+ Nuevo elemento'}</Btn>
      </div>

      {/* Formulario nuevo elemento */}
      {showNuevo && (
        <div style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 12, padding: '16px' }}>
          <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 14 }}>Nuevo elemento</div>

          <div style={{ marginBottom: 12 }}>
            <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 5 }}>Nombre del elemento *</label>
            <input autoFocus placeholder="p.ej. Forjados torre A" value={form.nombre} onChange={e => upd('nombre', e.target.value)} />
          </div>

          {/* Tipo */}
          <div style={{ marginBottom: 12 }}>
            <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 6 }}>Tipo de elemento</label>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
              {Object.entries(TIPOS_ELEMENTO).map(([k, v]) => (
                <div key={k} onClick={() => upd('tipo', k)} style={{ padding: '10px 12px', borderRadius: 10, border: `1.5px solid ${form.tipo === k ? '#18180F' : '#E0DFD9'}`, background: form.tipo === k ? '#F5F4F0' : '#fff', cursor: 'pointer', transition: 'all .15s' }}>
                  <div style={{ fontSize: 13, fontWeight: 500 }}>{v.label}</div>
                  <div style={{ fontSize: 11, color: '#A5A5A0', marginTop: 2 }}>{v.desc}</div>
                </div>
              ))}
            </div>
          </div>

          {/* m³ y m² */}
          <div style={{ display: 'grid', gridTemplateColumns: tipoForm.superficie ? '1fr 1fr' : '1fr', gap: 10, marginBottom: 12 }}>
            <div>
              <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 5 }}>Volumen total (m³) *</label>
              <input type="number" placeholder="m³" value={form.volumen} onChange={e => upd('volumen', e.target.value)} />
            </div>
            {tipoForm.superficie && (
              <div>
                <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 5 }}>Superficie total (m²)</label>
                <input type="number" placeholder="m²" value={form.superficie} onChange={e => upd('superficie', e.target.value)} />
              </div>
            )}
          </div>

          {/* DOR */}
          <div style={{ marginBottom: 14 }}>
            <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 6 }}>¿Hormigón con DOR?</label>
            <div style={{ display: 'flex', gap: 6 }}>
              {[[false, 'Sin DOR'], [true, 'Con DOR']].map(([val, label]) => (
                <button key={String(val)} onClick={() => upd('conDOR', val)} style={{ padding: '6px 14px', borderRadius: 20, border: `1.5px solid ${form.conDOR === val ? '#18180F' : '#E0DFD9'}`, background: form.conDOR === val ? '#18180F' : 'transparent', color: form.conDOR === val ? '#fff' : '#6B6B66', fontSize: 12, cursor: 'pointer', fontWeight: form.conDOR === val ? 600 : 400 }}>
                  {label}
                </button>
              ))}
            </div>
          </div>

          {/* Preview del cálculo */}
          {form.volumen && (
            <div style={{ background: '#F0F6F1', border: '1px solid #C5E3CE', borderRadius: 10, padding: '12px 14px', marginBottom: 14 }}>
              <div style={{ fontSize: 11, color: '#1C1C1A', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8, fontWeight: 600 }}>Lotificación resultante</div>
              <div style={{ display: 'flex', alignItems: 'baseline', gap: 6, marginBottom: 4 }}>
                <span style={{ fontSize: 26, fontWeight: 600, color: '#1C1C1A' }}>{preview.numLotes}</span>
                <span style={{ fontSize: 13, color: '#52524E' }}>lote{preview.numLotes > 1 ? 's' : ''}</span>
                <span style={{ fontSize: 14, color: '#A5A5A0', margin: '0 4px' }}>×</span>
                <span style={{ fontSize: 26, fontWeight: 600, color: '#1C1C1A' }}>{preview.seriesPorLote}</span>
                <span style={{ fontSize: 13, color: '#52524E' }}>series/lote</span>
                <span style={{ fontSize: 14, color: '#A5A5A0', margin: '0 4px' }}>=</span>
                <span style={{ fontSize: 26, fontWeight: 600, color: '#1C1C1A' }}>{preview.totalSeries}</span>
                <span style={{ fontSize: 13, color: '#52524E' }}>series totales</span>
              </div>
              <div style={{ fontSize: 12, color: '#6B6B66' }}>{preview.motivo}</div>
            </div>
          )}

          <div style={{ display: 'flex', gap: 8 }}>
            <Btn onClick={() => setShowNuevo(false)} full>Cancelar</Btn>
            <Btn primary onClick={crearElemento} disabled={!form.nombre.trim() || !form.volumen} full>Crear lotificación</Btn>
          </div>
        </div>
      )}

      {/* Lista de elementos */}
      {elementos.length === 0 && !showNuevo ? (
        <div style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 12, padding: '40px 20px', textAlign: 'center', color: '#A5A5A0' }}>
          <div style={{ fontSize: 32, marginBottom: 10 }}>🧱</div>
          <div style={{ fontSize: 13, marginBottom: 14 }}>Sin elementos de hormigón todavía</div>
          <Btn onClick={() => setShowNuevo(true)}>+ Crear primera lotificación</Btn>
        </div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {elementos.map(el => {
            const t = TIPOS_ELEMENTO[el.tipo] || TIPOS_ELEMENTO.flexion;
            const lotesEl = el.lotes || [];
            const numLotes = el.numLotes || lotesEl.length;
            const totalSeries = (numLotes || 0) * (el.seriesPorLote || 0);
            const seriesRellenas = lotesEl.reduce((s, l) => s + (l.series || []).filter(se => se.acta).length, 0);
            const abierto = expandido === el.id;
            return (
              <div key={el.id} style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 11, overflow: 'hidden' }}>
                {/* Cabecera elemento */}
                <div onClick={() => setExpandido(abierto ? null : el.id)} style={{ padding: '13px 15px', cursor: 'pointer' }}>
                  <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10 }}>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 14, fontWeight: 600, color: '#141412' }}>{el.nombre}</div>
                      <div style={{ fontSize: 12, color: '#9B9B97', marginTop: 2 }}>
                        {t.label} · {el.volumen} m³{el.superficie ? ` · ${el.superficie} m²` : ''} · {el.conDOR ? 'Con DOR' : 'Sin DOR'}
                      </div>
                    </div>
                    <span style={{ fontSize: 11, padding: '3px 9px', borderRadius: 20, background: '#EEEDE7', color: '#1C1C1A', fontWeight: 600, whiteSpace: 'nowrap' }}>
                      {numLotes} lote{numLotes > 1 ? 's' : ''} · {totalSeries} series
                    </span>
                    <button onClick={e => { e.stopPropagation(); eliminar(el.id); }} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#D4D3CE', fontSize: 16, padding: '0 2px', lineHeight: 1 }}>×</button>
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginTop: 10 }}>
                    <div style={{ flex: 1, height: 4, background: '#ECEAE4', borderRadius: 2 }}>
                      <div style={{ width: (totalSeries ? seriesRellenas / totalSeries * 100 : 0) + '%', height: 4, borderRadius: 2, background: '#52A124' }} />
                    </div>
                    <span style={{ fontSize: 11, color: '#A5A5A0' }}>{seriesRellenas}/{totalSeries} series con acta</span>
                    <span style={{ fontSize: 11, color: '#A5A5A0' }}>{abierto ? '▲' : '▼'}</span>
                  </div>
                </div>

                {/* Detalle lotes y series */}
                {abierto && (
                  <div style={{ borderTop: '1px solid #F2F1ED', padding: '12px 15px', background: '#FAFAF8', display: 'flex', flexDirection: 'column', gap: 8 }}>
                    {lotesEl.map(lote => (
                      <div key={lote.id} style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 9, padding: '10px 12px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
                          <span style={{ fontSize: 13, fontWeight: 500 }}>Lote {lote.num}</span>
                          <span style={{ fontSize: 11, color: '#A5A5A0' }}>{(lote.series || []).length} series</span>
                        </div>
                        <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                          {(lote.series || []).map(serie => (
                            <div key={serie.id} title={serie.acta ? 'Con acta' : 'Pendiente de acta'}
                              style={{ width: 34, height: 34, borderRadius: 8, border: `1.5px solid ${serie.acta ? '#52A124' : '#E0DFD9'}`, background: serie.acta ? '#E8F5E0' : '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, fontWeight: 500, color: serie.acta ? '#2D5E10' : '#A5A5A0' }}>
                              {serie.acta ? '✓' : serie.num}
                            </div>
                          ))}
                        </div>
                      </div>
                    ))}
                    <div style={{ padding: '8px 11px', background: '#F0EFEA', borderRadius: 8, fontSize: 12, color: '#9B9B97', textAlign: 'center' }}>
                      El volcado automático de actas de probeta sobre las series se añade en la siguiente fase
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ─── MÓDULO: Placeholder ──────────────────────────────────────────────────────

function ModuloProximamente({ icono, titulo, descripcion }) {
  return (
    <div style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 12, padding: '48px 20px', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 12, textAlign: 'center' }}>
      <div style={{ fontSize: 40 }}>{icono}</div>
      <div style={{ fontSize: 16, fontWeight: 600 }}>{titulo}</div>
      <div style={{ fontSize: 13, color: '#6B6B66', maxWidth: 380, lineHeight: 1.6 }}>{descripcion}</div>
      <div style={{ marginTop: 4, fontSize: 12, color: '#A5A5A0', background: '#F5F4F0', padding: '4px 12px', borderRadius: 20 }}>Próximamente</div>
    </div>
  );
}

// ─── Vista: Hoy ───────────────────────────────────────────────────────────────

// Calcula las alertas activas de todas las obras
function calcularAlertas(obras) {
  const incSinRevisar = obras.flatMap(o =>
    o.incidencias
      .filter(i => i.estado !== 'resuelta' && diasSinRevisar(i) >= 10)
      .map(i => ({ obra: o, inc: i, dias: diasSinRevisar(i) }))
  ).sort((a, b) => b.dias - a.dias);

  const tareasVencidas = obras.flatMap(o =>
    (o.apuntes || [])
      .filter(a => a.tipo === 'tarea' && !a.hecha && a.fechaLimite && diasHasta(a.fechaLimite) < 0)
      .map(a => ({ obra: o, tarea: a, dias: -diasHasta(a.fechaLimite) }))
  ).sort((a, b) => b.dias - a.dias);

  const tareasProximas = obras.flatMap(o =>
    (o.apuntes || [])
      .filter(a => a.tipo === 'tarea' && !a.hecha && a.fechaLimite && diasHasta(a.fechaLimite) >= 0 && diasHasta(a.fechaLimite) <= 2)
      .map(a => ({ obra: o, tarea: a, dias: diasHasta(a.fechaLimite) }))
  ).sort((a, b) => a.dias - b.dias);

  return { incSinRevisar, tareasVencidas, tareasProximas, total: incSinRevisar.length + tareasVencidas.length + tareasProximas.length };
}

function VistaAlertas({ obras, onIrObra, isMobile }) {
  const { incSinRevisar, tareasVencidas, tareasProximas, total } = calcularAlertas(obras);
  const pad = isMobile ? '14px' : '18px 22px';

  const SeccionHeader = ({ titulo, count, color }) => (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 11 }}>
      <span style={{ fontSize: 14, color: color, fontWeight: 400 }}>/</span>
      <span style={{ fontSize: 12.5, fontWeight: 600, color: '#141412', letterSpacing: '0.02em' }}>{titulo}</span>
      <span style={{ fontSize: 11, padding: '1px 8px', borderRadius: 4, background: color + '1A', color, fontWeight: 600 }}>{count}</span>
    </div>
  );

  const Fila = ({ obra, titulo, etiqueta, etColor, etBg, accent }) => (
    <div onClick={() => onIrObra(obra)} className="obra-card"
      style={{ background: '#fff', borderRadius: 10, boxShadow: '0 1px 3px rgba(0,0,0,0.06)', padding: '11px 14px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 12, borderLeft: `3px solid ${accent}` }}>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontSize: 13, fontWeight: 500, color: '#141412', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{titulo}</div>
        <div style={{ fontSize: 12, color: '#9B9B97', marginTop: 2, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{obra.nombre}</div>
      </div>
      <span style={{ fontSize: 11, color: etColor, fontWeight: 500, whiteSpace: 'nowrap', background: etBg, padding: '2px 8px', borderRadius: 20, flexShrink: 0 }}>{etiqueta}</span>
    </div>
  );

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
      {/* Topbar */}
      <div style={{ background: '#fff', borderBottom: '1px solid #ECEAE4', padding: isMobile ? '13px 16px' : '13px 22px', paddingTop: isMobile ? 'calc(13px + env(safe-area-inset-top))' : '13px', flexShrink: 0 }}>
        <div style={{ fontSize: 16, fontWeight: 600, color: '#141412' }}>Alertas</div>
        <div style={{ fontSize: 12, color: '#9B9B97', marginTop: 1 }}>{total > 0 ? `${total} aviso${total !== 1 ? 's' : ''} requieren tu atención` : 'Sin avisos pendientes'}</div>
      </div>

      <div style={{ flex: 1, overflow: 'auto', padding: pad, display: 'flex', flexDirection: 'column', gap: 20 }}>

        {total === 0 && (
          <div style={{ textAlign: 'center', padding: '80px 20px', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 14 }}>
            <div style={{ width: 56, height: 56, borderRadius: 12, background: '#1C1C1A', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 26, color: '#fff', fontWeight: 300 }}>✓</div>
            <div style={{ fontSize: 16, fontWeight: 600 }}>Todo al día</div>
            <div style={{ fontSize: 13, color: '#9B9B97' }}>No hay incidencias ni tareas que requieran atención</div>
          </div>
        )}

        {/* Tareas vencidas */}
        {tareasVencidas.length > 0 && (
          <div>
            <SeccionHeader titulo="Tareas vencidas" count={tareasVencidas.length} color="#8A1F1F" />
            <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              {tareasVencidas.map(({ obra, tarea, dias }) => (
                <Fila key={tarea.id} obra={obra} titulo={tarea.texto}
                  etiqueta={dias === 0 ? 'Vence hoy' : `Hace ${dias}d`} etColor="#8A1F1F" etBg="#FDECEC" accent="#E24B4A" />
              ))}
            </div>
          </div>
        )}

        {/* Tareas próximas a vencer */}
        {tareasProximas.length > 0 && (
          <div>
            <SeccionHeader titulo="Tareas próximas a vencer" count={tareasProximas.length} color="#7C4A00" />
            <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              {tareasProximas.map(({ obra, tarea, dias }) => (
                <Fila key={tarea.id} obra={obra} titulo={tarea.texto}
                  etiqueta={dias === 0 ? 'Vence hoy' : dias === 1 ? 'Mañana' : `En ${dias} días`} etColor="#7C4A00" etBg="#FEF3DB" accent="#D48A0C" />
              ))}
            </div>
          </div>
        )}

        {/* Incidencias sin revisar +10 días */}
        {incSinRevisar.length > 0 && (
          <div>
            <SeccionHeader titulo="Incidencias sin revisar +10 días" count={incSinRevisar.length} color="#7C4A00" />
            <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              {incSinRevisar.map(({ obra, inc, dias }) => (
                <Fila key={inc.id} obra={obra} titulo={inc.titulo}
                  etiqueta={`${dias} días`} etColor="#7C4A00" etBg="#FEF3DB" accent="#D48A0C" />
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Detalle de Obra ──────────────────────────────────────────────────────────

// ─── Módulo Acta de Visita de Obra ────────────────────────────────────────────

function getDefaultEquipo() {
  return [
    { id: uid(), nombre: 'PROJECT MANAGER',                    personas: [{ id: uid(), empresa: '', nombre: '', email: '', tel: '', asistio: false }] },
    { id: uid(), nombre: 'DIRECCIÓN DE OBRA (DO)',             personas: [{ id: uid(), empresa: '', nombre: '', email: '', tel: '', asistio: false }] },
    { id: uid(), nombre: 'DIRECCIÓN DE EJECUCIÓN (DEO)',       personas: [{ id: uid(), empresa: 'PLAAT', nombre: 'Xavier Pla', email: 'xpla@plaat.es', tel: '629 72 72 62', asistio: false }] },
    { id: uid(), nombre: 'ING. ESTRUCTURAS',                   personas: [{ id: uid(), empresa: '', nombre: '', email: '', tel: '', asistio: false }] },
    { id: uid(), nombre: 'ING. INSTALACIONES',                 personas: [{ id: uid(), empresa: '', nombre: '', email: '', tel: '', asistio: false }] },
    { id: uid(), nombre: 'COORDINADOR DE SEGURIDAD (CSS)',     personas: [{ id: uid(), empresa: '', nombre: '', email: '', tel: '', asistio: false }] },
    { id: uid(), nombre: 'CONTRATISTA (EC)',                   personas: [{ id: uid(), empresa: '', nombre: '', email: '', tel: '', asistio: false }] },
  ];
}

function getDefaultSecciones() {
  return [
    { id: uid(), codigo: '1',   titulo: 'TEMAS TRATADOS',                  temas: [] },
    { id: uid(), codigo: '2',   titulo: 'CONTROL DE CALIDAD',               temas: [] },
    { id: uid(), codigo: '3.1', titulo: 'INSTALACIONES',                    temas: [] },
    { id: uid(), codigo: '3.2', titulo: 'ACOMETIDAS',                       temas: [] },
    { id: uid(), codigo: '4',   titulo: 'SEGURIDAD Y SALUD',                temas: [] },
    { id: uid(), codigo: '5',   titulo: 'SEGUIMIENTO PLANIFICACIÓN',        temas: [] },
    { id: uid(), codigo: '6',   titulo: 'SEGUIMIENTO HITOS CONTRACTUALES',  temas: [] },
    { id: uid(), codigo: '7',   titulo: 'SEGUIMIENTO CONTRATACIÓN',         temas: [] },
    { id: uid(), codigo: '8',   titulo: 'SEGUIMIENTO PERSONAL',             temas: [] },
    { id: uid(), codigo: '9',   titulo: 'SEGUIMIENTO LEED / WELL / WIREDSCORE', temas: [] },
  ];
}

function migrateVO(raw) {
  let vo = raw ? { ...raw } : {};
  if (!vo.num) vo.num = 1;
  if (!vo.estadoObra) vo.estadoObra = { descripcion: '', fotos: [] };
  if (!Array.isArray(vo.equipo)) {
    const eq = getDefaultEquipo();
    // try to recover old object-style equipo
    if (vo.equipo && typeof vo.equipo === 'object') {
      const old = vo.equipo;
      eq.forEach(r => {
        const k = r.nombre.toLowerCase().includes('deo') ? 'deo' : r.nombre.toLowerCase().includes('do') ? 'do' : null;
        if (k && old[k]) { const d = old[k]; r.personas[0] = { ...r.personas[0], empresa: d.empresa||'', nombre: d.personas||'', email: d.contacto||'' }; }
      });
    }
    vo.equipo = eq;
  }
  if (!Array.isArray(vo.secciones)) {
    const secs = getDefaultSecciones();
    if (Array.isArray(vo.temas)) {
      vo.temas.forEach(t => {
        const sec = secs.find(s => s.codigo === t.seccion);
        if (sec) sec.temas.push({ ...t, entradas: (t.entradas||[]).map(e => ({ ...e, fin: e.fin||'' })) });
      });
    }
    vo.secciones = secs;
  }
  // ensure every entry has fin field
  vo.secciones = vo.secciones.map(s => ({ ...s, temas: (s.temas||[]).map(t => ({ ...t, entradas: (t.entradas||[]).map(e => ({ fin: '', ...e })) })) }));
  return vo;
}

// ── ModuloActaVO ─────────────────────────────────────────────────────────────
function ModuloActaVO({ obra, onSave }) {
  const isMobile = useIsMobile();
  const vo = migrateVO(obra.actaVO);
  const [showEquipo,    setShowEquipo]    = useState(false);
  const [showHistorico, setShowHistorico] = useState(false);
  const [borrar,        setBorrar]        = useState(null);
  const [generando,     setGenerando]     = useState(false);
  const [editandoSec,   setEditandoSec]   = useState(null); // id sección editando nombre

  function guardarVO(nuevo) { onSave({ ...obra, actaVO: nuevo }); }

  // Equipo
  function updRol(id, campo, val) {
    guardarVO({ ...vo, equipo: vo.equipo.map(r => r.id === id ? { ...r, [campo]: val } : r) });
  }
  function addPersona(rolId) {
    guardarVO({ ...vo, equipo: vo.equipo.map(r => r.id === rolId ? { ...r, personas: [...(r.personas||[]), { id: uid(), empresa: '', nombre: '', email: '', tel: '', asistio: false }] } : r) });
  }
  function updPersona(rolId, pId, campo, val) {
    guardarVO({ ...vo, equipo: vo.equipo.map(r => r.id !== rolId ? r : { ...r, personas: r.personas.map(p => p.id === pId ? { ...p, [campo]: val } : p) }) });
  }
  function delPersona(rolId, pId) {
    guardarVO({ ...vo, equipo: vo.equipo.map(r => r.id !== rolId ? r : { ...r, personas: r.personas.filter(p => p.id !== pId) }) });
  }
  function addRol() {
    guardarVO({ ...vo, equipo: [...vo.equipo, { id: uid(), nombre: 'NUEVO ROL', personas: [{ id: uid(), empresa: '', nombre: '', email: '', tel: '', asistio: false }] }] });
  }
  function delRol(id) { guardarVO({ ...vo, equipo: vo.equipo.filter(r => r.id !== id) }); }

  // Secciones
  function addSeccion() {
    const n = vo.secciones.length + 1;
    guardarVO({ ...vo, secciones: [...vo.secciones, { id: uid(), codigo: String(n), titulo: 'NUEVA SECCIÓN', temas: [] }] });
  }
  function updSeccion(id, campo, val) {
    guardarVO({ ...vo, secciones: vo.secciones.map(s => s.id === id ? { ...s, [campo]: val } : s) });
  }
  function delSeccion(id) { guardarVO({ ...vo, secciones: vo.secciones.filter(s => s.id !== id) }); setBorrar(null); }

  // Temas
  function nextNum(sec) {
    const nums = (sec.temas||[]).map(t => parseInt((t.num||'').split('.').pop()||'0', 10));
    const n = (nums.length ? Math.max(...nums) : 0) + 1;
    return `${sec.codigo}.${String(n).padStart(2,'0')}`;
  }
  function addTema(secId, texto) {
    if (!texto.trim()) return;
    guardarVO({ ...vo, secciones: vo.secciones.map(s => s.id !== secId ? s : { ...s, temas: [...(s.temas||[]), { id: uid(), num: nextNum(s), resuelto: false, resueltoEnActa: null, entradas: [{ id: uid(), texto: texto.trim(), estado: 'P', fecha: today(), fin: '', resp: '' }] }] }) });
  }
  function addEntrada(secId, temaId, texto) {
    if (!texto.trim()) return;
    guardarVO({ ...vo, secciones: vo.secciones.map(s => s.id !== secId ? s : { ...s, temas: s.temas.map(t => t.id !== temaId ? t : { ...t, entradas: [...t.entradas, { id: uid(), texto: texto.trim(), estado: 'P', fecha: today(), fin: '', resp: '' }] }) }) });
  }
  function updEntrada(secId, temaId, entId, campo, val) {
    guardarVO({ ...vo, secciones: vo.secciones.map(s => s.id !== secId ? s : { ...s, temas: s.temas.map(t => {
      if (t.id !== temaId) return t;
      const entradas = t.entradas.map(e => e.id !== entId ? e : { ...e, [campo]: val });
      const ult = entradas[entradas.length - 1];
      const resuelto = ult.estado === 'R';
      return { ...t, entradas, resuelto, resueltoEnActa: resuelto ? (t.resueltoEnActa || vo.num) : null };
    }) }) });
  }
  function delTema(secId, temaId) {
    guardarVO({ ...vo, secciones: vo.secciones.map(s => s.id !== secId ? s : { ...s, temas: s.temas.filter(t => t.id !== temaId) }) });
    setBorrar(null);
  }

  // Estado de obra
  function updEstado(campo, val) { guardarVO({ ...vo, estadoObra: { ...(vo.estadoObra||{}), [campo]: val } }); }
  function addFotoEstado() {
    pickFiles('image/*', f => guardarVO({ ...vo, estadoObra: { ...(vo.estadoObra||{}), fotos: [...((vo.estadoObra||{}).fotos||[]), { id: uid(), data: f.data }] } }));
  }
  function delFotoEstado(id) { guardarVO({ ...vo, estadoObra: { ...(vo.estadoObra||{}), fotos: (vo.estadoObra.fotos||[]).filter(f => f.id !== id) } }); }

  async function exportar() {
    setGenerando(true);
    try { await generarActaVO(obra, vo); guardarVO({ ...vo, num: vo.num + 1 }); }
    catch (e) { alert('Error al exportar: ' + e.message); }
    setGenerando(false);
  }

  // Activos = no resueltos en acta anterior; resueltos = para histórico
  const todosResueltos = vo.secciones.flatMap(s => (s.temas||[]).filter(t => t.resuelto));
  const activosPorSec = id => (vo.secciones.find(s => s.id === id)?.temas||[]).filter(t => !(t.resuelto && t.resueltoEnActa && t.resueltoEnActa < vo.num));

  return (
    <div>
      {/* Cabecera */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', marginBottom: 14 }}>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 16, fontWeight: 600, color: '#141412' }}>Acta de Visita de Obra</div>
          <div style={{ fontSize: 12, color: '#9B9B97' }}>Próxima exportación: Nº {String(vo.num).padStart(2,'0')}</div>
        </div>
        <Btn onClick={() => setShowHistorico(true)}>Resueltos ({todosResueltos.length})</Btn>
        <Btn primary disabled={generando} onClick={exportar}>{generando ? 'Generando...' : `↓ Exportar Acta Nº ${String(vo.num).padStart(2,'0')}`}</Btn>
      </div>

      {/* Fase + lugar */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginBottom: 14 }}>
        <Field label="Fase"><input value={vo.fase||''} onChange={e => guardarVO({...vo, fase: e.target.value})} placeholder="Estructura, acabados..." /></Field>
        <Field label="Lugar"><input value={vo.lugar||''} onChange={e => guardarVO({...vo, lugar: e.target.value})} placeholder="Obra / oficina" /></Field>
      </div>

      {/* Equipo técnico */}
      <div style={{ border: '1px solid #E8E7E1', borderRadius: 10, marginBottom: 14, overflow: 'hidden' }}>
        <button onClick={() => setShowEquipo(v => !v)} style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 8, padding: '11px 14px', background: '#FAFAF8', border: 'none', cursor: 'pointer', fontSize: 13, fontWeight: 500, color: '#141412' }}>
          <span style={{ fontSize: 10, color: '#A5A5A0', display: 'inline-block', transition: 'transform .2s', transform: showEquipo ? 'rotate(90deg)' : 'none' }}>▶</span>
          Equipo técnico y datos de contacto
        </button>
        {showEquipo && (
          <div className="fade" style={{ padding: '10px 14px' }}>
            {/* Encabezado columnas */}
            {!isMobile && (
              <div style={{ display: 'grid', gridTemplateColumns: '2fr 1.4fr 1.6fr 1.8fr 1.3fr 60px 26px', gap: 5, marginBottom: 4 }}>
                {['Rol','Empresa','Nombre','Email','Teléfono','Asistido',''].map((h,i) => <div key={i} style={{ fontSize: 10, color: '#A5A5A0', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em', padding: '0 3px' }}>{h}</div>)}
              </div>
            )}
            {(vo.equipo||[]).map(rol => (
              <div key={rol.id} style={{ marginBottom: 10, border: '1px solid #F2F1ED', borderRadius: 8, overflow: 'hidden' }}>
                {/* Nombre del rol (editable) */}
                <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '6px 10px', background: '#F5F4F0' }}>
                  <input value={rol.nombre} onChange={e => updRol(rol.id, 'nombre', e.target.value)} style={{ flex: 1, fontSize: 11, fontWeight: 600, background: 'transparent', border: 'none', padding: 0, boxShadow: 'none', color: '#141412' }} />
                  <button onClick={() => delRol(rol.id)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#D4D3CE', fontSize: 15, lineHeight: 1 }}>×</button>
                </div>
                {/* Personas */}
                {(rol.personas||[]).map(p => (
                  <div key={p.id} style={{ display: isMobile ? 'flex' : 'grid', gridTemplateColumns: isMobile ? undefined : '2fr 1.4fr 1.6fr 1.8fr 1.3fr 60px 26px', flexDirection: isMobile ? 'column' : undefined, gap: 5, padding: '6px 10px', borderTop: '1px solid #F2F1ED', alignItems: 'center' }}>
                    {isMobile && <div style={{ fontSize: 10, color: '#A5A5A0', marginBottom: 3 }}>Empresa / Nombre / Email / Tel</div>}
                    <input placeholder="Empresa" value={p.empresa||''} onChange={e => updPersona(rol.id, p.id, 'empresa', e.target.value)} style={{ fontSize: 12 }} />
                    <input placeholder="Nombre" value={p.nombre||''} onChange={e => updPersona(rol.id, p.id, 'nombre', e.target.value)} style={{ fontSize: 12 }} />
                    <input placeholder="Email" value={p.email||''} onChange={e => updPersona(rol.id, p.id, 'email', e.target.value)} style={{ fontSize: 12 }} />
                    <input placeholder="Teléfono" value={p.tel||''} onChange={e => updPersona(rol.id, p.id, 'tel', e.target.value)} style={{ fontSize: 12 }} />
                    <label style={{ display: 'flex', alignItems: 'center', gap: 5, cursor: 'pointer', fontSize: 12, color: '#52524E', paddingLeft: isMobile ? 0 : 4 }}>
                      <input type="checkbox" checked={!!p.asistio} onChange={e => updPersona(rol.id, p.id, 'asistio', e.target.checked)} style={{ width: 14, height: 14 }} />
                      {isMobile ? 'Asistido' : ''}
                    </label>
                    <button onClick={() => delPersona(rol.id, p.id)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#D4D3CE', fontSize: 15, lineHeight: 1, textAlign: 'center' }}>×</button>
                  </div>
                ))}
                <button onClick={() => addPersona(rol.id)} style={{ width: '100%', padding: '5px', background: 'none', border: 'none', cursor: 'pointer', fontSize: 11, color: '#9B9B97', borderTop: '1px dashed #E8E7E1' }}>+ Añadir persona</button>
              </div>
            ))}
            <button onClick={addRol} style={{ width: '100%', padding: '7px', borderRadius: 8, border: '1.5px dashed #E0DFD9', background: 'transparent', cursor: 'pointer', fontSize: 12, color: '#9B9B97', marginTop: 4 }}>+ Añadir rol</button>
          </div>
        )}
      </div>

      {/* Sección 0 — Estado de la obra */}
      <div style={{ border: '1px solid #E8E7E1', borderRadius: 10, marginBottom: 14, padding: '12px 14px' }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: '#141412', letterSpacing: '0.02em', marginBottom: 8 }}>
          <span style={{ color: '#52524E', marginRight: 6 }}>0</span>ESTADO DE LA OBRA
        </div>
        <textarea placeholder="Describe brevemente el estado general de la obra en esta visita..." value={vo.estadoObra?.descripcion||''} onChange={e => updEstado('descripcion', e.target.value)} style={{ minHeight: 64, marginBottom: 10 }} />
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 8 }}>
          {(vo.estadoObra?.fotos||[]).map(f => (
            <div key={f.id} style={{ position: 'relative', width: isMobile ? 80 : 110, height: isMobile ? 60 : 80 }}>
              <img src={f.data} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover', borderRadius: 7, display: 'block' }} />
              <button onClick={() => delFotoEstado(f.id)} style={{ position: 'absolute', top: 3, right: 3, background: 'rgba(0,0,0,0.55)', color: '#fff', border: 'none', borderRadius: '50%', width: 18, height: 18, cursor: 'pointer', fontSize: 11, display: 'flex', alignItems: 'center', justifyContent: 'center', lineHeight: 1 }}>×</button>
            </div>
          ))}
          <button onClick={addFotoEstado} style={{ width: isMobile ? 80 : 110, height: isMobile ? 60 : 80, borderRadius: 7, border: '1.5px dashed #E0DFD9', background: '#FAFAF8', cursor: 'pointer', fontSize: 22, color: '#D0D0CB', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>+</button>
        </div>
      </div>

      {/* Secciones editables */}
      {(vo.secciones||[]).map(sec => {
        const activos = activosPorSec(sec.id);
        return (
          <div key={sec.id} style={{ marginBottom: 16 }}>
            {/* Cabecera sección con editar nombre */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 7 }}>
              <span style={{ fontSize: 11, fontWeight: 700, color: '#52524E', flexShrink: 0 }}>{sec.codigo}</span>
              {editandoSec === sec.id
                ? <input autoFocus value={sec.titulo} onChange={e => updSeccion(sec.id, 'titulo', e.target.value)} onBlur={() => setEditandoSec(null)} style={{ flex: 1, fontSize: 12, fontWeight: 600 }} />
                : <span onClick={() => setEditandoSec(sec.id)} style={{ fontSize: 12, fontWeight: 600, color: '#141412', cursor: 'text', flex: 1 }}>{sec.titulo}</span>}
              <span style={{ fontSize: 11, color: '#A5A5A0' }}>{activos.length}</span>
              <button onClick={() => setBorrar({ tipo: 'seccion', id: sec.id, label: sec.titulo })} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#D4D3CE', fontSize: 15, lineHeight: 1 }}>×</button>
            </div>
            {/* Temas activos */}
            {activos.map(t => {
              const ult = t.entradas[t.entradas.length - 1];
              const est = ESTADOS_VO[ult.estado] || ESTADOS_VO.P;
              return (
                <TemaVO key={t.id} t={t} est={est} secId={sec.id}
                  onUpdEntrada={(tId,eId,campo,val) => updEntrada(sec.id, tId, eId, campo, val)}
                  onAddEntrada={(tId,txt) => addEntrada(sec.id, tId, txt)}
                  onDel={() => setBorrar({ tipo: 'tema', secId: sec.id, id: t.id, label: t.num })} />
              );
            })}
            <NuevoTema onAdd={txt => addTema(sec.id, txt)} />
          </div>
        );
      })}
      <button onClick={addSeccion} style={{ width: '100%', padding: '8px', borderRadius: 9, border: '1.5px dashed #E0DFD9', background: 'transparent', cursor: 'pointer', fontSize: 12, color: '#9B9B97', marginBottom: 14 }}>+ Añadir sección</button>

      {/* Modal histórico */}
      {showHistorico && (
        <Modal title="Histórico de temas resueltos" onClose={() => setShowHistorico(false)}>
          {todosResueltos.length === 0 ? <div style={{ fontSize: 13, color: '#A5A5A0' }}>Ningún tema resuelto todavía.</div>
            : todosResueltos.map(t => (
              <div key={t.id} style={{ padding: '9px 0', borderBottom: '1px solid #F2F1ED' }}>
                <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 3 }}>
                  <span style={{ fontSize: 11, fontWeight: 700, color: '#52524E' }}>{t.num}</span>
                  <Pill label="Resuelto" bg={ESTADOS_VO.R.bg} color={ESTADOS_VO.R.color} />
                  <span style={{ fontSize: 11, color: '#A5A5A0', marginLeft: 'auto' }}>Acta {String(t.resueltoEnActa||'—').padStart(2,'0')}</span>
                </div>
                <div style={{ fontSize: 13, color: '#18180F', lineHeight: 1.5 }}>{t.entradas[t.entradas.length-1].texto}</div>
              </div>
            ))}
        </Modal>
      )}

      {/* Confirmaciones */}
      {borrar && <ConfirmMini
        titulo={borrar.tipo === 'seccion' ? 'Eliminar sección' : 'Eliminar tema'}
        texto={borrar.tipo === 'seccion' ? `Vas a eliminar "${borrar.label}" y todos sus temas.` : `Vas a eliminar el tema ${borrar.label} y su historial.`}
        onSi={() => borrar.tipo === 'seccion' ? delSeccion(borrar.id) : delTema(borrar.secId, borrar.id)}
        onNo={() => setBorrar(null)} />}
    </div>
  );
}

// Componente de un tema (para evitar closures stale en los selects)
function TemaVO({ t, est, secId, onUpdEntrada, onAddEntrada, onDel }) {
  const [abierto, setAbierto] = useState(false);
  const ult = t.entradas[t.entradas.length - 1];
  return (
    <div style={{ border: `1px solid ${abierto ? '#C5C4BE' : '#E8E7E1'}`, borderRadius: 9, marginBottom: 6, overflow: 'hidden' }}>
      <div onClick={() => setAbierto(v => !v)} style={{ display: 'flex', alignItems: 'center', gap: 9, padding: '9px 12px', cursor: 'pointer', background: abierto ? '#FAFAF8' : '#fff' }}>
        <span style={{ fontSize: 11, fontWeight: 700, color: '#52524E', flexShrink: 0, minWidth: 34 }}>{t.num}</span>
        <span style={{ flex: 1, fontSize: 13, color: '#18180F', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{ult.texto}</span>
        <Pill label={est.label} bg={est.bg} color={est.color} />
      </div>
      {abierto && (
        <div className="fade" style={{ padding: '4px 12px 12px', borderTop: '1px solid #F2F1ED' }}>
          {t.entradas.map(en => (
            <div key={en.id} style={{ padding: '8px 0', borderBottom: '1px solid #F5F4F0' }}>
              <div style={{ display: 'flex', gap: 7, alignItems: 'flex-start' }}>
                <span style={{ fontSize: 11, color: '#A5A5A0', whiteSpace: 'nowrap', paddingTop: 2, minWidth: 36 }}>{fmtShort(en.fecha)}</span>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 13, color: '#18180F', lineHeight: 1.5, marginBottom: 5 }}>{en.texto}</div>
                  <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', alignItems: 'center' }}>
                    <select value={en.estado} onChange={ev => onUpdEntrada(t.id, en.id, 'estado', ev.target.value)}
                      style={{ width: 'auto', fontSize: 11, padding: '3px 7px', borderRadius: 6, border: `1px solid ${ESTADOS_VO[en.estado].color}40`, background: ESTADOS_VO[en.estado].bg, color: ESTADOS_VO[en.estado].color, fontWeight: 500 }}>
                      {Object.entries(ESTADOS_VO).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
                    </select>
                    {en.estado === 'R' && <input type="date" value={en.fin||''} onChange={ev => onUpdEntrada(t.id, en.id, 'fin', ev.target.value)} style={{ width: 'auto', fontSize: 11 }} />}
                    <select value={en.resp||''} onChange={ev => onUpdEntrada(t.id, en.id, 'resp', ev.target.value)}
                      style={{ width: 'auto', fontSize: 11, padding: '3px 7px', borderRadius: 6, border: '1px solid #E0DFD9' }}>
                      {RESP_VO.map(r => <option key={r} value={r}>{r ? `${r}` : '— resp.'}</option>)}
                    </select>
                  </div>
                </div>
              </div>
            </div>
          ))}
          <NuevaEntrada onAdd={txt => onAddEntrada(t.id, txt)} />
          <div style={{ marginTop: 8, textAlign: 'right' }}><Btn sm danger onClick={onDel}>Eliminar tema</Btn></div>
        </div>
      )}
    </div>
  );
}

function NuevoTema({ onAdd }) {
  const [open, setOpen] = useState(false);
  const [txt, setTxt] = useState('');
  if (!open) return <button onClick={() => setOpen(true)} style={{ width: '100%', padding: '6px', borderRadius: 8, border: '1.5px dashed #E0DFD9', background: 'transparent', cursor: 'pointer', fontSize: 12, color: '#9B9B97' }}>+ Nuevo tema</button>;
  return (
    <div className="fade" style={{ background: '#F9F8F5', borderRadius: 9, padding: 10, marginTop: 2 }}>
      <textarea autoFocus value={txt} onChange={e => setTxt(e.target.value)} placeholder="Describe el tema tratado..." style={{ minHeight: 50, marginBottom: 8 }} />
      <div style={{ display: 'flex', gap: 8 }}>
        <Btn primary full disabled={!txt.trim()} onClick={() => { onAdd(txt); setTxt(''); setOpen(false); }}>Añadir tema</Btn>
        <Btn onClick={() => { setTxt(''); setOpen(false); }}>✕</Btn>
      </div>
    </div>
  );
}
function NuevaEntrada({ onAdd }) {
  const [open, setOpen] = useState(false);
  const [txt, setTxt] = useState('');
  if (!open) return <button onClick={() => setOpen(true)} style={{ marginTop: 8, fontSize: 12, color: '#6B6B66', background: 'none', border: 'none', cursor: 'pointer', padding: 0 }}>+ Añadir seguimiento de esta visita</button>;
  return (
    <div className="fade" style={{ marginTop: 8 }}>
      <textarea autoFocus value={txt} onChange={e => setTxt(e.target.value)} placeholder="Novedad de esta visita..." style={{ minHeight: 50, marginBottom: 6 }} />
      <div style={{ display: 'flex', gap: 8 }}>
        <Btn sm primary disabled={!txt.trim()} onClick={() => { onAdd(txt); setTxt(''); setOpen(false); }}>Añadir</Btn>
        <Btn sm onClick={() => { setTxt(''); setOpen(false); }}>✕</Btn>
      </div>
    </div>
  );
}

// ── Generador PDF Acta VO ────────────────────────────────────────────────────
async function generarActaVO(obra, vo) {
  if (!window.jspdf) {
    await new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
      s.onload = res; s.onerror = () => rej(new Error('No se pudo cargar jsPDF'));
      document.head.appendChild(s);
    });
  }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const PW = 210, PH = 297, M = 14, CW = PW - M * 2;
  const NEGRO = [0,0,0], GRIS = [217,217,217], LH = 4;
  const num = String(vo.num).padStart(2,'0');
  const fecha = new Date().toLocaleDateString('es-ES');
  let y = 0;

  function pie() {
    doc.setTextColor(0,0,0); doc.setFont('helvetica','normal'); doc.setFontSize(7);
    doc.text('Plaat Arquitectura Técnica', M, PH-8);
    doc.text('Barcelona – Madrid', PW/2, PH-8, { align:'center' });
    doc.text('www.plaat.es', PW-M, PH-8, { align:'right' });
  }
  function cabecera() {
    doc.setFont('helvetica','bold'); doc.setFontSize(7.5); doc.setTextColor(0,0,0);
    doc.text(`ACTA DE VISITA DE OBRA N.º ${num}`, M, 10);
    doc.setFontSize(22); doc.text('Plaat.', PW-M, 12, { align:'right' });
  }
  function check(h) {
    if (y + h > PH - 18) { doc.addPage(); cabecera(); pie(); y = 20; }
  }
  function rectRow(x, yy, cols, h, fillHeader) {
    let cx = x;
    cols.forEach(([w]) => {
      if (fillHeader) { doc.setFillColor(...GRIS); doc.rect(cx, yy, w, h, 'F'); }
      doc.setDrawColor(...NEGRO); doc.setLineWidth(0.2); doc.rect(cx, yy, w, h);
      cx += w;
    });
  }
  function textInCell(x, yy, w, h, txt, bold, size) {
    doc.setFont('helvetica', bold ? 'bold' : 'normal');
    doc.setFontSize(size || 8);
    doc.setTextColor(0,0,0);
    const ll = doc.splitTextToSize(String(txt||''), w - 3);
    const ty = yy + h/2 - (ll.length-1)*(size||8)*0.352645/2 + (size||8)*0.352645/2;
    doc.text(ll, x+2, ty);
  }

  cabecera(); pie(); y = 18;

  // ── Título + Fecha/Lugar/Fase ────────────────────────────────────────────
  doc.setFillColor(...GRIS); doc.rect(M,y,CW,8,'F');
  doc.setDrawColor(...NEGRO); doc.rect(M,y,CW,8);
  doc.setFont('helvetica','bold'); doc.setFontSize(11); doc.setTextColor(0,0,0);
  doc.text(`ACTA DE VISITA DE OBRA N.º ${num}`, M+3, y+5.5);
  y += 8;
  [['FECHA', fecha], ['LUGAR', vo.lugar||'Obra'], ['FASE', vo.fase||'']].forEach(([k,v]) => {
    check(7);
    doc.rect(M,y,40,7); doc.rect(M+40,y,CW-40,7);
    doc.setFont('helvetica','bold'); doc.setFontSize(8.5); doc.text(k, M+2, y+4.7);
    doc.setFont('helvetica','normal'); doc.text(String(v), M+42, y+4.7);
    y += 7;
  });
  y += 4;

  // ── Equipo técnico ──────────────────────────────────────────────────────
  check(8);
  doc.setFillColor(...GRIS); doc.rect(M,y,CW,7,'F'); doc.rect(M,y,CW,7);
  doc.setFont('helvetica','bold'); doc.setFontSize(8.5); doc.setTextColor(0,0,0);
  doc.text('EQUIPO TÉCNICO Y DATOS DE CONTACTO', M+2, y+4.8); y += 7;

  // Col widths: Rol 44 | Empresa 24 | Nombre 30 | Email 34 | Tel 22 | Asistido 28
  const eW = [44,24,30,34,22,28];
  // Header de columnas
  check(6);
  let hx = M;
  [['ROL','bold'], ['EMPRESA','bold'], ['NOMBRE','bold'], ['EMAIL','bold'], ['TEL.','bold'], ['ASISTIDO','bold']].forEach(([t,b], i) => {
    doc.setFillColor(235,235,230); doc.rect(hx, y, eW[i], 6, 'F');
    doc.setDrawColor(...NEGRO); doc.rect(hx,y,eW[i],6);
    doc.setFont('helvetica','bold'); doc.setFontSize(6.5); doc.setTextColor(80,80,75);
    doc.text(t, hx+2, y+4); hx += eW[i];
  });
  y += 6;

  (vo.equipo||[]).forEach(rol => {
    const ps = rol.personas||[{ empresa:'', nombre:'', email:'', tel:'', asistio: false }];
    ps.forEach((p, pi) => {
      const h = 7;
      check(h);
      const rx = M;
      doc.setDrawColor(...NEGRO); doc.setLineWidth(0.2);
      if (pi === 0) {
        const rowH = h * ps.length;
        doc.setFillColor(...GRIS); doc.rect(rx, y, eW[0], rowH, 'F');
        doc.rect(rx, y, eW[0], rowH);
        doc.setFont('helvetica','bold'); doc.setFontSize(7.5); doc.setTextColor(0,0,0);
        const ll = doc.splitTextToSize(rol.nombre||'', eW[0]-3);
        doc.text(ll, rx+2, y + 4);
      } else {
        doc.rect(rx, y, eW[0], h);
      }
      let cx = rx + eW[0];
      [p.empresa, p.nombre, p.email, p.tel].forEach((v,i) => {
        doc.rect(cx, y, eW[i+1], h);
        doc.setFont('helvetica','normal'); doc.setFontSize(7.5); doc.setTextColor(0,0,0);
        const ll = doc.splitTextToSize(String(v||''), eW[i+1]-3);
        doc.text(ll, cx+2, y+4.2);
        cx += eW[i+1];
      });
      // Asistido checkbox
      doc.rect(cx, y, eW[5], h);
      const chx = cx + eW[5]/2 - 2.5, chy = y + h/2 - 2.5;
      doc.rect(chx, chy, 5, 5);
      if (p.asistio) { doc.setFont('helvetica','bold'); doc.setFontSize(8); doc.text('✓', chx+0.8, chy+4); }
      y += h;
    });
  });
  y += 5;

  // ── Sección 0: Estado de la obra ────────────────────────────────────────
  const eo = vo.estadoObra || {};
  if (eo.descripcion || (eo.fotos||[]).length > 0) {
    check(14);
    // Cabecera sección 0
    doc.setFillColor(...GRIS); doc.rect(M,y,16,7,'F'); doc.rect(M+16,y,CW-16,7,'F');
    doc.setDrawColor(...NEGRO); doc.rect(M,y,16,7); doc.rect(M+16,y,CW-16,7);
    doc.setFont('helvetica','bold'); doc.setFontSize(8); doc.setTextColor(0,0,0);
    doc.text('0', M+3, y+4.8); doc.text('ESTADO DE LA OBRA (FOTOGRAFÍAS)', M+19, y+4.8);
    y += 7;

    if (eo.descripcion) {
      const dl = doc.splitTextToSize(eo.descripcion, CW-4);
      const dh = dl.length * LH + 5;
      check(dh);
      doc.rect(M, y, CW, dh);
      doc.setFont('helvetica','normal'); doc.setFontSize(8.5);
      doc.text(dl, M+2, y+5);
      y += dh + 2;
    }

    // Fotos 2 por fila, proporción real
    const fotos = eo.fotos||[];
    const fW2 = (CW - 4) / 2;
    for (let fi = 0; fi < fotos.length; fi += 2) {
      let rowH = 0;
      const pair = [fotos[fi], fotos[fi+1]].filter(Boolean);
      const dims = pair.map(f => {
        try { const pr = doc.getImageProperties(f.data); const r = pr.height/pr.width; const h = Math.min(fW2 * r, 68); return { w: h/r, h }; }
        catch(e) { return { w: fW2, h: 55 }; }
      });
      rowH = Math.max(...dims.map(d => d.h));
      check(rowH + 3);
      pair.forEach((f, pi) => {
        const x = M + pi * (fW2 + 4);
        doc.addImage(f.data, 'JPEG', x, y, dims[pi].w, dims[pi].h);
      });
      y += rowH + 4;
    }
    y += 3;
  }

  // ── Secciones 1-N ────────────────────────────────────────────────────────
  const secciones = vo.secciones||[];
  secciones.forEach(sec => {
    const activos = (sec.temas||[]).filter(t => !(t.resuelto && t.resueltoEnActa && t.resueltoEnActa < vo.num));
    if (!activos.length) return;

    // Widths: Núm 14 | Descripción adaptada | Estado 14 | Inicio 18 | Fin 18 | Res 14
    const col = { n:14, e:14, ini:18, fin:18, res:14, desc: CW-14-14-18-18-14 };

    check(8);
    // Cabecera sección
    [M, M+col.n, M+col.n+col.desc, M+col.n+col.desc+col.e, M+col.n+col.desc+col.e+col.ini, M+col.n+col.desc+col.e+col.ini+col.fin].forEach((x,i) => {
      const w = [col.n, col.desc, col.e, col.ini, col.fin, col.res][i];
      doc.setFillColor(...GRIS); doc.rect(x, y, w, 7, 'F');
      doc.setDrawColor(...NEGRO); doc.rect(x, y, w, 7);
    });
    doc.setFont('helvetica','bold'); doc.setFontSize(8); doc.setTextColor(0,0,0);
    doc.text(sec.codigo, M+2, y+4.8);
    doc.text(sec.titulo, M+col.n+2, y+4.8);
    doc.setFontSize(6.5);
    ['Estado','Inicio','Fin','Res.'].forEach((t,i) => {
      const x = [M+col.n+col.desc, M+col.n+col.desc+col.e, M+col.n+col.desc+col.e+col.ini, M+col.n+col.desc+col.e+col.ini+col.fin][i];
      doc.text(t, x+1, y+4.5);
    });
    y += 7;

    // Temas
    activos.forEach(t => {
      t.entradas.forEach((en, ei) => {
        const txt = en.texto||'';
        const tl = doc.splitTextToSize(txt, col.desc-3);
        const h = Math.max(6, tl.length*LH + 2.5);
        check(h);
        doc.setDrawColor(...NEGRO); doc.setLineWidth(0.15);
        doc.rect(M, y, col.n, h);
        doc.rect(M+col.n, y, col.desc, h);
        doc.rect(M+col.n+col.desc, y, col.e, h);
        doc.rect(M+col.n+col.desc+col.e, y, col.ini, h);
        doc.rect(M+col.n+col.desc+col.e+col.ini, y, col.fin, h);
        doc.rect(M+col.n+col.desc+col.e+col.ini+col.fin, y, col.res, h);
        doc.setFont('helvetica','bold'); doc.setFontSize(7.5); doc.setTextColor(0,0,0);
        if (ei === 0) doc.text(t.num, M+2, y+4);
        doc.setFont('helvetica','normal');
        doc.text(tl, M+col.n+2, y+4);
        doc.text(en.estado||'', M+col.n+col.desc+1, y+4);
        const isR = en.estado === 'R';
        doc.text(isR ? '' : fmtFechaCorta(en.fecha), M+col.n+col.desc+col.e+1, y+4);
        doc.text(isR ? fmtFechaCorta(en.fin||en.fecha) : '', M+col.n+col.desc+col.e+col.ini+1, y+4);
        doc.text(en.resp||'', M+col.n+col.desc+col.e+col.ini+col.fin+1, y+4);
        y += h;
      });
    });
    y += 4;
  });

  // ── Seguimiento conforme / NOTA ─────────────────────────────────────────
  check(24);
  y += 4;
  doc.setFont('helvetica','italic'); doc.setFontSize(8); doc.setTextColor(0,0,0);
  const nota = 'NOTA: La presente acta se entenderá como conforme en caso de no manifestar comentarios en el plazo de 48 horas tras su difusión.';
  const notaL = doc.splitTextToSize(nota, CW);
  doc.text(notaL, M, y); y += notaL.length * LH + 3;
  doc.setFont('helvetica','normal');
  doc.text('Conforme, firma y fecha:', M, y); y += 8;

  // ── Cuadro de firmas ─────────────────────────────────────────────────────
  check(28);
  const firmas = ['PROMOTOR', 'DIRECCIÓN DE OBRA', 'DIRECCIÓN DE EJECUCIÓN\nCOORD. SEGURIDAD', 'CONTRATISTA'];
  const fw = CW/4;
  firmas.forEach((f, i) => {
    const fx = M + i*fw;
    doc.rect(fx, y, fw, 24);
    doc.setFont('helvetica','bold'); doc.setFontSize(7); doc.setTextColor(0,0,0);
    const fl = doc.splitTextToSize(f, fw-4);
    doc.text(fl, fx+2, y+4);
  });

  const total = doc.getNumberOfPages();
  for (let p=1; p<=total; p++) { doc.setPage(p); pie(); }
  doc.save(`Acta_VO_${num}_${(obra.nombre||'obra').replace(/\s+/g,'_')}.pdf`);
}

function fmtFechaCorta(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getFullYear()).slice(2)}`;
}



function DetalleObra({ obra, onBack, onSave, isMobile }) {
  const [tab, setTab]               = useState('inspecciones');
  const [editEstado, setEditEstado] = useState(false);

  const tabs = [
    { id: 'inspecciones', label: 'Inspecciones',   short: 'Insp.'   },
    { id: 'incidencias',  label: 'Incidencias',    short: 'Incid.'  },
    { id: 'calidad',      label: 'Calidad',         short: 'Calidad' },
    { id: 'actavo',       label: 'Acta VO',         short: 'Acta VO' },
    { id: 'anotaciones',  label: 'Notas y tareas',  short: 'Notas'   },
  ];

  const accentColor = STATUS_ACCENT[obra.estado] || STATUS_ACCENT.en_curso;
  const e           = ESTADOS_OBRA[obra.estado]  || ESTADOS_OBRA.en_curso;
  const pad         = isMobile ? '12px 14px' : '18px 22px';

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden' }}>

      {/* Header oscuro — mismo estilo que el sidebar */}
      <div style={{ background: '#1C1C1A', flexShrink: 0, paddingTop: isMobile ? 'env(safe-area-inset-top)' : 0 }}>

        {/* Breadcrumb + nombre + estado */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: isMobile ? '12px 14px' : '14px 22px 12px', borderBottom: isMobile ? 'none' : '1px solid rgba(255,255,255,0.07)' }}>
          <button onClick={onBack} style={{ fontSize: isMobile ? 14 : 12, color: 'rgba(255,255,255,0.55)', background: 'none', border: 'none', cursor: 'pointer', padding: 0, display: 'flex', alignItems: 'center', gap: 4, flexShrink: 0 }}>
            ←{isMobile ? '' : ' Mis obras'}
          </button>
          {!isMobile && <span style={{ color: 'rgba(255,255,255,0.15)' }}>/</span>}
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 14, fontWeight: 500, color: '#F2F1ED', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{obra.nombre}</div>
            {isMobile && <div style={{ fontSize: 11, color: 'rgba(255,255,255,0.35)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{obra.cliente}</div>}
          </div>
          {!isMobile && <span style={{ fontSize: 12, color: 'rgba(255,255,255,0.35)', whiteSpace: 'nowrap', flexShrink: 0 }}>{obra.cliente}</span>}

          {/* Estado editable */}
          <div style={{ position: 'relative', flexShrink: 0 }}>
            <button onClick={() => setEditEstado(v => !v)} style={{ fontSize: 11, padding: '3px 10px', borderRadius: 20, background: 'transparent', border: `1px solid ${accentColor}`, color: accentColor, cursor: 'pointer', fontWeight: 600, letterSpacing: '0.05em' }}>
              {e.label.toUpperCase()}
            </button>
            {editEstado && (
              <div style={{ position: 'absolute', right: 0, top: 'calc(100% + 6px)', background: '#2A2A28', border: '1px solid rgba(255,255,255,0.1)', borderRadius: 10, boxShadow: '0 12px 32px rgba(0,0,0,.4)', padding: 6, zIndex: 10, minWidth: 160 }}>
                {Object.entries(ESTADOS_OBRA).map(([k, v]) => (
                  <div key={k} onClick={() => { onSave({ ...obra, estado: k }); setEditEstado(false); }}
                    style={{ padding: '7px 12px', borderRadius: 7, cursor: 'pointer', fontSize: 13, fontWeight: obra.estado === k ? 500 : 400, background: obra.estado === k ? 'rgba(255,255,255,0.08)' : 'transparent', color: STATUS_ACCENT[k] || '#F2F1ED' }}>
                    {v.label}
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        {/* Tabs arriba — solo escritorio */}
        {!isMobile && (
          <div className="no-scrollbar" style={{ display: 'flex', padding: '0 22px', overflowX: 'auto' }}>
            {tabs.map(t => (
              <button key={t.id} onClick={() => setTab(t.id)} style={{
                padding: '11px 16px', background: 'none', border: 'none',
                borderBottom: `2px solid ${tab === t.id ? accentColor : 'transparent'}`,
                cursor: 'pointer', fontSize: 13, whiteSpace: 'nowrap', flexShrink: 0,
                fontWeight: tab === t.id ? 500 : 400,
                color: tab === t.id ? '#F2F1ED' : 'rgba(255,255,255,0.32)',
                transition: 'all .15s',
              }}>
                {t.label}
              </button>
            ))}
          </div>
        )}
      </div>

      {/* Contenido */}
      <div style={{ flex: 1, overflow: 'auto', padding: pad, background: '#F2F1ED', minHeight: 0 }}>
        <div key={tab} className="fade">
          {tab === 'inspecciones' && <ModuloInspecciones obra={obra} onSave={onSave} />}
          {tab === 'incidencias'  && <ModuloIncidencias  obra={obra} onSave={onSave} />}
          {tab === 'calidad'      && <ModuloCalidad obra={obra} onSave={onSave} />}
          {tab === 'actavo'       && <ModuloActaVO obra={obra} onSave={onSave} />}
          {tab === 'anotaciones'  && <ModuloApuntes obra={obra} onSave={onSave} />}
        </div>
      </div>

      {/* Tabs abajo — solo móvil */}
      {isMobile && (
        <div style={{ display: 'flex', background: '#1C1C1A', borderTop: '1px solid rgba(255,255,255,0.08)', flexShrink: 0, paddingBottom: 'env(safe-area-inset-bottom)' }}>
          {tabs.map(t => {
            const activo = tab === t.id;
            return (
              <button key={t.id} onClick={() => setTab(t.id)} style={{ flex: 1, background: 'none', border: 'none', borderTop: `2px solid ${activo ? accentColor : 'transparent'}`, cursor: 'pointer', padding: '11px 2px 12px', color: activo ? '#F2F1ED' : 'rgba(255,255,255,0.4)', fontSize: 11.5, fontWeight: activo ? 600 : 400, letterSpacing: '0.01em' }}>
                {t.short}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ─── Pantalla de login ────────────────────────────────────────────────────────

function LoginScreen() {
  const [email, setEmail]     = useState('');
  const [pass,  setPass]      = useState('');
  const [error, setError]     = useState('');
  const [loading, setLoading] = useState(false);

  async function entrar() {
    if (!email.trim() || !pass) return;
    setLoading(true); setError('');
    try {
      await window.auth.signIn(email.trim(), pass);
      // el cambio de sesión actualiza el estado en App automáticamente
    } catch (e) {
      setError('Correo o contraseña incorrectos.');
      setLoading(false);
    }
  }

  const onKey = e => { if (e.key === 'Enter') entrar(); };

  return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#1C1C1A', padding: 20 }}>
      <div className="fade" style={{ background: '#fff', borderRadius: 16, padding: '34px 30px', width: 370, maxWidth: '100%', boxShadow: '0 24px 64px rgba(0,0,0,0.35)' }}>
        <div style={{ fontSize: 26, fontWeight: 700, letterSpacing: '0.1em', color: '#141412' }}>PLAAT</div>
        <div style={{ fontSize: 12, letterSpacing: '0.18em', color: '#9B9B97', marginTop: 6, marginBottom: 26, fontWeight: 400 }}>/ DEO · ARQUITECTURA TÉCNICA</div>

        <div style={{ marginBottom: 12 }}>
          <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 5 }}>Correo</label>
          <input type="email" autoFocus value={email} onChange={e => setEmail(e.target.value)} onKeyDown={onKey} placeholder="tu@correo.com" />
        </div>
        <div style={{ marginBottom: 18 }}>
          <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 5 }}>Contraseña</label>
          <input type="password" value={pass} onChange={e => setPass(e.target.value)} onKeyDown={onKey} placeholder="••••••••" />
        </div>

        {error && <div style={{ fontSize: 12, color: '#8A1F1F', background: '#FDECEC', borderRadius: 8, padding: '8px 11px', marginBottom: 14 }}>{error}</div>}

        <Btn primary full onClick={entrar} disabled={loading || !email.trim() || !pass}>
          {loading ? 'Entrando…' : 'Entrar'}
        </Btn>

        <div style={{ fontSize: 11, color: '#A5A5A0', marginTop: 16, textAlign: 'center', lineHeight: 1.5 }}>
          Acceso restringido al equipo de PLAAT.<br />Si no tienes cuenta, pídela al administrador.
        </div>
      </div>
    </div>
  );
}

// ─── App principal ────────────────────────────────────────────────────────────

export default function App() {
  const isMobile = useIsMobile();
  const [user,       setUser]       = useState(undefined); // undefined=cargando, null=sin login, obj=dentro
  const [obras,      setObras]      = useState([]);
  const [loading,    setLoading]    = useState(true);
  const [nav,        setNav]        = useState('alertas');
  const [obraActiva, setObraActiva] = useState(null);
  const [showNueva,  setShowNueva]  = useState(false);
  const [obraEditar,   setObraEditar]   = useState(null);
  const [obraEliminar, setObraEliminar] = useState(null);

  // Sesión: si no hay sistema de auth (p.ej. dentro de Claude), entra directo
  useEffect(() => {
    if (!window.auth) { setUser({ local: true }); return; }
    window.auth.getUser().then(u => setUser(u || null));
    window.auth.onChange(u => setUser(u || null));
  }, []);

  // Cargar obras: primero intenta índice nuevo, luego migra del blob antiguo
  useEffect(() => {
    if (!user) return;
    let cancelado = false;
    (async () => {
      setLoading(true);
      try {
        if (window.storage) {
          const T = ms => new Promise((_, r) => setTimeout(() => r(new Error('timeout')), ms));
          let listaObras = [];

          // 1. Intenta índice nuevo
          const idxR = await Promise.race([window.storage.get(SK_IDX, true), T(8000)]).catch(() => null);
          if (idxR?.value) {
            const ids = JSON.parse(idxR.value);
            const resultados = await Promise.all(ids.map(id =>
              Promise.race([window.storage.get(SK_OBR(id), true), T(8000)])
                .then(r => r?.value ? JSON.parse(r.value) : null)
                .catch(() => null)
            ));
            listaObras = resultados.filter(Boolean);
          } else {
            // 2. Migra blob antiguo
            const viejoR = await Promise.race([window.storage.get(SK, true), T(8000)]).catch(() => null);
            if (viejoR?.value) {
              listaObras = JSON.parse(viejoR.value);
              // Guarda cada obra por separado
              await Promise.all(listaObras.map(o =>
                window.storage.set(SK_OBR(o.id), JSON.stringify(o), true).catch(() => null)
              ));
              await window.storage.set(SK_IDX, JSON.stringify(listaObras.map(o => o.id)), true).catch(() => null);
              console.log('Migración completada:', listaObras.length, 'obras');
            }
          }
          if (!cancelado) setObras(listaObras);
        }
      } catch (e) { console.error('Carga de obras:', e); }
      if (!cancelado) setLoading(false);
    })();
    return () => { cancelado = true; };
  }, [user]);

  // Guarda UNA obra + actualiza el índice (operación mínima de escritura)
  async function saveUnaObra(obra, lista) {
    const T = ms => new Promise((_, r) => setTimeout(() => r(new Error('timeout')), ms));
    const obraJSON = JSON.stringify(obra);
    const kb = Math.round(obraJSON.length / 1024);
    if (kb > 4000) {
      alert(`Esta obra ocupa ${kb} KB. Las fotos y planos consumen mucho espacio — considera reducirlas o eliminar las antiguas. Se intentará guardar igualmente.`);
    }
    try {
      if (!window.storage) return;
      await Promise.race([window.storage.set(SK_OBR(obra.id), obraJSON, true), T(10000)]);
      await Promise.race([window.storage.set(SK_IDX, JSON.stringify(lista.map(o => o.id)), true), T(10000)]);
    } catch (e) {
      console.error('Error guardando obra:', e);
      alert('No se pudieron guardar los cambios. Esta obra tiene demasiados datos (fotos/planos). Prueba a eliminar fotos antiguas de incidencias o planos para liberar espacio.');
    }
  }

  async function crearObra(data) {
    const obra = {
      id: uid(), nombre: data.nombre, cliente: data.cliente, direccion: data.direccion,
      responsable: data.responsable, diasVisita: data.diasVisita || [],
      emplazamiento: data.emplazamiento || '', propiedad: data.propiedad || '',
      proyectista: data.proyectista || '', direccionObra: data.direccionObra || '',
      constructora: data.constructora || '', deoFirmante: data.deoFirmante || '',
      numActaSeq: 0, estado: 'en_curso',
      disciplinas: [], lotes: [], incidencias: [], apuntes: [], creadaEn: now(),
    };
    const lista = [obra, ...obras];
    setObras(lista);
    await saveUnaObra(obra, lista);
    setShowNueva(false);
    setObraActiva(obra);
  }

  async function actualizarObra(updated) {
    const lista = obras.map(o => o.id === updated.id ? updated : o);
    setObras(lista);
    setObraActiva(updated);
    await saveUnaObra(updated, lista);
  }

  async function guardarEdicion(datos) {
    const updated = { ...obraEditar, ...datos };
    const lista = obras.map(o => o.id === obraEditar.id ? updated : o);
    setObras(lista);
    await saveUnaObra(updated, lista);
    setObraEditar(null);
  }

  async function eliminarObra() {
    const id = obraEliminar.id;
    const lista = obras.filter(o => o.id !== id);
    setObras(lista);
    try {
      if (window.storage) {
        await window.storage.delete(SK_OBR(id), true).catch(() => null);
        await window.storage.set(SK_IDX, JSON.stringify(lista.map(o => o.id)), true);
      }
    } catch (e) { console.error('Error eliminando obra:', e); }
    setObraEliminar(null);
  }

  // Calcular el contador del badge "Hoy"
  const stats = {
    total: obras.length,
    alertas: calcularAlertas(obras).total,
    incidencias: obras.reduce((s, o) => s + o.incidencias.filter(i => i.estado !== 'resuelta').length, 0),
  };

  // ── Portón de acceso ──
  if (user === undefined) {
    return (
      <>
        <style>{CSS}</style>
        <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#1C1C1A', color: 'rgba(255,255,255,0.5)', fontSize: 14 }}>Cargando…</div>
      </>
    );
  }
  if (user === null) {
    return (
      <>
        <style>{CSS}</style>
        <LoginScreen />
      </>
    );
  }

  // Vista detalle de obra
  if (obraActiva) {
    const fresh = obras.find(o => o.id === obraActiva.id) || obraActiva;
    return (
      <>
        <style>{CSS}</style>
        <div style={{ display: 'flex', flexDirection: isMobile ? 'column' : 'row', height: '100vh', overflow: 'hidden' }}>
          {!isMobile && <Sidebar nav={nav} setNav={setNav} stats={stats} user={user} />}
          <div style={{ flex: 1, overflow: 'hidden', display: 'flex', flexDirection: 'column', minHeight: 0 }}>
            <DetalleObra obra={fresh} onBack={() => setObraActiva(null)} onSave={actualizarObra} isMobile={isMobile} />
          </div>
        </div>
      </>
    );
  }

  return (
    <>
      <style>{CSS}</style>
      <div style={{ display: 'flex', flexDirection: isMobile ? 'column' : 'row', height: '100vh', overflow: 'hidden' }}>
        {!isMobile && <Sidebar nav={nav} setNav={setNav} stats={stats} user={user} />}
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', minHeight: 0 }}>

          {nav === 'alertas'   && <VistaAlertas obras={obras} onIrObra={o => setObraActiva(o)} isMobile={isMobile} />}
          {nav === 'tablero'   && (
            <>
              {/* Topbar */}
              <div style={{ background: '#fff', borderBottom: '1px solid #ECEAE4', padding: isMobile ? '13px 16px' : '13px 22px', paddingTop: isMobile ? 'calc(13px + env(safe-area-inset-top))' : '13px', display: 'flex', alignItems: 'center', gap: 12, flexShrink: 0 }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 16, fontWeight: 600, color: '#141412' }}>Mis obras</div>
                  {obras.length > 0 && <div style={{ fontSize: 12, color: '#9B9B97', marginTop: 1 }}>{obras.filter(o => o.estado === 'en_curso').length} en curso · {obras.length} en total</div>}
                </div>
                <Btn primary onClick={() => setShowNueva(true)}>+ Nueva obra</Btn>
              </div>

              <div style={{ flex: 1, overflow: 'auto', padding: '18px 22px' }}>
                {loading ? (
                  <div style={{ textAlign: 'center', padding: '48px', fontSize: 13, color: '#A5A5A0' }}>Cargando obras…</div>
                ) : obras.length === 0 ? (
                  <div style={{ textAlign: 'center', padding: '80px 20px', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 16 }}>
                    <div style={{ width: 56, height: 56, borderRadius: 12, background: '#1C1C1A', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 28, color: '#F2F1ED', fontWeight: 300 }}>/</div>
                    <div>
                      <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 6 }}>Sin obras todavía</div>
                      <div style={{ fontSize: 13, color: '#9B9B97', marginBottom: 18 }}>Crea tu primera obra para empezar el seguimiento</div>
                      <Btn primary onClick={() => setShowNueva(true)}>+ Crear primera obra</Btn>
                    </div>
                  </div>
                ) : (
                  <div className="list-in" style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                    {obras.map(o => <ObraCard key={o.id} obra={o} onClick={() => setObraActiva(o)} onEditar={setObraEditar} onEliminar={setObraEliminar} />)}
                  </div>
                )}
              </div>
            </>
          )}

        </div>
        {isMobile && <BottomNav nav={nav} setNav={setNav} stats={stats} />}
      </div>

      {showNueva && <ModalNuevaObra onClose={() => setShowNueva(false)} onCreate={crearObra} />}
      {obraEditar && <ModalNuevaObra obra={obraEditar} onClose={() => setObraEditar(null)} onCreate={guardarEdicion} />}

      {/* Confirmar eliminación */}
      {obraEliminar && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.2)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000, backdropFilter: 'blur(3px)' }} onClick={e => { if (e.target === e.currentTarget) setObraEliminar(null); }}>
          <div className="fade" style={{ background: '#fff', borderRadius: 14, width: 400, maxWidth: '95vw', border: '1px solid #E0DFD9', boxShadow: '0 24px 64px rgba(0,0,0,.14)', padding: '20px' }}>
            <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 8 }}>Eliminar obra</div>
            <div style={{ fontSize: 13, color: '#52524E', lineHeight: 1.5, marginBottom: 18 }}>
              Vas a eliminar <strong>{obraEliminar.nombre}</strong> y todo su contenido (incidencias, ensayos, notas...). Esta acción no se puede deshacer.
            </div>
            <div style={{ display: 'flex', gap: 8 }}>
              <Btn onClick={() => setObraEliminar(null)} full>Cancelar</Btn>
              <Btn danger full onClick={eliminarObra}>Eliminar</Btn>
            </div>
          </div>
        </div>
      )}

    </>
  );
}

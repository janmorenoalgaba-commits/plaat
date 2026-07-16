import { useState, useEffect, useRef, useMemo, memo } from "react";
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
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; -webkit-tap-highlight-color: transparent; }
html { -webkit-text-size-adjust: 100%; }
body { font-family: 'DM Sans', -apple-system, BlinkMacSystemFont, sans-serif; background: #F7F6F3; color: #16160F; font-size: 14px; -webkit-font-smoothing: antialiased; text-rendering: optimizeLegibility; }
button, input, select, textarea { font-family: inherit; }
input, select, textarea {
  display: block; width: 100%; padding: 10px 13px;
  border: 1px solid #E6E4DD; border-radius: 11px;
  font-size: 14px; color: #16160F; background: #fff;
  outline: none; transition: border-color .18s ease, box-shadow .18s ease;
}
input::placeholder, textarea::placeholder { color: #B5B4AE; }
input:focus, select:focus, textarea:focus {
  border-color: #16160F; box-shadow: 0 0 0 3.5px rgba(20,20,15,.06);
}
textarea { resize: vertical; min-height: 72px; line-height: 1.5; }
::-webkit-scrollbar { width: 4px; height: 4px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #D8D6CF; border-radius: 3px; }
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
  .auro, .login-card, .login-logo, .login-row, .stat-card, .shimmer-btn::after,
  .arch-grid, .arch-shape, .arch-dash, .scan, .splash-logo, .splash-sub, .splash-bar,
  .spl-ring, .spl-blueprint, .spl-draw, .spl-beam, .spl-flash, .spl-particle,
  .dash-banner .arch-grid, .dash-banner-ring, .dash-beam, .dash-title { animation: none !important; }
}

/* ── Movimiento "notable pero profesional" ───────────────── */
@keyframes auroraMove {
  0%   { transform: translate(-8%, -6%) scale(1); }
  50%  { transform: translate(8%, 6%) scale(1.15); }
  100% { transform: translate(-8%, -6%) scale(1); }
}
@keyframes riseIn {
  from { opacity: 0; transform: translateY(18px); }
  to   { opacity: 1; transform: none; }
}
@keyframes logoIn {
  from { opacity: 0; transform: translateY(10px); letter-spacing: 0.32em; }
  to   { opacity: 1; transform: none; letter-spacing: 0.1em; }
}
@keyframes underlineGrow {
  from { width: 0; }
  to   { width: 38px; }
}
@keyframes shimmer {
  0%   { transform: translateX(-120%); }
  60%  { transform: translateX(120%); }
  100% { transform: translateX(120%); }
}
@keyframes countUp {
  from { opacity: 0; transform: translateY(8px) scale(.96); }
  to   { opacity: 1; transform: none; }
}

/* Aurora de fondo en login */
.login-bg { position: relative; overflow: hidden; }
.auro {
  position: absolute; border-radius: 50%; filter: blur(70px);
  pointer-events: none; will-change: transform;
}
.auro-1 { width: 460px; height: 460px; background: rgba(82,161,36,0.16);  top: -120px; left: -100px; animation: auroraMove 16s ease-in-out infinite; }
.auro-2 { width: 380px; height: 380px; background: rgba(120,120,255,0.10); bottom: -140px; right: -80px; animation: auroraMove 20s ease-in-out infinite reverse; }
.auro-3 { width: 300px; height: 300px; background: rgba(255,255,255,0.05); top: 40%; left: 55%; animation: auroraMove 24s ease-in-out infinite; }

.login-card  { animation: riseIn .5s cubic-bezier(.2,.8,.2,1) both; }
.login-logo  { animation: logoIn .7s cubic-bezier(.2,.8,.2,1) both; }
.login-underline { animation: underlineGrow .6s cubic-bezier(.2,.8,.2,1) .35s both; }
.login-row   { animation: riseIn .5s ease both; }
.login-row.r1 { animation-delay: .12s; }
.login-row.r2 { animation-delay: .20s; }
.login-row.r3 { animation-delay: .28s; }
.login-row.r4 { animation-delay: .36s; }

/* Botón con barrido de brillo al hover */
.shimmer-btn { position: relative; overflow: hidden; }
.shimmer-btn::after {
  content: ''; position: absolute; top: 0; left: 0; width: 100%; height: 100%;
  background: linear-gradient(110deg, transparent 30%, rgba(255,255,255,0.18) 50%, transparent 70%);
  transform: translateX(-120%);
}
.shimmer-btn:hover::after { animation: shimmer 1.1s ease; }

/* Tarjetas de estadística (tablero) con entrada y elevación */
.stat-card { animation: countUp .5s cubic-bezier(.2,.8,.2,1) both; transition: transform .18s ease, box-shadow .2s; }
.stat-card:hover { transform: translateY(-2px); box-shadow: 0 6px 22px rgba(0,0,0,0.08); }
.stat-card.s1 { animation-delay: .04s; }
.stat-card.s2 { animation-delay: .10s; }
.stat-card.s3 { animation-delay: .16s; }
.stat-card.s4 { animation-delay: .22s; }

/* ── Movimiento arquitectónico (login + splash) ──────────── */
@keyframes drift1 { 0%{transform:translate(0,0) rotate(0deg);} 50%{transform:translate(40px,-30px) rotate(8deg);} 100%{transform:translate(0,0) rotate(0deg);} }
@keyframes drift2 { 0%{transform:translate(0,0) rotate(0deg);} 50%{transform:translate(-50px,40px) rotate(-10deg);} 100%{transform:translate(0,0) rotate(0deg);} }
@keyframes drift3 { 0%{transform:translate(0,0) scale(1);} 50%{transform:translate(30px,30px) scale(1.1);} 100%{transform:translate(0,0) scale(1);} }
@keyframes dashFlow { to { stroke-dashoffset: -1000; } }
@keyframes rotateSlow { to { transform: rotate(360deg); } }
@keyframes rotateSlowRev { to { transform: rotate(-360deg); } }
@keyframes floatY { 0%{transform:translateY(0);} 50%{transform:translateY(-22px);} 100%{transform:translateY(0);} }
@keyframes gridPan { 0%{transform:translate(0,0);} 100%{transform:translate(56px,56px);} }
@keyframes pulseGlow { 0%,100%{opacity:.4;} 50%{opacity:.9;} }
@keyframes scanLine { 0%{transform:translateY(-100%);} 100%{transform:translateY(2000%);} }
@keyframes splashLogo { 0%{opacity:0;transform:scale(.8) translateY(20px);letter-spacing:.5em;} 55%{opacity:1;letter-spacing:.1em;} 100%{opacity:1;transform:scale(1) translateY(0);letter-spacing:.12em;} }
@keyframes splashSub  { 0%{opacity:0;transform:translateY(14px);} 100%{opacity:.85;transform:none;} }
@keyframes splashOut  { 0%{opacity:1;} 100%{opacity:0;visibility:hidden;} }
@keyframes barGrow { from{width:0;} to{width:120px;} }
@keyframes blueprintIn { from{opacity:0;stroke-dashoffset:1200;} to{opacity:1;stroke-dashoffset:0;} }

/* Malla técnica de fondo que se desplaza */
.arch-grid {
  position: absolute; inset: -60px; pointer-events: none;
  background-image: linear-gradient(rgba(255,255,255,0.04) 1px, transparent 1px),
                    linear-gradient(90deg, rgba(255,255,255,0.04) 1px, transparent 1px);
  background-size: 56px 56px;
  animation: gridPan 9s linear infinite;
}
/* Formas geométricas (planos/contornos) flotando */
.arch-shape { position: absolute; pointer-events: none; will-change: transform; transform: translateZ(0); }
.arch-line  { stroke: rgba(255,255,255,0.10); fill: none; }
.arch-line-accent { stroke: rgba(138,168,138,0.32); fill: none; }
.arch-dash {
  stroke-dasharray: 8 10; animation: dashFlow 24s linear infinite;
}
/* Línea de escaneo tipo plano */
.scan { position:absolute; left:0; right:0; height:1px; background:linear-gradient(90deg, transparent, rgba(138,168,138,0.5), transparent); animation: scanLine 7s linear infinite; }

/* Splash de bienvenida */
.splash-wrap { position: fixed; inset: 0; z-index: 99998; display: flex; flex-direction: column; align-items: center; justify-content: center; background: radial-gradient(circle at 50% 35%, #232320 0%, #1A1A18 60%, #141412 100%); overflow: hidden; }
.splash-wrap.out { animation: splashOut .7s cubic-bezier(.4,0,.2,1) .05s forwards; }
.splash-logo { font-size: 56px; font-weight: 700; color: #F2F1ED; animation: splashLogo 1.1s cubic-bezier(.2,.8,.2,1) both; display:flex; align-items:baseline; justify-content:center; gap:12px; text-shadow: 0 4px 40px rgba(138,168,138,0.30); }
.splash-sub  { font-size: 12px; letter-spacing: .26em; color: rgba(255,255,255,0.6); margin-top: 18px; animation: splashSub .7s ease .9s both; }
.splash-bar  { height: 3px; background: linear-gradient(90deg, transparent, #5A7D5A, #8AA88A, #5A7D5A, transparent); border-radius: 2px; margin-top: 24px; animation: barGrow 1.4s cubic-bezier(.2,.8,.2,1) .4s both, barPulse 2s ease-in-out 1.6s infinite; }

/* ── Splash espectacular ─────────────────────────────────── */
@keyframes barPulse { 0%,100%{box-shadow:0 0 8px rgba(138,168,138,0.45);} 50%{box-shadow:0 0 22px rgba(138,168,138,0.95);} }
@keyframes ringExpand { 0%{transform:scale(0);opacity:0;} 30%{opacity:.8;} 100%{transform:scale(3);opacity:0;} }
@keyframes drawStroke { from{stroke-dashoffset:var(--len);} to{stroke-dashoffset:0;} }
@keyframes particleRise { 0%{transform:translateY(0) scale(1);opacity:0;} 10%{opacity:1;} 90%{opacity:1;} 100%{transform:translateY(-110vh) scale(.4);opacity:0;} }
@keyframes flashOut { 0%{opacity:0;} 92%{opacity:0;} 96%{opacity:.7;} 100%{opacity:0;} }
@keyframes sweepBeam { 0%{transform:translateX(-60vw) rotate(12deg);opacity:0;} 20%{opacity:.5;} 80%{opacity:.5;} 100%{transform:translateX(60vw) rotate(12deg);opacity:0;} }
@keyframes blueprintRotate { 0%{transform:rotate(0) scale(1);} 100%{transform:rotate(360deg) scale(1);} }
@keyframes textGlitch { 0%,100%{opacity:1;} 50%{opacity:.85;} }
@keyframes sp-pop { 0%{opacity:0;transform:scale(.5);} 60%{opacity:1;transform:scale(1.1);} 100%{opacity:1;transform:scale(1);} }

/* Anillos que se expanden detrás del logo */
.spl-ring { position:absolute; border:1.5px solid rgba(138,168,138,0.45); border-radius:50%; width:120px; height:120px; left:50%; top:50%; margin:-60px 0 0 -60px; }
.spl-ring.r1 { animation: ringExpand 3s ease-out infinite; }
.spl-ring.r2 { animation: ringExpand 3s ease-out .8s infinite; }
.spl-ring.r3 { animation: ringExpand 3s ease-out 1.6s infinite; border-color: rgba(255,255,255,0.15); }

/* Plano que se "dibuja" progresivamente */
.spl-blueprint { position:absolute; left:50%; top:50%; transform:translate(-50%,-50%); width:520px; height:520px; opacity:.5; animation: blueprintRotate 80s linear infinite; }
.spl-draw { stroke-dasharray: var(--len); stroke-dashoffset: var(--len); animation: drawStroke 2s cubic-bezier(.4,0,.2,1) forwards; }
.spl-draw.d1 { --len: 1600; animation-delay:.2s; }
.spl-draw.d2 { --len: 1200; animation-delay:.5s; }
.spl-draw.d3 { --len: 900;  animation-delay:.8s; }
.spl-draw.d4 { --len: 600;  animation-delay:1.1s; }

/* Haz de luz que barre */
.spl-beam { position:absolute; top:0; bottom:0; width:120px; background:linear-gradient(90deg, transparent, rgba(138,168,138,0.14), transparent); animation: sweepBeam 3.5s ease-in-out infinite; }

/* Destello final antes de salir */
.spl-flash { position:absolute; inset:0; background:radial-gradient(circle at 50% 45%, rgba(138,168,138,0.5), transparent 60%); animation: flashOut 2.6s ease forwards; pointer-events:none; }

/* Partículas ascendentes */
.spl-particle { position:absolute; width:3px; height:3px; border-radius:50%; background:rgba(138,168,138,0.7); bottom:-10px; will-change:transform; }

/* ── Banner del tablero (sutil) ───────────────────────────── */
.dash-banner { position: relative; overflow: hidden; background: linear-gradient(120deg, #1A1A17 0%, #232320 100%); }
.dash-banner .arch-grid { inset: -40px; background-size: 44px 44px; opacity: .55; will-change: transform; transform: translateZ(0); }
.dash-banner-ring { position:absolute; border:1px solid rgba(138,168,138,0.22); border-radius:50%; pointer-events:none; will-change: transform; transform: translateZ(0); }
.db-r1 { width:160px; height:160px; right:-40px; top:-70px; animation: floatY 9s ease-in-out infinite; }
.db-r2 { width:90px;  height:90px;  right:40px;  bottom:-50px; animation: floatY 7s ease-in-out infinite reverse; }
.dash-beam { position:absolute; top:0; bottom:0; width:80px; background:linear-gradient(90deg, transparent, rgba(138,168,138,0.10), transparent); animation: sweepBeam 6s ease-in-out infinite; will-change: transform; }
.dash-title { animation: riseIn .5s cubic-bezier(.2,.8,.2,1) both; }

/* En móvil: banner estático (sin animación permanente) para evitar lag */
@media (max-width: 760px) {
  .dash-banner .arch-grid, .dash-banner-ring, .dash-beam { animation: none !important; }
}



.hov-nav:hover { background: #ECEAE4 !important; }
.hov-row:hover   { background: #F9F8F5 !important; }
.hov-chip:hover  { background: #ECEAE4 !important; }
.no-scrollbar::-webkit-scrollbar { display: none; height: 0; }
.no-scrollbar { -ms-overflow-style: none; scrollbar-width: none; }
`;

// Detecta móvil: pantalla estrecha Y en vertical. En horizontal usa la interfaz de ordenador.
function useIsMobile() {
  const calc = () => typeof window !== 'undefined' && window.innerWidth < 760;
  const [m, setM] = useState(calc);
  useEffect(() => {
    // resize: respuesta inmediata
    const onResize = () => setM(calc());
    // orientationchange en iOS dispara antes de que cambien las dimensiones
    // esperamos 300ms para leer el ancho real
    const onOrient = () => setTimeout(() => setM(calc()), 300);
    window.addEventListener('resize', onResize);
    window.addEventListener('orientationchange', onOrient);
    return () => {
      window.removeEventListener('resize', onResize);
      window.removeEventListener('orientationchange', onOrient);
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
    padding: sm ? '6px 13px' : '9px 17px', borderRadius: 11, border: '1.5px solid',
    cursor: disabled ? 'not-allowed' : 'pointer', fontSize: 13, fontWeight: 600,
    transition: 'opacity .15s, transform .12s ease', opacity: disabled ? 0.45 : 1, width: full ? '100%' : 'auto',
    letterSpacing: '0.01em',
  };
  const style = primary ? { ...base, background: '#16160F', color: '#fff', borderColor: '#16160F' }
    : danger  ? { ...base, background: '#FDECEC', color: '#8A1F1F', borderColor: '#F9CACA' }
    : ghost   ? { ...base, background: 'transparent', color: '#6B6B66', borderColor: 'transparent' }
    : { ...base, background: 'transparent', color: '#16160F', borderColor: '#E6E4DD' };
  return <button className="tap" style={style} onClick={disabled ? undefined : onClick}>{children}</button>;
}

function ConfirmMini({ titulo, texto, onSi, onNo }) {
  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.25)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 99999, backdropFilter: 'blur(3px)' }}
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

// ── Modal Compartir obra ─────────────────────────────────────────────────────
function ModalCompartir({ obra, user, onClose }) {
  const [email,    setEmail]    = useState('');
  const [enviando, setEnviando] = useState(false);
  const [msg,      setMsg]      = useState('');
  const [usuarios, setUsuarios] = useState([]);
  const [perfiles, setPerfiles] = useState({});
  const [confirm,  setConfirm]  = useState(null);
  const userId = user?.id || user?.sub;
  const esOwner = obra._rol === 'owner';

  useEffect(() => {
    // Cargar accesos y perfiles en paralelo
    Promise.all([
      window.db?.getUsuariosObra(obra.id).catch(() => []),
      window.db?.getPerfiles().catch(() => []),
    ]).then(([accesos, profs]) => {
      setUsuarios(accesos);
      const map = {};
      profs.forEach(p => { map[p.user_id] = p.nombre; });
      setPerfiles(map);
    });
  }, [obra.id]);

  async function invitar() {
    if (!email.trim()) return;
    setEnviando(true); setMsg('');
    try {
      await window.db.invitarUsuario(obra.id, email.trim(), userId);
      setMsg('✓ Usuario añadido correctamente.');
      setEmail('');
      const [accesos, profs] = await Promise.all([
        window.db.getUsuariosObra(obra.id),
        window.db.getPerfiles(),
      ]);
      setUsuarios(accesos);
      const map = {};
      profs.forEach(p => { map[p.user_id] = p.nombre; });
      setPerfiles(map);
    } catch(e) { setMsg('Error: ' + e.message); }
    setEnviando(false);
  }

  async function quitarAcceso(uid) {
    try {
      await window.db.quitarAcceso(obra.id, uid);
      setUsuarios(prev => prev.filter(u => u.user_id !== uid));
      setMsg('✓ Acceso eliminado.');
    } catch(e) { setMsg('Error: ' + e.message); }
    setConfirm(null);
  }

  const nombreUsuario = (uid) => perfiles[uid] || (uid === userId ? 'Tú' : uid.slice(0,8) + '…');

  return (
    <Modal title={`Compartir — ${obra.nombre}`} onClose={onClose}>
      <p style={{ fontSize: 13, color: '#6B6B66', marginBottom: 16, lineHeight: 1.5 }}>
        Añade compañeros por su email. Podrán ver y editar esta obra. Solo el creador puede eliminarla.
      </p>

      {esOwner && (
        <Field label="Email del compañero">
          <div style={{ display: 'flex', gap: 8 }}>
            <input type="email" value={email} onChange={e => setEmail(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && invitar()}
              placeholder="correo@plaat.es" style={{ flex: 1 }} />
            <Btn primary disabled={enviando || !email.trim()} onClick={invitar}>
              {enviando ? '...' : 'Añadir'}
            </Btn>
          </div>
        </Field>
      )}

      {msg && (
        <div style={{ fontSize: 13, padding: '8px 12px', borderRadius: 9, marginBottom: 12,
          background: msg.startsWith('✓') ? '#E8F5E0' : '#FDECEC',
          color:      msg.startsWith('✓') ? '#2D5E10'  : '#8A1F1F' }}>
          {msg}
        </div>
      )}

      {usuarios.length > 0 && (
        <div>
          <div style={{ fontSize: 11, textTransform: 'uppercase', letterSpacing: '0.1em', color: '#A5A5A0', fontWeight: 600, marginBottom: 8 }}>
            Con acceso ({usuarios.length})
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
            {usuarios.map(u => {
              const esYo     = u.user_id === userId;
              const esCreador = u.rol === 'owner';
              const nombre   = esYo ? `${nombreUsuario(u.user_id)} (tú)` : nombreUsuario(u.user_id);
              return (
                <div key={u.user_id} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px', borderRadius: 11, border: '1px solid #ECEAE4', background: '#FAFAF8' }}>
                  {/* Avatar con iniciales */}
                  <div style={{ width: 34, height: 34, borderRadius: '50%', background: esCreador ? '#1A1A17' : '#ECEAE4', color: esCreador ? '#F2F1ED' : '#52524E', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700, flexShrink: 0 }}>
                    {(perfiles[u.user_id] || '?').slice(0,2).toUpperCase()}
                  </div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 13, fontWeight: 600, color: '#16160F', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{nombre}</div>
                    <div style={{ fontSize: 11, color: '#A5A5A0' }}>{esCreador ? '👑 Creador' : 'Editor'}</div>
                  </div>
                  {/* Quitar acceso — solo owner puede, no puede echarse a sí mismo ni al creador */}
                  {esOwner && !esYo && !esCreador && (
                    <button onClick={() => setConfirm(u.user_id)}
                      style={{ background: '#FDECEC', border: '1px solid #F9CACA', borderRadius: 8, padding: '4px 10px', cursor: 'pointer', fontSize: 12, color: '#8A1F1F', fontWeight: 500, flexShrink: 0, whiteSpace: 'nowrap' }}>
                      Quitar acceso
                    </button>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}

      {confirm && (
        <ConfirmMini
          titulo="Quitar acceso"
          texto={`Vas a quitar el acceso de ${nombreUsuario(confirm)} a esta obra. Podrás volver a invitarle en cualquier momento.`}
          onSi={() => quitarAcceso(confirm)}
          onNo={() => setConfirm(null)}
        />
      )}
    </Modal>
  );
}

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

// ─── Menú de perfil (⋯ junto al usuario) ────────────────────────────────────
function MenuPerfil({ onBackup, onSalir }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);

  useEffect(() => {
    function handleClick(e) { if (ref.current && !ref.current.contains(e.target)) setOpen(false); }
    if (open) document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, [open]);

  return (
    <div ref={ref} style={{ position: 'relative', flexShrink: 0 }}>
      <button onClick={() => setOpen(v => !v)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'rgba(255,255,255,0.35)', fontSize: 16, padding: '4px 2px', lineHeight: 1, letterSpacing: '0.1em', display: 'flex', alignItems: 'center' }}
        title="Opciones">
        ···
      </button>
      {open && (
        <div className="fade" style={{ position: 'absolute', bottom: 'calc(100% + 6px)', right: 0, background: '#2A2A27', border: '1px solid rgba(255,255,255,0.1)', borderRadius: 10, boxShadow: '0 8px 24px rgba(0,0,0,0.4)', minWidth: 180, zIndex: 200, overflow: 'hidden' }}>
          <button onClick={() => { setOpen(false); onBackup(); }}
            style={{ width: '100%', padding: '10px 14px', background: 'none', border: 'none', cursor: 'pointer', fontSize: 13, color: 'rgba(255,255,255,0.75)', textAlign: 'left', display: 'flex', alignItems: 'center', gap: 9 }}>
            <span>💾</span> Copia de seguridad
          </button>
          <div style={{ height: 1, background: 'rgba(255,255,255,0.07)', margin: '0 10px' }} />
          <button onClick={() => { setOpen(false); onSalir(); }}
            style={{ width: '100%', padding: '10px 14px', background: 'none', border: 'none', cursor: 'pointer', fontSize: 13, color: 'rgba(255,100,100,0.8)', textAlign: 'left', display: 'flex', alignItems: 'center', gap: 9 }}>
            <span>→</span> Cerrar sesión
          </button>
        </div>
      )}
    </div>
  );
}

function Sidebar({ nav, setNav, stats, user, onBackup }) {
  const navItems = [
    { id: 'alertas',      label: 'Alertas',      badge: stats.alertas, alert: stats.alertas > 0 },
    { id: 'tablero',      label: 'Tablero',       badge: stats.total, alert: false },
    { id: 'seguimiento',  label: 'Seguimiento',   badge: 0, alert: false },
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

        {/* Fin nav items */}
      </nav>

      {/* Usuario */}
      <div style={{ padding: '12px 14px', borderTop: '1px solid rgba(255,255,255,0.06)' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: email ? 8 : 0 }}>
          <div style={{ width: 32, height: 32, borderRadius: '50%', background: 'rgba(255,255,255,0.1)', color: '#F2F1ED', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 500, flexShrink: 0, letterSpacing: '0.02em' }}>{iniciales}</div>
          <div style={{ minWidth: 0, flex: 1 }}>
            <div style={{ fontSize: 12, fontWeight: 500, color: '#F2F1ED', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{email || 'PLAAT'}</div>
            <div style={{ fontSize: 11, color: 'rgba(255,255,255,0.35)' }}>Arquitectura Técnica</div>
          </div>
          <MenuPerfil onBackup={onBackup} onSalir={salir} />
        </div>
      </div>
    </div>
  );
}

// ─── Barra inferior (móvil) ───────────────────────────────────────────────────

function BottomNav({ nav, setNav, stats }) {
  const ICONS = {
    alertas:     'M10 4a3 3 0 0 0-3 3c0 4-1.5 5-2 6h10c-.5-1-2-2-2-6a3 3 0 0 0-3-3Z M8.5 16a1.5 1.5 0 0 0 3 0',
    tablero:     'M3 17V8l5-3 5 3v9 M7 17v-4h2v4',
    seguimiento: 'M4 6h12 M4 10h12 M4 14h7 M14 13l2 2 3-3',
    salir:       'M7 4H4v12h3 M10 10h7 M14 7l3 3-3 3',
  };
  const items = [
    { id: 'alertas',     label: 'Alertas',     badge: stats.alertas, alert: true },
    { id: 'tablero',     label: 'Obras',        badge: stats.total,   alert: false },
    { id: 'seguimiento', label: 'Seguimiento',  badge: 0,             alert: false },
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
            <span style={{ fontSize: 10.5, fontWeight: activo ? 500 : 400, letterSpacing: '0.02em' }}>{it.label}</span>
            {it.badge > 0 && (
              <span style={{ position: 'absolute', top: 6, left: '50%', marginLeft: 7, background: it.alert ? '#E24B4A' : 'rgba(255,255,255,0.2)', color: '#fff', fontSize: 9, minWidth: 16, height: 16, borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '0 4px', fontWeight: 600 }}>{it.badge}</span>
            )}
          </button>
        );
      })}
      <button onClick={salir} style={{ flex: 1, background: 'none', border: 'none', cursor: 'pointer', padding: '10px 0 12px', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 4, color: 'rgba(255,255,255,0.4)' }}>
        <Icono d={ICONS.salir} activo={false} />
        <span style={{ fontSize: 10.5, letterSpacing: '0.02em' }}>Salir</span>
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
      style={{ background: '#fff', borderRadius: 16, boxShadow: '0 1px 2px rgba(0,0,0,0.04), 0 4px 16px rgba(0,0,0,0.045)', overflow: 'hidden', borderLeft: `3px solid ${accentColor}`, position: 'relative' }}>

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
                      {obra._rol === 'owner' && <div onClick={ev => { ev.stopPropagation(); setMenu(false); onEliminar(obra); }} style={{ padding: '7px 11px', borderRadius: 6, cursor: 'pointer', fontSize: 13, color: '#8A1F1F' }} className="hov-row">Eliminar</div>}
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
      {/* Fecha CFO — sota el footer */}
      {obra.fechaCFO && (
        <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '6px 16px', borderTop: '1px solid #F2F1ED', background: '#F9F8F5' }}>
          <span style={{ fontSize: 10.5, color: '#A5A5A0', fontWeight: 500, letterSpacing: '0.06em', textTransform: 'uppercase' }}>CFO</span>
          <span style={{ fontSize: 11.5, fontWeight: 500, color: '#52524E' }}>{fmtDate(obra.fechaCFO)}</span>
        </div>
      )}
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
    ? { nombre: obra.nombre || '', cliente: obra.cliente || '', direccion: obra.direccion || '', responsable: obra.responsable || RESPONSABLES[0], diasVisita: obra.diasVisita || [], emplazamiento: obra.emplazamiento || '', propiedad: obra.propiedad || '', proyectista: obra.proyectista || '', direccionObra: obra.direccionObra || '', constructora: obra.constructora || '', deoFirmante: obra.deoFirmante || '', fechaCFO: obra.fechaCFO || '' }
    : { nombre: '', cliente: '', direccion: '', responsable: RESPONSABLES[0], diasVisita: [], emplazamiento: '', propiedad: '', proyectista: '', direccionObra: '', constructora: '', deoFirmante: '', fechaCFO: '' });
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
        <Field label="Fecha CFO (Certificat Final d'Obra)"><input type="date" value={form.fechaCFO} onChange={e => upd('fechaCFO', e.target.value)} /></Field>
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

// ─── Actas de inspección genéricas (historial por obra) ─────────────────────
const ESTADOS_INSP_ACTA = {
  ok:         { label: 'Correcto',    bg: '#E8F5E0', color: '#2D5E10' },
  incidencia: { label: 'Incidencia',  bg: '#FDECEC', color: '#8A1F1F' },
  info:       { label: 'Informativo', bg: '#EEEDE7', color: '#52524E' },
};

function FormActaInspeccion({ obra, acta, onGuardar, onCerrar, onExportar }) {
  const isMobile = useIsMobile();
  const nextNum = acta ? acta.num : ((obra.actasInsp || []).length + 1);
  const [a, setA] = useState(() => acta || {
    id: uid(), num: nextNum, fecha: today(), temas: [], creadaEn: now(),
  });
  const [confirmacion, setConfirmacion] = useState(null);

  function upd(k, v) { setA(prev => ({ ...prev, [k]: v })); }

  function addTema() {
    setA(prev => ({ ...prev, temas: [...prev.temas, {
      id: uid(), num: `N${prev.temas.length + 1}`, titulo: '', descripcion: '', estado: 'incidencia', fotos: [],
    }]}));
  }
  function updTema(id, k, v) {
    setA(prev => ({ ...prev, temas: prev.temas.map(t => t.id === id ? { ...t, [k]: v } : t) }));
  }
  function delTema(id) { setA(prev => ({ ...prev, temas: prev.temas.filter(t => t.id !== id) })); setConfirmacion(null); }
  function addFoto(temaId) {
    pickFiles('image/*', f => setA(prev => ({ ...prev, temas: prev.temas.map(t => t.id === temaId ? { ...t, fotos: [...(t.fotos||[]), f] } : t) })), obra?.id);
  }
  function delFoto(temaId, fotoId) {
    setA(prev => ({ ...prev, temas: prev.temas.map(t => t.id === temaId ? { ...t, fotos: (t.fotos||[]).filter(f => f.id !== fotoId) } : t) }));
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', background: '#FFFFFF', borderRadius: 16, overflow: 'hidden', boxShadow: '0 1px 3px rgba(0,0,0,0.06)' }}>
      {/* Cabecera */}
      <div style={{ background: '#1A1A17', padding: '16px 20px', display: 'flex', alignItems: 'center', gap: 12, flexShrink: 0 }}>
        <button onClick={onCerrar} style={{ background: 'none', border: 'none', color: 'rgba(255,255,255,0.6)', fontSize: 15, cursor: 'pointer', padding: 0, flexShrink: 0 }}>←</button>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 15, fontWeight: 700, color: '#F2F1ED', display: 'flex', alignItems: 'baseline', gap: 6 }}>
            Acta de inspección <span style={{ color: '#8AA88A' }}>Nº {String(a.num).padStart(2,'0')}</span>
          </div>
          <div style={{ fontSize: 11, color: 'rgba(255,255,255,0.4)', marginTop: 2 }}>{obra.nombre}</div>
        </div>
        <button onClick={() => onGuardar(a)} style={{ padding: '8px 18px', borderRadius: 11, border: 'none', background: '#5A7D5A', color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Guardar</button>
      </div>

      {/* Cuerpo */}
      <div style={{ flex: 1, overflowY: 'auto', padding: '18px 20px', display: 'flex', flexDirection: 'column', gap: 16 }}>

        {/* Nº y fecha */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
          <Field label="Nº de acta">
            <input type="number" min="1" value={a.num} onChange={e => upd('num', Math.max(1, parseInt(e.target.value||'1',10)))} />
          </Field>
          <Field label="Fecha de inspección">
            <input type="date" value={a.fecha} onChange={e => upd('fecha', e.target.value)} />
          </Field>
        </div>

        {/* Tabla resumen de temas tratados */}
        <div style={{ background: '#F7F6F3', borderRadius: 12, border: '1px solid #E6E4DD', overflow: 'hidden' }}>
          <div style={{ background: '#1A1A17', padding: '10px 14px', display: 'flex', alignItems: 'center', gap: 8 }}>
            <span style={{ fontSize: 12, fontWeight: 700, color: '#F2F1ED', letterSpacing: '0.05em' }}>TEMAS TRATADOS</span>
            <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.4)', marginLeft: 'auto' }}>{a.temas.length} tema{a.temas.length !== 1 ? 's' : ''}</span>
          </div>
          {a.temas.length === 0 && <div style={{ padding: '14px', fontSize: 13, color: '#A5A5A0', textAlign: 'center' }}>Sin temas todavía — añade uno abajo</div>}
          {a.temas.map((t, i) => (
            <div key={t.id} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 14px', borderTop: i > 0 ? '1px solid #E6E4DD' : 'none', background: '#fff' }}>
              <span style={{ width: 32, height: 32, borderRadius: 8, background: '#1A1A17', color: '#F2F1ED', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700, flexShrink: 0 }}>{t.num}</span>
              <input value={t.titulo} onChange={e => updTema(t.id,'titulo',e.target.value)} placeholder="Título del tema..." style={{ flex: 1, fontWeight: 600, fontSize: 13 }} />
              <select value={t.estado} onChange={e => updTema(t.id,'estado',e.target.value)} style={{ fontSize: 11, padding: '4px 8px', borderRadius: 20, border: `1px solid ${ESTADOS_INSP_ACTA[t.estado]?.color}30`, background: ESTADOS_INSP_ACTA[t.estado]?.bg, color: ESTADOS_INSP_ACTA[t.estado]?.color, fontWeight: 600, width: 'auto' }}>
                {Object.entries(ESTADOS_INSP_ACTA).map(([k,v]) => <option key={k} value={k}>{v.label}</option>)}
              </select>
              <button onClick={() => setConfirmacion({ titulo: 'Eliminar tema', texto: `Vas a eliminar "${t.titulo||'este tema'}".`, onSi: () => delTema(t.id) })} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#C4C3BE', fontSize: 18, lineHeight: 1, flexShrink: 0 }}>×</button>
            </div>
          ))}
          <div style={{ padding: '10px 14px', borderTop: a.temas.length > 0 ? '1px solid #E6E4DD' : 'none' }}>
            <button onClick={addTema} style={{ width: '100%', padding: '9px', borderRadius: 9, border: '1.5px dashed #E0DFD9', background: 'transparent', cursor: 'pointer', fontSize: 13, color: '#6B6B66', fontWeight: 500 }}>+ Añadir tema</button>
          </div>
        </div>

        {/* Desarrollo de cada tema */}
        {a.temas.map((t, i) => (
          <div key={t.id} style={{ background: '#fff', borderRadius: 14, border: '1px solid #E6E4DD', overflow: 'hidden' }}>
            <div style={{ background: '#F0EFE9', padding: '10px 14px', display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ width: 28, height: 28, borderRadius: 7, background: '#1A1A17', color: '#F2F1ED', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700, flexShrink: 0 }}>{t.num}</span>
              <span style={{ fontSize: 13, fontWeight: 700, color: '#16160F', flex: 1 }}>{t.titulo || 'Sin título'}</span>
              <span style={{ fontSize: 11, padding: '3px 10px', borderRadius: 20, background: ESTADOS_INSP_ACTA[t.estado]?.bg, color: ESTADOS_INSP_ACTA[t.estado]?.color, fontWeight: 600 }}>{ESTADOS_INSP_ACTA[t.estado]?.label}</span>
            </div>
            <div style={{ padding: '14px' }}>
              <Field label="Descripción detallada">
                <textarea value={t.descripcion} onChange={e => updTema(t.id,'descripcion',e.target.value)} placeholder="Describe la incidencia, qué se observa, qué se solicita a la EC..." style={{ minHeight: 90 }} />
              </Field>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginTop: 4 }}>
                {(t.fotos||[]).map(ft => (
                  <div key={ft.id} style={{ position: 'relative', width: 80, height: 60 }}>
                    <img src={fotoSrc(ft)} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover', borderRadius: 9, display: 'block' }} />
                    <button onClick={() => delFoto(t.id, ft.id)} style={{ position: 'absolute', top: 3, right: 3, background: 'rgba(0,0,0,0.55)', color: '#fff', border: 'none', borderRadius: '50%', width: 18, height: 18, cursor: 'pointer', fontSize: 11, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>×</button>
                  </div>
                ))}
                <button onClick={() => addFoto(t.id)} style={{ width: 80, height: 60, borderRadius: 9, border: '1.5px dashed #E0DFD9', background: '#FAFAF8', cursor: 'pointer', fontSize: 11, color: '#9B9B97' }}>+ foto</button>
              </div>
            </div>
          </div>
        ))}
      </div>

      {/* Pie */}
      <div style={{ borderTop: '1px solid #ECEAE4', padding: '14px 20px', display: 'flex', gap: 10, flexShrink: 0, background: '#fff' }}>
        <Btn onClick={onCerrar} full>Cerrar</Btn>
        <Btn primary full onClick={() => { onGuardar(a); onExportar(a); }}>↓ Exportar PDF</Btn>
      </div>
      {confirmacion && <ConfirmMini titulo={confirmacion.titulo} texto={confirmacion.texto} onSi={confirmacion.onSi} onNo={() => setConfirmacion(null)} />}
    </div>
  );
}

// ─── Generador PDF Acta Inspección ────────────────────────────────────────────
async function generarActaInspeccion(obra, acta) {
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
  const LW = 0.25;
  const GRIS_CAB = [230, 230, 225];
  const GRIS_TEMA = [217, 217, 212];
  const num = String(acta.num).padStart(2, '0');
  const fechaStr = acta.fecha ? new Date(acta.fecha).toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '';
  const temas = acta.temas || [];

  // ── Utilidades ────────────────────────────────────────────────────────────
  function sl() { doc.setLineWidth(LW); doc.setDrawColor(0, 0, 0); }

  function cabecera() {
    // Textos cabecera izquierda
    doc.setFont('helvetica', 'bold'); doc.setFontSize(7.5); doc.setTextColor(0, 0, 0);
    doc.text(`PARTE DE INSPECCIÓN Nº: ${num}`, M, 8);
    doc.setFont('helvetica', 'normal');
    doc.text(`FECHA PARTE INSPECCIÓN: ${fechaStr}`, M, 12);
    // Logo derecha — solo texto, sin caja ni fondo
    doc.setFont('helvetica', 'bold'); doc.setFontSize(22); doc.setTextColor(0, 0, 0);
    doc.text('Plaat.', PW - M, 12, { align: 'right' });
    // Línea separadora
    sl(); doc.line(M, 15, PW - M, 15);
  }

  function pie() {
    sl(); doc.line(M, PH - 10, PW - M, PH - 10);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(7); doc.setTextColor(80, 80, 80);
    doc.text('Plaat Arquitectura Técnica', M, PH - 6);
    doc.text('Barcelona \u2013 Madrid', PW / 2, PH - 6, { align: 'center' });
    doc.text('www.plaat.es', PW - M, PH - 6, { align: 'right' });
  }

  function calcH(txt, w, sz, minH) {
    doc.setFontSize(sz || 8.5);
    const lines = doc.splitTextToSize(String(txt || ''), w - 4);
    return Math.max(minH || 7, lines.length * ((sz || 8.5) * 0.3528 + 0.8) + 4);
  }

  function celda(x, y, w, h, txt, opts = {}) {
    const sz = opts.sz || 8.5;
    const bold = opts.bold || false;
    const fill = opts.fill;
    sl();
    if (fill) { doc.setFillColor(...fill); doc.rect(x, y, w, h, 'FD'); }
    else doc.rect(x, y, w, h);
    if (!txt && txt !== 0) return;
    doc.setFont('helvetica', bold ? 'bold' : 'normal');
    doc.setFontSize(sz); doc.setTextColor(0, 0, 0);
    const lines = doc.splitTextToSize(String(txt), w - 4);
    const lineH = sz * 0.3528 + 0.8;
    const totalH = lines.length * lineH;
    let ty = y + (h - totalH) / 2 + sz * 0.3528;
    lines.forEach(l => { doc.text(l, x + 2, ty); ty += lineH; });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // PÁGINA 1 — Portada + datos de obra
  // ══════════════════════════════════════════════════════════════════════════
  cabecera(); pie();
  let y = 20;

  // Banda título
  doc.setFillColor(...GRIS_TEMA); sl();
  doc.rect(M, y, CW - 32, 9, 'FD');
  doc.rect(M + CW - 32, y, 32, 9, 'FD');
  doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(0, 0, 0);
  doc.text('ACTA DE INSPECCIÓN DE OBRA', M + 3, y + 6.2);
  doc.text(`NÚM.: ${num}`, M + CW - 30, y + 6.2);
  y += 9;

  // OBRA + EMPLAZAMIENTO
  [['OBRA', obra.nombre || ''], ['EMPLAZAMIENTO', obra.emplazamiento || obra.direccion || '']].forEach(([k, v]) => {
    const h = calcH(v, CW - 36, 8.5, 8);
    celda(M, y, 36, h, k, { bold: true, fill: GRIS_CAB });
    celda(M + 36, y, CW - 36, h, v);
    y += h;
  });
  y += 6;

  // Datos de la obra
  doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(0, 0, 0);
  doc.text('DATOS DE LA OBRA:', M, y); y += 5;
  [
    ['PROPIEDAD',    obra.propiedad || ''],
    ['PROYECTISTA',  obra.proyectista || ''],
    ['DO',           obra.direccionObra || ''],
    ['DEO',          'PLAAT ARQUITECTURA TÉCNICA S.L.'],
    ['CONSTRUCTORA', obra.constructora || ''],
  ].forEach(([k, v]) => {
    const h = calcH(v, CW - 36, 8.5, 8);
    celda(M, y, 36, h, k, { bold: true, fill: GRIS_CAB });
    celda(M + 36, y, CW - 36, h, v);
    y += h;
  });

  // Fecha firma
  y += 8;
  doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(0, 0, 0);
  doc.text('FECHA FIRMA:', M, y); y += 5;
  const soloFecha = acta.fecha ? new Date(acta.fecha).toLocaleDateString('es-ES', { day: 'numeric', month: 'long', year: 'numeric' }) : '';
  doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5);
  doc.text(soloFecha, M, y); y += 14;

  // Firma portada
  doc.setFont('helvetica', 'bold'); doc.setFontSize(8.5); doc.setTextColor(0, 0, 0);
  doc.text('DIRECTOR DE EJECUCIÓN DE OBRA', M, y); y += 5;
  if (obra.deoFirmante) { doc.setFont('helvetica', 'normal'); doc.text(obra.deoFirmante, M, y); }

  // ══════════════════════════════════════════════════════════════════════════
  // PÁGINA 2 — Aspecto revisado + tabla temas + intro incidencias
  // ══════════════════════════════════════════════════════════════════════════
  doc.addPage(); cabecera(); pie();
  y = 20;

  // Aspecto revisado
  doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(0, 0, 0);
  doc.text('Aspecto revisado:', M, y); y += 5;
  doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5);
  const introAsp = 'Durante la visita realizada, se han revisado los siguientes aspectos:';
  doc.text(doc.splitTextToSize(introAsp, CW), M, y); y += 7;

  // Tabla temas tratados — cabecera
  const COL_N = 18, COL_T = CW - COL_N;
  // Cabecera tabla Pg2 — N.º centrado
  doc.setFillColor(...GRIS_CAB); doc.setDrawColor(0,0,0); doc.setLineWidth(LW);
  doc.rect(M, y, COL_N, 7, 'FD');
  doc.setFont('helvetica','bold'); doc.setFontSize(8.5); doc.setTextColor(0,0,0);
  const nwCab = doc.getTextWidth('N.º');
  doc.text('N.º', M + (COL_N - nwCab) / 2, y + 5);
  celda(M + COL_N, y, COL_T, 7, 'TEMAS TRATADOS', { bold: true, fill: GRIS_CAB });
  y += 7;

  temas.forEach(t => {
    const h = calcH(t.titulo, COL_T, 8.5, 9);
    // Nº centrado horizontal y vertical
    doc.setFillColor(255,255,255); doc.setDrawColor(0,0,0); doc.setLineWidth(LW);
    doc.rect(M, y, COL_N, h);
    doc.setFont('helvetica', 'bold'); doc.setFontSize(8.5); doc.setTextColor(0, 0, 0);
    const nw = doc.getTextWidth(t.num || '');
    doc.text(t.num || '', M + (COL_N - nw) / 2, y + h / 2 + 1.5);
    celda(M + COL_N, y, COL_T, h, t.titulo, { bold: true });
    y += h;
  });
  y += 8;

  // Intro incidencias
  doc.setFont('helvetica', 'bold'); doc.setFontSize(8.5); doc.setTextColor(0, 0, 0);
  doc.text('Incidencias detectadas:', M, y); y += 5;
  doc.setFont('helvetica', 'normal');
  [
    'Las incidencias detectadas, deberán subsanarse y dar respuesta a la DEO.',
    'Para ello deberán rellenar los datos solicitados en cada una de las incidencias detectadas y enviar fotografías a la DEO con las rectificaciones.',
    'Se adjuntan tablas con las incidencias detectadas.',
  ].forEach(txt => {
    const ll = doc.splitTextToSize(txt, CW);
    doc.text(ll, M, y); y += ll.length * 4.6 + 2;
  });
  y += 4;
  const txtBase = 'Este Acta de Inspección de obra se ha llevado a cabo en base a las inspecciones y muestreos realizados por el DEO en la fecha indicada y en base a Partes de Inspección procedimentados.';
  const llBase = doc.splitTextToSize(txtBase, CW);
  doc.text(llBase, M, y);

  // ══════════════════════════════════════════════════════════════════════════
  // PÁGINA 3+ — Zonas revisadas (desarrollo por tema)
  // ══════════════════════════════════════════════════════════════════════════
  if (temas.length > 0) {
    doc.addPage(); cabecera(); pie();
    y = 20;

    doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(0, 0, 0);
    doc.text('1. ZONAS REVISADAS', M, y); y += 8;

    // Dimensiones estándar fijas
    const FOTO_W = (CW - 10) / 2;
    const FOTO_H = 58;
    const FOTO_GAP = 6;
    const PAD = 4;
    const lineHTit = 9 * 0.3528 + 0.8;
    const lineHDesc = 8.5 * 0.3528 + 0.8;

    temas.forEach((t, tIdx) => {
      // 2 temas por página — salto cada 2
      if (tIdx > 0 && tIdx % 2 === 0) {
        doc.addPage(); cabecera(); pie(); y = 20;
      }

      const fotos = t.fotos || [];
      const filasFoto = Math.ceil(fotos.length / 2);

      doc.setFont('helvetica', 'bold'); doc.setFontSize(9);
      const titLines = doc.splitTextToSize(t.titulo || '', COL_T - 6);
      const hTit = Math.max(9, titLines.length * lineHTit + 4);

      doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5);
      const descLines = t.descripcion ? doc.splitTextToSize(t.descripcion, CW - PAD * 2) : [];
      const hDesc = descLines.length > 0 ? descLines.length * lineHDesc + PAD * 2 : PAD * 2;
      const hFotos = filasFoto > 0 ? filasFoto * (FOTO_H + PAD) + PAD : 0;
      const hCuerpo = hDesc + hFotos;

      // ── Cabecera gris ────────────────────────────────────────────────────
      doc.setFillColor(230, 230, 225); doc.setDrawColor(0, 0, 0); doc.setLineWidth(LW);
      doc.rect(M, y, COL_N, hTit, 'FD');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(0, 0, 0);
      const nw = doc.getTextWidth(t.num || '');
      doc.text(t.num || '', M + (COL_N - nw) / 2, y + hTit / 2 + 1.8);

      doc.setFillColor(230, 230, 225); doc.setDrawColor(0, 0, 0); doc.setLineWidth(LW);
      doc.rect(M + COL_N, y, COL_T, hTit, 'FD');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(0, 0, 0);
      let tyTit = y + (hTit - titLines.length * lineHTit) / 2 + 9 * 0.3528;
      titLines.forEach(l => { doc.text(l, M + COL_N + 3, tyTit); tyTit += lineHTit; });
      y += hTit;

      // ── Cuerpo: UN solo rectángulo con texto + fotos ──────────────────────
      doc.setFillColor(255, 255, 255); doc.setDrawColor(0, 0, 0); doc.setLineWidth(LW);
      doc.rect(M, y, CW, hCuerpo);
      let cy = y + PAD;

      // Texto justificado
      if (descLines.length > 0) {
        doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(0, 0, 0);
        descLines.forEach((line, idx) => {
          const isLast = idx === descLines.length - 1;
          const words = line.split(' ');
          if (!isLast && words.length > 1) {
            const wordsW = words.reduce((s, w) => s + doc.getTextWidth(w), 0);
            const spaceW = (CW - PAD * 2 - wordsW) / (words.length - 1);
            let wx = M + PAD;
            words.forEach(w => { doc.text(w, wx, cy + 8.5 * 0.3528); wx += doc.getTextWidth(w) + spaceW; });
          } else {
            doc.text(line, M + PAD, cy + 8.5 * 0.3528);
          }
          cy += lineHDesc;
        });
        cy += PAD;
      }

      // Fotos centradas sin marcos propios
      for (let fi = 0; fi < fotos.length; fi += 2) {
        const pair = [fotos[fi], fotos[fi + 1]].filter(Boolean);
        const totalW = pair.length === 2 ? FOTO_W * 2 + FOTO_GAP : FOTO_W;
        const startX = M + (CW - totalW) / 2;
        pair.forEach((f, pi) => {
          try {
            const pr = doc.getImageProperties(f.url||f.data);
            const ratio = pr.width / pr.height;
            let iw = FOTO_W, ih = FOTO_H;
            if (ratio > FOTO_W / FOTO_H) ih = FOTO_W / ratio;
            else iw = FOTO_H * ratio;
            const fx = startX + pi * (FOTO_W + FOTO_GAP);
            doc.addImage(f.data, 'JPEG', fx + (FOTO_W - iw) / 2, cy + (FOTO_H - ih) / 2, iw, ih);
          } catch (e) {}
        });
        cy += FOTO_H + PAD;
      }

      y += hCuerpo + 6;
    });

    // ── Tabla firma — SIEMPRE al final de todo ─────────────────────────────
    // Forzar nueva página si no cabe
    if (y + 50 > PH - 16) { doc.addPage(); cabecera(); pie(); y = 20; }
    y += 4;

    // Nota contratista
    doc.setFont('helvetica', 'italic'); doc.setFontSize(8); doc.setTextColor(0, 0, 0);
    const notaTxt = 'NOTA: El contratista ha sido informado por la DF y se compromete a la ejecución de las medidas correctoras necesarias.';
    const notaLL = doc.splitTextToSize(notaTxt, CW);
    doc.text(notaLL, M, y); y += notaLL.length * 4.2 + 6;

    // Tabla rectificación
    const colsRect = [['Rectificado Día', 40], ['Responsable \u2013 Nombre', 80], ['Firma', CW - 120]];
    let cx = M;
    colsRect.forEach(([label, w]) => {
      celda(cx, y, w, 7, label, { bold: true, fill: GRIS_CAB });
      cx += w;
    });
    y += 7;
    cx = M;
    colsRect.forEach(([, w]) => { sl(); doc.rect(cx, y, w, 18); cx += w; });
    y += 24;

    // Firma DEO
    y += 6;
    doc.setFont('helvetica', 'bold'); doc.setFontSize(8.5); doc.setTextColor(0, 0, 0);
    doc.text('DIRECTOR DE EJECUCIÓN DE OBRA', M, y); y += 5;
    if (obra.deoFirmante) { doc.setFont('helvetica', 'normal'); doc.text(obra.deoFirmante, M, y); }
  }

  // Pie en todas las páginas
  const total = doc.getNumberOfPages();
  for (let p = 1; p <= total; p++) { doc.setPage(p); pie(); }

  // iOS Safari no permite descargas programáticas — abrir en nueva pestaña
  const pdfBlob = doc.output('blob');
  const url = URL.createObjectURL(pdfBlob);
  window.open(url, '_blank');
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}


function ModuloInspecciones({ obra, onSave }) {
  const isMobile = useIsMobile();
  const [disciplinaActiva,    setDisciplinaActiva]    = useState(null);
  const [showNuevaDisciplina, setShowNuevaDisciplina] = useState(false);
  const [showNuevoPunto,      setShowNuevoPunto]      = useState(false);
  const [nombreDisciplina,    setNombreDisciplina]    = useState('');
  const [nombrePunto,         setNombrePunto]         = useState('');
  const [confirmacion,        setConfirmacion]         = useState(null); // {titulo,texto,onSi}
  const [menuDisciplina,      setMenuDisciplina]       = useState(false);
  const [actaActiva,          setActaActiva]           = useState(null); // acta abierta en el formulario (objeto o 'nueva')
  const [generandoActa,       setGenerandoActa]        = useState(false);

  const actas = obra.actasInsp || [];

  async function guardarActa(acta) {
    const existe = actas.some(x => x.id === acta.id);
    const nuevas = existe ? actas.map(x => x.id === acta.id ? acta : x) : [acta, ...actas];
    // Guardar en Supabase directamente
    try {
      await window.db.upsertModulo('actas_insp', { id: acta.id, obra_id: obra.id, data: acta, updated_at: now() });
    } catch(e) { console.error('Error guardando acta:', e); }
    onSave({ ...obra, actasInsp: nuevas });
    setActaActiva(acta);
  }

  async function eliminarActa(id) {
    // Eliminar de Supabase directamente
    try {
      await window.db.deleteModulo('actas_insp', id);
    } catch(e) { console.error('Error eliminando acta:', e); }
    onSave({ ...obra, actasInsp: actas.filter(x => x.id !== id) });
    setConfirmacion(null);
  }
  async function exportarActa(acta) {
    setGenerandoActa(true);
    try { await generarActaInspeccion(obra, acta); }
    catch (e) { alert('Error al exportar: ' + e.message); }
    setGenerandoActa(false);
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

  // Si hay un acta abierta, el formulario reemplaza toda la vista (como Acta VO)
  if (actaActiva) {
    return (
      <>
        <FormActaInspeccion
          obra={obra}
          acta={actaActiva === 'nueva' ? null : actaActiva}
          onGuardar={guardarActa}
          onCerrar={() => setActaActiva(null)}
          onExportar={exportarActa}
        />
        {confirmacion && <ConfirmMini titulo={confirmacion.titulo} texto={confirmacion.texto} onSi={confirmacion.onSi} onNo={() => setConfirmacion(null)} />}
      </>
    );
  }

  return (
    <>
      {/* ── Actas de inspección ── */}
      <div style={{ background: '#fff', borderRadius: 16, boxShadow: '0 1px 2px rgba(0,0,0,0.04), 0 3px 12px rgba(0,0,0,0.04)', padding: isMobile ? '14px 15px' : '16px 18px', marginBottom: 12 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: actas.length ? 12 : 0 }}>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 13, fontWeight: 600, color: '#16160F' }}>Actas de inspección</div>
            <div style={{ fontSize: 11.5, color: '#9B9B97', marginTop: 1 }}>{actas.length ? `${actas.length} acta${actas.length !== 1 ? 's' : ''} registrada${actas.length !== 1 ? 's' : ''}` : 'Aún no hay actas'}</div>
          </div>
          <button onClick={() => setActaActiva('nueva')} className="tap" style={{ display: 'inline-flex', alignItems: 'center', gap: 5, padding: '8px 15px', borderRadius: 11, border: '1.5px solid #5A7D5A', background: '#5A7D5A', color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer', whiteSpace: 'nowrap' }}>
            + Crear acta de inspección
          </button>
        </div>
        {actas.length > 0 && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
            {actas.map(act => {
              const nInc = (act.temas || []).filter(t => t.estado === 'incidencia').length;
              return (
                <div key={act.id} className="hov-row" style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px', borderRadius: 11, border: '1px solid #ECEAE4', cursor: 'pointer' }}
                  onClick={() => setActaActiva(act)}>
                  <span style={{ width: 30, height: 30, borderRadius: 9, background: '#1A1A17', color: '#F2F1ED', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 600, flexShrink: 0 }}>{String(act.num).padStart(2, '0')}</span>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 13, fontWeight: 500, color: '#16160F' }}>Acta Nº {String(act.num).padStart(2, '0')}</div>
                    <div style={{ fontSize: 11.5, color: '#9B9B97' }}>{fmtDate(act.fecha)} · {(act.temas || []).length} incidencia{(act.temas || []).length !== 1 ? 's' : ''}{nInc > 0 ? ` · ${nInc} abierta${nInc !== 1 ? 's' : ''}` : ''}</div>
                  </div>
                  <button onClick={e => { e.stopPropagation(); exportarActa(act); }} disabled={generandoActa} title="Exportar PDF" style={{ background: 'none', border: '1px solid #E6E4DD', borderRadius: 9, padding: '5px 10px', cursor: 'pointer', fontSize: 12, color: '#6B6B66', flexShrink: 0 }}>↓ PDF</button>
                  <button onClick={e => { e.stopPropagation(); setConfirmacion({ titulo: 'Eliminar acta', texto: `Vas a eliminar el Acta Nº ${String(act.num).padStart(2, '0')}. Esta acción no se puede deshacer.`, onSi: () => eliminarActa(act.id) }); }} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#C4C3BE', fontSize: 18, padding: '0 2px', lineHeight: 1, flexShrink: 0 }}>×</button>
                </div>
              );
            })}
          </div>
        )}
      </div>

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
                const est = ESTADOS_INSP[p.estado] || ESTADOS_INSP.pendiente;
                return (
                  <div key={p.id} style={{ border: '1px solid #E8E7E1', borderRadius: 11, overflow: 'hidden' }}>
                    <div style={{ display: 'flex', flexDirection: isMobile ? 'column' : 'row', alignItems: isMobile ? 'stretch' : 'center', gap: isMobile ? 8 : 10, padding: '10px 12px', background: '#fff' }}>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: 13, fontWeight: 400, color: '#18180F' }}>{p.nombre}</div>
                        {p.fecha && <div style={{ fontSize: 11, color: '#A5A5A0', marginTop: 2 }}>{fmtDate(p.fecha)}</div>}
                      </div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8, justifyContent: isMobile ? 'space-between' : 'flex-end' }}>
                        <select value={p.estado} onChange={e => updatePuntoEstado(p.id, e.target.value)}
                          style={{ width: isMobile ? '100%' : 'auto', fontSize: 12, padding: '5px 9px', borderRadius: 20, border: `1px solid ${est.color}30`, background: est.bg, color: est.color, fontWeight: 500, cursor: 'pointer' }}>
                          {Object.entries(ESTADOS_INSP).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
                        </select>
                        <button onClick={() => setConfirmacion({ titulo: 'Eliminar punto de control', texto: `Vas a eliminar "${p.nombre}". Esta acción no se puede deshacer.`, onSi: () => { deletePunto(p.id); setConfirmacion(null); } })} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#C4C3BE', fontSize: 18, padding: '0 2px', lineHeight: 1, flexShrink: 0 }}>×</button>
                      </div>
                    </div>
                    {p.notas !== undefined && (
                      <div style={{ padding: '0 12px 10px', background: '#fff' }}>
                        <textarea value={p.notas || ''} onChange={e => {
                          const disciplinas = obra.disciplinas.map(d => d.id === disciplinaActiva ? { ...d, puntos: d.puntos.map(pp => pp.id === p.id ? { ...pp, notas: e.target.value } : pp) } : d);
                          onSave({ ...obra, disciplinas });
                        }} placeholder="Notas del punto de control..." style={{ minHeight: 44, fontSize: 12.5 }} />
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
      {confirmacion && <ConfirmMini titulo={confirmacion.titulo} texto={confirmacion.texto} onSi={confirmacion.onSi} onNo={() => setConfirmacion(null)} />}
    </div>
    </>
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
async function subirFotoStorage(obraId, fotoId, base64) {
  if (!window.db?.subirFoto || !obraId) return { id: fotoId, data: base64 };
  try {
    const { path, url } = await window.db.subirFoto(obraId, fotoId, base64);
    return { id: fotoId, path, url, data: base64 }; // guardamos base64 como fallback temporal
  } catch(e) {
    console.error('Error subiendo foto a Storage:', e);
    return { id: fotoId, data: base64 }; // fallback: guardar base64
  }
}

// Renovar URLs firmadas (caducan, hay que renovarlas al cargar)
async function renovarUrlsFotos(fotos) {
  if (!window.db?.getFotoUrl || !fotos?.length) return fotos;
  return Promise.all(fotos.map(async f => {
    if (f.path && (!f.url || f.url.includes('token='))) {
      try {
        const url = await window.db.getFotoUrl(f.path);
        return { ...f, url };
      } catch(e) { return f; }
    }
    return f;
  }));
}

function pickFiles(accept, cb, obraId) {
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
            const maxW = 1080, ratio = Math.min(1, maxW / img.width);
            const c = document.createElement('canvas');
            c.width = Math.round(img.width * ratio);
            c.height = Math.round(img.height * ratio);
            c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
            const base64 = c.toDataURL('image/jpeg', 0.75);
            const fotoId = uid();
            // Subir a Storage si hay obraId
            subirFotoStorage(obraId, fotoId, base64).then(foto => {
              cb({ ...foto, nombre: f.name, tipo: 'imagen' });
            });
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

// Helper: obtener src de una foto (URL de Storage o base64 fallback)
function fotoSrc(foto) {
  if (!foto) return '';
  return foto.url || foto.data || '';
}
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
            ? <img src={fotoSrc(foto)} alt="" style={{ width: '100%', height: '100%', minHeight: isMobile ? 78 : 80, objectFit: 'cover', display: 'block' }} />
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
function FormNuevaIncidencia({ onClose, onCrear, obraId }) {
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
              <img src={fotoSrc(f)} alt="" style={{ width: 70, height: 56, objectFit: 'cover', borderRadius: 7, border: '1px solid #E0DFD9', display: 'block' }} />
              <button onClick={() => setFotos(p => p.filter(x => x.id !== f.id))} style={{ position: 'absolute', top: -5, right: -5, width: 16, height: 16, borderRadius: '50%', background: '#8A1F1F', color: '#fff', border: 'none', cursor: 'pointer', fontSize: 10, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>×</button>
            </div>
          ))}
          <button onClick={() => pickFiles('image/*', f => setFotos(p => [...p, f]), obraId)} style={{ width: 70, height: 56, borderRadius: 7, border: '1.5px dashed #E0DFD9', background: 'transparent', cursor: 'pointer', fontSize: 20, color: '#A5A5A0', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>📷</button>
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
function DetalleIncidencia({ inc, onClose, onActualizar, onEliminar, obraId }) {
  const [nota,       setNota]       = useState('');
  const [adjuntos,   setAdjuntos]   = useState([]);
  const [estado,     setEstado]     = useState(inc.estado);
  const [preview,    setPreview]    = useState(null);
  const [menu,       setMenu]       = useState(false);
  const [confirmar,  setConfirmar]  = useState(false);
  const [editTitulo, setEditTitulo] = useState(false);
  const [titulo,     setTitulo]     = useState(inc.titulo);
  const [editH,      setEditH]      = useState(null); // id de entrada del historial en edición

  function guardar() {
    if (!nota.trim() && adjuntos.length === 0 && estado === inc.estado) return;
    const estadoCambio = estado !== inc.estado;
    const entrada = { id: uid(), tipo: estadoCambio ? 'cambio_estado' : 'nota', estado, nota: nota.trim(), adjuntos, fecha: now() };
    onActualizar({ ...inc, titulo, estado, historial: [...(inc.historial || []), entrada], ultimaActualizacion: now() });
    setNota(''); setAdjuntos([]);
  }

  function guardarTitulo() {
    if (!titulo.trim()) return;
    onActualizar({ ...inc, titulo: titulo.trim(), ultimaActualizacion: now() });
    setEditTitulo(false);
  }

  function guardarEntrada(entradaId, nuevaNota, nuevosAdjuntos) {
    const historial = (inc.historial || []).map(h =>
      h.id === entradaId ? { ...h, nota: nuevaNota, adjuntos: nuevosAdjuntos } : h
    );
    onActualizar({ ...inc, historial, ultimaActualizacion: now() });
    setEditH(null);
  }

  function eliminarEntrada(entradaId) {
    const historial = (inc.historial || []).filter(h => h.id !== entradaId);
    onActualizar({ ...inc, historial, ultimaActualizacion: now() });
  }

  const est = ESTADOS_INC[inc.estado] || ESTADOS_INC.detectada;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden' }}>
      {/* Header */}
      <div style={{ padding: '12px 16px', borderBottom: '1px solid #E8E7E1', display: 'flex', alignItems: 'center', gap: 10, flexShrink: 0 }}>
        <button onClick={onClose} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 13, color: '#6B6B66', display: 'flex', alignItems: 'center', gap: 4 }}>← Volver</button>
        <span style={{ color: '#D4D3CE' }}>/</span>
        {editTitulo ? (
          <input value={titulo} onChange={e => setTitulo(e.target.value)} onBlur={guardarTitulo} onKeyDown={e => e.key === 'Enter' && guardarTitulo()} autoFocus style={{ flex: 1, fontWeight: 500, fontSize: 13 }} />
        ) : (
          <>
            <span style={{ fontSize: 13, fontWeight: 500, flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{inc.titulo}</span>
            <button onClick={() => setEditTitulo(true)} title="Editar título" style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#A5A5A0', fontSize: 13, padding: '0 2px', flexShrink: 0 }}>✏️</button>
          </>
        )}
        <Pill label={est.label} bg={est.bg} color={est.color} />
        <div style={{ position: 'relative', flexShrink: 0 }}>
          <button onClick={() => setMenu(m => !m)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#A5A5A0', fontSize: 18, lineHeight: 1, padding: '0 2px' }}>⋮</button>
          {menu && (
            <>
              <div onClick={() => setMenu(false)} style={{ position: 'fixed', inset: 0, zIndex: 20 }} />
              <div style={{ position: 'absolute', right: 0, top: '100%', background: '#fff', border: '1px solid #E0DFD9', borderRadius: 9, boxShadow: '0 8px 24px rgba(0,0,0,.12)', padding: 5, zIndex: 21, minWidth: 130 }}>
                <div onClick={() => { setMenu(false); setEditTitulo(true); }} className="hov-row" style={{ padding: '7px 11px', borderRadius: 6, cursor: 'pointer', fontSize: 13, color: '#16160F' }}>✏️ Editar título</div>
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
        {/* Estado */}
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
        <div style={{ fontSize: 11, textTransform: 'uppercase', letterSpacing: '0.1em', color: '#A5A5A0', fontWeight: 500, marginBottom: 10 }}>Historial</div>

        {(inc.historial || []).slice().reverse().map(h => (
          <div key={h.id} style={{ padding: h.tipo === 'revision' ? '7px 12px' : '10px 12px', background: h.tipo === 'revision' ? 'transparent' : '#fff', border: `1px solid ${h.tipo === 'revision' ? '#EEECEA' : '#E8E7E1'}`, borderRadius: 9, marginBottom: 6 }}>

            {editH === h.id ? (
              // ── Modo edición inline ──────────────────────────────────────
              <EntradaEditor
                entrada={h}
                obraId={obraId}
                onGuardar={(nota, adjs) => guardarEntrada(h.id, nota, adjs)}
                onCancelar={() => setEditH(null)}
              />
            ) : (
              // ── Modo lectura ─────────────────────────────────────────────
              <>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: (h.nota || h.adjuntos?.length) ? 6 : 0 }}>
                  <span style={{ fontSize: 11, color: '#A5A5A0' }}>{fmtDate(h.fecha)}</span>
                  {h.tipo === 'creacion'     && <span style={{ fontSize: 11, color: '#6B6B66', background: '#F0EFEA', padding: '1px 7px', borderRadius: 20 }}>Registrada</span>}
                  {h.tipo === 'cambio_estado'&& <Pill label={ESTADOS_INC[h.estado]?.label || h.estado} bg={ESTADOS_INC[h.estado]?.bg || '#eee'} color={ESTADOS_INC[h.estado]?.color || '#333'} />}
                  {h.tipo === 'nota'         && <span style={{ fontSize: 11, color: '#A5A5A0', fontStyle: 'italic' }}>Actualización</span>}
                  {h.tipo === 'revision'     && <span style={{ fontSize: 11, color: '#A5A5A0' }}>✓ Revisada · Sin cambios</span>}
                  {/* Botones editar/eliminar entrada */}
                  {(h.tipo === 'nota' || h.tipo === 'creacion' || h.tipo === 'cambio_estado') && (
                    <div style={{ marginLeft: 'auto', display: 'flex', gap: 4 }}>
                      <button onClick={() => setEditH(h.id)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#A5A5A0', fontSize: 13, padding: '0 3px' }} title="Editar">✏️</button>
                      <button onClick={() => eliminarEntrada(h.id)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#C4C3BE', fontSize: 15, lineHeight: 1, padding: '0 3px' }} title="Eliminar entrada">×</button>
                    </div>
                  )}
                </div>
                {h.nota && <div style={{ fontSize: 13, lineHeight: 1.5, marginBottom: h.adjuntos?.length ? 8 : 0 }}>{h.nota}</div>}
                {h.adjuntos?.length > 0 && (
                  <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                    {h.adjuntos.map(a => a.tipo === 'imagen'
                      ? <img key={a.id} src={fotoSrc(a)} alt={a.nombre} onClick={() => setPreview(a)} style={{ width: 140, height: 110, objectFit: 'cover', borderRadius: 8, border: '1px solid #E0DFD9', cursor: 'zoom-in' }} />
                      : <a key={a.id} href={fotoSrc(a)} download={a.nombre} style={{ display: 'inline-flex', alignItems: 'center', gap: 4, fontSize: 12, padding: '4px 9px', borderRadius: 7, background: '#F5F4F0', border: '1px solid #E0DFD9', color: '#18180F', textDecoration: 'none' }}>📄 {a.nombre}</a>
                    )}
                  </div>
                )}
              </>
            )}
          </div>
        ))}

        {/* Añadir actualización */}
        <div style={{ background: '#F9F8F5', borderRadius: 10, padding: '12px 14px', marginTop: 4 }}>
          <div style={{ fontSize: 11, color: '#A5A5A0', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8 }}>Añadir actualización</div>
          <textarea placeholder="Nota de seguimiento..." value={nota} onChange={e => setNota(e.target.value)} style={{ marginBottom: 8, minHeight: 64 }} />
          {adjuntos.length > 0 && (
            <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginBottom: 8 }}>
              {adjuntos.map(a => (
                <div key={a.id} style={{ position: 'relative' }}>
                  {a.tipo === 'imagen'
                    ? <img src={fotoSrc(a)} alt="" style={{ width: 60, height: 48, objectFit: 'cover', borderRadius: 6, border: '1px solid #E0DFD9' }} />
                    : <div style={{ fontSize: 11, padding: '4px 8px', background: '#fff', border: '1px solid #E0DFD9', borderRadius: 6 }}>📄 {a.nombre}</div>}
                  <button onClick={() => setAdjuntos(p => p.filter(x => x.id !== a.id))} style={{ position: 'absolute', top: -4, right: -4, width: 14, height: 14, borderRadius: '50%', background: '#8A1F1F', color: '#fff', border: 'none', cursor: 'pointer', fontSize: 9, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>×</button>
                </div>
              ))}
            </div>
          )}
          <div style={{ display: 'flex', gap: 8 }}>
            <button onClick={() => pickFiles('image/*,.pdf,.doc,.docx', f => setAdjuntos(p => [...p, f]), obraId)} style={{ padding: '7px 12px', borderRadius: 8, border: '1px solid #E0DFD9', background: '#fff', cursor: 'pointer', fontSize: 13, whiteSpace: 'nowrap' }}>📎 Adjuntar</button>
            <Btn primary full onClick={guardar} disabled={!nota.trim() && adjuntos.length === 0 && estado === inc.estado}>Guardar</Btn>
          </div>
        </div>
      </div>

      {preview && (
        <div onClick={() => setPreview(null)} style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.85)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 2000, padding: 30, cursor: 'zoom-out' }}>
          <img src={fotoSrc(preview)} alt={preview.nombre} style={{ maxWidth: '100%', maxHeight: '100%', objectFit: 'contain', borderRadius: 6 }} />
          <button onClick={() => setPreview(null)} style={{ position: 'absolute', top: 20, right: 24, background: 'rgba(255,255,255,0.15)', border: 'none', color: '#fff', fontSize: 24, width: 40, height: 40, borderRadius: '50%', cursor: 'pointer', lineHeight: 1 }}>×</button>
          <a href={fotoSrc(preview)} download={preview.nombre} onClick={e => e.stopPropagation()} style={{ position: 'absolute', bottom: 24, background: 'rgba(255,255,255,0.15)', color: '#fff', fontSize: 13, padding: '8px 16px', borderRadius: 8, textDecoration: 'none' }}>↓ Descargar</a>
        </div>
      )}
    </div>
  );
}

// ── Editor inline de entrada del historial ────────────────────────────────────
function EntradaEditor({ entrada, obraId, onGuardar, onCancelar }) {
  const [nota,     setNota]     = useState(entrada.nota || '');
  const [adjuntos, setAdjuntos] = useState(entrada.adjuntos || []);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
      <textarea value={nota} onChange={e => setNota(e.target.value)} style={{ minHeight: 64, fontSize: 13 }} autoFocus />
      {/* Fotos actuales */}
      <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
        {adjuntos.map(a => a.tipo === 'imagen' ? (
          <div key={a.id} style={{ position: 'relative' }}>
            <img src={fotoSrc(a)} alt="" style={{ width: 60, height: 48, objectFit: 'cover', borderRadius: 6, border: '1px solid #E0DFD9' }} />
            <button onClick={() => setAdjuntos(p => p.filter(x => x.id !== a.id))} style={{ position: 'absolute', top: -4, right: -4, width: 14, height: 14, borderRadius: '50%', background: '#8A1F1F', color: '#fff', border: 'none', cursor: 'pointer', fontSize: 9, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>×</button>
          </div>
        ) : (
          <div key={a.id} style={{ position: 'relative', display: 'flex', alignItems: 'center', gap: 4, fontSize: 11, padding: '4px 8px', background: '#fff', border: '1px solid #E0DFD9', borderRadius: 6 }}>
            📄 {a.nombre}
            <button onClick={() => setAdjuntos(p => p.filter(x => x.id !== a.id))} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#C4C3BE', fontSize: 13, lineHeight: 1 }}>×</button>
          </div>
        ))}
        <button onClick={() => pickFiles('image/*,.pdf,.doc,.docx', f => setAdjuntos(p => [...p, f]), obraId)} style={{ width: 60, height: 48, borderRadius: 6, border: '1.5px dashed #E0DFD9', background: '#FAFAF8', cursor: 'pointer', fontSize: 18, color: '#A5A5A0', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>+</button>
      </div>
      <div style={{ display: 'flex', gap: 8 }}>
        <Btn sm onClick={onCancelar} full>Cancelar</Btn>
        <Btn sm primary onClick={() => onGuardar(nota, adjuntos)} full>Guardar cambios</Btn>
      </div>
    </div>
  );
}

// ── Modal de revisión en visita ───────────────────────────────────────────────
function ModalRevision({ inc, onSinCambios, onConCambios, onClose, obraId }) {
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
                        ? <img src={fotoSrc(a)} alt="" style={{ width: 60, height: 48, objectFit: 'cover', borderRadius: 6, border: '1px solid #E0DFD9' }} />
                        : <div style={{ fontSize: 11, padding: '4px 8px', background: '#F5F4F0', border: '1px solid #E0DFD9', borderRadius: 6 }}>📄 {a.nombre}</div>}
                      <button onClick={() => setAdjuntos(p => p.filter(x => x.id !== a.id))} style={{ position: 'absolute', top: -4, right: -4, width: 14, height: 14, borderRadius: '50%', background: '#8A1F1F', color: '#fff', border: 'none', cursor: 'pointer', fontSize: 9 }}>×</button>
                    </div>
                  ))}
                </div>
              )}

              <button onClick={() => pickFiles('image/*,.pdf,.doc,.docx', f => setAdjuntos(p => [...p, f]), obraId)} style={{ width: '100%', padding: '8px', borderRadius: 8, border: '1.5px dashed #E0DFD9', background: 'transparent', cursor: 'pointer', fontSize: 12, color: '#6B6B66', marginBottom: 14 }}>📎 Adjuntar foto o documento</button>

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

  async function borrarInc(incId) {
    // 1. Actualizar estado local inmediatamente
    const nuevas = obra.incidencias.filter(i => i.id !== incId);
    onSave({ ...obra, incidencias: nuevas });
    setIncActiva(null);
    // 2. Borrar de Supabase directamente
    try {
      await window.db.deleteModulo('incidencias', incId);
    } catch(e) { console.error('Error borrando incidencia:', e); }
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
          obraId={obra.id}
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
      {showNueva && <FormNuevaIncidencia onClose={() => setShowNueva(false)} onCrear={crearInc} obraId={obra.id} />}

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
          obraId={obra.id}
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

  async function eliminar(id) {
    onSave({ ...obra, apuntes: apuntes.filter(a => a.id !== id) });
    try { await window.db.deleteModulo('notas', id); } catch(e) { console.error('Error borrando nota:', e); }
  }
  function editarTexto(id, texto) {
    onSave({ ...obra, apuntes: apuntes.map(a => a.id === id ? { ...a, texto } : a) });
  }
  function addComentario(id, texto) {
    if (!texto.trim()) return;
    onSave({ ...obra, apuntes: apuntes.map(a => a.id === id ? { ...a, comentarios: [...(a.comentarios || []), { id: uid(), texto: texto.trim(), creadaEn: now() }] } : a) });
  }
  function delComentario(id, comId) {
    onSave({ ...obra, apuntes: apuntes.map(a => a.id === id ? { ...a, comentarios: (a.comentarios || []).filter(c => c.id !== comId) } : a) });
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
          {mostrados.map(item => (
            <ApunteItem key={item.id} item={item} vencida={isVencida(item)}
              onToggleHecha={() => toggleHecha(item.id)}
              onEditarTexto={txt => editarTexto(item.id, txt)}
              onAddComentario={txt => addComentario(item.id, txt)}
              onConfirmar={setConfirmacion}
              onDelComentario={comId => delComentario(item.id, comId)}
              onEliminar={() => eliminar(item.id)} />
          ))}
        </div>
      )}
      {confirmacion && <ConfirmMini titulo={confirmacion.titulo} texto={confirmacion.texto} onSi={confirmacion.onSi} onNo={() => setConfirmacion(null)} />}
    </div>
  );
}

// Tarjeta de tarea/nota con edición de texto y comentarios de seguimiento
function ApunteItem({ item, vencida, onToggleHecha, onEditarTexto, onAddComentario, onConfirmar, onDelComentario, onEliminar }) {
  const [abierto, setAbierto] = useState(false);
  const [editando, setEditando] = useState(false);
  const [txt, setTxt] = useState(item.texto);
  const [coment, setComent] = useState('');
  const comentarios = item.comentarios || [];

  return (
    <div style={{ background: '#fff', border: `1px solid ${vencida ? '#F4ABAB' : '#E8E7E1'}`, borderLeft: vencida ? '3px solid #E24B4A' : '1px solid #E8E7E1', borderRadius: 10, padding: '10px 13px' }}>
      <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10 }}>
        {/* Checkbox (solo tareas) */}
        {item.tipo === 'tarea' && (
          <button onClick={onToggleHecha} style={{ width: 20, height: 20, borderRadius: 5, border: `2px solid ${item.hecha ? '#52A124' : '#D4D3CE'}`, background: item.hecha ? '#52A124' : 'transparent', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0, marginTop: 1, fontSize: 11, color: '#fff', transition: 'all .15s' }}>
            {item.hecha ? '✓' : ''}
          </button>
        )}
        {item.tipo === 'nota' && <span style={{ fontSize: 15, flexShrink: 0, marginTop: 1 }}>📝</span>}

        {/* Contenido */}
        <div style={{ flex: 1, minWidth: 0 }}>
          {editando ? (
            <div style={{ marginBottom: 6 }}>
              <textarea autoFocus value={txt} onChange={e => setTxt(e.target.value)} style={{ minHeight: 60, marginBottom: 6 }} />
              <div style={{ display: 'flex', gap: 6 }}>
                <Btn sm primary disabled={!txt.trim()} onClick={() => { onEditarTexto(txt.trim()); setEditando(false); }}>Guardar</Btn>
                <Btn sm onClick={() => { setTxt(item.texto); setEditando(false); }}>✕</Btn>
              </div>
            </div>
          ) : (
            <div style={{ display: 'flex', alignItems: 'flex-start', gap: 6, marginBottom: 5 }}>
              <div style={{ flex: 1, fontSize: 13, color: item.hecha ? '#A5A5A0' : '#18180F', textDecoration: item.hecha ? 'line-through' : 'none', lineHeight: 1.4, whiteSpace: 'pre-wrap' }}>
                {item.texto}
              </div>
              <button onClick={() => setEditando(true)} title="Editar" style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#C4C3BE', fontSize: 13, padding: '0 2px', flexShrink: 0, marginTop: 1 }}>✏️</button>
            </div>
          )}
          <div style={{ display: 'flex', gap: 6, alignItems: 'center', flexWrap: 'wrap' }}>
            <span style={{ fontSize: 11, padding: '1px 7px', borderRadius: 20, background: '#F0EFEA', color: '#6B6B66' }}>{item.categoria}</span>
            {item.tipo === 'tarea' && item.fechaLimite && (
              <span style={{ fontSize: 11, color: vencida ? '#8A1F1F' : '#A5A5A0', fontWeight: vencida ? 500 : 400 }}>
                {vencida ? '⚠ Vencida · ' : '📅 '}{fmtDate(item.fechaLimite)}
              </span>
            )}
            {item.tipo === 'nota' && <span style={{ fontSize: 11, color: '#A5A5A0' }}>{fmtShort(item.creadaEn)}</span>}
            <button onClick={() => setAbierto(v => !v)} style={{ marginLeft: 'auto', background: 'none', border: 'none', cursor: 'pointer', fontSize: 11, color: '#6B6B66', display: 'flex', alignItems: 'center', gap: 3 }}>
              💬 {comentarios.length > 0 ? comentarios.length : ''} {abierto ? '▲' : '▼'}
            </button>
          </div>
        </div>

        {/* Eliminar */}
        <button onClick={() => onConfirmar({ titulo: item.tipo === 'tarea' ? 'Eliminar tarea' : 'Eliminar nota', texto: 'Esta acción no se puede deshacer.', onSi: () => { onEliminar(); onConfirmar(null); } })} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#D4D3CE', fontSize: 16, padding: '0 2px', lineHeight: 1, flexShrink: 0 }}>×</button>
      </div>

      {/* Comentarios de seguimiento */}
      {abierto && (
        <div className="fade" style={{ marginTop: 10, paddingTop: 10, borderTop: '1px solid #F2F1ED' }}>
          {comentarios.map(c => (
            <div key={c.id} style={{ display: 'flex', gap: 8, alignItems: 'flex-start', padding: '6px 0', borderBottom: '1px solid #F5F4F0' }}>
              <span style={{ fontSize: 11, color: '#A5A5A0', whiteSpace: 'nowrap', paddingTop: 1, minWidth: 40 }}>{fmtShort(c.creadaEn)}</span>
              <span style={{ flex: 1, fontSize: 13, color: '#18180F', lineHeight: 1.45 }}>{c.texto}</span>
              <button onClick={() => onConfirmar({ titulo: 'Eliminar comentario', texto: 'Vas a eliminar este comentario de seguimiento.', onSi: () => { onDelComentario(c.id); onConfirmar(null); } })} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#D4D3CE', fontSize: 14, lineHeight: 1 }}>×</button>
            </div>
          ))}
          <div style={{ display: 'flex', gap: 6, marginTop: 8 }}>
            <input placeholder="Añadir seguimiento..." value={coment} onChange={e => setComent(e.target.value)} onKeyDown={e => { if (e.key === 'Enter' && coment.trim()) { onAddComentario(coment); setComent(''); } }} style={{ flex: 1, fontSize: 12 }} />
            <Btn sm primary disabled={!coment.trim()} onClick={() => { onAddComentario(coment); setComent(''); }}>Añadir</Btn>
          </div>
        </div>
      )}
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
          <img src={fotoSrc(preview)} alt={preview.nombre} style={{ maxWidth: '100%', maxHeight: '100%', objectFit: 'contain', borderRadius: 6 }} />
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
                          ? <img src={fotoSrc(a)} alt="" onClick={() => onPreview(a)} style={{ width: 60, height: 48, objectFit: 'cover', borderRadius: 6, border: '1px solid #E0DFD9', cursor: 'zoom-in' }} />
                          : <div style={{ fontSize: 11, padding: '4px 8px', background: '#fff', border: '1px solid #E0DFD9', borderRadius: 6 }}>📄 {a.nombre}</div>}
                        <button onClick={() => setAdjuntos(p => p.filter(x => x.id !== a.id))} style={{ position: 'absolute', top: -4, right: -4, width: 14, height: 14, borderRadius: '50%', background: '#8A1F1F', color: '#fff', border: 'none', cursor: 'pointer', fontSize: 9 }}>×</button>
                      </div>
                    ))}
                  </div>
                )}
                <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                  <button onClick={() => pickFiles('image/*,.pdf,.doc,.docx', f => setAdjuntos(p => [...p, f]), obraId)} style={{ padding: '7px 12px', borderRadius: 8, border: '1px solid #E0DFD9', background: '#fff', cursor: 'pointer', fontSize: 13, whiteSpace: 'nowrap' }}>📎 Adjuntar</button>
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
                      ? <img key={a.id} src={fotoSrc(a)} alt={a.nombre} onClick={() => onPreview(a)} style={{ width: 120, height: 92, objectFit: 'cover', borderRadius: 8, border: '1px solid #E0DFD9', cursor: 'zoom-in' }} />
                      : <a key={a.id} href={fotoSrc(a)} download={a.nombre} style={{ display: 'inline-flex', alignItems: 'center', gap: 4, fontSize: 12, padding: '4px 9px', borderRadius: 7, background: '#F5F4F0', border: '1px solid #E0DFD9', color: '#18180F', textDecoration: 'none' }}>📄 {a.nombre}</a>
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
                    ? <img src={fotoSrc(a)} alt="" onClick={() => onPreview(a)} style={{ width: 60, height: 48, objectFit: 'cover', borderRadius: 6, border: '1px solid #E0DFD9', cursor: 'zoom-in' }} />
                    : <div style={{ fontSize: 11, padding: '4px 8px', background: '#fff', border: '1px solid #E0DFD9', borderRadius: 6 }}>📄 {a.nombre}</div>}
                  <button onClick={() => setAdjuntos(p => p.filter(x => x.id !== a.id))} style={{ position: 'absolute', top: -4, right: -4, width: 14, height: 14, borderRadius: '50%', background: '#8A1F1F', color: '#fff', border: 'none', cursor: 'pointer', fontSize: 9 }}>×</button>
                </div>
              ))}
            </div>
          )}
          <div style={{ display: 'flex', gap: 8 }}>
            <button onClick={() => pickFiles('image/*,.pdf,.doc,.docx', f => setAdjuntos(p => [...p, f]), obraId)} style={{ padding: '7px 12px', borderRadius: 8, border: '1px solid #E0DFD9', background: '#fff', cursor: 'pointer', fontSize: 13, whiteSpace: 'nowrap' }}>📎 Adjuntar acta</button>
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
  const [sub, setSub] = useState('hormigon'); // hormigon | materiales | ensayos

  const subTabs = [
    { id: 'hormigon',   label: 'Control de hormigón' },
    { id: 'materiales', label: 'Materiales' },
    { id: 'ensayos',    label: 'Ensayos' },
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

      {sub === 'hormigon'   && <ControlHormigon obra={obra} onSave={onSave} />}
      {sub === 'materiales' && <ModuloMateriales obra={obra} onSave={onSave} />}
      {sub === 'ensayos'    && <ModuloEnsayos obra={obra} onSave={onSave} />}
    </div>
  );
}

// ─── Materiales: contenedor con dos sub-apartados ─────────────────────────────
function ModuloMateriales({ obra, onSave }) {
  const [sub, setSub] = useState('pendientes'); // pendientes | seguimiento
  const subTabs = [
    { id: 'pendientes',  label: 'Materiales pendientes' },
    { id: 'seguimiento', label: 'Seguimiento CQ' },
  ];
  return (
    <div>
      <div style={{ display: 'flex', gap: 4, marginBottom: 14, borderBottom: '1px solid #E8E7E1' }}>
        {subTabs.map(t => (
          <button key={t.id} onClick={() => setSub(t.id)} style={{ padding: '7px 14px', border: 'none', borderBottom: `2px solid ${sub === t.id ? '#1C1C1A' : 'transparent'}`, background: 'transparent', color: sub === t.id ? '#141412' : '#9B9B97', fontSize: 13, cursor: 'pointer', fontWeight: sub === t.id ? 600 : 400, marginBottom: -1 }}>
            {t.label}
          </button>
        ))}
      </div>
      {sub === 'pendientes'  && <MaterialesPendientes obra={obra} onSave={onSave} />}
      {sub === 'seguimiento' && <SeguimientoCQ obra={obra} onSave={onSave} />}
    </div>
  );
}

// ─── Materiales pendientes (sin documentación de CQ) ──────────────────────────
function MaterialesPendientes({ obra, onSave }) {
  const materiales = obra.materiales || [];
  const [nombre, setNombre]   = useState('');
  const [filtro, setFiltro]   = useState('todos'); // todos | falta | pedido | recibido
  const [confirmacion, setConfirmacion] = useState(null);

  function guardar(lista) { onSave({ ...obra, materiales: lista }); }
  function add() {
    if (!nombre.trim()) return;
    guardar([{ id: uid(), nombre: nombre.trim(), pedido: false, recibido: false, nota: '', creadaEn: now() }, ...materiales]);
    setNombre('');
  }
  function toggle(id, campo) {
    guardar(materiales.map(m => {
      if (m.id !== id) return m;
      const v = !m[campo];
      // Si se marca recibido, también queda pedido
      if (campo === 'recibido' && v) return { ...m, recibido: true, pedido: true };
      // Si se desmarca pedido, también deja de estar recibido
      if (campo === 'pedido' && !v) return { ...m, pedido: false, recibido: false };
      return { ...m, [campo]: v };
    }));
  }
  function updNota(id, nota) { guardar(materiales.map(m => m.id === id ? { ...m, nota } : m)); }
  function eliminar(id) { guardar(materiales.filter(m => m.id !== id)); setConfirmacion(null); }

  const filtrados = materiales.filter(m =>
    filtro === 'todos'    ? true :
    filtro === 'falta'    ? !m.pedido :
    filtro === 'pedido'   ? (m.pedido && !m.recibido) :
    filtro === 'recibido' ? m.recibido : true
  );
  const stats = {
    total:    materiales.length,
    falta:    materiales.filter(m => !m.pedido).length,
    pedido:   materiales.filter(m => m.pedido && !m.recibido).length,
    recibido: materiales.filter(m => m.recibido).length,
  };

  const Check = ({ on, onClick, label, color }) => (
    <button onClick={onClick} style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '5px 10px', borderRadius: 8, border: `1.5px solid ${on ? color : '#E0DFD9'}`, background: on ? color + '18' : 'transparent', cursor: 'pointer', fontSize: 12, fontWeight: 500, color: on ? color : '#9B9B97' }}>
      <span style={{ width: 16, height: 16, borderRadius: 5, border: `1.5px solid ${on ? color : '#C5C4BE'}`, background: on ? color : 'transparent', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontSize: 11 }}>{on ? '✓' : ''}</span>
      {label}
    </button>
  );

  return (
    <div>
      <p style={{ fontSize: 13, color: '#6B6B66', lineHeight: 1.5, marginBottom: 12 }}>
        Apunta los materiales que ves en obra y de los que falta documentación de control de calidad. Marca si se ha <strong>pedido</strong> y si se ha <strong>recibido</strong>.
      </p>

      {/* Añadir material */}
      <div style={{ display: 'flex', gap: 8, marginBottom: 14 }}>
        <input placeholder="Ej. Acero corrugado B500S, mortero cola C2..." value={nombre} onChange={e => setNombre(e.target.value)} onKeyDown={e => { if (e.key === 'Enter') add(); }} style={{ flex: 1 }} />
        <Btn primary disabled={!nombre.trim()} onClick={add}>+ Añadir</Btn>
      </div>

      {/* Resumen + filtros */}
      {materiales.length > 0 && (
        <div style={{ display: 'flex', gap: 6, marginBottom: 12, flexWrap: 'wrap' }}>
          {[['todos', `Todos (${stats.total})`], ['falta', `Sin pedir (${stats.falta})`], ['pedido', `Pedidos (${stats.pedido})`], ['recibido', `Recibidos (${stats.recibido})`]].map(([id, label]) => (
            <button key={id} onClick={() => setFiltro(id)} style={{ padding: '4px 12px', borderRadius: 20, border: `1px solid ${filtro === id ? '#1C1C1A' : '#E0DFD9'}`, background: filtro === id ? '#1C1C1A' : 'transparent', color: filtro === id ? '#F2F1ED' : '#6B6B66', fontSize: 12, cursor: 'pointer', fontWeight: filtro === id ? 500 : 400 }}>
              {label}
            </button>
          ))}
        </div>
      )}

      {/* Lista */}
      {materiales.length === 0 ? (
        <div style={{ textAlign: 'center', padding: '40px 20px', color: '#A5A5A0' }}>
          <div style={{ fontSize: 13 }}>Sin materiales apuntados todavía.</div>
        </div>
      ) : filtrados.length === 0 ? (
        <div style={{ textAlign: 'center', padding: '30px', color: '#A5A5A0', fontSize: 13 }}>Ningún material con este filtro.</div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {filtrados.map(m => {
            const accent = m.recibido ? '#2D5E10' : m.pedido ? '#C47610' : '#8A1F1F';
            return (
              <div key={m.id} style={{ background: '#fff', border: '1px solid #E8E7E1', borderLeft: `3px solid ${accent}`, borderRadius: 10, padding: '12px 14px' }}>
                <div style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 10 }}>
                  <div style={{ flex: 1, fontSize: 14, fontWeight: 500, color: '#18180F', lineHeight: 1.4 }}>{m.nombre}</div>
                  <button onClick={() => setConfirmacion({ titulo: 'Eliminar material', texto: `Vas a eliminar "${m.nombre}".`, onSi: () => eliminar(m.id) })} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#D4D3CE', fontSize: 17, lineHeight: 1, padding: '0 2px', flexShrink: 0 }}>×</button>
                </div>
                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 9 }}>
                  <Check on={m.pedido} onClick={() => toggle(m.id, 'pedido')} label="Pedido" color="#C47610" />
                  <Check on={m.recibido} onClick={() => toggle(m.id, 'recibido')} label="Recibido" color="#2D5E10" />
                </div>
                <input placeholder="Nota (proveedor, fecha, referencia...)" value={m.nota || ''} onChange={e => updNota(m.id, e.target.value)} style={{ fontSize: 12 }} />
              </div>
            );
          })}
        </div>
      )}

      {confirmacion && <ConfirmMini titulo={confirmacion.titulo} texto={confirmacion.texto} onSi={confirmacion.onSi} onNo={() => setConfirmacion(null)} />}
    </div>
  );
}

// ─── Seguimiento de paquetes de documentación de CQ ───────────────────────────
const ESTADOS_CQ = {
  pendiente: { label: 'Pendiente', bg: '#FDECEC', color: '#8A1F1F', dot: '#E24B4A' },
  parcial:   { label: 'Parcial',   bg: '#FEF3DB', color: '#7C4A00', dot: '#D48A0C' },
  completo:  { label: 'Completo',  bg: '#E8F5E0', color: '#2D5E10', dot: '#52A124' },
};

function getDefaultSeguimientoCQ() {
  const it = (codigo, nombre) => ({ id: uid(), codigo, nombre, estado: 'pendiente', nota: '' });
  const cap = (codigo, titulo, items) => ({ id: uid(), codigo, titulo, items });
  return [
    cap('04', 'CIMENTACIÓN Y CONTENCIÓN', [ it('', 'Cimentación y contención') ]),
    cap('05', 'SISTEMA ESTRUCTURAL', [ it('05.01', 'Estructura de hormigón'), it('05.02', 'Estructura metálica') ]),
    cap('06', 'FACHADAS Y CERRAMIENTOS', [ it('06.02', 'Acabados y revestimientos'), it('06.03', 'Carpintería de aluminio') ]),
    cap('08', 'IMPERMEABILIZACIONES', [ it('08.02', 'Impermeabilizaciones bajo rasante'), it('08.03', 'Impermeabilizaciones sobre rasante') ]),
    cap('09', 'ALBAÑILERÍA Y OBRA SECA', [ it('09.01', 'Albañilería'), it('09.02', 'Obra seca') ]),
    cap('12', 'FALSOS TECHOS', [ it('12', 'Falsos techos (12.06 a 12.11)') ]),
    cap('13', 'PAVIMENTOS Y REVESTIMIENTOS', [ it('13', 'Pavimentos y revestimientos (13.02 y 13.03)') ]),
    cap('14', 'CARPINTERÍA INTERIOR / MADERA', [ it('14.01', 'Puertas de madera y registros'), it('14.02', 'Mostradores'), it('14.03', 'Panelados') ]),
    cap('15', 'CERRAJERÍA Y PUERTAS METÁLICAS', [ it('15', 'Cerrajería y puertas metálicas (15.15 a 15.18)') ]),
    cap('16', 'METALISTERÍA, VIDRIOS Y VARIOS', [ it('16.02', 'Barandillas'), it('16.03', 'Sombreretes cubierta y registros'), it('16.04', 'Rejas y divisorias'), it('16.05', 'Mamparas y divisorias') ]),
    cap('17', 'APARATOS SANITARIOS Y GRIFERÍA', [ it('17', 'Aparatos sanitarios y grifería (17.11 a 17.14)') ]),
    cap('18', 'SEÑALIZACIÓN Y EQUIPAMIENTO', [ it('18.01', 'Señalización de incendios y pictogramas'), it('18.02', 'Equipamientos') ]),
    cap('21', 'JARDINERÍA Y CUBIERTAS VERDES', [ it('21', 'Jardinería y cubiertas verdes (21.07 a 21.11)') ]),
    cap('22', 'PROVISIONALES', [ it('22.01', 'Provisionales de obra') ]),
    cap('23', 'CONTROL DE CALIDAD', [ it('23.01', 'Control de calidad') ]),
    cap('24', 'SEGURIDAD E HIGIENE', [ it('24.01', 'Seguridad e higiene') ]),
    cap('25', 'ACOMETIDAS', [ it('25.01', 'Acometidas de fontanería y contraincendios'), it('25.02', 'Acometida de red de fecales'), it('25.03', 'Acometida de red de pluviales') ]),
  ];
}

function SeguimientoCQ({ obra, onSave }) {
  const capitulos = (obra.seguimientoCQ && obra.seguimientoCQ.length) ? obra.seguimientoCQ : null;
  const [abierto, setAbierto] = useState({});
  const [confirmacion, setConfirmacion] = useState(null);
  const [editItem, setEditItem] = useState(null);  // id item en edición
  const [editCap, setEditCap]   = useState(null);  // id capítulo en edición

  function guardar(caps) { onSave({ ...obra, seguimientoCQ: caps }); }
  function inicializar() { guardar(getDefaultSeguimientoCQ()); }

  if (!capitulos) {
    return (
      <div style={{ textAlign: 'center', padding: '36px 20px' }}>
        <p style={{ fontSize: 13, color: '#6B6B66', lineHeight: 1.5, marginBottom: 16, maxWidth: 420, margin: '0 auto 16px' }}>
          Seguimiento general de los paquetes de documentación de control de calidad de la obra. Carga la estructura de paquetes por defecto y ve marcando el estado de cada uno.
        </p>
        <Btn primary onClick={inicializar}>Cargar paquetes de documentación</Btn>
      </div>
    );
  }

  // Stats globales
  const allItems = capitulos.flatMap(c => c.items);
  const total = allItems.length;
  const completos = allItems.filter(i => i.estado === 'completo').length;
  const parciales = allItems.filter(i => i.estado === 'parcial').length;
  const pct = total ? Math.round(completos / total * 100) : 0;

  function setEstado(capId, itemId, estado) {
    guardar(capitulos.map(c => c.id !== capId ? c : { ...c, items: c.items.map(i => i.id === itemId ? { ...i, estado } : i) }));
  }
  function setNota(capId, itemId, nota) {
    guardar(capitulos.map(c => c.id !== capId ? c : { ...c, items: c.items.map(i => i.id === itemId ? { ...i, nota } : i) }));
  }
  function setNombre(capId, itemId, nombre) {
    guardar(capitulos.map(c => c.id !== capId ? c : { ...c, items: c.items.map(i => i.id === itemId ? { ...i, nombre } : i) }));
  }
  function setTitulo(capId, titulo) {
    guardar(capitulos.map(c => c.id === capId ? { ...c, titulo } : c));
  }
  function addItem(capId) {
    guardar(capitulos.map(c => c.id !== capId ? c : { ...c, items: [...c.items, { id: uid(), codigo: '', nombre: 'Nuevo elemento', estado: 'pendiente', nota: '' }] }));
  }
  function delItem(capId, itemId) {
    guardar(capitulos.map(c => c.id !== capId ? c : { ...c, items: c.items.filter(i => i.id !== itemId) }));
    setConfirmacion(null);
  }
  function addCap() {
    const nuevo = { id: uid(), codigo: '', titulo: 'NUEVO CAPÍTULO', items: [{ id: uid(), codigo: '', nombre: 'Nuevo elemento', estado: 'pendiente', nota: '' }] };
    guardar([...capitulos, nuevo]);
    setAbierto(a => ({ ...a, [nuevo.id]: true }));
    setEditCap(nuevo.id);
  }
  function delCap(capId) {
    guardar(capitulos.filter(c => c.id !== capId));
    setConfirmacion(null);
  }

  // Estado agregado del capítulo
  function estadoCap(c) {
    if (!c.items.length) return 'pendiente';
    if (c.items.every(i => i.estado === 'completo')) return 'completo';
    if (c.items.some(i => i.estado !== 'pendiente')) return 'parcial';
    return 'pendiente';
  }

  const Seg = ({ estado, onChange }) => (
    <div style={{ display: 'flex', gap: 0, border: '1px solid #E0DFD9', borderRadius: 7, overflow: 'hidden', flexShrink: 0 }}>
      {Object.entries(ESTADOS_CQ).map(([k, v]) => (
        <button key={k} onClick={() => onChange(k)} style={{ padding: '4px 9px', border: 'none', borderRight: k !== 'completo' ? '1px solid #E0DFD9' : 'none', background: estado === k ? v.dot : 'transparent', color: estado === k ? '#fff' : '#9B9B97', fontSize: 11, cursor: 'pointer', fontWeight: estado === k ? 600 : 400, whiteSpace: 'nowrap' }}>
          {v.label}
        </button>
      ))}
    </div>
  );

  return (
    <div>
      {/* Resumen global */}
      <div style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 12, padding: '14px 16px', marginBottom: 14 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 10 }}>
          <span style={{ fontSize: 13, fontWeight: 600, color: '#141412', flex: 1 }}>Estado general de la documentación</span>
          <span style={{ fontSize: 20, fontWeight: 700, color: pct === 100 ? '#2D5E10' : '#141412' }}>{pct}%</span>
        </div>
        <div style={{ height: 6, background: '#ECEAE4', borderRadius: 3, overflow: 'hidden', marginBottom: 8 }}>
          <div className="bar-fill" style={{ width: pct + '%', height: 6, background: '#52A124', borderRadius: 3 }} />
        </div>
        <div style={{ display: 'flex', gap: 14, fontSize: 12, color: '#6B6B66' }}>
          <span><strong style={{ color: '#2D5E10' }}>{completos}</strong> completos</span>
          <span><strong style={{ color: '#7C4A00' }}>{parciales}</strong> parciales</span>
          <span><strong style={{ color: '#8A1F1F' }}>{total - completos - parciales}</strong> pendientes</span>
        </div>
      </div>

      {/* Capítulos */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 7 }}>
        {capitulos.map(c => {
          const est = ESTADOS_CQ[estadoCap(c)];
          const open = abierto[c.id];
          const doneItems = c.items.filter(i => i.estado === 'completo').length;
          return (
            <div key={c.id} style={{ background: '#fff', border: '1px solid #E8E7E1', borderRadius: 10, overflow: 'hidden' }}>
              {/* Cabecera capítulo */}
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '11px 14px' }}>
                <span onClick={() => setAbierto(a => ({ ...a, [c.id]: !a[c.id] }))} style={{ width: 9, height: 9, borderRadius: '50%', background: est.dot, flexShrink: 0, cursor: 'pointer' }} />
                {c.codigo && <span style={{ fontSize: 11, fontWeight: 700, color: '#A5A5A0', flexShrink: 0 }}>{c.codigo}</span>}
                {editCap === c.id
                  ? <input autoFocus value={c.titulo} onChange={e => setTitulo(c.id, e.target.value)} onBlur={() => setEditCap(null)}
                      style={{ flex: 1, fontSize: 13, fontWeight: 600, padding: '3px 7px' }} />
                  : <span onClick={() => setAbierto(a => ({ ...a, [c.id]: !a[c.id] }))} style={{ flex: 1, fontSize: 13, fontWeight: 500, color: '#141412', cursor: 'pointer' }}>{c.titulo}</span>}
                <span style={{ fontSize: 11, color: '#A5A5A0' }}>{doneItems}/{c.items.length}</span>
                <button onClick={e => { e.stopPropagation(); setEditCap(c.id); }} title="Editar nombre" style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#C5C4BE', fontSize: 12, lineHeight: 1, padding: '0 2px' }}>✎</button>
                <button onClick={e => { e.stopPropagation(); setConfirmacion({ titulo: 'Eliminar capítulo', texto: `Vas a eliminar "${c.titulo}" y todos sus elementos.`, onSi: () => delCap(c.id) }); }} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#D4D3CE', fontSize: 16, lineHeight: 1, padding: '0 2px' }}>×</button>
                <span onClick={() => setAbierto(a => ({ ...a, [c.id]: !a[c.id] }))} style={{ fontSize: 11, color: '#C5C4BE', cursor: 'pointer' }}>{open ? '▲' : '▼'}</span>
              </div>

              {open && (
                <div className="fade" style={{ padding: '0 14px 12px', borderTop: '1px solid #F2F1ED' }}>
                  {c.items.map(i => (
                    <div key={i.id} style={{ padding: '10px 0', borderBottom: '1px solid #F5F4F0' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 7 }}>
                        {i.codigo && <span style={{ fontSize: 11, fontWeight: 600, color: '#A5A5A0', flexShrink: 0 }}>{i.codigo}</span>}
                        {editItem === i.id
                          ? <input autoFocus value={i.nombre} onChange={e => setNombre(c.id, i.id, e.target.value)} onBlur={() => setEditItem(null)} style={{ flex: 1, fontSize: 13, padding: '3px 7px' }} />
                          : <><span style={{ flex: 1, fontSize: 13, color: '#18180F' }}>{i.nombre}</span><button onClick={() => setEditItem(i.id)} title="Editar" style={{ background:'none', border:'none', cursor:'pointer', color:'#C4C3BE', fontSize:12, padding:'0 2px', flexShrink:0 }}>✏️</button></>}
                        <button onClick={() => setConfirmacion({ titulo: 'Eliminar elemento', texto: `Vas a eliminar "${i.nombre}".`, onSi: () => delItem(c.id, i.id) })} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#D4D3CE', fontSize: 15, lineHeight: 1, flexShrink: 0 }}>×</button>
                      </div>
                      <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                        <Seg estado={i.estado} onChange={e => setEstado(c.id, i.id, e)} />
                        <input placeholder="Nota (qué falta, referencia...)" value={i.nota || ''} onChange={e => setNota(c.id, i.id, e.target.value)} style={{ flex: 1, minWidth: 140, fontSize: 12 }} />
                      </div>
                    </div>
                  ))}
                  <button onClick={() => addItem(c.id)} style={{ width: '100%', padding: '6px', marginTop: 8, borderRadius: 8, border: '1.5px dashed #E0DFD9', background: 'transparent', cursor: 'pointer', fontSize: 12, color: '#9B9B97' }}>+ Añadir elemento</button>
                </div>
              )}
            </div>
          );
        })}
      </div>

      {/* Añadir capítulo */}
      <button onClick={addCap} style={{ width: '100%', padding: '9px', borderRadius: 9, border: '1.5px dashed #E0DFD9', background: 'transparent', cursor: 'pointer', fontSize: 12, color: '#9B9B97', marginTop: 8 }}>+ Añadir capítulo</button>

      {confirmacion && <ConfirmMini titulo={confirmacion.titulo} texto={confirmacion.texto} onSi={confirmacion.onSi} onNo={() => setConfirmacion(null)} />}
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
  const [confirmacion, setConfirmacion] = useState(null);
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
                    <button onClick={e => { e.stopPropagation(); setConfirmacion({ titulo: 'Eliminar lotificación', texto: `Vas a eliminar "${el.nombre}" y todos sus datos. Esta acción no se puede deshacer.`, onSi: () => { eliminar(el.id); setConfirmacion(null); } }); }} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#D4D3CE', fontSize: 16, padding: '0 2px', lineHeight: 1 }}>×</button>
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
      {confirmacion && <ConfirmMini titulo={confirmacion.titulo} texto={confirmacion.texto} onSi={confirmacion.onSi} onNo={() => setConfirmacion(null)} />}
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
      {/* Banner con fondo arquitectónico animado */}
      <div className="dash-banner" style={{ borderBottom: '1px solid #ECEAE4', padding: isMobile ? '18px 16px' : '22px 22px', paddingTop: isMobile ? 'calc(18px + env(safe-area-inset-top))' : '22px', flexShrink: 0 }}>
        <div className="arch-grid" />
        <div className="dash-banner-ring db-r1" />
        <div className="dash-banner-ring db-r2" />
        <div className="dash-beam" />
        <div className="dash-title" style={{ position: 'relative', zIndex: 2 }}>
          <div style={{ fontSize: isMobile ? 18 : 20, fontWeight: 700, color: '#F2F1ED', letterSpacing: '0.01em', display: 'flex', alignItems: 'baseline', gap: 7 }}>
            Alertas <span style={{ fontSize: 13, color: '#8AA88A' }}>.</span>
          </div>
          <div style={{ fontSize: 12, color: 'rgba(255,255,255,0.5)', marginTop: 3 }}>{total > 0 ? `${total} aviso${total !== 1 ? 's' : ''} requieren tu atención` : 'Sin avisos pendientes'}</div>
        </div>
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
  const [voLocal, setVoLocal] = useState(null);
  const [cargandoVO, setCargandoVO] = useState(false);

  // DEBUG TEMPORAL — esborrar després
  useEffect(() => {
    console.log('[ActaVO] obra.id:', obra.id, 'obra.actaVO:', obra.actaVO ? `OK (num:${obra.actaVO.num}, seccions:${obra.actaVO.secciones?.length}, temes:${obra.actaVO.secciones?.reduce((a,s)=>a+(s.temas?.length||0),0)})` : 'NULL');
    console.log('[ActaVO] voLocal:', voLocal ? `OK (num:${voLocal.num}, seccions:${voLocal.secciones?.length}, temes:${voLocal.secciones?.reduce((a,s)=>a+(s.temas?.length||0),0)})` : 'NULL');
  }, [obra.actaVO, voLocal]);

  // Si obra.actaVO és null (Fase 1 o mòdul no carregat), el carrega directament de Supabase
  useEffect(() => {
    if (obra.actaVO !== null && obra.actaVO !== undefined) {
      setVoLocal(migrateVO(obra.actaVO));
    } else if (window.db && obra.id) {
      setCargandoVO(true);
      window.db.getModulo('actas_vo', obra.id).then(rows => {
        const raw = rows?.[0]?.data || null;
        if (raw) {
          // Netejar fotos amb base64 antic (massa pesades) — conservar només path i url
          const netejat = {
            ...raw,
            estadoObra: {
              ...( raw.estadoObra || {} ),
              fotos: (raw.estadoObra?.fotos || []).map(f => ({
                id: f.id, path: f.path,
                // Regenerar URL pública si té path, sinó conservar URL existent
                url: f.path
                  ? `${window._supabaseUrl}/storage/v1/object/public/plaat-fotos/${f.path}`
                  : (f.url || ''),
                // Eliminar el base64 (data) que pesa molt
              })),
            },
          };
          const migrated = migrateVO(netejat);
          setVoLocal(migrated);
        } else {
          setVoLocal(migrateVO(null));
        }
        setCargandoVO(false);
      }).catch(err => {
        console.error('Error carregant actaVO:', err);
        setVoLocal(migrateVO(null));
        setCargandoVO(false);
      });
    } else {
      setVoLocal(migrateVO(obra.actaVO));
    }
  }, [obra.id, obra.actaVO]);

  // Sincronitzar quan obra.actaVO canvia des de fora (Fase 2)
  useEffect(() => {
    if (obra.actaVO !== null && obra.actaVO !== undefined) {
      setVoLocal(migrateVO(obra.actaVO));
    }
  }, [obra.actaVO]);

  const vo = voLocal || migrateVO(null);
  const [showEquipo,    setShowEquipo]    = useState(false);
  const [showHistorico, setShowHistorico] = useState(false);
  const [borrar,        setBorrar]        = useState(null);
  const [generando,     setGenerando]     = useState(false);
  const [editandoSec,   setEditandoSec]   = useState(null);
  const [confirmacion,  setConfirmacion]  = useState(null); // id sección editando nombre

  function guardarVO(nuevo) {
    setVoLocal(nuevo);
    onSave({ ...obra, actaVO: nuevo });
  }

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
    guardarVO({ ...vo, secciones: vo.secciones.map(s => s.id !== secId ? s : { ...s, temas: [...(s.temas||[]), { id: uid(), num: nextNum(s), resuelto: false, resueltoEnActa: null, entradas: [{ id: uid(), texto: texto.trim(), estado: 'P', fecha: today(), fin: '', resp: '', actaNum: vo.num }] }] }) });
  }
  function addEntrada(secId, temaId, texto) {
    if (!texto.trim()) return;
    guardarVO({ ...vo, secciones: vo.secciones.map(s => s.id !== secId ? s : { ...s, temas: s.temas.map(t => t.id !== temaId ? t : { ...t, entradas: [...t.entradas, { id: uid(), texto: texto.trim(), estado: 'P', fecha: today(), fin: '', resp: '', actaNum: vo.num }] }) }) });
  }
  function updEntrada(secId, temaId, entId, campo, val) {
    guardarVO({ ...vo, secciones: vo.secciones.map(s => s.id !== secId ? s : { ...s, temas: s.temas.map(t => {
      if (t.id !== temaId) return t;
      const entradas = t.entradas.map(e => e.id !== entId ? e : { ...e, [campo]: val });
      // El tema solo está resuelto cuando TODOS sus comentarios están en R
      const resuelto = entradas.length > 0 && entradas.every(e => e.estado === 'R');
      // Si alguna entrada es nueva de este acta, se ve "N" ahora y "R" el siguiente → desaparece un acta después
      const hayNueva = entradas.some(e => e.actaNum === vo.num);
      const resueltoEnActa = resuelto ? (t.resueltoEnActa || (hayNueva ? vo.num + 1 : vo.num)) : null;
      return { ...t, entradas, resuelto, resueltoEnActa };
    }) }) });
  }
  function updTema(secId, temaId, campo, val) {
    guardarVO({ ...vo, secciones: vo.secciones.map(s => s.id !== secId ? s : { ...s, temas: s.temas.map(t => t.id === temaId ? { ...t, [campo]: val } : t) }) });
  }
  function addFotoEntrada(secId, temaId, entId) {
    pickFiles('image/*', f => guardarVO({ ...vo, secciones: vo.secciones.map(s => s.id !== secId ? s : { ...s, temas: s.temas.map(t => t.id !== temaId ? t : { ...t, entradas: t.entradas.map(e => e.id !== entId ? e : { ...e, fotos: [...(e.fotos||[]), f] }) }) }) }), obra?.id);
  }
  function delFotoEntrada(secId, temaId, entId, fotoId) {
    guardarVO({ ...vo, secciones: vo.secciones.map(s => s.id !== secId ? s : { ...s, temas: s.temas.map(t => t.id !== temaId ? t : { ...t, entradas: t.entradas.map(e => e.id !== entId ? e : { ...e, fotos: (e.fotos||[]).filter(ft => ft.id !== fotoId) }) }) }) });
  }
  function delEntrada(secId, temaId, entId) {
    guardarVO({ ...vo, secciones: vo.secciones.map(s => s.id !== secId ? s : { ...s, temas: s.temas.map(t => t.id !== temaId ? t : { ...t, entradas: t.entradas.filter(e => e.id !== entId) }) }) });
  }
  function reabrirTema(secId, temaId) {
    guardarVO({ ...vo, secciones: vo.secciones.map(s => s.id !== secId ? s : { ...s, temas: s.temas.map(t => {
      if (t.id !== temaId) return t;
      // Vuelve a pendiente: última entrada pasa a P, se limpia resuelto
      const entradas = t.entradas.map((e, i) => i === t.entradas.length - 1 ? { ...e, estado: 'P', fin: '' } : e);
      return { ...t, entradas, resuelto: false, resueltoEnActa: null };
    }) }) });
  }
  function delTema(secId, temaId) {
    guardarVO({ ...vo, secciones: vo.secciones.map(s => s.id !== secId ? s : { ...s, temas: s.temas.filter(t => t.id !== temaId) }) });
    setBorrar(null);
  }
  function moverTema(secOrigenId, temaId, secDestinoId) {
    if (secOrigenId === secDestinoId) return;
    const secOrigen = vo.secciones.find(s => s.id === secOrigenId);
    const tema = secOrigen?.temas.find(t => t.id === temaId);
    if (!tema) return;
    guardarVO({
      ...vo,
      secciones: vo.secciones.map(s => {
        if (s.id === secOrigenId) return { ...s, temas: s.temas.filter(t => t.id !== temaId) };
        if (s.id === secDestinoId) return { ...s, temas: [...s.temas, tema] };
        return s;
      }),
    });
  }
  function reordenarTema(secId, temaId, direccion) {
    guardarVO({
      ...vo,
      secciones: vo.secciones.map(s => {
        if (s.id !== secId) return s;
        const idx = s.temas.findIndex(t => t.id === temaId);
        const nuevoIdx = idx + direccion;
        if (nuevoIdx < 0 || nuevoIdx >= s.temas.length) return s;
        const temas = [...s.temas];
        [temas[idx], temas[nuevoIdx]] = [temas[nuevoIdx], temas[idx]];
        return { ...s, temas };
      }),
    });
  }

  // Estado de obra
  function updEstado(campo, val) { guardarVO({ ...vo, estadoObra: { ...(vo.estadoObra||{}), [campo]: val } }); }
  function addFotoEstado() {
    pickFiles('image/*', f => guardarVO({ ...vo, estadoObra: { ...(vo.estadoObra||{}), fotos: [...((vo.estadoObra||{}).fotos||[]), f] } }), obra?.id);
  }
  function delFotoEstado(id) { guardarVO({ ...vo, estadoObra: { ...(vo.estadoObra||{}), fotos: (vo.estadoObra.fotos||[]).filter(f => f.id !== id) } }); }

  const [showIdioma, setShowIdioma] = useState(false);
  const [versionExport, setVersionExport] = useState(null); // 'v1' | 'v2'

  async function exportar(idioma) {
    setShowIdioma(false);
    setGenerando(true);
    guardarVO({ ...vo, num: vo.num + 1 });
    try {
      if (versionExport === 'v2') {
        await generarActaVO_v2(obra, { ...vo, num: vo.num }, idioma);
      } else {
        await generarActaVO(obra, { ...vo, num: vo.num }, idioma);
      }
    } catch (e) {
      if (!e.message?.includes('Load failed') && !e.message?.includes('fetch')) {
        alert('Error al exportar: ' + e.message);
      }
    }
    setGenerando(false);
    setVersionExport(null);
  }

  // Activos = no resueltos en acta anterior; resueltos = para histórico
  const todosResueltos = vo.secciones.flatMap(s => (s.temas||[]).filter(t => t.resuelto).map(t => ({ ...t, _secId: s.id })));
  const activosPorSec = id => (vo.secciones.find(s => s.id === id)?.temas||[]).filter(t => !(t.resuelto && t.resueltoEnActa && t.resueltoEnActa < vo.num));

  if (cargandoVO) return (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '60px 20px', flexDirection: 'column', gap: 12 }}>
      <div style={{ width: 28, height: 28, border: '2.5px solid #E0DFD9', borderTopColor: '#1C1C1A', borderRadius: '50%', animation: 'spin .7s linear infinite' }} />
      <div style={{ fontSize: 13, color: '#A5A5A0' }}>Carregant acta...</div>
    </div>
  );

  return (
    <div>
      {/* Cabecera */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', marginBottom: 14 }}>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 16, fontWeight: 600, color: '#141412' }}>Acta de Visita de Obra</div>
          <div style={{ fontSize: 12, color: '#9B9B97', display: 'flex', alignItems: 'center', gap: 6, marginTop: 3 }}>
            Nº de acta:
            <input type="number" min="1" value={vo.num} onChange={e => guardarVO({ ...vo, num: Math.max(1, parseInt(e.target.value || '1', 10)) })}
              style={{ width: 60, padding: '3px 7px', fontSize: 12, fontWeight: 600, textAlign: 'center' }} />
          </div>
        </div>
        <Btn onClick={() => setShowHistorico(true)}>Resueltos ({todosResueltos.length})</Btn>
        <Btn primary disabled={generando} onClick={() => setShowIdioma(true)}>{generando ? 'Generando...' : `↓ Exportar Acta Nº ${String(vo.num).padStart(2,'0')}`}</Btn>
      </div>

      {showIdioma && !versionExport && (
        <Modal title="Format de l'acta" onClose={() => setShowIdioma(false)} footer={<Btn onClick={() => setShowIdioma(false)}>Cancel·lar</Btn>}>
          <p style={{ fontSize: 13, color: '#6B6B66', marginBottom: 16 }}>Tria el format per exportar l'Acta Nº {String(vo.num).padStart(2,'0')}.</p>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            <div onClick={() => setVersionExport('v2')} style={{ padding: '14px 16px', borderRadius: 10, border: '1.5px solid #18180F', background: '#F5F4F0', cursor: 'pointer' }}>
              <div style={{ fontSize: 13, fontWeight: 600 }}>✦ Format nou PLAAT 2026</div>
              <div style={{ fontSize: 11, color: '#6B6B66', marginTop: 3 }}>Brandbook juny 2026 — Arial, banda negra, capçalera corporativa</div>
            </div>
            <div onClick={() => setVersionExport('v1')} style={{ padding: '14px 16px', borderRadius: 10, border: '1.5px solid #E0DFD9', background: '#fff', cursor: 'pointer' }}>
              <div style={{ fontSize: 13, fontWeight: 500, color: '#6B6B66' }}>Format anterior</div>
              <div style={{ fontSize: 11, color: '#A5A5A0', marginTop: 3 }}>Versió original de l'acta</div>
            </div>
          </div>
        </Modal>
      )}

      {showIdioma && versionExport && (
        <Modal title="Idioma de l'acta" onClose={() => { setShowIdioma(false); setVersionExport(null); }} footer={<Btn onClick={() => { setShowIdioma(false); setVersionExport(null); }}>Cancel·lar</Btn>}>
          <p style={{ fontSize: 13, color: '#6B6B66', marginBottom: 16 }}>Tria l'idioma per exportar l'Acta Nº {String(vo.num).padStart(2,'0')} en format {versionExport === 'v2' ? 'nou PLAAT 2026' : 'anterior'}.</p>
          <div style={{ display: 'flex', gap: 10 }}>
            <Btn full onClick={() => exportar('ca')}>Català</Btn>
            <Btn primary full onClick={() => exportar('es')}>Castellano</Btn>
          </div>
        </Modal>
      )}

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
            {/* Cabecera — mismo grid que filas de datos para alineado perfecto */}
            {!isMobile && (
              <div style={{ display: 'grid', gridTemplateColumns: '2fr 1.4fr 1.6fr 1.8fr 1.3fr 50px 22px', gap: 4, padding: '2px 0 6px' }}>
                {['Rol','Empresa','Nombre','Email','Teléfono','Asistido',''].map((h,i) => (
                  <div key={i} style={{ fontSize: 10, color: '#A5A5A0', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em', paddingLeft: 4 }}>{h}</div>
                ))}
              </div>
            )}
            {(vo.equipo||[]).map(rol => (
              <div key={rol.id} style={{ marginBottom: 10, border: '1px solid #F2F1ED', borderRadius: 8, overflow: 'hidden' }}>
                {/* Nombre del rol — separado, fondo tenue */}
                <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '5px 8px', background: '#F5F4F0', borderBottom: '1px solid #ECEAE4' }}>
                  <input value={rol.nombre} onChange={e => updRol(rol.id, 'nombre', e.target.value)} style={{ flex: 1, fontSize: 11, fontWeight: 600, background: 'transparent', border: 'none', padding: 0, boxShadow: 'none', color: '#141412' }} />
                  <button onClick={() => setConfirmacion({ titulo: 'Eliminar rol', texto: `Vas a eliminar el rol "${rol.nombre}" y todas sus personas.`, onSi: () => { delRol(rol.id); setConfirmacion(null); } })} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#D4D3CE', fontSize: 15, lineHeight: 1 }}>×</button>
                </div>
                {/* Filas de personas — mismo grid que el header */}
                {(rol.personas||[]).map(p => (
                  <div key={p.id} style={{
                    display: isMobile ? 'flex' : 'grid',
                    gridTemplateColumns: isMobile ? undefined : '2fr 1.4fr 1.6fr 1.8fr 1.3fr 50px 22px',
                    flexDirection: isMobile ? 'column' : undefined,
                    gap: 4, padding: '5px 8px',
                    borderBottom: '1px solid #F9F8F5', alignItems: 'center'
                  }}>
                    {isMobile && <div style={{ fontSize: 10, color: '#A5A5A0', marginBottom: 3 }}>Empresa / Nombre / Email / Tel</div>}
                    {/* Primera columna: vacía (el rol ya está arriba) */}
                    {!isMobile && <div />}
                    <input placeholder="Empresa" value={p.empresa||''} onChange={e => updPersona(rol.id, p.id, 'empresa', e.target.value)} style={{ fontSize: 12 }} />
                    <input placeholder="Nombre" value={p.nombre||''} onChange={e => updPersona(rol.id, p.id, 'nombre', e.target.value)} style={{ fontSize: 12 }} />
                    <input placeholder="Email" value={p.email||''} onChange={e => updPersona(rol.id, p.id, 'email', e.target.value)} style={{ fontSize: 12 }} />
                    <input placeholder="Teléfono" value={p.tel||''} onChange={e => updPersona(rol.id, p.id, 'tel', e.target.value)} style={{ fontSize: 12 }} />
                    <label style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 4, cursor: 'pointer', fontSize: 12, color: '#52524E' }}>
                      <input type="checkbox" checked={!!p.asistio} onChange={e => updPersona(rol.id, p.id, 'asistio', e.target.checked)} style={{ width: 14, height: 14 }} />
                      {isMobile ? 'Asistido' : ''}
                    </label>
                    <button onClick={() => setConfirmacion({ titulo: 'Eliminar persona', texto: `Vas a eliminar a "${p.nombre||'esta persona'}".`, onSi: () => { delPersona(rol.id, p.id); setConfirmacion(null); } })} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#D4D3CE', fontSize: 15, lineHeight: 1, textAlign: 'center' }}>×</button>
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
              <img src={fotoSrc(f)} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover', borderRadius: 7, display: 'block' }} />
              <button onClick={() => setConfirmacion({ titulo: 'Eliminar foto', texto: 'Vas a eliminar esta foto del estado de obra.', onSi: () => { delFotoEstado(f.id); setConfirmacion(null); } })} style={{ position: 'absolute', top: 3, right: 3, background: 'rgba(0,0,0,0.55)', color: '#fff', border: 'none', borderRadius: '50%', width: 18, height: 18, cursor: 'pointer', fontSize: 11, display: 'flex', alignItems: 'center', justifyContent: 'center', lineHeight: 1 }}>×</button>
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
                : <><span style={{ fontSize: 12, fontWeight: 600, color: '#141412', flex: 1 }}>{sec.titulo}</span><button onClick={() => setEditandoSec(sec.id)} title="Editar" style={{ background:'none', border:'none', cursor:'pointer', color:'#C4C3BE', fontSize:12, padding:'0 2px', flexShrink:0 }}>✏️</button></>}
              <span style={{ fontSize: 11, color: '#A5A5A0' }}>{activos.length}</span>
              <button onClick={() => setBorrar({ tipo: 'seccion', id: sec.id, label: sec.titulo })} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#D4D3CE', fontSize: 15, lineHeight: 1 }}>×</button>
            </div>
            {/* Temas activos */}
            {activos.map(t => {
              // Estado agregado del tema: Resuelto solo si TODOS en R; si hay algún P → Pendiente; resto → Informativo
              const allR = t.entradas.length > 0 && t.entradas.every(e => e.estado === 'R');
              const anyP = t.entradas.some(e => e.estado === 'P');
              const estKey = allR ? 'R' : anyP ? 'P' : 'I';
              const est = ESTADOS_VO[estKey];
              return (
                <TemaVO key={t.id} t={t} est={est} secId={sec.id} voNum={vo.num}
                  secciones={vo.secciones}
                  onUpdEntrada={(tId,eId,campo,val) => updEntrada(sec.id, tId, eId, campo, val)}
                  onUpdTema={(tId,campo,val) => updTema(sec.id, tId, campo, val)}
                  onAddEntrada={(tId,txt) => addEntrada(sec.id, tId, txt)}
                  onAddFoto={(tId,eId) => addFotoEntrada(sec.id, tId, eId)}
                  onDelFoto={(tId,eId,fId) => delFotoEntrada(sec.id, tId, eId, fId)}
                  onDelEntrada={(tId,eId) => delEntrada(sec.id, tId, eId)}
                  onMover={(secDestId) => moverTema(sec.id, t.id, secDestId)}
                  onReordenar={(dir) => reordenarTema(sec.id, t.id, dir)}
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
                  <Btn sm onClick={() => { reabrirTema(t._secId, t.id); }}>Reabrir</Btn>
                  <span style={{ fontSize: 11, color: '#A5A5A0', marginLeft: 'auto' }}>Acta {String(t.resueltoEnActa||'—').padStart(2,'0')}</span>
                </div>
                <div style={{ fontSize: 13, color: '#18180F', lineHeight: 1.5 }}>{t.entradas[t.entradas.length-1]?.texto || ''}</div>
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
      {confirmacion && <ConfirmMini titulo={confirmacion.titulo} texto={confirmacion.texto} onSi={confirmacion.onSi} onNo={() => setConfirmacion(null)} />}
    </div>
  );
}

// Componente de un tema (para evitar closures stale en los selects)
function TemaVO({ t, est, secId, voNum, secciones, onUpdEntrada, onUpdTema, onAddEntrada, onAddFoto, onDelFoto, onDelEntrada, onMover, onReordenar, onDel }) {
  const [abierto, setAbierto] = useState(false);
  const [editNum, setEditNum] = useState(false);
  const [editEnt, setEditEnt] = useState(null); // id entrada en edición
  const [txtEdit, setTxtEdit] = useState('');
  const [confirmFoto, setConfirmFoto] = useState(null);
  const ult = t.entradas[t.entradas.length - 1] || { texto: '', actaNum: null, estado: 'P' };
  const ultEsNueva = ult.actaNum === voNum;
  return (
    <div style={{ border: `1px solid ${abierto ? '#C5C4BE' : '#E8E7E1'}`, borderRadius: 9, marginBottom: 6, overflow: 'hidden' }}>
      <div onClick={() => setAbierto(v => !v)} style={{ display: 'flex', alignItems: 'center', gap: 9, padding: '9px 12px', cursor: 'pointer', background: abierto ? '#FAFAF8' : '#fff' }}>
        <span style={{ fontSize: 11, fontWeight: 700, color: '#52524E', flexShrink: 0, minWidth: 34 }}>{t.num}</span>
        <span style={{ flex: 1, fontSize: 13, color: '#18180F', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{ult.texto}</span>
        {ultEsNueva
          ? <Pill label="N — Nueva" bg="#F2F1ED" color="#52524E" />
          : <Pill label={est.label} bg={est.bg} color={est.color} />}
      </div>
      {abierto && (
        <div className="fade" style={{ padding: '4px 12px 12px', borderTop: '1px solid #F2F1ED' }}>
          {/* Editar número de tema */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 8, paddingTop: 6, flexWrap: 'wrap' }}>
            <span style={{ fontSize: 11, color: '#9B9B97' }}>Nº tema:</span>
            {editNum
              ? <input autoFocus value={t.num} onChange={e => onUpdTema(t.id, 'num', e.target.value)} onBlur={() => setEditNum(false)}
                  style={{ width: 70, fontSize: 12, padding: '3px 6px' }} />
              : <><span style={{ fontSize: 12, fontWeight: 600, color: '#141412', padding: '2px 6px', borderRadius: 5, border: '1px solid #E0DFD9' }}>{t.num}</span><button onClick={() => setEditNum(true)} title="Editar" style={{ background:'none', border:'none', cursor:'pointer', color:'#C4C3BE', fontSize:11, padding:'0 2px' }}>✏️</button></>}

            {/* Reordenar dentro de la sección */}
            <button onClick={() => onReordenar(-1)} title="Subir" style={{ background:'none', border:'1px solid #E0DFD9', borderRadius:6, cursor:'pointer', color:'#6B6B66', fontSize:12, padding:'2px 7px', marginLeft:8 }}>↑</button>
            <button onClick={() => onReordenar(1)} title="Bajar" style={{ background:'none', border:'1px solid #E0DFD9', borderRadius:6, cursor:'pointer', color:'#6B6B66', fontSize:12, padding:'2px 7px' }}>↓</button>

            {/* Mover a otra sección */}
            {secciones && secciones.length > 1 && (
              <>
                <span style={{ fontSize: 11, color: '#9B9B97', marginLeft: 10 }}>Mover a:</span>
                <select value={secId} onChange={e => onMover(e.target.value)} style={{ width: 'auto', fontSize: 11, padding: '3px 7px', borderRadius: 6 }}>
                  {secciones.map(s => <option key={s.id} value={s.id}>{s.codigo} {s.titulo}</option>)}
                </select>
              </>
            )}
          </div>

          {t.entradas.map(en => {
            const esNueva = en.actaNum === voNum;
            return (
              <div key={en.id} style={{ padding: '8px 0', borderBottom: '1px solid #F5F4F0' }}>
                <div style={{ display: 'flex', gap: 7, alignItems: 'flex-start' }}>
                  <span style={{ fontSize: 11, color: '#A5A5A0', whiteSpace: 'nowrap', paddingTop: 2, minWidth: 36 }}>{fmtShort(en.fecha)}</span>
                  <div style={{ flex: 1 }}>
                    {/* Texto editable */}
                    {editEnt === en.id
                      ? <div style={{ marginBottom: 6 }}>
                          <textarea autoFocus value={txtEdit} onChange={e => setTxtEdit(e.target.value)} style={{ minHeight: 54, marginBottom: 6 }} />
                          <div style={{ display: 'flex', gap: 6 }}>
                            <Btn sm primary onClick={() => { onUpdEntrada(t.id, en.id, 'texto', txtEdit.trim()); setEditEnt(null); }}>Guardar</Btn>
                            <Btn sm onClick={() => setEditEnt(null)}>✕</Btn>
                          </div>
                        </div>
                      : <div style={{ display: 'flex', alignItems: 'flex-start', gap: 5, marginBottom: 5 }}>
                          <div style={{ flex: 1, fontSize: 13, color: '#18180F', lineHeight: 1.5 }}>{en.texto}</div>
                          <button onClick={() => { setEditEnt(en.id); setTxtEdit(en.texto); }} title="Editar" style={{ background:'none', border:'none', cursor:'pointer', color:'#C4C3BE', fontSize:12, padding:'0 2px', flexShrink:0, marginTop:1 }}>✏️</button>
                        </div>}
                    <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', alignItems: 'center' }}>
                      {esNueva && <span style={{ fontSize: 10, padding: '2px 7px', borderRadius: 6, background: '#F2F1ED', color: '#52524E', fontWeight: 700, border: '1px solid #E0DFD9' }}>N</span>}
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
                    {esNueva && <div style={{ fontSize: 10.5, color: '#9B9B97', marginTop: 4 }}>En esta acta aparece como "N". En la siguiente mostrará el estado elegido.</div>}
                    {/* Fotos del comentario */}
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 8 }}>
                      {(en.fotos||[]).map(ft => (
                        <div key={ft.id} style={{ position: 'relative', width: 64, height: 48 }}>
                          <img src={fotoSrc(ft)} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover', borderRadius: 6, display: 'block' }} />
                          <button onClick={() => setConfirmFoto({ eId: en.id, fId: ft.id })} style={{ position: 'absolute', top: 2, right: 2, background: 'rgba(0,0,0,0.55)', color: '#fff', border: 'none', borderRadius: '50%', width: 16, height: 16, cursor: 'pointer', fontSize: 10, display: 'flex', alignItems: 'center', justifyContent: 'center', lineHeight: 1 }}>×</button>
                        </div>
                      ))}
                      <button onClick={() => onAddFoto(t.id, en.id)} style={{ width: 64, height: 48, borderRadius: 6, border: '1.5px dashed #E0DFD9', background: '#FAFAF8', cursor: 'pointer', fontSize: 11, color: '#9B9B97', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 3 }}>+ foto</button>
                    </div>
                  </div>
                  {/* Botón eliminar entrada */}
                  <button onClick={() => onDelEntrada(t.id, en.id)} title="Eliminar comentario" style={{ background:'none', border:'none', cursor:'pointer', color:'#C4C3BE', fontSize:16, lineHeight:1, padding:'2px 3px', flexShrink:0, marginTop:1 }}>×</button>
                </div>
              </div>
            );
          })}
          <NuevaEntrada onAdd={txt => onAddEntrada(t.id, txt)} />
          <div style={{ marginTop: 8, textAlign: 'right' }}><Btn sm danger onClick={onDel}>Eliminar tema</Btn></div>
        </div>
      )}
      {confirmFoto && <ConfirmMini titulo="Eliminar foto" texto="Vas a eliminar esta foto del comentario." onSi={() => { onDelFoto(t.id, confirmFoto.eId, confirmFoto.fId); setConfirmFoto(null); }} onNo={() => setConfirmFoto(null)} />}
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
const TEXTOS_VO = {
  ca: {
    titulo:    n => `ACTA DE VISITA D'OBRA N.\u00ba ${n}`,
    fecha:     'DATA',
    lugar:     'LLOC',
    fase:      'FASE',
    equipo:    'EQUIP T\u00c8CNIC I DADES DE CONTACTE',
    colRol:    'ROL', colEmp: 'EMPRESA', colNom: 'NOM', colEmail: 'EMAIL', colTel: 'TEL.', colAsis: 'ASSISTIT',
    sec0:      'ESTAT DE L\'OBRA (FOTOGRAFIES)',
    colDesc:   'DESCRIPCI\u00d3', colEst: 'Estat', colIni: 'Inici', colFi: 'Fi', colRes: 'Res.',
    nota:      'NOTA: La present acta s\'entendr\u00e0 com a conforme en cas de no manifestar comentaris en el termini de 48 hores despr\u00e9s de la seva difusi\u00f3.',
    conforme:  'Conforme, signatura i data:',
    firmas:    ['PROMOTOR', 'DIRECCI\u00d3 D\'OBRA', 'DIRECCI\u00d3 D\'EXECUCI\u00d3\nCOORD. SEGURETAT', 'CONTRACTISTA'],
    pie:       ['Plaat Arquitectura T\u00e8cnica', 'Barcelona \u2013 Madrid', 'www.plaat.es'],
  },
  es: {
    titulo:    n => `ACTA DE VISITA DE OBRA N.\u00ba ${n}`,
    fecha:     'FECHA',
    lugar:     'LUGAR',
    fase:      'FASE',
    equipo:    'EQUIPO T\u00c9CNICO Y DATOS DE CONTACTO',
    colRol:    'ROL', colEmp: 'EMPRESA', colNom: 'NOMBRE', colEmail: 'EMAIL', colTel: 'TEL.', colAsis: 'ASISTIDO',
    sec0:      'ESTADO DE LA OBRA (FOTOGRAF\u00cdAS)',
    colDesc:   'DESCRIPCI\u00d3N', colEst: 'Estado', colIni: 'Inicio', colFi: 'Fin', colRes: 'Res.',
    nota:      'NOTA: La presente acta se entender\u00e1 como conforme en caso de no manifestar comentarios en el plazo de 48 horas tras su difusi\u00f3n.',
    conforme:  'Conforme, firma y fecha:',
    firmas:    ['PROMOTOR', 'DIRECCI\u00d3N DE OBRA', 'DIRECCI\u00d3N DE EJECUCI\u00d3N\nCOORD. SEGURIDAD', 'CONTRATISTA'],
    pie:       ['Plaat Arquitectura T\u00e9cnica', 'Barcelona \u2013 Madrid', 'www.plaat.es'],
  },
};

async function generarActaVO(obra, vo, idioma = 'ca') {
  if (!window.jspdf) {
    await new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
      s.onload = res; s.onerror = () => rej(new Error('No se pudo cargar jsPDF'));
      document.head.appendChild(s);
    });
  }
  const { jsPDF } = window.jspdf;
  const T = TEXTOS_VO[idioma] || TEXTOS_VO.ca;
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const PW = 210, PH = 297, M = 14, CW = PW - M * 2;
  const LW = 0.25; // line width uniforme para todos los bordes
  const NEGRO = [0,0,0], GRIS = [217,217,217], GRIS_H = [232,232,228];
  const C_P = [255,235,200], C_R = [200,240,200], C_I = [210,228,255];
  const num = String(vo.num).padStart(2,'0');
  const fecha = new Date().toLocaleDateString('es-ES');
  let y = 0;

  // ── Helpers ─────────────────────────────────────────────────────────────────
  function setLW() { doc.setLineWidth(LW); doc.setDrawColor(...NEGRO); }
  function rectF(x,yy,w,h,fill) { if(fill){ doc.setFillColor(...fill); doc.rect(x,yy,w,h,'F'); } setLW(); doc.rect(x,yy,w,h); }
  // Texto envuelto con padding y alineación
  function wText(x, yy, w, h, txt, opts={}) {
    if (txt===null || txt===undefined || txt==='') return;
    const sz = opts.size||8; const bold = opts.bold||false; const center = opts.center||false;
    doc.setFont('helvetica', bold ? 'bold':'normal'); doc.setFontSize(sz); doc.setTextColor(0,0,0);
    const ll = doc.splitTextToSize(String(txt), w - 3.5);
    if (center) {
      const bh = ll.length*(sz*0.352645+0.5);
      let ty = yy + h/2 - bh/2 + sz*0.352645;
      ll.forEach(l => { doc.text(l, x+w/2, ty, {align:'center'}); ty += sz*0.352645+0.5; });
    } else {
      let ty = yy + 3 + sz*0.352645;
      ll.forEach(l => { doc.text(l, x+2, ty); ty += sz*0.352645+0.5; });
    }
  }
  // Calcula altura necesaria para un texto en ancho w (fuente activa)
  function calcH(txt, w, sz) {
    if (!txt) return 0;
    doc.setFontSize(sz||8);
    return doc.splitTextToSize(String(txt), w-3.5).length * ((sz||8)*0.352645+0.5);
  }
  function rowH(pairs, minH=7) { // pairs = [{txt,w,sz}]
    return Math.max(minH, Math.max(...pairs.map(p=>calcH(p.txt,p.w,p.sz||8)))+4);
  }
  function checkbox(cx, yy, h, checked) {
    const s=4, x=cx-s/2, y2=yy+h/2-s/2;
    setLW(); doc.rect(x,y2,s,s);
    if (checked) { doc.setLineWidth(0.5); doc.line(x+0.7,y2+s/2,x+s*0.4,y2+s*0.85); doc.line(x+s*0.4,y2+s*0.85,x+s-0.5,y2+0.5); setLW(); }
  }
  function checkPage(h) { if (y+h > PH-18) { doc.addPage(); cabecera(); pie(); y=20; } }
  function cabecera() {
    doc.setFont('helvetica','bold'); doc.setFontSize(7.5); doc.setTextColor(0,0,0);
    doc.text(T.titulo(num), M, 10);
    doc.setFontSize(22); doc.text('Plaat.', PW-M, 12, {align:'right'});
  }
  function pie() {
    doc.setFont('helvetica','normal'); doc.setFontSize(7); doc.setTextColor(0,0,0);
    doc.text(T.pie[0], M, PH-8); doc.text(T.pie[1], PW/2, PH-8, {align:'center'}); doc.text(T.pie[2], PW-M, PH-8, {align:'right'});
  }

  cabecera(); pie(); y=18;

  // ── 1. Título ────────────────────────────────────────────────────────────
  rectF(M, y, CW, 9, GRIS);
  doc.setFont('helvetica','bold'); doc.setFontSize(12); doc.setTextColor(0,0,0);
  doc.text(T.titulo(num), M+3, y+6.2);
  y += 9;

  // Fecha / Lugar / Fase
  [[ T.fecha, fecha], [T.lugar, vo.lugar||'Obra'], [T.fase, vo.fase||'']].forEach(([k,v]) => {
    const h = Math.max(8, calcH(v, CW-42)+4);
    checkPage(h); rectF(M,y,42,h,GRIS); rectF(M+42,y,CW-42,h,null);
    wText(M,y,42,h,k,{bold:true,size:9}); wText(M+42,y,CW-42,h,v,{size:9});
    y+=h;
  });
  y+=5;

  // ── 2. Equipo técnico ────────────────────────────────────────────────────
  // Cabecera equipo (sin fila extra de columnas negras — solo título)
  checkPage(10);
  rectF(M,y,CW,8,GRIS);
  doc.setFont('helvetica','bold'); doc.setFontSize(9); doc.setTextColor(0,0,0);
  doc.text(T.equipo, M+2, y+5.3); y+=8;

  // Anchos: Rol 38 | Empresa 26 | Nombre 28 | Email 44 | Tel 24 | Asist 22 = 182
  const eW=[38,26,28,44,24,22]; const ex=[M]; eW.forEach((w,i)=>ex.push(ex[i]+w));
  // Cabecera columnas (fila con fondo gris claro)
  checkPage(6);
  [T.colRol,T.colEmp,T.colNom,T.colEmail,T.colTel,T.colAsis].forEach((t,i)=>{
    rectF(ex[i],y,eW[i],6,GRIS_H);
    doc.setFont('helvetica','bold'); doc.setFontSize(6.5); doc.setTextColor(80,80,75);
    doc.text(t, ex[i]+eW[i]/2, y+4, {align:'center'});
  }); y+=6;

  (vo.equipo||[]).forEach(rol=>{
    const ps=(rol.personas&&rol.personas.length)?rol.personas:[{empresa:'',nombre:'',email:'',tel:'',asistio:false}];
    const rowHs=ps.map(p=>rowH([{txt:p.nombre,w:eW[2]},{txt:p.email,w:eW[3]},{txt:p.tel,w:eW[4]}],7));
    const rolH=rowHs.reduce((a,b)=>a+b,0);
    checkPage(rolH);
    rectF(ex[0],y,eW[0],rolH,GRIS); wText(ex[0],y,eW[0],rolH,rol.nombre||'',{bold:true,size:7.5,center:true});
    // Empresa: span hasta que cambie
    let py=y, gi=0;
    while(gi<ps.length){
      let gj=gi;
      while(gj+1<ps.length&&(ps[gj+1].empresa||'')===(ps[gi].empresa||''))gj++;
      const gh=rowHs.slice(gi,gj+1).reduce((a,b)=>a+b,0);
      rectF(ex[1],py,eW[1],gh,null); wText(ex[1],py,eW[1],gh,ps[gi].empresa||'',{size:7.5,center:true});
      py+=gh; gi=gj+1;
    }
    py=y;
    ps.forEach((p,pi)=>{
      const rh=rowHs[pi];
      rectF(ex[2],py,eW[2],rh,null); wText(ex[2],py,eW[2],rh,p.nombre||'',{size:7.5});
      rectF(ex[3],py,eW[3],rh,null); wText(ex[3],py,eW[3],rh,p.email||'',{size:7.5});
      rectF(ex[4],py,eW[4],rh,null); wText(ex[4],py,eW[4],rh,p.tel||'',{size:7.5});
      rectF(ex[5],py,eW[5],rh,null); checkbox(ex[5]+eW[5]/2,py,rh,!!p.asistio);
      py+=rh;
    });
    y+=rolH;
  }); y+=6;

  // ── 3. Sección 0: Estado de la obra ────────────────────────────────────
  const eo=vo.estadoObra||{};
  if(eo.descripcion||(eo.fotos||[]).length>0){
    checkPage(14);
    rectF(M,y,16,8,GRIS); rectF(M+16,y,CW-16,8,GRIS);
    doc.setFont('helvetica','bold'); doc.setFontSize(9); doc.setTextColor(0,0,0);
    doc.text('0',M+8,y+5.3,{align:'center'}); doc.text(T.sec0,M+19,y+5.3);
    y+=8;
    if(eo.descripcion){
      doc.setFontSize(8.5); const dl=doc.splitTextToSize(eo.descripcion,CW-4);
      const dh=Math.max(10,dl.length*4.2+5); checkPage(dh);
      rectF(M,y,CW,dh,null);
      doc.setFont('helvetica','normal'); doc.text(dl,M+2,y+5); y+=dh+2;
    }
    const fotos=eo.fotos||[], fW=(CW-4)/2;
    for(let fi=0;fi<fotos.length;fi+=2){
      const pair=[fotos[fi],fotos[fi+1]].filter(Boolean);
      const dims=pair.map(f=>{try{const pr=doc.getImageProperties(f.url||f.data);const r=pr.height/pr.width;const h=Math.min(fW*r,72);return{w:h/r,h};}catch(e){return{w:fW,h:56};}});
      const rh=Math.max(...dims.map(d=>d.h)); checkPage(rh+4);
      pair.forEach((f,pi)=>doc.addImage(f.url||f.data,'JPEG',M+pi*(fW+4),y,dims[pi].w,dims[pi].h));
      y+=rh+4;
    } y+=4;
  }

  // ── 4. Secciones de temas ────────────────────────────────────────────────
  // Anchos: Núm 16 | Desc variable | Estado 16 | Inicio 18 | Fin 18 | Res 14
  const cN=16, cE=16, cIn=18, cFi=18, cR=14, cD=CW-cN-cE-cIn-cFi-cR;
  const cx=[M,M+cN,M+cN+cD,M+cN+cD+cE,M+cN+cD+cE+cIn,M+cN+cD+cE+cIn+cFi];

  (vo.secciones||[]).forEach(sec=>{
    const activos=(sec.temas||[]).filter(t=>!(t.resuelto&&t.resueltoEnActa&&t.resueltoEnActa<vo.num));
    if(!activos.length) return;
    checkPage(18);
    // Fila cabecera sección: Núm | Título (span todo ancho restante)
    rectF(M,y,cN,8,GRIS); rectF(M+cN,y,CW-cN,8,GRIS);
    doc.setFont('helvetica','bold'); doc.setFontSize(9); doc.setTextColor(0,0,0);
    doc.text(sec.codigo,M+cN/2,y+5.3,{align:'center'}); doc.text(sec.titulo,M+cN+2,y+5.3);
    y+=8;
    // Sub-cabecera: celdas individuales para cada columna
    const subH=6;
    rectF(M,y,cN,subH,GRIS_H);
    rectF(M+cN,y,cD,subH,GRIS_H);   doc.setFont('helvetica','bold'); doc.setFontSize(7); doc.setTextColor(0,0,0); doc.text(T.colDesc,M+cN+cD/2,y+4,{align:'center'});
    rectF(M+cN+cD,y,cE,subH,GRIS_H);   doc.text(T.colEst, M+cN+cD+cE/2, y+4,{align:'center'});
    rectF(M+cN+cD+cE,y,cIn,subH,GRIS_H); doc.text(T.colIni, M+cN+cD+cE+cIn/2,y+4,{align:'center'});
    rectF(M+cN+cD+cE+cIn,y,cFi,subH,GRIS_H); doc.text(T.colFi, M+cN+cD+cE+cIn+cFi/2,y+4,{align:'center'});
    rectF(M+cN+cD+cE+cIn+cFi,y,cR,subH,GRIS_H); doc.text(T.colRes,M+cN+cD+cE+cIn+cFi+cR/2,y+4,{align:'center'});
    y+=subH;

    activos.forEach(t=>{
      const fW2=(cD-5)/2;
      const ed=t.entradas.map(en=>{
        const esNueva=en.actaNum===vo.num;
        const estado=esNueva?'N':(en.estado||'P');
        const fill=esNueva?null:(estado==='R'?C_R:estado==='I'?C_I:C_P);
        doc.setFontSize(8); const lines=doc.splitTextToSize(en.texto||'',cD-3.5);
        const textH=lines.length*(8*0.352645+0.5)+4;
        const fotos=en.fotos||[]; const fotoRows=[]; let fotosH=0;
        for(let i=0;i<fotos.length;i+=2){
          const pair=[fotos[i],fotos[i+1]].filter(Boolean);
          const dims=pair.map(f=>{try{const pr=doc.getImageProperties(f.url||f.data);const r=pr.height/pr.width;const h=Math.min(fW2*r,38);return{w:h/r,h};}catch(e){return{w:fW2,h:30};}});
          const rh=Math.max(...dims.map(d=>d.h));
          fotoRows.push({pair,dims,rh}); fotosH+=rh+2;
        }
        const h=Math.max(7,textH+(fotosH>0?fotosH+2:0));
        return{en,esNueva,estado,fill,lines,textH,fotoRows,h};
      });
      const temaH=ed.reduce((a,e)=>a+e.h,0);
      checkPage(temaH);

      // Fondo color por entrada
      let ey=y;
      ed.forEach(e=>{if(e.fill){doc.setFillColor(...e.fill);doc.rect(M+cN,ey,CW-cN,e.h,'F');}ey+=e.h;});

      // Num centrado en el bloque del tema
      doc.setFont('helvetica','bold'); doc.setFontSize(8); doc.setTextColor(0,0,0);
      doc.text(t.num,M+cN/2,y+ed[0].h/2+1.2,{align:'center'});

      ey=y;
      ed.forEach(e=>{
        const bold=e.esNueva;
        doc.setFont('helvetica',bold?'bold':'normal'); doc.setFontSize(8); doc.setTextColor(0,0,0);
        let ty=ey+3+8*0.352645;
        e.lines.forEach(l=>{doc.text(l,M+cN+2,ty);ty+=8*0.352645+0.5;});
        let fy=ey+e.textH;
        e.fotoRows.forEach(row=>{row.pair.forEach((f,pi)=>{const src=f.url||f.data;if(src)try{doc.addImage(src,'JPEG',M+cN+2+pi*(fW2+1),fy,row.dims[pi].w,row.dims[pi].h);}catch(e){}});fy+=row.rh+2;});
        const midY=ey+Math.min(e.h,e.textH)/2+1.2;
        doc.setFont('helvetica','bold'); doc.setFontSize(8);
        doc.text(e.estado,M+cN+cD+cE/2,midY,{align:'center'});
        const isR=e.en.estado==='R'&&!e.esNueva;
        doc.setFont('helvetica',bold?'bold':'normal'); doc.setFontSize(7.5);
        doc.text(isR?'':fmtFechaCorta(e.en.fecha),M+cN+cD+cE+cIn/2,midY,{align:'center'});
        doc.text(isR?fmtFechaCorta(e.en.fin||e.en.fecha):'',M+cN+cD+cE+cIn+cFi/2,midY,{align:'center'});
        doc.setFont('helvetica','bold'); doc.setFontSize(8);
        doc.text(e.en.resp||'',M+cN+cD+cE+cIn+cFi+cR/2,midY,{align:'center'});
        ey+=e.h;
      });

      // Bordes del tema: solo exteriores + líneas verticales de columnas (sin horizontales intermedias)
      setLW();
      cx.forEach(x=>doc.line(x,y,x,y+temaH));
      doc.line(M+CW,y,M+CW,y+temaH);
      doc.line(M,y,M+CW,y); doc.line(M,y+temaH,M+CW,y+temaH);
      y+=temaH;
    }); y+=5;
  });

  // ── 5. NOTA + Firmas ──────────────────────────────────────────────────────
  checkPage(36); y+=4;
  doc.setFont('helvetica','italic'); doc.setFontSize(7.5); doc.setTextColor(0,0,0);
  const notaL=doc.splitTextToSize(T.nota,CW); doc.text(notaL,M,y); y+=notaL.length*4.2+4;
  doc.setFont('helvetica','normal'); doc.text(T.conforme,M,y); y+=8;
  const fw=CW/4;
  T.firmas.forEach((f,i)=>{
    setLW(); doc.rect(M+i*fw,y,fw,26);
    doc.setFont('helvetica','bold'); doc.setFontSize(7); doc.setTextColor(0,0,0);
    doc.text(doc.splitTextToSize(f,fw-4),M+i*fw+2,y+4.5);
  });

  const total=doc.getNumberOfPages();
  for(let p=1;p<=total;p++){doc.setPage(p);pie();}
  const lang=idioma==='ca'?'Cat':'Cast';
  // iOS Safari no permite descargas programáticas — abrir en nueva pestaña
  const pdfBlob = doc.output('blob');
  const url = URL.createObjectURL(pdfBlob);
  window.open(url, '_blank');
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}


function fmtFechaCorta(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getFullYear()).slice(2)}`;
}

// ── ACTA VO v2 — FORMAT PLAAT BRANDBOOK 2026 ─────────────────────────────────
async function generarActaVO_v2(obra, vo, idioma = 'ca') {
  if (!window.jspdf) {
    await new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
      s.onload = res; s.onerror = () => rej(new Error('No es pot carregar jsPDF'));
      document.head.appendChild(s);
    });
  }
  const { jsPDF } = window.jspdf;
  const esCA = idioma === 'ca';

  // ── Textos bilingues ──────────────────────────────────────────────────────
  const T = {
    tipusDoc:    esCA ? 'ACTA DE VISITA D\'OBRA' : 'ACTA DE VISITA DE OBRA',
    acta:        esCA ? 'ACTA' : 'ACTA',
    data:        esCA ? 'DATA' : 'FECHA',
    lloc:        esCA ? 'LLOC' : 'LUGAR',
    fase:        esCA ? 'FASE' : 'FASE',
    equip:       esCA ? 'Equip tècnic i dades de contacte' : 'Equipo técnico y datos de contacto',
    pm:          esCA ? 'PROJECT MANAGER (PM)' : 'PROJECT MANAGER (PM)',
    do:          esCA ? 'DIRECCIÓ D\'OBRA (DO)' : 'DIRECCIÓN DE OBRA (DO)',
    deo:         esCA ? 'DIRECCIÓ D\'EXECUCIÓ (DEO)' : 'DIRECCIÓN DE EJECUCIÓN (DEO)',
    atEst:       esCA ? 'ASSISTÈNCIA TÈCNICA ESTRUCTURES' : 'ASISTENCIA TÉCNICA ESTRUCTURAS',
    atInst:      esCA ? 'ASSISTÈNCIA TÈCNICA INSTAL·LACIONS' : 'ASISTENCIA TÉCNICA INSTALACIONES',
    css:         esCA ? 'COORDINACIÓ DE SEGURETAT (CSS)' : 'COORDINACIÓN DE SEGURIDAD (CSS)',
    ec:          esCA ? 'CONTRACTISTA (EC)' : 'CONTRATISTA (EC)',
    estat0:      esCA ? 'ESTAT DE L\'OBRA (FOTOGRAFIES)' : 'ESTADO DE LA OBRA (FOTOGRAFÍAS)',
    desc:        esCA ? 'DESCRIPCIÓ' : 'DESCRIPCIÓN',
    es:          'ES',
    inici:       esCA ? 'INICI' : 'INICIO',
    fi:          esCA ? 'FI' : 'FIN',
    res:         esCA ? 'RES.' : 'RES.',
    nota:        esCA
      ? 'NOTA: La present acta s\'entendrà com a conforme en cas de no manifestar comentaris en el termini de 48 hores després de la seva difusió.'
      : 'NOTA: La presente acta se entenderá como conforme en caso de no manifestar comentarios en el plazo de 48 horas tras su difusión.',
    llegenda:    esCA
      ? 'Estat: (R) Resolt; (P) Pendent; (N) Nou; (INF) Informatiu'
      : 'Estado: (R) Resuelto; (P) Pendiente; (N) Nuevo; (INF) Informativo',
    conforme:    esCA ? 'Conforme, signatura i data' : 'Conforme, firma y fecha',
    promotor:    esCA ? 'PROMOTOR' : 'PROMOTOR',
    pm_f:        esCA ? 'PROJECT MANAGER' : 'PROJECT MANAGER',
    do_f:        esCA ? 'DIRECCIÓ D\'OBRA' : 'DIRECCIÓN DE OBRA',
    deo_f:       esCA ? 'DIRECCIÓ D\'EXECUCIÓ\nCOORDINADOR DE SEGURETAT' : 'DIRECCIÓN EJECUCIÓN OBRA\nCOORDINADOR DE SEGURIDAD',
    ec_f:        esCA ? 'CONTRACTISTA' : 'CONTRATISTA',
    peu:         'Plaat Arquitectura Tècnica  |  Barcelona - Madrid  |  plaat.es',
    peuAlt:      esCA ? 'Acta de visita d\'obra' : 'Acta de visita de obra',
  };

  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

  // ── Mides brandbook ───────────────────────────────────────────────────────
  const PW = 210, PH = 297;
  const ML = 15, MR = 15, MT = 12.5, MB = 15;
  const CW = PW - ML - MR; // 180mm
  const NEGRO = [0,0,0], BLANC = [255,255,255], GRIS15 = [217,217,217];
  const C_P = [255,246,215], C_R = [230,246,236], C_I = [235,243,255], C_A = [252,194,191];

  const num = String(vo.num).padStart(2,'0');
  const dataAvui = (() => {
    const d = new Date();
    const dd = String(d.getDate()).padStart(2,'0');
    const mm = String(d.getMonth()+1).padStart(2,'0');
    const yyyy = d.getFullYear();
    return `${dd}-${mm}-${yyyy}`;
  })();
  let y = 0;
  let pagActual = 1;

  // ── Helpers ───────────────────────────────────────────────────────────────
  const LW = 0.3;
  function setLW(w) { doc.setLineWidth(w || LW); doc.setDrawColor(0,0,0); }

  function text(x, yy, txt, opts = {}) {
    if (txt === null || txt === undefined || txt === '') return;
    doc.setFont('helvetica', opts.bold ? 'bold' : opts.italic ? 'italic' : 'normal');
    doc.setFontSize(opts.size || 9);
    doc.setTextColor(0,0,0);
    doc.text(String(txt), x, yy, { align: opts.align || 'left', baseline: 'middle' });
  }

  function wrappedH(txt, w, sz) {
    if (!txt) return 0;
    doc.setFontSize(sz || 8.5);
    return doc.splitTextToSize(String(txt), w - 3).length * ((sz || 8.5) * 0.3528 + 0.6);
  }

  function wrappedText(x, yy, w, h, txt, opts = {}) {
    if (!txt) return;
    doc.setFont('helvetica', opts.bold ? 'bold' : opts.italic ? 'italic' : 'normal');
    doc.setFontSize(opts.size || 8.5);
    doc.setTextColor(0,0,0);
    const lines = doc.splitTextToSize(String(txt), w - 3);
    const lh = (opts.size || 8.5) * 0.3528 + 0.6;
    const totalH = lines.length * lh;
    if (opts.center) {
      let ty = yy + h/2 - totalH/2 + lh * 0.8;
      lines.forEach(l => { doc.text(l, x + w/2, ty, { align: 'center', baseline: 'middle' }); ty += lh; });
    } else {
      let ty = yy + Math.max(3, h/2 - totalH/2) + lh * 0.8;
      lines.forEach(l => { doc.text(l, x + 2, ty, { baseline: 'middle' }); ty += lh; });
    }
  }

  function fillRect(x, yy, w, h, fill) {
    if (fill) { doc.setFillColor(...fill); doc.rect(x, yy, w, h, 'F'); }
    setLW(); doc.rect(x, yy, w, h, 'S');
  }

  function hLine(yy) { setLW(); doc.line(ML, yy, ML + CW, yy); }

  function checkPage(h) {
    if (y + h > PH - MB - 12) {
      doc.addPage();
      pagActual++;
      dibuixarCapçalera(false);
      dibuixarPeu();
      y = MT + 14;
    }
  }

  // Logo Plaat PNG en base64
  // Logo carregat via fetch a la funció dibuixarCapçalera

  function dibuixarCapçalera(primeraPag = true) {
    const nomObra = obra.nombre || '';
    const localitzacio = (obra.emplazamiento || obra.direccion || '').toUpperCase();
    const promotor = (obra.propiedad || obra.cliente || '').toUpperCase();

    // Nom obra Arial 16p minúscules — sense línies ni recuadres
    doc.setFont('helvetica', 'normal'); doc.setFontSize(16);
    doc.setTextColor(0,0,0);
    doc.text(nomObra, ML, MT + 5);

    // Localització Arial 7p MAJ
    doc.setFontSize(7); doc.setFont('helvetica', 'normal');
    doc.text(localitzacio, ML, MT + 9);

    // Promotor Arial 7p MAJ
    doc.text(promotor, ML, MT + 12.5);

    // Logo PNG real — 12mm alçada, 15mm marge dret
    // Ratio PNG: 414×128 = 3.234375
    // Logo Plaat comprimit (PNG escala de grisos, 515×152px, ~10KB)
    const LOGO_B64_SMALL = 'iVBORw0KGgoAAAANSUhEUgAAAgMAAACYCAAAAABcjrUiAAAdLUlEQVR42u1deYCUxZV/71X1wDAMAwIDcgsoihjucHjifQSDhBBdgzF47MYVNcYjyq5BsxqMurreRo0r6w1oCEHiAYiCCkiAhAAigsCAnMMxwDB8VfX2j0Hoqu6Z/q5hepiuv5im+3tf1fvVu+sVMhzRwbhr0C48TFSo3tMYITdqccgjTtGUlFl/t84xod5hABKYLAd0IseE+ocBZuCkPzjHhFoelFuCHAZyS5DDQG4JcvZAKHWcc+aOKgyEY6dhBkTMYeGowMDe4IJAJORBFWIMUg4HdR4DffYE9u0SDfMLW7Zvf0LXTo0IQEMOBnUcA2srQv9WtD351DP6FABozJmWdRkDDQ6EsQgREFivWzcdOpw5/NzGYHIORt0dWFgW/reIqBmg48iru/tFAeOOLjusWHH/+bmckZ+FM7ZPH+OaUaTX0opJ0tqHev/kcyKVi/rW5GYV1ohz20TOFxgDRAfemnjF2O6gRY5XNSQFcOPX4vAeQ929eXzSM46ckTFI5rUpt9/ZUMkcu2pkaPn6bdYHk35kYttw8ZhyrFnsHTfoM6lz+qCmjHdK4KEhY3XEYnuWRrn4zEcFmhy7asgm5OQB2YgBYEXq1p9VUA4EdW3EKVMMJCZcWEo6t6r1FwPAnvzo/G9FThLUMVPD/cCX51mlRlJy4SUfHmNyQcM6jQG/klygSbffVWLRsA8SuchfHcYA8q0n6up2Mao9qvTbktUbNAChSZUGXuKT0a/m4gR1GgNX9PPzs/I1C2fMWgcgdBoQvNb7tlzEsC7rgp0qM/8Q87t3H1U26+UpGlJ9QSV/fdrAHAjqsF8gZOYhiI3ShZdOXjAKU0OWzHp0OeQChke7b4gkBWvTe8LcU3VKVaGWy+/POYj1Ij6AgoweNPs+YPchmh7+MhcvrA8YAAASmv5z2jGuPmCsGIs5ZVA/MAAgQF00q53rTmp6Z6HIxYzrCQYApfreB63dyCCZ3+XWtt5gAECqE/+U7xiGGqeuzJmF9QcDIL0BzzjagMWBFyGHgfqDAUh4P7vKCQoZeKNc5szC+oMBEOa/W9kmgRHrPsoJgvqEATLN73NyhYjv5Fa3PmEAhPn5iY4g4BkVIqcM6hEG0CR+aQsCg2uW5ZRBfcIACB7ZQqP9ydxc4ii+wdmPATRNh4MTMv7siPUrYaOVUlobYyr/ZfjIUjY1TZl31Nwc4iv44eF/MLZ3uJTpCBSVGQOU5gQm65puj8FsQKRpxVIjjTnwi5qTBLFhgHBgy63JiSKGb7YWc03vQpZEAGUbvilZv3XHvj0GGhYUtWjXsXO7phLAGKL0P3PfPRRlJIAdG9auK9m+c+9e4PyCZs07dujcvqCaxhyulEhfaZMiS3TeP2fap3eM1hnYUQsYQFPUb3ry2QLG3euLjahRCSAFwOr58xat+tbhKhZ363/agGMpPTMw6lsZFgLM8vmfL1m91V38NicNOL1/MwCdjsH+VG9qYc/ay53DO42EyD45AIb6T7eWm/S6vjWIAA2C9Lzp7y+pAAAgOmQ3IQDrzZs/fqRo0CVDOwJoBwWMazdYXiuf0jiIzmJDBOWfTp2xzAAcLMZPolxS8gEUn3rZBcXAqTtg2U6LMvXKS7eSiyuSX4epbPYftlhigOHvzb3qQBBsSoUOAmey4lBD8UQb+RIeZ8/9luHSZsmmooD+zkk6n9QM84pxpwAACClS9jqikAIACke+a9iZkcdjnG/PCzJnxcwLf9X5IOUUewBJSAKAlqPnMBv7uZpPd3bgWtapK7Tv2Oi2e5ApxVkEfpx74HRLjWkBEDDj2WnlgMKwTutJaQAkKHvrrf5jrpBslzoJlCqJa0GEgBbCTHl6BgNRFZQZAJBg6x//ePZNPxSOEEpgUlkFcl56ygm0BT+SdiMt1b9yoCnFeNYMobihQ7q0hixBTTT97HMnlUtiVY0vxlqjoAVXfX8SknIYFe4QL2sBbw4c/iFLMhkp48xhQ2YK1NUQ5qo8Dkc0pLZ4ybS3oZYwUNTY+WhPjQQINIovhl48iwQqk5lphsSiH1+yRLKJhfIHgy9fQAL8UGZBH50zar0w2R4pi/NkYCLf4jnDnhoRAmLnrYP/QsL47HdhNIl3Bz7Akc9DsxElo87/TJDx+SRtBL7S90VCXX8wkNeo5uObBsT0/o96wgRYV6PF/rHnrhIqohCgCf1eoVTVXN1vWG699ifbI1KuSxgwqsYxoEndfvEqGXRnaZSzBk+N1ClHiZ0//dlmaQKqFIXyrUELpKovGPD2u4Ih7rdVYu25D4fpgsdKbL300QiqWcmFp74qMDgvWcmvznozq0EQJwb277V2PjqxhxggIOee+nHgrfidBKFbbxdhG/koOfHMZSHliKJ9lz+RzSCIMWeEpa4N2CRebaDk5FHlVetWRKiM2HDaW5IMyod3P6dDpXOU/J9bwA9lhrSU8abdY7P3PH6MGIAtnnO4qEXMUuCla6qy7pHQWIIeBRiHHawSf6BnVIg4u5b3jqvKCiVMQ9m4SyP+I+/2rAVBnBj4CiwOMbSLFwIvXEdpW9+hMEYDyOKWxUWJoj1e2fYtW/coAJBsGwBe4tlj7lchKI+7N70eIDIGAPKKW7Vskte4zNu9bcvWcgUAjsJiI+845hrlV/Gmux2E3WhM9bGa2sEAwGKbNkP7GGNESr55XVqbjlAraNFvcN/j2xQcYvaWdf+cP2+ZArB558kHOl6/Xwal/Lv0EBCgDbT9/qDenY/N/+6zik1rl8yfv1IBUvLLshbXdzjvgE8juYJ15h1X/f9yrWCAwC5zQJOIEQNazhhFaSBAYKDVJcMHNq/UvMwAQJho23bQtbDs/XfmKrCAo8W/9xgcFAIv3p0OAsJoOO6HP+jf5DBlRGrQseMZ4C15709faIsyM1+xoJMvIQAN/lLhKFWd9/rjyec3yTwwxKtWqnQPYu3HlTc0vKnQYjlBh/LUlGDIvKHmL49JMysUAH2e28rM2lP68HOMVp7HzLz4l63sQg2CTtu8MRb0EeZXM2fFM9J1CScCOPuNMmZWKZQVM/Pc0Y1tygIGe96ZyR8h5K9LzRumH09bryzgXY5vxIYBz7xhp44FnJdmguEwYHRZjzQlGQKgz0TNrFTaR2jPMG++rxiSu/tKuJzv8I8BzWtbpmkOLADO+ZCZvbSUjfaY+etfFlgFLBLu5XN8YUArZ+xXjzgYmKQqVHXD1AYGFI9wywfGppYPhMSAx/+SqrVQQMsnvMpSgipZ6DFvuBGT34zg7jMtkVIdBoxSg1PBhwSd32Q21VJWzMtHJosCpLx7uyZT9i8HPH7UwcDbYcs8ahADxmy0VQEImJbmUaEwoPjFVAgQwvC1nPFtjcc8u2dy682Uqo9qMODxr9NQBrihlHUmytpjfrt9NRZXtmAgrjihxpfLrDOmqJt+P6YopKE1t6TEBYSRj03uoCCTu4+S1Rmf/kIffheW5N8SnfOgTKXc7M2nminKRJmk0ZfNH6YOww9lVvbujAkDTHuftn13wkEtTDxTZvxFmdvaRuji927Wxo9bg1I3evo5cfgonDK+Ce+/nl1nROqT54xUvgqwSKjW7/ynPiR5ODuv+4kJA5qeWu/UP/HQmI6aafHae25jG6G7fjxECZ9vL1hdP7UweBdlQ+OXu500pDp1dnfld0NLo+97HrP8buCYpLXYON6BgCq4FEQ8UmD3nalS4PgPugWIvaL0LpwWGASGvv49pUDg7Peaa/+UidS1r1N2XwocEwbwph02nwjPa6tjmbmhR0qczUi6w/udgoXfE97pf2kYcD8yji0nG31Cff9PBYEOTaD0Rr5sBB7tGPDkC5Md04n5+nhe0NCGxxyDEKHwz52CZmAS3hmvmkBJQy0WTHRqUUl3mlKoKSjlKx9R4ijHgEp8fqOjsIl7nMcxqYJHdtubEYWZ0DN4Ei7hDX8gUBdlhHuN21cjf3Lr4J2YE96t12RzJ/cYMKDkyssqnDwt8m0ylkpKps1/dCrHSN0xLMySSnXX0AD7UdPf/uoIINKP9wmzo6V+qqeioxgDSn55/ibHdCJzwuXxiAGNz++y25kI3e/+UF3Rkfj54iDu6qOO1Bd6xLWh9jNig5fzstcujIoBo+WnZ611HSjk+xrEYhGy2Pe8HXdASDwvw60n6VaP+Q6KGbH+bVsAITd/MmRMjVTPsZqOTgywIvHM2ZsoxXs/faSJRQEanL7OljGkx/QKK1elvuI8vyLEwCv7bAFE5oFWYRkp9J3dsvaWpyjvZRTK1SNuSLnSECHxREyCD+FFe8+Taf0f4Q+0I/zeb9NEoSbYMS6he18T+mYOhAa/y9pLnsJigI0yJLf/tu/k1KtNhR7bMx7JZ2jtR+wI5DubhQ9Bk+71Y38CSuNnK2xsM/xWhA/4CT1scLbe7ZKyIMbHCR4GIiRY/srLG9LcaCTU4LiUn6Ep5VZVNpl210YxNpHvmujLX2F4y76mR+gBl0TqqIF3DYU6goFG5Kvobf+Kj6Z9fADSVNuSaf6qjClbRPBnuzKO1L81juJqkz7loqk+av1ZHpjupjt+BVE0ujAX9lianRd7pKzn4oSXAe3lpVtWrfjnGgCQaZriIMJrnWKSekwbP7c4gapwdESfk/99qg9GGLFotaXjSHceGo2ykf86pk5gwMANvn/JJs1+QqGeOz+uoJgRn+y1dI1QPzg2Gr4IhnRdlZkVDB+yo4RGNYw2LQEj7i7LyttdKHUfZxxCSkGg0pVbI6nx18cWF2WYZVthBkZFXETUeSN8GMIEM+0iaS1/EtGPRt36AshK9zDlpXzUHmmlqjiAjajH3xlfaFyYz2xVYNqegRGXEWFY5lPLTDsWW5SJ+54UteaKefjRGCNyecbm8RghYLBkpbUbBZxXEDX8KKB354zcNLC0lOzzs5dC1PyHwLMaazzKMSB14cQxKs7Ta//Yb3GC4UKOrFB13hkZZ83wN/s7Gs+NXAyE5ti+WakM4nsnIvW92SPizJEyLLGLwHX+IKToTz3Lx0msRfbcuN0p0ZfKwOlwNMsBlMb865zeOs40OcI/nL+7tY9eCE3QT2QSyQKWO60U+uVHF+MIA4GPWgygZNVtyrOF8bamJVhlrRlB7xjaOyF0bpthOzLuWWcbo9AvBu4h9GiYjQZBHBggyarongWXao5V2zHuKXFidb3jwKtpeHwmDMDGbY4U7x2DFEdo2x6ORgygIKOajFl0b6GOuXCSYUupw4mT4lhCA90yYuBblWyMopFd4sCAkZ2yEQOR9DcisdbQ+crRnUBT3Fkxhq0H0OZEx3iWsGtGyiUOoZZtY0Efdf3gaMEAAiAiGMMGWgy5/IIC0FgTedHNdmMTaNYqDk4gtM+o3DfZx+j1sQXxZP87ZKFJGA4DfLAPRsNug887vQWAoppAAIOtlJFbNo7nya0zYmCb83cxxGPuHpuNjkE4XZDXpGmrLiee0qMTAGigmqqb3mHtRoRmFEc9FkJTMhmSNzucX7SOydltUQfkAJnnB2Su1ssrKCiq/KFCqsHiGLfXXdOYdlHjvP0ZeFVWQ5Sb1gk5cPIpPiW1BkSs2ZMTLgYKY9qNjTJgAKDc4VWTmGbUoE5gYJ/xk6BHwCNwcMZzeRfTc/MyzjCl50BM6MsXOvtKCFKbbBBnTV4DU23RmKzNTHR314gEgjpSP5AblTLbHhVH82TrFAbiqsbL1L6cU7ROPBhg2G+ysJosqzHQ2FnBspg4sS8jS0UNyYEKzsmBYMPdjWUxRdv3HMj0jSYOpV0xzWgXYA4DAb1pe+yMZQUZdmYwztnxBRm2x0R5Sw4DAb2C5mBXkpV68WjTTRmn3dz5e3NMvNuUw0DA0dLBwNad8dgDazNyotj5xWZDsaBvTc4vCCgHisk6s4Zl38YTIvgqI+W2Dvo27YxnsVdlY5wwqzHQygrNMPHqOJaQnGLBdKOt0/hi5/oYKDN5X+cwEHA0O9ZJHP4zFk7sW5nhMQRtCpPrBZj4yzgow/qSHAaCyQEtjnMuTl0Ug0nF8NXGjKZl83YO+hbFwDwDfz8gchgIyK3uDgYWe9HX0MDnnOEpqEVXh/L8WFbqUzjazxnFP3rZphl+szL6dkT4yAf6elrcMrB4R3THQMBsyMUJg77b96xT4iz0x5GvL2ex7+OMiQeEftZ3mEoXRs5WGFyzBEwOAwF37AntrEQ2w/TIXf4Mf7qRMmOgV0NjX840PfIWNvx+hczJgaBGYcO+jkievSWyQYCTMk+auP1JDuV3I7ccJpwER/P9BTVlFJ5tm2Zi9/SIR8BZlE3xIZA1nmadbjW0Yl5EyoZWfYImh4HAL3e22/ZsQsQ31jB1k8jMCYTzbcuDYEJELWTglQqRkwOBX45PPNkyCDTOXhrt2CnB0/6+dlpzyyDQMGl7pI5SLPa9lJUWYbb7hpoutt9Q6Ccx2gM/mUs+QIS66VnW0SmWpRMiSXKNb64T8WEA6w0GEEbYHNP4akm0Bm/3+1s+hpF2yxODT5RHCBEwHXgwzjIyVW8wILhPb6srIIs94yMspRYz3yNfukTAha2sXgGG1rxA4dWQpv/9Mgp43TnvidHFyPKaUo2j7H2r6YXloVnBoO/wKUVRNfmRXVXI+EBpaC4a2nlvJDHQyJGPe+qNHAABVxRZ25Gx4lcYXgw8u1D4BBDCtTbHDW0aFwED92yMpMOKnH2/pf5gAHXxFXaLWC2m/58IJwiMXHO3b0YI03uIcSg/NSfkBT1aznlKRPBn3LI6gHUxWoXZfr4A+aaE3QDb0C3hDGw2cM1u/wKZ4VcpD7hubyiBbrBstImmv1ugXdi0qh5hgMxJI4yjl0uvMmHaFCp5z6wA+1iYC/rbRy+NWHEDhTHIDf3bV5H8QoRWjZKLWgys2ElcXzAAwGOdTvhazr45hGBViUm/lSoQ5+5JEekTHkt4gSl78vevRbvlDaG4DViFTdsXsak3GBDm5KudHiBKPvmw9AIzYu5VwRwKaX5wpnMGW4tfTgoMAi/x2p0RL/pLaYpFMBXrsBwI2mkUeVyR035Ei9tfSAS7Q1wlvvhheUCfmuEhstUuG7pyWkAQeIlpV5GOyDADbgZ10r7wyQejVXLb8SOPgbx8S6pBWYZDP2Ta3uNggI247gkZxCZQ8vMLA8f7he5/rSMIGA4MnxIEfqwSk4ebyAEdhIFWssGI9RPCJk40k5BSCuLvkFnoiIWZrLgGh+GKLsnAQyjcwqb6nyivr3sIFAnGsu831Yr/3LhKuCPMr+JJWm9v44otAnqGjV/KyvDTWJUJj5C/jrXPZdtsH4EkbLNd6xAM0Mz89dsP3//wpJUH/+IjjQFm7mNxg+C9TCuq+G95bv9LFDBiG3vGD0GP+UGwgecPA6z47ZQuHYgwpsIfZeOxdzMkXbRtV0EFwABrvsBu/ydgBKvgIFDMb5zTsFIgn/ayZlUbGFB8kbWpJYzkAxmZOD61YYqE499nH7xQmkuGQzU3nleDAVZ8XSoIBAxY5Isy8z8GQ9XdW4NgwOPnHFko4Fb2uQmSX2n5EAAQUkoBAKcuZVULGPD4RntVSXzIFdpk+NHFqSAQADdsYvZ0tXpEM7/U2lo+gmL0iwGj9vRIbUYkoeE9uzMwwCjD+x8sSH5tRCoOiwHDm5o4tZQCfrqD2VMmCARmHQPyICqRJBTNYlUbGHjWZidi8ZzKC3KUqpIVZlvnVH1OCK0f2smsqloG7THzrCHgCJ77dhQmq4PqMMCalxWmejIEcMKL+6umbLTHzJN6uZRfX5GsiIJggBX/3N0GArpMKK80mJyhqprMfAuUIKHxQtZHHAOav0BXHze4fUlFRpOgII1TKQA6/ddaZtaeMsbmgvKYWb97MdjSWMDpvLupbwyw4ilpxDlKgJOf2MLMyt2IBynvfW2QQ1nCVbycwmNgsaDU6Z9w1+ztvu1xvatzikI5YY/GwjLbEZs5pGavZGXc3825xhoZsEu7hgBAz7avoi+wklN/iKm1WEgaCi8ecVaLSqfn8DwIAGDZn99YAmhFhsi0WtB+2wk7kvxR5Pn9q5mzkk+Okan+PaGG4qE/HtTEpoxIAMB/mzxxFRDYF6X2nJu/vGeSJ4yc/2V7351XtfiX193wKIEBaNGlTXGRdTLOtL8x/RPuGu9GSqUaf+eR9wsUX+32Njxs8S6tcl94/FJa84oEALQc+sjHm62vH1gz5bZ+4uB/J2MmMZu5tJl/OcDs8TiQVVFud/kz8+yNuP/Lt248BQCcXUvQYhXzMhFWDrA2XzdKXQBK92bQLZ27bXhr05TjGYQddx+BTpMp46f/6+xn1ogIgCyq3o5SXa2uo9RgiwEk2Dp1KhR16NShVdMiCd6uHRvWrS4pBwBpjCs2Xj7DSwQMFanf7H0ojSQwgMKUvPEGNO/UvmOrokIBFbtKS9auKfEAQDqXABM3mNwlKGX7CbrzuDtSEh7GADrBTBC6WVoxIKftTMmzGFw7oxbiA+ZAj/Sd7hHEsmr2hccvYhWBHhSp20HIFH8QBTzGngkoB9h4PO6gdkkjDHxRJsK32dNR5AAbZc721WNcQP90csDj0Wmaywq8uRZyRiZxe6jLAKQa/Xaj9DfXs1aMlRFQISoDoaiVm7JH0vffHOLiNRT6N4+b9M3rjS/KxPzGZVGvfEOE/2sX4bQTwVdpwuvMK2sBA8Jc2TfUVKQaNqNTlSvJlZkQrSsTIpxmDfRjd4diBAo15p0iHZqyMHkTR0a/9Y9MmylNwlvsCOVpP99WG7ljFs9KCCcJBs49X4W84lCYxCs3h2WEVMPm9FQiHGWpW07/URwXPwrdZ2rTCNcHpv9l49rAgND9HtIyFAh0m7+OIxPityh1+w+uDM8IqXvMvV6Ho6x6fTIknrs/hTpjZmclw5WRGUh3KRNih1qpIRHqlpu8ULegCQO/+aiX4qACUbC66LMzojBCmILnJh0XivLVn3SL6+5PqXt/OlRxKBQwDEj3KQ+sBb+AmY3iO1JFUwa/4JB5W35/E4AgECIB+Q8eSjUH9gsOv/T2myQEur2LCFq+xN9NKppfcCgR9VQLAEmB/QLNy9IEPLHwm9rBABvNLxYBCgyOAVbMX11NAH43AwmA8xez0RwNA8yKecFQcF+7espXfMOHgslxYICN5vU3NwGgKmFQBQZY86Up+07CKK4lDDAr/nIkAAgpDsWufGKAjce84PIEgMjMDBQCoPsrzF5SvCwsBtgo5pkXAYAkH5QJYNBfkytdYsEAs2Je/etOlZEIQamdWarCgDJ/TzgrhqLhVwaKKHlImnWEMMCKee7VLb4TmUREghLL/K2JVsx/v6kVwKFEaHo2kASAE5/cy8klN4ZLm5M4PGdBC/zPWWvmOaOaZKQsJAAMfI0typqXJZIoCyoIhQE2irlsys+7Hp6nNRI0IH1pluL/gYSVWUjA06xSAm8fHCkMsNbMm978xcDWDQ9TX+q7rkYzb3n+nDwAICnS7EokIQCAzn293Ck6M7w9z/7u50HmrDTz2kcGUjWUpQCA/GF/YYey5qXOd9eGwkDlJuCKRS+MOadbs7xUCPaoojxP8V3JOlQg3MWK8SNlp/D6NOMj1kPPsACAPdtK93iVxHGA//tEDQuAlR9On7f1u2KCg1dvIiAwMwBQr0sv+x6AdhmlPvWsvGH/JoHmrEEALH33rwt3V1pVNmUDAJDXf9iwrgBuRSruncfJeUMxsGHY1WZTmQbWu3buqtjHNhOb9K8y+fjMbfuAABgBNTQYf4sWUMtXq7Dh8DdlsxYIULpk/oLlJe4lVFB4XM/TTzsJgA3Fj2ljJAB8u2jeohXfppwAbtql72mDun6H8JpcPDZVpTGqxq9Y/rs/HSwXyP/B2J5aAKBKSSYecRwkRbGDUjeGCADUxtXfrN64uXRvOQI2KmzVtssJx3cgAFDp10hHn/NByvs3rPn6m01bSvftR8CCotbtuh5/fFsAYJ2Wsns4KDpKGNLV2FdXDgEr3/3k6/IGnU675ORKQfX/G9JzoxtSvssAAAAASUVORK5CYII=';
    const logoH = 11;
    const logoW = logoH * (515/152); // ratio comprimit
    try {
      doc.addImage(LOGO_B64_SMALL, 'PNG', PW - MR - logoW, MT + 0.5, logoW, logoH);
    } catch(e) {
      doc.setFont('helvetica','bold'); doc.setFontSize(20);
      doc.setTextColor(0,0,0);
      doc.text('Plaat.', PW - MR, MT + 8, { align:'right' });
    }

    // Banda negra TIPUS DOCUMENT — SOLO PÀGINA 1
    if (primeraPag) {
      const bandaY = MT + 12 + 9;
      doc.setFillColor(0,0,0);
      doc.rect(ML, bandaY, CW, 8, 'F');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(12);
      doc.setTextColor(255,255,255);
      // Centrat perfecte: mig de la banda
      doc.text(T.tipusDoc, ML + CW/2, bandaY + 4, { align: 'center', baseline: 'middle' });
      doc.setTextColor(0,0,0);
      y = bandaY + 8 + 4;
    } else {
      y = MT + 15;
    }
  }

  // ── Peu de pàgina brandbook — SENSE línia superior ────────────────────────
  function dibuixarPeu() {
    const peuY = PH - MB + 1;
    const boldPart = esCA ? 'Plaat Arquitectura Tècnica' : 'Plaat Arquitectura Técnica';
    const normalPart = '  |  Barcelona - Madrid  |  plaat.es';
    doc.setFont('helvetica','bold'); doc.setFontSize(6.5); doc.setTextColor(0,0,0);
    doc.text(boldPart, ML, peuY);
    const boldW = doc.getTextWidth(boldPart);
    doc.setFont('helvetica','normal'); doc.setFontSize(6.5);
    doc.text(normalPart, ML + boldW, peuY);
    // Dreta: numeració correcta sense solapament
    // S'escriu al final un cop el doc té totes les pàgines
    // Per ara deixem un placeholder que s'actualitza al final
    doc.setFont('helvetica','normal'); doc.setFontSize(6.5);
    doc.text(`${T.peuAlt}  |  ${pagActual}`, PW - MR, peuY, { align: 'right' });
  }

  // ── PÀGINA 1 ──────────────────────────────────────────────────────────────
  dibuixarCapçalera(true);
  dibuixarPeu();

  // ── FILA ACTA / DATA / LLOC / FASE ───────────────────────────────────────
  // Alçada 0.5cm = 5mm. 0.5pt. MAJÚSCULES. FASE dreta.
  const filaH = 5;
  setLW(0.5); doc.line(ML, y, ML+CW, y); // línia superior 0.5pt

  const filaY = y + filaH/2;
  // Les tres primeres parelles s'escriuen des de l'esquerra
  // FASE s'escriu alineat a la dreta
  const tresParelles = [
    [T.acta.toUpperCase(), String(num).toUpperCase()],
    [T.data.toUpperCase(), dataAvui.toUpperCase()],
    [T.lloc.toUpperCase(), (vo.lloc || 'OBRA').toUpperCase()],
  ];
  doc.setFontSize(7.5);
  let fx = ML + 3;
  tresParelles.forEach(([label, val]) => {
    doc.setFont('helvetica','bold'); doc.setTextColor(0,0,0);
    doc.text(label, fx, filaY, { baseline:'middle' });
    fx += doc.getTextWidth(label) + 2;
    doc.setFont('helvetica','normal');
    doc.text(val, fx, filaY, { baseline:'middle' });
    fx += doc.getTextWidth(val) + 8;
  });
  // FASE alineat a la dreta
  const faseLabelW = doc.getStringUnitWidth(T.fase.toUpperCase()) * 7.5 / doc.internal.scaleFactor;
  const faseValW   = doc.getStringUnitWidth((vo.fase || '').toUpperCase()) * 7.5 / doc.internal.scaleFactor;
  let rfx = ML + CW - 3;
  doc.setFont('helvetica','normal'); doc.setTextColor(0,0,0);
  doc.text((vo.fase || '').toUpperCase(), rfx, filaY, { baseline:'middle', align:'right' });
  rfx -= faseValW + 2;
  doc.setFont('helvetica','bold');
  doc.text(T.fase.toUpperCase(), rfx, filaY, { baseline:'middle', align:'right' });

  setLW(0.5); doc.line(ML, y + filaH, ML+CW, y + filaH); // línia inferior 0.5pt
  y += filaH + 5;

  // ── TAULA EQUIP TÈCNIC — format Word exacte ──────────────────────────────
  // Columnes ajustades per evitar solapaments:
  // ROL: 52mm | EMP: 14mm | NOM: 28mm | EMAIL: 58mm | TEL: resta (~28mm)
  const eRol=50, eEmp=22, eNom=28, eEmail=52, eTel=CW-eRol-eEmp-eNom-eEmail; // Tel=28mm OK
  const xRol=ML, xEmp=ML+eRol, xNom=ML+eRol+eEmp, xEmail=ML+eRol+eEmp+eNom, xTel=ML+eRol+eEmp+eNom+eEmail;
  const equipRols = vo.equipo || [];
  const RH = 6; // alçada fila persona

  checkPage(10);

  // Títol "Equipo técnico y datos de contacto" — SENSE línia superior, SENSE fons
  // Alçada 7mm + 2mm espai = 9mm total
  doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.setTextColor(0,0,0);
  doc.text(T.equip, ML + 2, y + 4.5, { baseline:'middle' });
  y += 9; // 7mm alçada text + 2mm espai sota

  // Grups de rol: els detectem per nom
  const GRUP_PM  = (n) => n && n.toUpperCase().includes('PROJECT MANAGER');
  const GRUP_DF  = (n) => n && (n.toUpperCase().includes('OBRA (DO)') || n.toUpperCase().includes('FACULTATIVA'));
  const GRUP_DEO = (n) => n && (n.toUpperCase().includes('EJECUCIÓN') || n.toUpperCase().includes('EXECUCIÓ') ||
                                n.toUpperCase().includes('ESTRUCTURA') || n.toUpperCase().includes('INSTALACION') ||
                                n.toUpperCase().includes('SEGURIDAD') || n.toUpperCase().includes('SEGURETAT'));
  const GRUP_EC  = (n) => n && (n.toUpperCase().includes('CONTRATISTA') || n.toUpperCase().includes('CONTRACTISTA'));

  // Fila de grup (gris, sense recuadre — solo text bold sobre línia horitzontal)
  function dibuixaFilaGrup(textGrup) {
    checkPage(6);
    const gh = 5.5;
    doc.setFillColor(...GRIS15);
    doc.rect(ML, y, CW, gh, 'F');
    // SENSE rect de contorn — seguint format Word (sense línies que rodegen la fila de grup)
    doc.setFont('helvetica','bold'); doc.setFontSize(7.5); doc.setTextColor(0,0,0);
    doc.text(textGrup.toUpperCase(), ML + 2, y + gh/2, { baseline:'middle' });
    y += gh;
  }

  // Fila de persona (seguint format Word)
  // - Línies de separació: NO abracen tot l'ample
  //   · Primera persona d'un rol: línia des de xEmp fins a ML+CW
  //   · Persones addicionals del mateix rol: línia des de xNom fins a ML+CW
  // - Primera persona: ROL (bold) + EMPRESA + NOM + EMAIL + TEL
  // - Persones addicionals: [buit ROL] + [buit EMPRESA] + NOM + EMAIL + TEL
  // isLastOfGroup: l'última persona de l'últim rol d'un grup NO té línia inferior
  // (per no solapar amb la fila grisa del grup següent)
  function dibuixaFilaRol(rol, persones, isLastOfGroup = false) {
    persones.forEach((p, pi) => {
      checkPage(RH);
      const isFirst = pi === 0;
      const isLastPer = pi === persones.length - 1;
      const mateixaEmpresa = pi > 0 && (p.empresa||'') === (persones[pi-1]?.empresa||'');
      const midY = y + RH/2;

      // ROL — bold, solo primera persona
      if (isFirst) {
        doc.setFont('helvetica','bold'); doc.setFontSize(7.5); doc.setTextColor(0,0,0);
        const rolLines = doc.splitTextToSize(rol.nombre||'', eRol-3);
        const rolLH = 7.5*0.3528+0.5;
        let ry = y + RH/2 - (rolLines.length*rolLH)/2 + rolLH*0.8;
        rolLines.forEach(l => { doc.text(l, xRol+2, ry, {baseline:'middle'}); ry+=rolLH; });
      }

      // EMPRESA — primera persona o si és nova empresa
      if (isFirst || !mateixaEmpresa) {
        doc.setFont('helvetica','normal'); doc.setFontSize(7.5); doc.setTextColor(0,0,0);
        doc.text(p.empresa||'', xEmp+1, midY, {baseline:'middle'});
      }

      // NOM
      doc.setFont('helvetica','normal'); doc.setFontSize(7.5); doc.setTextColor(0,0,0);
      doc.text(p.nombre||'', xNom+2, midY, {baseline:'middle'});

      // EMAIL
      doc.setFontSize(7);
      const emailStr = p.email||'';
      const emailLines = doc.splitTextToSize(emailStr, eEmail-3);
      if (emailLines.length > 1) {
        const eLH = 7*0.3528+0.4;
        let ey = y + RH/2 - (emailLines.length*eLH)/2 + eLH*0.8;
        emailLines.forEach(l => { doc.text(l, xEmail+2, ey, {baseline:'middle'}); ey+=eLH; });
      } else {
        doc.text(emailStr, xEmail+2, midY, {baseline:'middle'});
      }

      // TEL — truncat per no sortir dels marges
      doc.setFontSize(7.5);
      const telStr = doc.splitTextToSize(p.tel||'', eTel-3)[0] || '';
      doc.text(telStr, xTel+2, midY, {baseline:'middle'});

      // Línia INFERIOR — NO si és l'última persona del grup (evita solapament amb fila grisa)
      const omitirLinia = isLastOfGroup && isLastPer;
      if (!omitirLinia) {
        setLW(0.5);
        const xIniLinia = (!isFirst && mateixaEmpresa) ? xNom : xEmp;
        doc.line(xIniLinia, y + RH, ML+CW, y + RH);
      }

      y += RH;
    });
  }

  // Dibuixar tota la taula d'equip
  if (equipRols.length > 0) {
    let pmRols  = equipRols.filter(r => GRUP_PM(r.nombre));
    let dfRols  = equipRols.filter(r => GRUP_DF(r.nombre));
    let deoRols = equipRols.filter(r => GRUP_DEO(r.nombre));
    let ecRols  = equipRols.filter(r => GRUP_EC(r.nombre));
    let altres  = equipRols.filter(r => !GRUP_PM(r.nombre) && !GRUP_DF(r.nombre) && !GRUP_DEO(r.nombre) && !GRUP_EC(r.nombre));

    const hiHaDF = dfRols.length > 0, hiHaDEO = deoRols.length > 0, hiHaEC = ecRols.length > 0;
    if (pmRols.length > 0) {
      dibuixaFilaGrup('PROJECT MANAGER');
      pmRols.forEach((r, ri) => {
        const isLast = ri === pmRols.length - 1 && (hiHaDF || hiHaDEO || hiHaEC);
        dibuixaFilaRol(r, r.personas?.length ? r.personas : [{}], isLast);
      });
    }
    if (hiHaDF) {
      dibuixaFilaGrup(esCA ? 'DIRECCIÓ FACULTATIVA' : 'DIRECCIÓN FACULTATIVA');
      dfRols.forEach((r, ri) => {
        const isLast = ri === dfRols.length - 1 && (hiHaDEO || hiHaEC);
        dibuixaFilaRol(r, r.personas?.length ? r.personas : [{}], isLast);
      });
    }
    if (hiHaDEO) {
      dibuixaFilaGrup(esCA ? 'DIRECCIÓ D\'EXECUCIÓ' : 'DIRECCIÓN DE EJECUCIÓN');
      deoRols.forEach((r, ri) => {
        const isLast = ri === deoRols.length - 1 && hiHaEC;
        dibuixaFilaRol(r, r.personas?.length ? r.personas : [{}], isLast);
      });
    }
    altres.forEach(r => dibuixaFilaRol(r, r.personas?.length ? r.personas : [{}], false));
    if (hiHaEC) {
      dibuixaFilaGrup(esCA ? 'CONTRACTISTA' : 'CONTRATISTA');
      ecRols.forEach((r, ri) => {
        dibuixaFilaRol(r, r.personas?.length ? r.personas : [{}], false); // últim grup, no cal ometre
      });
    }
  } else {
    y += 8;
  }

  // Línia inferior tanca la taula equip
  setLW(0.5); doc.line(ML, y, ML+CW, y);
  y += 6;

  // NOTA 48h + LLEGENDA
  checkPage(10);
  setLW(0.3);
  doc.setFont('helvetica','normal'); doc.setFontSize(7); doc.setTextColor(0,0,0);
  const notaLines = doc.splitTextToSize(T.nota, CW);
  doc.text(notaLines, ML, y + 3); y += notaLines.length * 3.5 + 3;
  doc.setFont('helvetica','normal'); doc.setFontSize(7);
  // Llegenda amb paraules subratllades com a SUBRAYADOR (fons de color)
  // Colors brandbook: Pendent RGB(255,246,215) | Resolt RGB(230,246,236) | Informatiu RGB(235,243,255)
  const llegY = y + 0.5;
  const lh7 = 7 * 0.3528 + 0.3; // line height 7pt
  const llegHigh = lh7 + 1; // alçada del subrayat
  doc.setTextColor(0,0,0);
  let lx = ML;

  function subrayat(text, bgColor) {
    const tw = doc.getTextWidth(text);
    // Fons de color (subrayador)
    doc.setFillColor(...bgColor); doc.rect(lx, llegY - lh7 + 0.5, tw, llegHigh, 'F');
    // Text en negre per sobre
    doc.setTextColor(0,0,0); doc.setFont('helvetica','normal'); doc.setFontSize(7);
    doc.text(text, lx, llegY);
    lx += tw;
  }

  const prefixLleg = esCA ? 'Estat: (R) ' : 'Estado: (R) ';
  const wordResolt  = esCA ? 'Resolt'      : 'Resuelto';
  const sep1        = '; (P) ';
  const wordPendent = esCA ? 'Pendent'     : 'Pendiente';
  const sep2        = '; (N) ' + (esCA ? 'Nou' : 'Nuevo') + '; (INF) ';
  const wordInfo    = esCA ? 'Informatiu'  : 'Informativo';

  doc.text(prefixLleg, lx, llegY); lx += doc.getTextWidth(prefixLleg);
  subrayat(wordResolt, [230,246,236]);   // Resolt — verd brandbook
  doc.text(sep1, lx, llegY); lx += doc.getTextWidth(sep1);
  subrayat(wordPendent, [255,246,215]);  // Pendent — groc brandbook
  doc.text(sep2, lx, llegY); lx += doc.getTextWidth(sep2);
  subrayat(wordInfo, [235,243,255]);     // Informatiu — blau brandbook
  y += 6;

  // ── SECCIÓ 0: ESTAT DE L'OBRA ─────────────────────────────────────────────
  // Format Word: cabecera gris amb Nº | Títol, sense recuadres
  // Fotos: 2 per fila, amb descripció de text
  const eo = vo.estadoObra || {};
  if (eo.descripcion || (eo.fotos||[]).length > 0) {
    checkPage(20);

    // Capçalera secció 0 — IGUAL format que DF/Contratista (gris, sense bordes, font 8 bold)
    const eoH = 5.5;
    doc.setFillColor(...GRIS15);
    doc.rect(ML, y, CW, eoH, 'F');
    // SENSE rect de contorn ni línies
    doc.setFont('helvetica','bold'); doc.setFontSize(8); doc.setTextColor(0,0,0);
    doc.text('0', ML + 2, y + eoH/2, { baseline:'middle' });
    doc.text('  ' + T.estat0, ML + 2, y + eoH/2, { baseline:'middle' });
    y += eoH;

    // Text de descripció (si n'hi ha)
    if (eo.descripcion) {
      doc.setFont('helvetica','normal'); doc.setFontSize(8.5); doc.setTextColor(0,0,0);
      const dl = doc.splitTextToSize(eo.descripcion, CW - 4);
      const dh = Math.max(10, dl.length * 4.2 + 5);
      checkPage(dh);
      // Sense recuadre — format Word
      // Text dins dels marges (CW total) alineat amb l'inici del títol
      doc.text(dl, ML + 2, y + 4);
      y += dh - 3; // reduir 5mm l'espai entre text i fotos
    }

    // Fotos de 2 en 2 — SENSE salt de pàgina, han de quedar a la pàg 1
    const fotos = eo.fotos || [];
    const fW2 = (CW - 6) / 2;
    for (let fi = 0; fi < fotos.length; fi += 2) {
      const pair = [fotos[fi], fotos[fi+1]].filter(Boolean);
      const dims = pair.map(f => {
        try {
          const pr = doc.getImageProperties(f.url||f.data);
          const r = pr.height / pr.width;
          const h = Math.min(fW2 * r, 75);
          return { w: h/r, h };
        } catch(e) { return { w: fW2, h: 55 }; }
      });
      const rh = Math.max(...dims.map(d => d.h));
      // NO checkPage — les fotos han de quedar sempre a la mateixa pàgina
      pair.forEach((f, pi) => {
        const src = f.url || f.data;
        if (src) try { doc.addImage(src, 'JPEG', ML + pi * (fW2 + 6), y, dims[pi].w, dims[pi].h); } catch(e) {}
      });
      y += rh + 3;
    }
    y += 3; // espai sota les fotos, sense línia de tancament
  }

  // Seccions de temes tractats
  const cNum=14, cEs=14, cIni=20, cFi=20, cRes=14, cDesc=CW-cNum-cEs-cIni-cFi-cRes;

  (vo.secciones||[]).forEach(sec => {
    const actius=(sec.temas||[]).filter(t=>!(t.resuelto&&t.resueltoEnActa&&t.resueltoEnActa<vo.num));
    if (!actius.length) return;

    checkPage(20);
    // Capçalera secció: número + títol (fondo gris 15%)
    doc.setFillColor(...GRIS15);
    doc.rect(ML, y, cNum, 8, 'F'); setLW(); doc.rect(ML, y, cNum, 8, 'S');
    doc.setFillColor(...GRIS15);
    doc.rect(ML+cNum, y, CW-cNum, 8, 'F'); setLW(); doc.rect(ML+cNum, y, CW-cNum, 8, 'S');
    doc.setFont('helvetica','bold'); doc.setFontSize(9); doc.setTextColor(0,0,0);
    doc.text(sec.codigo||'', ML+cNum/2, y+4.5, { align:'center', baseline:'middle' });
    doc.text(sec.titulo||'', ML+cNum+2, y+4.5, { baseline:'middle' });
    y+=8;

    // Sub-capçalera columnes (línies horitzontals 0.5p, sense verticals — brandbook)
    const shH=5.5;
    setLW(0.5); doc.line(ML, y, ML+CW, y);
    [[ML, cNum, ''], [ML+cNum, cDesc, T.desc], [ML+cNum+cDesc, cEs, T.es],
     [ML+cNum+cDesc+cEs, cIni, T.inici], [ML+cNum+cDesc+cEs+cIni, cFi, T.fi],
     [ML+cNum+cDesc+cEs+cIni+cFi, cRes, T.res]].forEach(([x,w,t]) => {
      doc.setFont('helvetica','bold'); doc.setFontSize(6.5); doc.setTextColor(0,0,0);
      if (t) doc.text(t, x+w/2, y+shH/2+0.5, { align:'center', baseline:'middle' });
    });
    y += shH;
    setLW(0.5); doc.line(ML, y, ML+CW, y); setLW(LW);

    actius.forEach(t => {
      const fW3=(cDesc-5)/2;
      const ed = (t.entradas||[]).map(en => {
        const esNova = en.actaNum === vo.num;
        const estat = esNova ? 'N' : (en.estado||'P');
        const fill = esNova ? null : (estat==='R'?C_R:estat==='I'?C_I:estat==='A'?C_A:C_P);
        doc.setFontSize(8.5);
        const lines = doc.splitTextToSize(en.texto||'', cDesc-3);
        const lh85 = 8.5*0.3528+0.6;
        const textH = lines.length*lh85+5;
        const fotos=en.fotos||[]; const fotoRows=[]; let fotosH=0;
        for(let i=0;i<fotos.length;i+=2){
          const pair=[fotos[i],fotos[i+1]].filter(Boolean);
          const dims=pair.map(f=>{try{const pr=doc.getImageProperties(f.url||f.data);const r=pr.height/pr.width;const h=Math.min(fW3*r,42);return{w:h/r,h};}catch(e){return{w:fW3,h:32};}});
          const rh=Math.max(...dims.map(d=>d.h));
          fotoRows.push({pair,dims,rh}); fotosH+=rh+2;
        }
        const h=Math.max(8,textH+fotosH);
        return{en,esNova,estat,fill,lines,textH,fotoRows,h,lh:lh85};
      });

      const temaH=ed.reduce((a,e)=>a+e.h,0);
      checkPage(temaH);

      // Fons de color per estat (brandbook)
      let ey=y;
      ed.forEach(e=>{
        if(e.fill){doc.setFillColor(...e.fill);doc.rect(ML+cNum,ey,CW-cNum,e.h,'F');}
        ey+=e.h;
      });

      // Número del tema (centrat en el bloc)
      doc.setFont('helvetica','bold'); doc.setFontSize(8.5); doc.setTextColor(0,0,0);
      doc.text(t.num||'', ML+cNum/2, y+temaH/2, { align:'center', baseline:'middle' });

      ey=y;
      ed.forEach(e=>{
        // Text descripció
        doc.setFont('helvetica', e.esNova?'bold':'normal');
        doc.setFontSize(8.5); doc.setTextColor(0,0,0);
        let ty=ey+3+e.lh*0.8;
        e.lines.forEach(l=>{doc.text(l,ML+cNum+2,ty,{baseline:'middle'});ty+=e.lh;});
        // Fotos
        let fy=ey+e.textH;
        e.fotoRows.forEach(row=>{
          row.pair.forEach((f,pi)=>{const src=f.url||f.data;if(src)try{doc.addImage(src,'JPEG',ML+cNum+2+pi*(fW3+1),fy,row.dims[pi].w,row.dims[pi].h);}catch(er){}});
          fy+=row.rh+2;
        });
        // Valors columnes (centrats)
        const midY=ey+e.h/2;
        doc.setFont('helvetica','bold'); doc.setFontSize(8.5);
        doc.text(e.estat, ML+cNum+cDesc+cEs/2, midY, { align:'center', baseline:'middle' });
        doc.setFont('helvetica','normal'); doc.setFontSize(7.5);
        const isR=e.en.estado==='R'&&!e.esNova;
        doc.text(isR?'':fmtFechaCorta(e.en.fecha), ML+cNum+cDesc+cEs+cIni/2, midY, { align:'center', baseline:'middle' });
        doc.text(isR?fmtFechaCorta(e.en.fin||e.en.fecha):'', ML+cNum+cDesc+cEs+cIni+cFi/2, midY, { align:'center', baseline:'middle' });
        doc.setFont('helvetica','bold'); doc.setFontSize(8.5);
        doc.text(e.en.resp||'', ML+cNum+cDesc+cEs+cIni+cFi+cRes/2, midY, { align:'center', baseline:'middle' });
        ey+=e.h;
      });

      // Línies horitzontals separadores (brandbook: sense verticals en temes)
      setLW(0.5); doc.line(ML, y, ML+CW, y); doc.line(ML, y+temaH, ML+CW, y+temaH); setLW(LW);
      // Línies verticals de columnes (en taula complex de temes sí n'hi ha)
      const vxs=[ML,ML+cNum,ML+cNum+cDesc,ML+cNum+cDesc+cEs,ML+cNum+cDesc+cEs+cIni,ML+cNum+cDesc+cEs+cIni+cFi,ML+CW];
      vxs.forEach(x=>{ setLW(0.3); doc.line(x,y,x,y+temaH); });
      y+=temaH;
    });
    y+=4;
  });

  // ── NOTA + TAULA FIRMES ────────────────────────────────────────────────────
  checkPage(45);
  y+=4;
  doc.setFont('helvetica','italic'); doc.setFontSize(7.5); doc.setTextColor(0,0,0);
  const notaFirmes=doc.splitTextToSize(T.nota,CW);
  doc.text(notaFirmes,ML,y); y+=notaFirmes.length*3.8+4;
  doc.setFont('helvetica','normal'); doc.setFontSize(7.5);
  doc.text(T.conforme,ML,y); y+=8;

  // 4 caselles de firma (brandbook: línies horitzontals 0.5p, sense verticals)
  const fw=CW/4;
  const firmesLabels=[T.promotor, T.pm_f, T.do_f, T.deo_f];
  firmesLabels.forEach((f,i)=>{
    setLW(0.5); doc.rect(ML+i*fw,y,fw,28,'S');
    doc.setFont('helvetica','bold'); doc.setFontSize(7); doc.setTextColor(0,0,0);
    const fl=doc.splitTextToSize(f,fw-4);
    doc.text(fl,ML+i*fw+2,y+4);
  });
  y+=28+4;
  // 5a casella: Contractista (ample mig)
  checkPage(30);
  setLW(0.5); doc.rect(ML,y,fw,28,'S');
  doc.setFont('helvetica','bold'); doc.setFontSize(7); doc.setTextColor(0,0,0);
  doc.text(T.ec_f,ML+2,y+4);

  // ── PEU FINAL TOTES LES PÀGINES ───────────────────────────────────────────
  const totalPags = doc.getNumberOfPages();
  for (let p=1; p<=totalPags; p++) {
    doc.setPage(p);
    pagActual = p;
    // Reescriure la part dreta del peu amb el total correcte (cobreix el placeholder)
    const peuY2 = PH - MB + 1;
    // Cobrir el text anterior amb blanc
    doc.setFillColor(255,255,255);
    doc.rect(PW/2, peuY2 - 4, PW/2 - MR, 6, 'F');
    // Escriure numeració correcta
    doc.setFont('helvetica','normal'); doc.setFontSize(6.5); doc.setTextColor(0,0,0);
    doc.text(`${T.peuAlt}  |  ${p} de ${totalPags}`, PW - MR, peuY2, { align: 'right' });
  }

  // Obrir en nova pestanya (compatible iOS Safari)
  const pdfBlob = doc.output('blob');
  const url = URL.createObjectURL(pdfBlob);
  window.open(url,'_blank');
  setTimeout(()=>URL.revokeObjectURL(url),10000);
}



// ─── MÓDULO: Seguimiento global de obra ──────────────────────────────────────
const TEMATICAS = ['ESTRUCTURAS','INSTALACIONES','OBRA CIVIL','ACABADOS','COSTES OBRA','PLANIFICACIÓN','OTROS'];
const VIAS = ['VT','VO','MAIL','LL','p'];
const ESTADOS_SEG = ['En curso','Cerrado','Pendiente'];

const SEG_KEY = 'plaat_seguimiento_v1';

function VistaSeguimiento({ obras, isMobile }) {
  const [puntos,      setPuntos]      = useState([]);
  const [loading,     setLoading]     = useState(true);
  const [showForm,    setShowForm]    = useState(false);
  const [editando,    setEditando]    = useState(null);
  const [filtroObra,  setFiltroObra]  = useState('todas');
  const [filtroEst,   setFiltroEst]   = useState('todos');
  const [exportando,  setExportando]  = useState(false);
  const [confirmacion,setConfirmacion]= useState(null);

  // Cargar desde Supabase con migración desde storage antiguo
  useEffect(() => {
    (async () => {
      try {
        if (window.db?.getSeguimiento) {
          const lista = await window.db.getSeguimiento();
          if (lista.length) { setPuntos(lista); setLoading(false); return; }
          // Migración: leer storage antiguo y subir
          const r = await window.storage?.get(SEG_KEY, true).catch(() => null);
          if (r?.value) {
            const vieja = JSON.parse(r.value);
            for (const p of vieja) await window.db.upsertPuntoSeg(p).catch(() => {});
            setPuntos(vieja);
          }
        } else {
          const r = await window.storage?.get(SEG_KEY, true);
          if (r?.value) setPuntos(JSON.parse(r.value));
        }
      } catch(e) { console.error('Seguimiento load:', e); }
      setLoading(false);
    })();
  }, []);

  // Realtime: actualizar cuando otro usuario guarda
  useEffect(() => {
    if (!window.db?.subscribeSeguimiento) return;
    const unsub = window.db.subscribeSeguimiento(async () => {
      try {
        const lista = await window.db.getSeguimiento();
        setPuntos(lista);
      } catch(e) {}
    });
    return unsub;
  }, []);

  async function guardar(lista) {
    setPuntos(lista);
    try {
      if (window.db?.upsertPuntoSeg) {
        for (const p of lista) await window.db.upsertPuntoSeg(p);
      } else {
        await window.storage?.set(SEG_KEY, JSON.stringify(lista), true);
      }
    } catch(e) { console.error('Seguimiento save:', e); }
  }

  function nextNum(obraId) {
    const deEstaObra = puntos.filter(p => p.obraId === obraId);
    return deEstaObra.length + 1;
  }

  async function guardarPunto(punto) {
    const existe = puntos.some(p => p.id === punto.id);
    const lista = existe ? puntos.map(p => p.id === punto.id ? punto : p) : [...puntos, punto];
    await guardar(lista);
    setShowForm(false); setEditando(null);
  }

  async function eliminarPunto(id) {
    const lista = puntos.filter(p => p.id !== id);
    setPuntos(lista);
    try {
      if (window.db?.deletePuntoSeg) await window.db.deletePuntoSeg(id);
      else await window.storage?.set(SEG_KEY, JSON.stringify(lista), true);
    } catch(e) { console.error('Seguimiento delete:', e); }
    setConfirmacion(null);
  }

  // Exportar a Excel con SheetJS
  async function exportarExcel() {
    setExportando(true);
    try {
      // Cargar SheetJS si no está
      if (!window.XLSX) {
        await new Promise((res, rej) => {
          const s = document.createElement('script');
          s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
          s.onload = res; s.onerror = rej;
          document.head.appendChild(s);
        });
      }
      const { utils, writeFile } = window.XLSX;

      // Cabeceras sin columnas vacías
      const cab = ['NUM','TEMAS TRATADOS','Responsable','Fecha','Fecha límite resolución','Fecha resolución','Estado','TEMÁTICA','VIA'];
      const filas = [cab];

      const puntosExp = filtroObra !== 'todas' ? puntos.filter(p => p.obraId === filtroObra) : puntos;
      puntosExp.forEach(p => {
        const fmtD = iso => iso ? new Date(iso).toLocaleDateString('es-ES') : '';
        filas.push([
          p.num, p.tema,
          p.responsable, fmtD(p.fecha), p.fechaLimite || '', fmtD(p.fechaResolucion),
          p.estado, p.tematica, p.via
        ]);
      });

      const ws = utils.aoa_to_sheet(filas);
      ws['!cols'] = [
        {wch:6},{wch:80},{wch:20},{wch:12},{wch:22},{wch:20},{wch:12},{wch:18},{wch:8}
      ];
      const wb = utils.book_new();
      utils.book_append_sheet(wb, ws, 'Seguimiento');
      const fecha = new Date().toLocaleDateString('es-ES').replace(/\//g,'-');
      writeFile(wb, `Seguimiento_PLAAT_${fecha}.xlsx`);
    } catch(e) { alert('Error exportando: ' + e.message); }
    setExportando(false);
  }

  const obrasFiltro = filtroObra !== 'todas' ? puntos.filter(p => p.obraId === filtroObra) : puntos;
  const puntosFiltrados = filtroEst !== 'todos' ? obrasFiltro.filter(p => p.estado === filtroEst) : obrasFiltro;

  const fmtD = iso => iso ? new Date(iso).toLocaleDateString('es-ES', {day:'2-digit',month:'2-digit',year:'2-digit'}) : '—';

  const ESTADO_COLORS = {
    'En curso':  { bg: '#FEF3DB', color: '#7C4A00' },
    'Cerrado':   { bg: '#E8F5E0', color: '#2D5E10' },
    'Pendiente': { bg: '#EEEDE7', color: '#52524E' },
  };

  if (showForm || editando) {
    return isMobile ? (
      <div style={{ position: 'fixed', inset: 0, zIndex: 9999, background: '#F7F6F3' }}>
        <FormSeguimiento
          punto={editando}
          obras={obras}
          nextNum={nextNum}
          onGuardar={guardarPunto}
          onCerrar={() => { setShowForm(false); setEditando(null); }}
          isMobile={isMobile}
        />
      </div>
    ) : (
      <FormSeguimiento
        punto={editando}
        obras={obras}
        nextNum={nextNum}
        onGuardar={guardarPunto}
        onCerrar={() => { setShowForm(false); setEditando(null); }}
        isMobile={isMobile}
      />
    );
  }

  return (
    <div style={{ display:'flex', flexDirection:'column', height:'100%', overflow:'hidden' }}>
      {/* Banner */}
      <div className="dash-banner" style={{ borderBottom:'1px solid #ECEAE4', padding: isMobile ? '18px 16px' : '22px 22px', paddingTop: isMobile ? 'calc(18px + env(safe-area-inset-top))' : '22px', flexShrink:0 }}>
        <div className="arch-grid" />
        <div className="dash-banner-ring db-r1" />
        <div className="dash-beam" />
        <div className="dash-title" style={{ position:'relative', zIndex:2, display:'flex', alignItems:'center', gap:12 }}>
          <div style={{ flex:1 }}>
            <div style={{ fontSize: isMobile?18:20, fontWeight:700, color:'#F2F1ED', display:'flex', alignItems:'baseline', gap:7 }}>
              Seguimiento <span style={{ fontSize:13, color:'#8AA88A' }}>.</span>
            </div>
            <div style={{ fontSize:12, color:'rgba(255,255,255,0.5)', marginTop:3 }}>{puntos.length} punto{puntos.length!==1?'s':''} registrado{puntos.length!==1?'s':''}</div>
          </div>
          <button onClick={exportarExcel} disabled={exportando} className="tap" style={{ padding:'7px 14px', borderRadius:9, border:'1px solid rgba(255,255,255,0.2)', background:'rgba(255,255,255,0.08)', color:'rgba(255,255,255,0.8)', fontSize:12, fontWeight:500, cursor:'pointer' }}>
            {exportando ? '...' : '↓ Excel'}
          </button>
          <button onClick={() => setShowForm(true)} className="tap shimmer-btn" style={{ padding:'8px 16px', borderRadius:11, border:'none', background:'#5A7D5A', color:'#fff', fontSize:13, fontWeight:600, cursor:'pointer' }}>
            + Nuevo punto
          </button>
        </div>
      </div>

      {/* Filtros */}
      <div style={{ padding: isMobile?'10px 16px':'12px 22px', borderBottom:'1px solid #ECEAE4', display:'flex', gap:8, flexWrap:'wrap', flexShrink:0, background:'#fff' }}>
        <select value={filtroObra} onChange={e => setFiltroObra(e.target.value)} style={{ fontSize:12, padding:'5px 10px', borderRadius:20, width:'auto' }}>
          <option value="todas">Todas las obras</option>
          {obras.map(o => <option key={o.id} value={o.id}>{o.nombre}</option>)}
        </select>
        <select value={filtroEst} onChange={e => setFiltroEst(e.target.value)} style={{ fontSize:12, padding:'5px 10px', borderRadius:20, width:'auto' }}>
          <option value="todos">Todos los estados</option>
          {ESTADOS_SEG.map(e => <option key={e}>{e}</option>)}
        </select>
      </div>

      {/* Tabla / Lista */}
      <div style={{ flex:1, overflowY:'auto', padding: isMobile?'12px 16px':'14px 22px' }}>
        {loading && <div style={{ textAlign:'center', padding:40, color:'#A5A5A0', fontSize:13 }}>Cargando...</div>}
        {!loading && puntosFiltrados.length === 0 && (
          <div style={{ textAlign:'center', padding:'60px 20px' }}>
            <div style={{ fontSize:36, marginBottom:14 }}>📋</div>
            <div style={{ fontSize:15, fontWeight:500, color:'#16160F', marginBottom:8 }}>Sin puntos de seguimiento</div>
            <div style={{ fontSize:13, color:'#9B9B97', marginBottom:20 }}>Registra el primer punto de acción de una reunión</div>
            <button onClick={() => setShowForm(true)} style={{ padding:'9px 20px', borderRadius:11, border:'none', background:'#5A7D5A', color:'#fff', fontSize:13, fontWeight:600, cursor:'pointer' }}>+ Nuevo punto</button>
          </div>
        )}
        {!loading && puntosFiltrados.length > 0 && (
          isMobile ? (
            <div style={{ display:'flex', flexDirection:'column', gap:8 }}>
              {puntosFiltrados.map(p => {
                const obraNombre = obras.find(o => o.id === p.obraId)?.nombre || '—';
                const ec = ESTADO_COLORS[p.estado] || ESTADO_COLORS['Pendiente'];
                return (
                  <div key={p.id} className="hov-card" style={{ background:'#fff', borderRadius:14, boxShadow:'0 1px 2px rgba(0,0,0,0.04), 0 3px 12px rgba(0,0,0,0.04)', padding:'13px 15px', cursor:'pointer' }}
                    onClick={() => setEditando(p)}>
                    <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:6 }}>
                      <span style={{ width:28, height:28, borderRadius:8, background:'#1A1A17', color:'#F2F1ED', display:'flex', alignItems:'center', justifyContent:'center', fontSize:11, fontWeight:700, flexShrink:0 }}>{p.num}</span>
                      <span style={{ fontSize:11, padding:'2px 8px', borderRadius:20, background:ec.bg, color:ec.color, fontWeight:600 }}>{p.estado}</span>
                      <span style={{ fontSize:11, color:'#9B9B97', marginLeft:'auto' }}>{p.tematica}</span>
                    </div>
                    <div style={{ fontSize:13, fontWeight:500, color:'#16160F', lineHeight:1.4, marginBottom:5 }}>{p.tema}</div>
                    <div style={{ fontSize:11, color:'#9B9B97', display:'flex', gap:10 }}>
                      <span>{obraNombre}</span>
                      <span>Resp: {p.responsable}</span>
                      <span>Límite: {fmtD(p.fechaLimite)}</span>
                    </div>
                  </div>
                );
              })}
            </div>
          ) : (
            <div style={{ background:'#fff', borderRadius:14, boxShadow:'0 1px 2px rgba(0,0,0,0.04), 0 3px 12px rgba(0,0,0,0.04)', overflow:'hidden' }}>
              <table style={{ width:'100%', borderCollapse:'collapse', fontSize:12.5 }}>
                <thead>
                  <tr style={{ background:'#F7F6F3', borderBottom:'1px solid #ECEAE4' }}>
                    {['Nº','Obra','Tema','Responsable','Fecha','Límite','Resolución','Estado','Temática','Via',''].map(h => (
                      <th key={h} style={{ padding:'9px 12px', textAlign:'left', fontWeight:600, color:'#52524E', fontSize:11, letterSpacing:'0.04em', whiteSpace:'nowrap' }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {puntosFiltrados.map((p, i) => {
                    const obraNombre = obras.find(o => o.id === p.obraId)?.nombre || '—';
                    const ec = ESTADO_COLORS[p.estado] || ESTADO_COLORS['Pendiente'];
                    return (
                      <tr key={p.id} style={{ borderBottom:'1px solid #F0EFEA', background: i%2===0?'#fff':'#FAFAF8' }}>
                        <td style={{ padding:'8px 12px', fontWeight:700, color:'#16160F' }}>{p.num}</td>
                        <td style={{ padding:'8px 12px', color:'#6B6B66', maxWidth:120, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{obraNombre}</td>
                        <td style={{ padding:'8px 12px', maxWidth:260, color:'#16160F' }}>{p.tema}</td>
                        <td style={{ padding:'8px 12px', whiteSpace:'nowrap' }}>{p.responsable}</td>
                        <td style={{ padding:'8px 12px', whiteSpace:'nowrap', color:'#6B6B66' }}>{fmtD(p.fecha)}</td>
                        <td style={{ padding:'8px 12px', whiteSpace:'nowrap', color:'#6B6B66' }}>{fmtD(p.fechaLimite)}</td>
                        <td style={{ padding:'8px 12px', whiteSpace:'nowrap', color:'#6B6B66' }}>{fmtD(p.fechaResolucion)}</td>
                        <td style={{ padding:'8px 12px' }}><span style={{ fontSize:11, padding:'2px 9px', borderRadius:20, background:ec.bg, color:ec.color, fontWeight:600, whiteSpace:'nowrap' }}>{p.estado}</span></td>
                        <td style={{ padding:'8px 12px', color:'#6B6B66', whiteSpace:'nowrap' }}>{p.tematica}</td>
                        <td style={{ padding:'8px 12px', color:'#6B6B66' }}>{p.via}</td>
                        <td style={{ padding:'8px 12px' }}>
                          <div style={{ display:'flex', gap:6 }}>
                            <button onClick={() => setEditando(p)} style={{ background:'none', border:'1px solid #E6E4DD', borderRadius:7, padding:'3px 9px', cursor:'pointer', fontSize:11, color:'#52524E' }}>Editar</button>
                            <button onClick={() => setConfirmacion({ titulo:'Eliminar punto', texto:`Vas a eliminar el punto ${p.num}: "${p.tema?.slice(0,40)}..."`, onSi: () => eliminarPunto(p.id) })} style={{ background:'none', border:'none', cursor:'pointer', color:'#C4C3BE', fontSize:17, lineHeight:1 }}>×</button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )
        )}
      </div>
      {confirmacion && <ConfirmMini titulo={confirmacion.titulo} texto={confirmacion.texto} onSi={confirmacion.onSi} onNo={() => setConfirmacion(null)} />}
    </div>
  );
}

function FormSeguimiento({ punto, obras, nextNum, onGuardar, onCerrar, isMobile }) {
  const obraDefecto = obras[0]?.id || '';
  const [f, setF] = useState(() => punto || {
    id: uid(), obraId: obraDefecto, num: '', tema: '',
    responsable: '', fecha: today(), fechaLimite: '', fechaResolucion: '',
    estado: 'En curso', tematica: 'OTROS', via: 'VT',
  });
  const upd = (k,v) => setF(prev => ({ ...prev, [k]: v }));

  useEffect(() => {
    if (!punto) upd('num', nextNum(f.obraId));
  }, [f.obraId]);

  const canSave = f.tema?.trim();

  return (
    <div style={{ display:'flex', flexDirection:'column', height:'100%', background:'#F7F6F3', borderRadius: isMobile?0:16, overflow:'hidden' }}>
      {/* Cabecera */}
      <div style={{ background:'#1A1A17', padding: isMobile?'16px 16px':'16px 20px', paddingTop: isMobile?'calc(16px + env(safe-area-inset-top))':'16px', display:'flex', alignItems:'center', gap:12, flexShrink:0 }}>
        <button onClick={onCerrar} style={{ background:'none', border:'none', color:'rgba(255,255,255,0.6)', fontSize:18, cursor:'pointer', padding:'0 4px', lineHeight:1 }}>←</button>
        <div style={{ flex:1 }}>
          <div style={{ fontSize:15, fontWeight:700, color:'#F2F1ED' }}>{punto ? 'Editar punto' : 'Nuevo punto'}</div>
          <div style={{ fontSize:11, color:'rgba(255,255,255,0.4)', marginTop:1 }}>Seguimiento de obra</div>
        </div>
        <button onClick={() => canSave && onGuardar(f)} style={{ padding:'9px 20px', borderRadius:11, border:'none', background: canSave?'#5A7D5A':'#3A3A38', color:'#fff', fontSize:14, fontWeight:700, cursor: canSave?'pointer':'not-allowed', opacity: canSave?1:0.5 }}>
          Guardar
        </button>
      </div>

      {/* Cuerpo — orden optimizado para móvil */}
      <div style={{ flex:1, overflowY:'auto', padding: isMobile?'16px':'20px 22px', display:'flex', flexDirection:'column', gap:14 }}>

        {/* TEMA — lo más importante, primero y grande */}
        <Field label="Tema tratado *">
          <textarea value={f.tema} onChange={e => upd('tema', e.target.value)}
            placeholder="Describe el punto de acción o tema tratado..."
            style={{ minHeight: isMobile?120:100, fontSize: isMobile?15:14 }} autoFocus />
        </Field>

        {/* Obra + Responsable — lo más frecuente */}
        <div style={{ display:'grid', gridTemplateColumns: isMobile?'1fr':'1fr 1fr', gap:12 }}>
          <Field label="Obra">
            <select value={f.obraId} onChange={e => upd('obraId', e.target.value)} style={{ fontSize: isMobile?15:14 }}>
              {obras.map(o => <option key={o.id} value={o.id}>{o.nombre}</option>)}
            </select>
          </Field>
          <Field label="Responsable">
            <input value={f.responsable} onChange={e => upd('responsable', e.target.value)}
              placeholder="Empresa o persona" style={{ fontSize: isMobile?15:14 }} />
          </Field>
        </div>

        {/* Estado + Temática + Vía */}
        <div style={{ display:'grid', gridTemplateColumns: isMobile?'1fr 1fr':'1fr 1fr 1fr', gap:12 }}>
          <Field label="Estado">
            <select value={f.estado} onChange={e => upd('estado', e.target.value)} style={{ fontSize: isMobile?15:14 }}>
              {ESTADOS_SEG.map(e => <option key={e}>{e}</option>)}
            </select>
          </Field>
          <Field label="Temática">
            <select value={f.tematica} onChange={e => upd('tematica', e.target.value)} style={{ fontSize: isMobile?15:14 }}>
              {TEMATICAS.map(t => <option key={t}>{t}</option>)}
            </select>
          </Field>
          <Field label="Vía">
            <select value={f.via} onChange={e => upd('via', e.target.value)} style={{ fontSize: isMobile?15:14 }}>
              {VIAS.map(v => <option key={v}>{v}</option>)}
            </select>
          </Field>
        </div>

        {/* Fechas — menos urgentes, al final */}
        <div style={{ display:'grid', gridTemplateColumns: isMobile?'1fr 1fr':'repeat(3,1fr)', gap:12 }}>
          <Field label="Fecha reunión">
            <input type="date" value={f.fecha} onChange={e => upd('fecha', e.target.value)} style={{ fontSize: isMobile?15:14 }} />
          </Field>
          <Field label="Fecha límite (texto)">
            <input value={f.fechaLimite} onChange={e => upd('fechaLimite', e.target.value)}
              placeholder="ASAP / dd/mm/aa / FO..." style={{ fontSize: isMobile?15:14 }} />
          </Field>
          <Field label="Fecha resolución">
            <input type="date" value={f.fechaResolucion} onChange={e => upd('fechaResolucion', e.target.value)} style={{ fontSize: isMobile?15:14 }} />
          </Field>
        </div>

        {/* Nº — al final, autonumérico */}
        <Field label="Nº de punto">
          <input type="number" min="1" value={f.num} onChange={e => upd('num', parseInt(e.target.value||'1',10))} style={{ fontSize: isMobile?15:14 }} />
        </Field>
      </div>

      {/* Pie fijo en móvil */}
      <div style={{ borderTop:'1px solid #ECEAE4', padding: isMobile?'12px 16px calc(12px + env(safe-area-inset-bottom))':'14px 20px', display:'flex', gap:10, flexShrink:0, background:'#fff' }}>
        <Btn onClick={onCerrar} full>Cancelar</Btn>
        <Btn primary full onClick={() => canSave && onGuardar(f)} disabled={!canSave}>Guardar punto</Btn>
      </div>
    </div>
  );
}

function DetalleObra({ obra, onBack, onSave, isMobile, user }) {
  const [tab, setTab]               = useState('inspecciones');
  const [editEstado, setEditEstado] = useState(false);
  const [showCompartir, setShowCompartir] = useState(false);
  const esOwner = obra._rol === 'owner';

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
    <>
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden' }}>

      {/* Header oscuro — mismo estilo que el sidebar */}
      <div className="dash-banner" style={{ background: '#1C1C1A', flexShrink: 0, paddingTop: isMobile ? 'env(safe-area-inset-top)' : 0 }}>
        <div className="arch-grid" style={{ opacity: 0.4 }} />
        <div className="dash-banner-ring db-r1" style={{ opacity: 0.6 }} />
        <div className="dash-beam" />

        {/* Breadcrumb + nombre + estado */}
        <div style={{ position: 'relative', zIndex: 2, display: 'flex', alignItems: 'center', gap: 10, padding: isMobile ? '12px 14px' : '14px 22px 12px', borderBottom: isMobile ? 'none' : '1px solid rgba(255,255,255,0.07)' }}>
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

          {/* Botón compartir */}
          <button onClick={() => setShowCompartir(true)} style={{ flexShrink: 0, padding: '5px 12px', borderRadius: 9, border: '1px solid rgba(255,255,255,0.15)', background: 'rgba(255,255,255,0.07)', color: 'rgba(255,255,255,0.7)', fontSize: 12, cursor: 'pointer', fontWeight: 500, display: 'flex', alignItems: 'center', gap: 5 }}>
            👥{!isMobile && ' Compartir'}
          </button>
        </div>

        {/* Tabs arriba — solo escritorio */}
        {!isMobile && (
          <div className="no-scrollbar" style={{ position: 'relative', zIndex: 2, display: 'flex', padding: '0 22px', overflowX: 'auto' }}>
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
    {showCompartir && <ModalCompartir obra={obra} user={user} onClose={() => setShowCompartir(false)} />}
  </>
  );
}

// ─── Fondo arquitectónico animado (login + splash) ──────────────────────────
const FondoArquitectonico = memo(function FondoArquitectonico() {
  return (
    <>
      {/* Auroras de color */}
      <div className="auro auro-1" />
      <div className="auro auro-2" />
      <div className="auro auro-3" />
      {/* Malla técnica desplazándose */}
      <div className="arch-grid" />
      {/* Línea de escaneo */}
      <div className="scan" style={{ top: 0 }} />

      {/* Plano técnico grande girando lento (arriba derecha) */}
      <svg className="arch-shape" style={{ top: '-80px', right: '-80px', width: 380, height: 380, animation: 'rotateSlow 60s linear infinite' }} viewBox="0 0 200 200">
        <circle className="arch-line arch-dash" cx="100" cy="100" r="90" strokeWidth="0.6" />
        <circle className="arch-line" cx="100" cy="100" r="62" strokeWidth="0.5" />
        <circle className="arch-line-accent" cx="100" cy="100" r="34" strokeWidth="0.7" />
        <line className="arch-line" x1="10" y1="100" x2="190" y2="100" strokeWidth="0.4" />
        <line className="arch-line" x1="100" y1="10" x2="100" y2="190" strokeWidth="0.4" />
      </svg>

      {/* Estructura de planos (abajo izquierda) flotando */}
      <svg className="arch-shape" style={{ bottom: '-40px', left: '-30px', width: 320, height: 320, animation: 'floatY 11s ease-in-out infinite' }} viewBox="0 0 200 200">
        <rect className="arch-line" x="30" y="30" width="120" height="120" strokeWidth="0.6" />
        <rect className="arch-line" x="50" y="50" width="120" height="120" strokeWidth="0.5" />
        <rect className="arch-line-accent" x="40" y="40" width="120" height="120" strokeWidth="0.6" />
        <line className="arch-line arch-dash" x1="30" y1="30" x2="170" y2="170" strokeWidth="0.4" />
      </svg>

      {/* Compás / triángulo girando (centro) */}
      <svg className="arch-shape" style={{ top: '38%', left: '18%', width: 200, height: 200, animation: 'rotateSlowRev 48s linear infinite', opacity: 0.7 }} viewBox="0 0 200 200">
        <polygon className="arch-line" points="100,20 180,170 20,170" strokeWidth="0.5" />
        <polygon className="arch-line-accent arch-dash" points="100,55 150,150 50,150" strokeWidth="0.6" />
        <circle className="arch-line" cx="100" cy="20" r="5" strokeWidth="0.6" />
      </svg>

      {/* Acotaciones flotando (derecha) */}
      <svg className="arch-shape" style={{ top: '20%', right: '12%', width: 160, height: 240, animation: 'drift1 14s ease-in-out infinite' }} viewBox="0 0 100 160">
        <line className="arch-line" x1="20" y1="10" x2="20" y2="150" strokeWidth="0.5" />
        <line className="arch-line" x1="15" y1="10" x2="25" y2="10" strokeWidth="0.5" />
        <line className="arch-line" x1="15" y1="150" x2="25" y2="150" strokeWidth="0.5" />
        <line className="arch-line-accent" x1="20" y1="80" x2="80" y2="80" strokeWidth="0.5" />
        <circle className="arch-line-accent" cx="80" cy="80" r="3" strokeWidth="0.6" />
      </svg>

      {/* Hexágono lento (abajo derecha) */}
      <svg className="arch-shape" style={{ bottom: '8%', right: '20%', width: 140, height: 140, animation: 'drift2 18s ease-in-out infinite' }} viewBox="0 0 100 100">
        <polygon className="arch-line" points="50,8 88,29 88,71 50,92 12,71 12,29" strokeWidth="0.6" />
        <polygon className="arch-line-accent arch-dash" points="50,22 76,36 76,64 50,78 24,64 24,36" strokeWidth="0.5" />
      </svg>
    </>
  );
});

// ─── Pantalla de bienvenida (splash al abrir) ───────────────────────────────
function SplashScreen({ saliendo }) {
  // Partículas ascendentes con posiciones/tiempos variados
  const particulas = Array.from({ length: 14 }, (_, i) => ({
    left: (i * 7 + (i % 3) * 4) % 100,
    delay: (i % 7) * 0.4,
    dur: 4 + (i % 5),
    size: 2 + (i % 3),
  }));

  return (
    <div className={`splash-wrap${saliendo ? ' out' : ''}`}>
      <FondoArquitectonico />

      {/* Haz de luz que barre */}
      <div className="spl-beam" />

      {/* Partículas ascendentes */}
      {particulas.map((p, i) => (
        <div key={i} className="spl-particle" style={{ left: p.left + '%', width: p.size, height: p.size, animation: `particleRise ${p.dur}s linear ${p.delay}s infinite` }} />
      ))}

      {/* Plano técnico dibujándose y rotando */}
      <svg className="spl-blueprint" viewBox="0 0 200 200" fill="none">
        <circle className="spl-draw d1 arch-line" cx="100" cy="100" r="80" strokeWidth="0.5" />
        <rect className="spl-draw d2 arch-line-accent" x="55" y="55" width="90" height="90" strokeWidth="0.6" />
        <polygon className="spl-draw d3 arch-line" points="100,30 165,140 35,140" strokeWidth="0.5" />
        <circle className="spl-draw d4 arch-line-accent" cx="100" cy="100" r="45" strokeWidth="0.6" />
        <line className="arch-line arch-dash" x1="20" y1="100" x2="180" y2="100" strokeWidth="0.4" />
        <line className="arch-line arch-dash" x1="100" y1="20" x2="100" y2="180" strokeWidth="0.4" />
      </svg>

      {/* Anillos expandiéndose detrás del logo */}
      <div className="spl-ring r1" />
      <div className="spl-ring r2" />
      <div className="spl-ring r3" />

      {/* Logo central */}
      <div style={{ position: 'relative', textAlign: 'center', zIndex: 3 }}>
        <div className="splash-logo">
          PLAAT<span style={{ fontSize: 32, color: '#8AA88A', animation: 'sp-pop .6s ease .8s both' }}>.</span>
        </div>
        <div className="splash-bar" style={{ width: 160, margin: '24px auto 0' }} />
        <div className="splash-sub">BIENVENIDO · DEO · ARQUITECTURA TÉCNICA</div>
      </div>

      {/* Destello final */}
      <div className="spl-flash" />
    </div>
  );
}

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
    <div className="login-bg" style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'radial-gradient(circle at 30% 20%, #232320 0%, #1C1C1A 55%, #161614 100%)', padding: 20 }}>
      <style>{CSS}</style>
      <FondoArquitectonico />

      <div className="login-card" style={{ position: 'relative', zIndex: 2, background: 'rgba(255,255,255,0.97)', borderRadius: 18, padding: '38px 32px', width: 380, maxWidth: '100%', boxShadow: '0 30px 80px rgba(0,0,0,0.45), 0 2px 0 rgba(255,255,255,0.05) inset', backdropFilter: 'blur(8px)' }}>
        <div className="login-logo" style={{ fontSize: 28, fontWeight: 700, letterSpacing: '0.1em', color: '#141412', display: 'flex', alignItems: 'baseline', gap: 8 }}>
          PLAAT
          <span style={{ fontSize: 14, fontWeight: 400, color: '#5A7D5A' }}>.</span>
        </div>
        <div className="login-underline" style={{ height: 3, background: 'linear-gradient(90deg, #5A7D5A, #8AA88A)', borderRadius: 2, margin: '8px 0 6px' }} />
        <div style={{ fontSize: 11.5, letterSpacing: '0.18em', color: '#9B9B97', marginBottom: 28, fontWeight: 400 }}>DEO · ARQUITECTURA TÉCNICA</div>

        <div className="login-row r1" style={{ marginBottom: 12 }}>
          <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 5 }}>Correo</label>
          <input type="email" autoFocus value={email} onChange={e => setEmail(e.target.value)} onKeyDown={onKey} placeholder="tu@correo.com" />
        </div>
        <div className="login-row r2" style={{ marginBottom: 18 }}>
          <label style={{ fontSize: 12, fontWeight: 500, color: '#52524E', display: 'block', marginBottom: 5 }}>Contraseña</label>
          <input type="password" value={pass} onChange={e => setPass(e.target.value)} onKeyDown={onKey} placeholder="••••••••" />
        </div>

        {error && <div className="fade" style={{ fontSize: 12, color: '#8A1F1F', background: '#FDECEC', borderRadius: 8, padding: '8px 11px', marginBottom: 14 }}>{error}</div>}

        <div className="login-row r3 shimmer-btn" style={{ borderRadius: 8 }}>
          <Btn primary full onClick={entrar} disabled={loading || !email.trim() || !pass}>
            {loading ? 'Entrando…' : 'Entrar'}
          </Btn>
        </div>

        <div className="login-row r4" style={{ fontSize: 11, color: '#A5A5A0', marginTop: 18, textAlign: 'center', lineHeight: 1.5 }}>
          Acceso restringido al equipo de PLAAT.<br />Si no tienes cuenta, pídela al administrador.
        </div>
      </div>
    </div>
  );
}

// ─── App principal ────────────────────────────────────────────────────────────

export default function App() {
  const isMobile = useIsMobile();
  const [user,       setUser]       = useState(undefined);
  const [obras,      setObras]      = useState([]);
  const [loading,    setLoading]    = useState(true);
  const [nav,        setNav]        = useState('alertas');
  const [obraActiva, setObraActiva] = useState(null);
  const [newVersion, setNewVersion] = useState(false);

  // Detectar nou deployment de la PWA
  useEffect(() => {
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.ready.then(reg => {
        reg.addEventListener('updatefound', () => {
          const newSW = reg.installing;
          if (newSW) {
            newSW.addEventListener('statechange', () => {
              if (newSW.state === 'installed' && navigator.serviceWorker.controller) {
                setNewVersion(true);
              }
            });
          }
        });
        // Comprovar si ja hi ha una actualització pendent
        if (reg.waiting && navigator.serviceWorker.controller) {
          setNewVersion(true);
        }
      });
    }
  }, []);
  const [showNueva,  setShowNueva]  = useState(false);
  const [obraEditar,   setObraEditar]   = useState(null);
  const [obraEliminar, setObraEliminar] = useState(null);
  const [showBackup,   setShowBackup]   = useState(false);
  const [splash,       setSplash]       = useState(true);   // pantalla de bienvenida
  const [splashOut,    setSplashOut]    = useState(false);  // fase de desvanecido

  // Splash de bienvenida: visible ~3s, luego se desvanece
  useEffect(() => {
    const t1 = setTimeout(() => setSplashOut(true), 2900);
    const t2 = setTimeout(() => setSplash(false), 3600);
    return () => { clearTimeout(t1); clearTimeout(t2); };
  }, []);
  const [importando,   setImportando]   = useState(false);
  const [backupMsg,    setBackupMsg]    = useState('');
  const [driveToken,   setDriveToken]   = useState(null);
  const [backupAuto,   setBackupAuto]   = useState(false); // true mientras hace backup automático

  const GDRIVE_CLIENT_ID = '517770541554-rbp3cnlas227d38svonvc5cpt608cr2p.apps.googleusercontent.com';
  const GDRIVE_SCOPE     = 'https://www.googleapis.com/auth/drive.file';
  const BACKUP_KEY       = 'plaat_last_backup'; // localStorage para recordar fecha

  // ── Lógica Google Drive ───────────────────────────────────────────────────
  function cargarGoogleAPI() {
    return new Promise((res, rej) => {
      if (window.google?.accounts?.oauth2) { res(); return; }
      const s = document.createElement('script');
      s.src = 'https://accounts.google.com/gsi/client';
      s.onload = () => res(); s.onerror = () => rej(new Error('No se pudo cargar Google API'));
      document.head.appendChild(s);
    });
  }

  async function obtenerToken() {
    if (driveToken) return driveToken;
    await cargarGoogleAPI();
    return new Promise((res, rej) => {
      const client = window.google.accounts.oauth2.initTokenClient({
        client_id: GDRIVE_CLIENT_ID,
        scope: GDRIVE_SCOPE,
        callback: (resp) => {
          if (resp.error) { rej(new Error(resp.error)); return; }
          setDriveToken(resp.access_token);
          res(resp.access_token);
        },
      });
      client.requestAccessToken();
    });
  }

  async function subirADrive(token, nombreArchivo, contenidoJson) {
    // 1. Busca o crea carpeta "PLAAT DEO Backups"
    const buscarCarpeta = await fetch(
      `https://www.googleapis.com/drive/v3/files?q=name%3D'PLAAT+DEO+Backups'+and+mimeType%3D'application%2Fvnd.google-apps.folder'+and+trashed%3Dfalse&fields=files(id,name)`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const resCarpeta = await buscarCarpeta.json();
    let carpetaId;
    if (resCarpeta.files && resCarpeta.files.length > 0) {
      carpetaId = resCarpeta.files[0].id;
    } else {
      const crearCarpeta = await fetch('https://www.googleapis.com/drive/v3/files', {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: 'PLAAT DEO Backups', mimeType: 'application/vnd.google-apps.folder' }),
      });
      const nuevaCarpeta = await crearCarpeta.json();
      carpetaId = nuevaCarpeta.id;
    }

    // 2. Sube el archivo JSON
    const metadata = { name: nombreArchivo, mimeType: 'application/json', parents: [carpetaId] };
    const form = new FormData();
    form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
    form.append('file', new Blob([contenidoJson], { type: 'application/json' }));
    const subir = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}` },
      body: form,
    });
    if (!subir.ok) throw new Error(`Error al subir: ${subir.status}`);
    return await subir.json();
  }

  async function exportarBackup(auto = false) {
    try {
      const timestamp = new Date().toISOString().slice(0, 10);
      const json = JSON.stringify({ version: 2, fecha: new Date().toISOString(), obras }, null, 2);
      const nombreArchivo = `PLAAT_DEO_backup_${timestamp}.json`;

      // Siempre descarga local también
      const blob = new Blob([json], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = nombreArchivo;
      document.body.appendChild(a); a.click();
      document.body.removeChild(a); URL.revokeObjectURL(url);

      // Sube a Google Drive
      const token = await obtenerToken();
      await subirADrive(token, nombreArchivo, json);

      // Guarda fecha del último backup
      localStorage.setItem(BACKUP_KEY, new Date().toISOString());

      const kb = Math.round(json.length / 1024);
      setBackupMsg(`✓ Backup guardado en Google Drive y descargado localmente. ${obras.length} obras, ${kb} KB`);
    } catch (e) {
      if (e.message === 'popup_closed_by_user' || e.message?.includes('popup')) {
        setBackupMsg('Cancelado — ventana de Google cerrada.');
      } else {
        // Si falla Drive, al menos la descarga local ya se hizo
        setBackupMsg(`✓ Descargado localmente (Google Drive falló: ${e.message})`);
      }
    }
    setBackupAuto(false);
  }

  // Sesión
  useEffect(() => {
    if (!window.auth) { setUser({ local: true, id: 'local' }); return; }
    window.auth.getUser().then(u => setUser(u || null));
    window.auth.onChange(u => {
      setUser(u || null);
      // Limpiar estado al cambiar de usuario
      if (!u) {
        setObras([]);
        setObraActiva(null);
        setNav('alertas');
      }
    });
  }, []);

  // Cargar obras cuando hay usuario
  useEffect(() => {
    if (!user) return;
    const userId = user.id || user.sub || 'local';
    cargarObras(userId);
  }, [user]);



  async function importarBackup(file) {
    setImportando(true); setBackupMsg('');
    try {
      const text = await file.text();
      const data = JSON.parse(text);
      const lista = data.obras || (Array.isArray(data) ? data : []);
      if (!lista.length) { setBackupMsg('El archivo no contiene obras válidas.'); setImportando(false); return; }
      const userId = user?.id || user?.sub || 'local';
      for (const o of lista) {
        await window.db.upsertObra(obraRow(o, userId));
        for (const inc of (o.incidencias || [])) await window.db.upsertModulo('incidencias', { id: inc.id, obra_id: o.id, data: inc, updated_at: now() });
        if (o.actaVO) await window.db.upsertModulo('actas_vo', { id: o.id + '_vo', obra_id: o.id, data: o.actaVO, updated_at: now() });
        for (const acta of (o.actasInsp || [])) await window.db.upsertModulo('actas_insp', { id: acta.id, obra_id: o.id, data: acta, updated_at: now() });
        for (const nota of (o.apuntes || [])) await window.db.upsertModulo('notas', { id: nota.id, obra_id: o.id, data: nota, updated_at: now() });
        if (o.materiales?.length || o.seguimientoCQ?.length) await window.db.upsertModulo('calidad', { id: o.id + '_cal', obra_id: o.id, data: { materiales: o.materiales || [], seguimientoCQ: o.seguimientoCQ || [] }, updated_at: now() });
      }
      await cargarObras(userId);
      localStorage.setItem(BACKUP_KEY, new Date().toISOString());
      setBackupMsg(`✓ Restauradas ${lista.length} obras correctamente.`);
    } catch(e) { setBackupMsg('Error al importar: ' + e.message); }
    setImportando(false);
  }

  // ── Helpers para separar datos por módulo ──────────────────────────────────
  function obraRow(o, userId) {
    return {
      id: o.id,
      user_id: userId,
      data: {
        nombre: o.nombre, cliente: o.cliente, direccion: o.direccion,
        responsable: o.responsable, estado: o.estado, tipo: o.tipo,
        diasVisita: o.diasVisita, emplazamiento: o.emplazamiento,
        propiedad: o.propiedad, proyectista: o.proyectista,
        direccionObra: o.direccionObra, constructora: o.constructora,
        deoFirmante: o.deoFirmante, numActaSeq: o.numActaSeq,
        fechaCFO: o.fechaCFO || '',
        fases: o.fases, disciplinas: o.disciplinas, lotes: o.lotes,
        creadaEn: o.creadaEn,
      },
      updated_at: now(),
    };
  }

  function rowToObra(row, modulos) {
    const d = row.data || {};
    return {
      id: row.id,
      nombre: d.nombre, cliente: d.cliente, direccion: d.direccion,
      responsable: d.responsable, estado: d.estado, tipo: d.tipo,
      diasVisita: d.diasVisita, emplazamiento: d.emplazamiento,
      propiedad: d.propiedad, proyectista: d.proyectista,
      direccionObra: d.direccionObra, constructora: d.constructora,
      deoFirmante: d.deoFirmante, numActaSeq: d.numActaSeq,
      fechaCFO: d.fechaCFO || '',
      fases: d.fases || [], disciplinas: d.disciplinas || [],
      lotes: d.lotes || [], creadaEn: d.creadaEn,
      incidencias: (modulos?.incidencias || []).map(r => r.data),
      actaVO: modulos?.actas_vo?.[0]?.data || null,
      actasInsp: (modulos?.actas_insp || []).map(r => r.data),
      apuntes: (modulos?.notas || []).map(r => r.data),
      materiales: modulos?.calidad?.[0]?.data?.materiales || [],
      seguimientoCQ: modulos?.calidad?.[0]?.data?.seguimientoCQ || [],
    };
  }

  // ── Migración automática: storage antiguo → tablas nuevas ─────────────────
  async function migrarSiNecesario(userId) {
    // Migración desactivada — los datos ya están en Supabase.
    return false;
  }

  // ── Cargar todas las obras del usuario ────────────────────────────────────
  // OPTIMIZADO: carga rápida en 2 fases
  // Fase 1 (inmediata): metadatos de obras → mostrar tablero
  // Fase 2 (segundo plano): módulos de cada obra → enriquecer cards
  async function cargarObras(userId) {
    if (!window.db) return;
    setLoading(true);
    try {
      await migrarSiNecesario(userId);
      const rows = await window.db.getObras();
      // FASE 1: mostrar obras inmediatamente con datos básicos
      const obrasParciales = rows.map(row => ({
        ...rowToObra(row, {}),
        _rol: 'deo',
        _cargando: true,
      }));
      setObras(obrasParciales);
      setLoading(false);

      // FASE 2: enriquecer cada obra con sus módulos en paralelo (segundo plano)
      const obrasCompletas = await Promise.all(rows.map(async row => {
        const [incs, vos, insps, notas, cal, rol] = await Promise.all([
          window.db.getModulo('incidencias', row.id),
          window.db.getModulo('actas_vo', row.id),
          window.db.getModulo('actas_insp', row.id),
          window.db.getModulo('notas', row.id),
          window.db.getModulo('calidad', row.id),
          window.db.getRolUsuario(row.id, userId),
        ]);
        return { ...rowToObra(row, { incidencias: incs, actas_vo: vos, actas_insp: insps, notas, calidad: cal }), _rol: rol };
      }));
      setObras(obrasCompletas);
      setTimeout(() => migrarFotosAntiguas(obrasCompletas), 2000);
    } catch (e) { console.error('Error cargando obras:', e); setLoading(false); }
  }

  // ── Migración de fotos antiguas (base64 → Storage) en segundo plano ────────
  async function migrarFotosAntiguas(listaObras) {
    if (!window.db?.subirFoto) return;
    let hayActualizaciones = false;
    const basePublic = `${window._supabaseUrl}/storage/v1/object/public/plaat-fotos/`;

    // Helper: comprova si una URL és signada (caduca) o pública (no caduca)
    const esUrlSignada = url => url && (url.includes('/sign/') || url.includes('token='));
    // Helper: reconstrueix URL pública a partir del path
    const urlPublica = path => path ? basePublic + path : null;

    for (const obra of listaObras) {
      let obraModificada = false;
      const incidenciasNuevas = [];

      for (const inc of (obra.incidencias || [])) {
        let incModificada = false;
        const historialNuevo = await Promise.all((inc.historial || []).map(async h => {
          const adjuntosNuevos = await Promise.all((h.adjuntos || []).map(async a => {
            // Si té path, actualitzar a URL pública
            if (a.path && esUrlSignada(a.url || '')) {
              incModificada = true;
              return { ...a, url: urlPublica(a.path), data: undefined };
            }
            // Si té base64 però no path, pujar a Storage
            if (!a.path && !a.url && a.data && a.data.startsWith('data:image')) {
              try {
                const { path, url } = await window.db.subirFoto(obra.id, a.id || uid(), a.data);
                incModificada = true;
                return { ...a, path, url, data: undefined };
              } catch(e) { return a; }
            }
            return a;
          }));
          return incModificada ? { ...h, adjuntos: adjuntosNuevos } : h;
        }));

        if (incModificada) {
          const incActualizada = { ...inc, historial: historialNuevo };
          await window.db.upsertModulo('incidencias', { id: inc.id, obra_id: obra.id, data: incActualizada, updated_at: now() });
          incidenciasNuevas.push(incActualizada);
          obraModificada = true;
        } else {
          incidenciasNuevas.push(inc);
        }
      }

      // Migrar fotos de Acta VO — regenerar URLs signades caducades
      if (obra.actaVO?.secciones) {
        let voModificado = false;
        const seccionesNuevas = (obra.actaVO.secciones || []).map(sec => ({
          ...sec,
          temas: (sec.temas || []).map(t => ({
            ...t,
            entradas: (t.entradas || []).map(en => ({
              ...en,
              fotos: (en.fotos || []).map(f => {
                // Si té path i URL signada caducada → URL pública
                if (f.path && esUrlSignada(f.url || '')) {
                  voModificado = true;
                  return { id: f.id, path: f.path, url: urlPublica(f.path) };
                }
                // Si té base64 però no path → marcar per pujar (ho ignorem ara, massa lent)
                return f;
              }),
            })),
          })),
        }));

        // Migrar fotos d'estat de l'obra
        const fotosEstat = (obra.actaVO.estadoObra?.fotos || []).map(f => {
          if (f.path && esUrlSignada(f.url || '')) {
            voModificado = true;
            return { id: f.id, path: f.path, url: urlPublica(f.path) };
          }
          return f;
        });

        if (voModificado) {
          const voActualizado = {
            ...obra.actaVO,
            secciones: seccionesNuevas,
            estadoObra: { ...(obra.actaVO.estadoObra || {}), fotos: fotosEstat },
          };
          await window.db.upsertModulo('actas_vo', { id: obra.id + '_vo', obra_id: obra.id, data: voActualizado, updated_at: now() });
          obraModificada = true;
        }
      }

      if (obraModificada) hayActualizaciones = true;
    }

    if (hayActualizaciones) {
      console.log('URLs de fotos migrades a públiques OK');
      const userId = user?.id || user?.sub;
      if (userId) cargarObras(userId);
    }
  }

  // ── Guardar obra: solo la parte que cambió ─────────────────────────────────
  async function saveUnaObra(obra, lista, anterior) {
    if (!window.db || !user) return;
    // PROTECCIÓ CRÍTICA: mai guardar si l'obra encara s'està carregant (Fase 1)
    // En Fase 1 els mòduls són null/buits i sobreescriurien dades reals a Supabase
    if (obra._cargando) {
      console.warn('saveUnaObra bloquejat: obra en fase de càrrega', obra.id);
      return;
    }
    try {
      const userId = user.id || user.sub;
      // Datos generales de la obra — siempre se guarda (es ligero, sin módulos)
      await window.db.upsertObra(obraRow(obra, userId));

      // Si no hay referencia "anterior", guardamos todo (compatibilidad)
      const prev = anterior || {};

      // Incidencias: solo las que cambiaron (comparación por referencia)
      const incActuales = obra.incidencias || [];
      const incAnteriores = prev.incidencias || [];
      const incCambiadas = incActuales.filter(inc => {
        const old = incAnteriores.find(i => i.id === inc.id);
        return !old || old !== inc;
      });
      for (const inc of incCambiadas) {
        await window.db.upsertModulo('incidencias', { id: inc.id, obra_id: obra.id, data: inc, updated_at: now() });
      }

      // Acta VO: solo si cambió respecto a la anterior
      // PROTECCIÓ: mai guardar si l'obra està en fase de càrrega (_cargando)
      // PROTECCIÓ: mai sobreescriure amb null/undefined (significaria que els mòduls no s'han carregat encara)
      if (!obra._cargando && obra.actaVO !== undefined && obra.actaVO !== null && obra.actaVO !== prev.actaVO) {
        await window.db.upsertModulo('actas_vo', { id: obra.id + '_vo', obra_id: obra.id, data: obra.actaVO, updated_at: now() });
      }

      // Actas Inspección: solo las que cambiaron
      const actasInsp = obra.actasInsp || [];
      const actasAnteriores = prev.actasInsp || [];
      const actasCambiadas = actasInsp.filter(acta => {
        const old = actasAnteriores.find(a => a.id === acta.id);
        return !old || old !== acta;
      });
      for (const acta of actasCambiadas) {
        await window.db.upsertModulo('actas_insp', { id: acta.id, obra_id: obra.id, data: acta, updated_at: now() });
      }

      // Notas/Apuntes: solo las que cambiaron
      const apuntes = obra.apuntes || [];
      const apuntesAnteriores = prev.apuntes || [];
      const apuntesCambiados = apuntes.filter(nota => {
        const old = apuntesAnteriores.find(n => n.id === nota.id);
        return !old || old !== nota;
      });
      for (const nota of apuntesCambiados) {
        await window.db.upsertModulo('notas', { id: nota.id, obra_id: obra.id, data: nota, updated_at: now() });
      }

      // Calidad: solo si cambió
      const calCambio = obra.materiales !== prev.materiales || obra.seguimientoCQ !== prev.seguimientoCQ;
      if (calCambio && (obra.materiales !== undefined || obra.seguimientoCQ !== undefined)) {
        await window.db.upsertModulo('calidad', { id: obra.id + '_cal', obra_id: obra.id, data: { materiales: obra.materiales || [], seguimientoCQ: obra.seguimientoCQ || [] }, updated_at: now() });
      }
    } catch (e) {
      console.error('Error guardando obra:', e);
      alert('No se pudieron guardar los cambios: ' + e.message);
    }
  }

  async function crearObra(data) {
    const obra = {
      id: uid(), nombre: data.nombre, cliente: data.cliente, direccion: data.direccion,
      responsable: data.responsable, diasVisita: data.diasVisita || [],
      emplazamiento: data.emplazamiento || '', propiedad: data.propiedad || '',
      proyectista: data.proyectista || '', direccionObra: data.direccionObra || '',
      constructora: data.constructora || '', deoFirmante: data.deoFirmante || '',
      fechaCFO: data.fechaCFO || '',
      numActaSeq: 0, estado: 'en_curso',
      disciplinas: [], lotes: [], incidencias: [], apuntes: [],
      materiales: [], seguimientoCQ: [], actasInsp: [],
      creadaEn: now(), _rol: 'owner',
    };
    const lista = [obra, ...obras];
    setObras(lista);
    const userId = user?.id || user?.sub;
    try {
      // RPC atómica: crea obra + registra owner en una sola transacción
      const obraRow = {
        id: obra.id, user_id: userId,
        data: {
          nombre: obra.nombre, cliente: obra.cliente, direccion: obra.direccion,
          responsable: obra.responsable, estado: obra.estado, tipo: obra.tipo,
          diasVisita: obra.diasVisita, emplazamiento: obra.emplazamiento,
          propiedad: obra.propiedad, proyectista: obra.proyectista,
          direccionObra: obra.direccionObra, constructora: obra.constructora,
          deoFirmante: obra.deoFirmante, numActaSeq: obra.numActaSeq,
          fechaCFO: obra.fechaCFO || '',
          fases: obra.fases, disciplinas: obra.disciplinas, lotes: obra.lotes,
          creadaEn: obra.creadaEn,
        },
        updated_at: now(),
      };
      await window.db.crearObraConOwner(obra.id, obraRow.data);
    } catch(e) {
      console.error('Error creando obra:', e);
      alert('No se pudo guardar la obra: ' + e.message);
    }
    setShowNueva(false);
    setObraActiva(obra);
  }

  async function actualizarObra(updated) {
    const anterior = obras.find(o => o.id === updated.id);
    const lista = obras.map(o => o.id === updated.id ? updated : o);
    setObras(lista);
    setObraActiva(updated);
    await saveUnaObra(updated, lista, anterior);
  }

  async function guardarEdicion(datos) {
    const updated = { ...obraEditar, ...datos };
    const anterior = obraEditar;
    const lista = obras.map(o => o.id === obraEditar.id ? updated : o);
    setObras(lista);
    await saveUnaObra(updated, lista, anterior);
    setObraEditar(null);
  }

  async function eliminarObra() {
    const id = obraEliminar.id;
    const lista = obras.filter(o => o.id !== id);
    setObras(lista);
    try {
      await window.db.deleteObra(id);
    } catch (e) { console.error('Error eliminando obra:', e); }
    setObraEliminar(null);
  }

  // ── Realtime: escuchar cambios mientras la app está abierta ───────────────
  useEffect(() => {
    if (!user || !window.db) return;
    const userId = user.id || user.sub;
    const unsub = window.db.subscribeListaObras(async () => {
      const rows = await window.db.getObras();
      const obrasCompletas = await Promise.all(rows.map(async row => {
        const [incs, vos, insps, notas, cal] = await Promise.all([
          window.db.getModulo('incidencias', row.id),
          window.db.getModulo('actas_vo', row.id),
          window.db.getModulo('actas_insp', row.id),
          window.db.getModulo('notas', row.id),
          window.db.getModulo('calidad', row.id),
        ]);
        const rol = await window.db.getRolUsuario(row.id, userId);
        return { ...rowToObra(row, { incidencias: incs, actas_vo: vos, actas_insp: insps, notas, calidad: cal }), _rol: rol };
      }));
      setObras(obrasCompletas);
    });
    return unsub;
  }, [user]);

  // Realtime solo para lista de obras (nuevas obras de compañeros)
  // La obra activa NO usa Realtime — se actualiza solo desde estado local
  // Esto evita el bug de "elemento vuelve a aparecer tras borrar"

  // Calcular el contador del badge "Hoy"
  const stats = {
    total: obras.length,
    alertas: calcularAlertas(obras).total,
    incidencias: obras.reduce((s, o) => s + o.incidencias.filter(i => i.estado !== 'resuelta').length, 0),
  };

  // ── Portón de acceso ──
  // Pantalla de bienvenida al abrir la app
  if (splash) {
    return (
      <>
        <style>{CSS}</style>
        <SplashScreen saliendo={splashOut} />
      </>
    );
  }

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
  const UpdateBanner = () => newVersion ? (
    <div style={{ position: 'fixed', top: 0, left: 0, right: 0, zIndex: 9999,
      background: '#1C1C1A', color: '#fff', padding: '10px 18px',
      display: 'flex', alignItems: 'center', gap: 12, fontSize: 13 }}>
      <span style={{ flex: 1 }}>🔄 Nova versió de PLAAT disponible</span>
      <button onClick={() => { setNewVersion(false); window.location.reload(); }}
        style={{ background: '#fff', color: '#1C1C1A', border: 'none', borderRadius: 8,
          padding: '5px 14px', fontWeight: 600, cursor: 'pointer', fontSize: 12 }}>
        Actualitzar ara
      </button>
      <button onClick={() => setNewVersion(false)}
        style={{ background: 'none', border: 'none', color: '#A5A5A0', cursor: 'pointer', fontSize: 18, lineHeight: 1 }}>×</button>
    </div>
  ) : null;

  if (obraActiva) {
    const fresh = obras.find(o => o.id === obraActiva.id) || obraActiva;
    // Si l'obra encara s'està carregant (Fase 1), mostrem un spinner lleuger
    // fins que la Fase 2 la tingui completa amb tots els mòduls
    return (
      <>
        <style>{CSS}</style>
        <div style={{ display: 'flex', flexDirection: isMobile ? 'column' : 'row', height: '100vh', overflow: 'hidden' }}>
          {!isMobile && <Sidebar nav={nav} setNav={setNav} stats={stats} user={user} onBackup={() => { setShowBackup(true); setBackupMsg(""); }} />}
          <div style={{ flex: 1, overflow: 'hidden', display: 'flex', flexDirection: 'column', minHeight: 0 }}>
            <UpdateBanner />
      <DetalleObra obra={fresh} onBack={() => setObraActiva(null)} onSave={actualizarObra} isMobile={isMobile} user={user} />
          </div>
        </div>
      </>
    );
  }

  return (
    <>
      <style>{CSS}</style>
      <div style={{ display: 'flex', flexDirection: isMobile ? 'column' : 'row', height: '100vh', overflow: 'hidden' }}>
        {!isMobile && <Sidebar nav={nav} setNav={setNav} stats={stats} user={user} onBackup={() => { setShowBackup(true); setBackupMsg(""); }} />}
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', minHeight: 0 }}>

          <UpdateBanner />
      {nav === 'alertas'      && <VistaAlertas obras={obras} onIrObra={o => setObraActiva(o)} isMobile={isMobile} />}
          {nav === 'seguimiento'  && <VistaSeguimiento obras={obras} isMobile={isMobile} />}
          {nav === 'tablero'   && (
            <>
              {/* Banner con fondo arquitectónico animado */}
              <div className="dash-banner" style={{ borderBottom: '1px solid #ECEAE4', padding: isMobile ? '18px 16px' : '22px 22px', paddingTop: isMobile ? 'calc(18px + env(safe-area-inset-top))' : '22px', display: 'flex', alignItems: 'center', gap: 12, flexShrink: 0 }}>
                <div className="arch-grid" />
                <div className="dash-banner-ring db-r1" />
                <div className="dash-banner-ring db-r2" />
                <div className="dash-beam" />
                <div className="dash-title" style={{ flex: 1, position: 'relative', zIndex: 2 }}>
                  <div style={{ fontSize: isMobile ? 18 : 20, fontWeight: 700, color: '#F2F1ED', letterSpacing: '0.01em', display: 'flex', alignItems: 'baseline', gap: 7 }}>
                    Mis obras <span style={{ fontSize: 13, color: '#8AA88A' }}>.</span>
                  </div>
                  {obras.length > 0 && <div style={{ fontSize: 12, color: 'rgba(255,255,255,0.5)', marginTop: 3 }}>{obras.filter(o => o.estado === 'en_curso').length} en curso · {obras.length} en total</div>}
                </div>
                <div style={{ position: 'relative', zIndex: 2 }}>
                  <button onClick={() => setShowNueva(true)} className="shimmer-btn tap" style={{ display: 'inline-flex', alignItems: 'center', gap: 5, padding: '8px 16px', borderRadius: 8, border: '1.5px solid #7A9D7A', background: '#5A7D5A', color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer', whiteSpace: 'nowrap' }}>
                    + Nueva obra
                  </button>
                </div>
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
                  <>
                    {/* Resumen con tarjetas animadas */}
                    <div style={{ display: 'grid', gridTemplateColumns: isMobile ? 'repeat(2,1fr)' : 'repeat(4,1fr)', gap: 10, marginBottom: 18 }}>
                      {[
                        { n: obras.length, l: 'Obras totales', c: '#141412', s: 's1' },
                        { n: obras.filter(o => o.estado === 'en_curso').length, l: 'En curso', c: '#C47610', s: 's2' },
                        { n: obras.filter(o => o.estado === 'acabada').length, l: 'Acabadas', c: '#2D5E10', s: 's3' },
                        { n: stats.alertas, l: 'Con alertas', c: '#8A1F1F', s: 's4' },
                      ].map(st => (
                        <div key={st.l} className={`stat-card ${st.s}`} style={{ background: '#fff', borderRadius: 16, boxShadow: '0 1px 2px rgba(0,0,0,0.04), 0 3px 12px rgba(0,0,0,0.04)', padding: isMobile ? '14px 16px' : '18px 20px' }}>
                          <div style={{ fontSize: isMobile ? 26 : 32, fontWeight: 700, color: st.c, lineHeight: 1, letterSpacing: '-0.03em' }}>{st.n}</div>
                          <div style={{ fontSize: 11.5, color: '#9B9B97', marginTop: 7, fontWeight: 500, letterSpacing: '0.02em' }}>{st.l}</div>
                        </div>
                      ))}
                    </div>
                    <div className="list-in" style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                      {obras.map(o => <ObraCard key={o.id} obra={o} onClick={() => {
                        // Si _cargando, espera Fase 2 — busca la versió completa
                        const completa = obras.find(x => x.id === o.id && !x._cargando);
                        setObraActiva(completa || o);
                      }} onEditar={setObraEditar} onEliminar={setObraEliminar} />)}
                    </div>
                  </>
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

      {/* Modal Backup */}
      {showBackup && (
        <Modal title="Copia de seguridad" onClose={() => { setShowBackup(false); setBackupMsg(''); }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, background: '#F0F6FF', border: '1px solid #C8DEFF', borderRadius: 10, padding: '10px 14px', marginBottom: 16 }}>
            <span style={{ fontSize: 20 }}>🔗</span>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: '#0C447C' }}>Google Drive conectado</div>
              <div style={{ fontSize: 11, color: '#4A7AB5' }}>Los backups se guardan en "PLAAT DEO Backups" en tu Drive</div>
            </div>
          </div>

          <div style={{ background: '#F5F4F0', borderRadius: 10, padding: '14px 16px', marginBottom: 14 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: '#52524E', marginBottom: 6 }}>↑ Hacer backup ahora</div>
            <div style={{ fontSize: 12, color: '#6B6B66', marginBottom: 10, lineHeight: 1.5 }}>
              Sube <strong>{obras.length} obras</strong> a Google Drive y descarga una copia local simultáneamente. La primera vez pedirá permiso a Google.
            </div>
            <Btn primary full onClick={() => exportarBackup(false)}>
              {backupAuto ? '⏰ Hacer backup ahora (recomendado)' : '☁️ Guardar en Google Drive'}
            </Btn>
          </div>

          <div style={{ background: '#F5F4F0', borderRadius: 10, padding: '14px 16px', marginBottom: backupMsg ? 12 : 0 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: '#52524E', marginBottom: 6 }}>↓ Restaurar backup</div>
            <p style={{ fontSize: 12, color: '#8A1F1F', background: '#FDECEC', border: '1px solid #F9CACA', borderRadius: 8, padding: '8px 12px', marginBottom: 10, lineHeight: 1.5 }}>
              ⚠️ Esto <strong>reemplaza todos los datos actuales</strong> con los del archivo. Haz un backup primero si tienes datos nuevos.
            </p>
            <label style={{ display: 'block', width: '100%', padding: '9px 14px', borderRadius: 8, border: '1.5px dashed #E0DFD9', background: '#fff', cursor: 'pointer', fontSize: 13, color: '#6B6B66', textAlign: 'center' }}>
              {importando ? 'Restaurando...' : '📂 Seleccionar archivo .json de backup'}
              <input type="file" accept=".json" style={{ display: 'none' }} disabled={importando}
                onChange={e => { if (e.target.files[0]) importarBackup(e.target.files[0]); }} />
            </label>
          </div>

          {backupMsg && (
            <div style={{ marginTop: 12, padding: '10px 14px', borderRadius: 9, background: backupMsg.startsWith('✓') ? '#E8F5E0' : backupMsg.startsWith('⏰') ? '#FEF3DB' : '#FDECEC', color: backupMsg.startsWith('✓') ? '#2D5E10' : backupMsg.startsWith('⏰') ? '#7C4A00' : '#8A1F1F', fontSize: 13, fontWeight: 500, lineHeight: 1.5 }}>
              {backupMsg}
            </div>
          )}

          <div style={{ marginTop: 14, padding: '10px 14px', borderRadius: 9, background: '#F9F8F5', fontSize: 11, color: '#A5A5A0', lineHeight: 1.6 }}>
            <strong>Backup automático:</strong> la app te avisará cuando lleven más de 7 días sin backup y te pedirá hacerlo al abrir la app.
          </div>
        </Modal>
      )}
    </>
  );
}

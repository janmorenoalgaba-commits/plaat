# PLAAT DEO — Context del projecte

PWA de gestió d'obres per a PLAAT Arquitectura Tècnica (React + Supabase + Vercel).
Comunicació habitual amb Jan Moreno: espanyol/català barrejat. Comentaris de codi i UI en català/castellà.

## Stack i estructura

- **Frontend**: React (Vite), un únic fitxer principal `src/App.jsx` (~8.000 línies) + `src/main.jsx` (capa de dades)
- **Backend**: Supabase (Postgres + Auth + Storage)
- **Desplegament**: Vercel, domini `plaat.vercel.app`
- **PDF**: jsPDF, carregat dinàmicament via CDN (`window.jspdf`)
- **Mapes**: Leaflet, carregat dinàmicament via CDN (`window.L`)
- **Clima**: API Open-Meteo (gratuïta, sense clau) — es va provar AEMET però es va revertir per simplicitat

### Fitxers clau
- `src/App.jsx` — tota la lògica de la UI, els mòduls (Incidencias, Calidad, Acta VO, Notas y tareas, Inspecciones) i el generador de PDF
- `src/main.jsx` — capa de dades: client Supabase, `window.db.*` (CRUD), `window.storage.*` (KV legacy), cache offline amb IndexedDB, cua de mutacions pendents, auth
- `src/supabase.js` — credencials del client

## Arquitectura de dades important

- **`obraRow()` / `rowToObra()`** (a `App.jsx`, dins del component `App`): serialitzen/deserialitzen l'obra a/des de Supabase. **Qualsevol camp nou a nivell d'obra (com `climaLat`, `diasVisita`, etc.) s'ha d'afegir explícitament a totes dues funcions**, si no es perd en recarregar — ja ha passat un cop (bug de `climaLat`/`climaLon` no persistents).
- **Guardat amb debounce**: `actualizarObra()` (a `App`) actualitza l'estat local a l'instant i espera 700ms abans d'escriure a Supabase (`saveUnaObra`). Es fa *flush* forçat (`desarPendentsAra()`) en: sortir de l'obra, canviar de pestanya dins de `DetalleObra` (prop `onFlush`), i tancar la pestanya del navegador. **Important**: `actualizarObra` només toca `obraActiva` si l'obra editada ja és l'activa — si no, es podria navegar-hi sense voler en fer una edició ràpida des del tauler.
- **Mode offline**: `main.jsx` cacheja lectures a IndexedDB i encua escriptures quan no hi ha xarxa; les reintenta en recuperar connexió.

## Acta de Visita d'Obra (mòdul més treballat)

Aquest ha estat el focus de la majoria de sessions recents. Estructura de `vo` (objecte `actaVO` de l'obra):

- `num` — número d'acta actual (s'incrementa en exportar)
- `fechaActa` — data de l'acta, **editable** (no sempre coincideix amb el dia en què s'omple)
- `fase`, `lloc` — camps lliures
- `equipo` — array de rols, cadascun amb `grupo` (editable, agrupa visualment) i `personas[]`
- `estadoObra.ubicacions[]` — fotos agrupades per ubicació (secció "A")
- `trabajosEnCurso[]` — secció "B"
- `secciones[]` — les 5 seccions de temes (per defecte: Temas generales, Instalaciones, Control de calidad, LEED & BREAM, Planificación), cadascuna amb `temas[]`
- Cada **tema** té `titulo` + `entradas[]` (seguiments cronològics). Cada entrada té `estado` (P/R/I), `resp[]` (array de responsables), `fecha` (editable), `fotos[]`, i **`nueva`** (boolean manual — decideix si surt en blanc/negreta al PDF; **ja NO es calcula automàticament comparant `actaNum`**, perquè es van important actes antigues fora d'ordre)
- `hitos.{esenciales[], intermedios[]}` — secció numerada dinàmicament (sempre just després de l'última secció de temes)
- `contrataciones[]` — mateixa numeració dinàmica
- `clima[]` — fins a 7 dies; la data triada per l'usuari és sempre **l'últim** dia del rang (els 6 anteriors + aquest)
- `firmasSeleccionadas[]` — array d'ids (`ROLES_FIRMA`) que trien quins rols surten al quadre de firmes; per defecte tots (retrocompatible)
- `climaLat`/`climaLon` — **viuen a nivell d'OBRA, no de `vo`** (es reutilitzen entre actes)

### Generador de PDF
- `generarActaVO_v2()` és l'únic generador actiu (hi ha un `generarActaVO()` v1 sense ús, no tocar-lo)
- Numeració de seccions especials (Hitos/Contrataciones/Clima) és **dinàmica**: `numBaseFinal = vo.secciones.length`, mai fixa
- Text llarg es justifica amb `dibuixarLiniesJustificades()` — té un límit màxim d'espai extra per paraula (evita forats enormes en línies curtes)
- Cerca d'empresa per a firmes (`getEmpresa()`) usa prefixos curts en MAJÚSCULES perquè coincideixin tant en català com en castellà (compte: "SEGUR" no "SEGURE", ja va fallar un cop per aquest motiu)

## UI de l'Acta VO (mòdul `ModuloActaVO`)

- Tres pestanyes internes: **Temes** (contingut real, per defecte), **Equip**, **Dades** (Trabajos en curso + Hitos + Contrataciones + Clima)
- Capçalera *sticky* (compte amb marges negatius verticals combinats amb `position: sticky` — ja va causar un bug de solapament; només s'usen marges negatius horitzontals)
- Botó flotant "+" (crear tema nou) i tots els modals (`Modal`, `ConfirmMini`) usen **`createPortal` a `document.body`** — necessari per YouTube evitar bugs de `position: fixed` dins de contenidors amb scroll niat en mòbil real. **Advertència**: l'entorn de previsualització d'artefactes de Claude (dins del xat) NO permet `import ... from "react-dom"` i mostrarà error — això només afecta la previsualització en aquest xat, no el desplegament real a Vercel.
- Totes les seccions de temes es poden plegar; els temes es llisten en **ordre d'inserció** (no ordenats per estat — es va provar i es va revertir perquè feia que els punts nous no aparegueguessin al final)
- Data per defecte en crear un tema/seguiment: **no és "avui"**, és el dia de visita configurat més recent (`obra.diasVisita`), calculat amb `dataVisitaRecent()`

## Convencions i patrons a seguir

- Comentaris de codi nous: català, seguint l'estil ja present
- Mòbil: sempre `isMobile` (hook `useIsMobile()`) per adaptar mides; inputs han de tenir `font-size: 16px` mínim en mòbil (ja hi ha una regla CSS global que ho força, evita el zoom automàtic d'iOS)
- Hover: classes CSS globals (`.chip-link`, `.dia-toggle`, `.hov-*`) en lloc d'estils inline, ja que `:hover` no funciona en `style={{}}`
- Abans de donar per bo un canvi al PDF, **generar-lo de veritat** amb un arnès de test Node.js + jsPDF i renderitzar-lo a imatge per revisar-lo visualment — no assumir que compila = que es veu bé

## Pendent / coses a vigilar

- El fitxer `App.jsx` és molt gran; **bona oportunitat per dividir-lo en mòduls** ara que es treballa des de Claude Code (per exemple: `pdf/generarActaVO.js`, `components/ObraCard.jsx`, `modules/ModuloActaVO.jsx`, etc.)
- Fer servir git per a l'historial en lloc de còpies manuals de fitxers
- Si es canvia l'esquema de Supabase (columnes, RLS), verificar que no trenqui l'app d'algú que tingui la versió anterior carregada al navegador

// Central app state + config — the single source of truth that every other
// module imports. Extracted from main.js so the app is no longer one giant
// module scope; main.js (and future modules like blocks.js / ui.js) read and
// mutate state exclusively through this module.
//
// SHARING CONTRACT (read before adding to this file):
//
//   * `state` is an OBJECT, exported by reference. Mutating its fields
//     (`state.blocks = ...`, `state.settings.theme = ...`) is visible
//     everywhere that imported it. Never REASSIGN the `state` binding itself
//     (`state = ...`) from another module — that would not propagate. Mutate
//     its fields instead. `load()` replaces the contents via a one-time init
//     in this module only.
//
//   * Plain CONSTANTS (INTERVALS, PALETTE, HOUR_PX, …) are exported `const`.
//     They are never reassigned, so live-binding imports are safe.
//
//   * The three values that get REASSIGNED at runtime — the active block size,
//     the visible anchor day, and the column count — are NOT exported as bare
//     `let`. A `let` reassigned from another module does not propagate back to
//     importers (ES module live bindings are read-only for the importer). So
//     they live behind accessor pairs: getSlotMin/setSlotMin, getAnchor/
//     setAnchor, getDayCols/setDayCols. This is the only contract that stays
//     correct once reads/writes are spread across multiple modules, and it is
//     hard to misuse (you cannot accidentally shadow them with a local copy).

import { startOfDay } from './time.js';

/* ---------- localStorage key + config constants ---------- */
export const KEY = 'timelog.v1';
export const INTERVALS = [60, 30, 20, 15, 10, 6];   // = 60 / n  (n = 1,2,3,4,6,10)
export const CATCHUP_CAP_MS = 2 * 60 * 60 * 1000;   // only ask back ~2h
/* Morgen-Modus boundary: gaps that reach back across today 06:00 local are a
   night's sleep, not forgotten work — see morningMode() in blocks.js. */
export const MORNING_BOUNDARY_HOUR = 6;
/* Calm, Kideon-harmonious block hues — mid-tone, readable on off-white and navy. */
export const PALETTE = ['#c49a6c', '#5f8f8c', '#c0766a', '#6b86b0', '#a07fa6',
  '#8aa173', '#c2935a', '#5b9aa0', '#d0a85c', '#9488c0', '#c98bab', '#79a3c4'];
export const HOUR_PX = 116;
export const PX_PER_MIN = HOUR_PX / 60;
export const DOW = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
export const MON = ['Jan', 'Feb', 'Mär', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez'];

/* ---------- persisted state (mutate fields, never reassign the binding) ---------- */
// First-run theme default: honor the OS preference so dark-mode users don't get
// flash-banged by a light UI on their very first visit. Only affects a FRESH
// state with no stored settings — any persisted `theme` still wins on merge, so
// returning users (incl. those who picked light) are never overridden.
function defaultTheme() {
  try {
    if (window.matchMedia && window.matchMedia('(prefers-color-scheme:dark)').matches) return 'dark';
  } catch (e) { /* no matchMedia → fall through to light */ }
  return 'light';
}

function defaultSettings(slotMin) {
  return {
    // reminderMin defaults to the block size and reminderCoupled to true, so a
    // fresh (and, via the default-fill in parseRaw, an existing) user gets
    // EXACTLY today's behaviour: the reminder cadence equals the block size.
    intervalMin: slotMin, reminderMin: slotMin, reminderCoupled: true,
    soundOn: true, notifyOn: false, introSeen: false,
    notifyNudgeDismissed: false, exportReminderDay: '', exportNotifyDay: '', theme: defaultTheme(),
  };
}

/* Parse one raw localStorage payload into a normalized {blocks, recentLabels,
   settings, savedAt} bundle, or null when missing/corrupt. `savedAt` is the
   writer's Date.now() stamp (undefined for legacy payloads) — see save(). */
function parseRaw(raw, slotMin) {
  try {
    const s = JSON.parse(raw);
    if (s && s.blocks) {
      // Default-fill reminderMin from the payload's OWN intervalMin (not the
      // caller's slotMin) so an existing user with intervalMin=20 and no
      // reminder fields inherits reminder cadence = 20 — i.e. exactly today's
      // behaviour. defaultSettings also seeds reminderCoupled ← true.
      const stored = s.settings || {};
      const base = defaultSettings(INTERVALS.includes(stored.intervalMin) ? stored.intervalMin : slotMin);
      return {
        blocks: s.blocks,
        recentLabels: s.recentLabels || [],
        settings: Object.assign(base, stored),
        savedAt: Number.isFinite(s.savedAt) ? s.savedAt : undefined,
      };
    }
  } catch (e) { /* corrupt storage → fall through to a fresh default state */ }
  return null;
}

// Initial load. SLOT_MIN starts at 15 (the default block size) and is then
// reconciled with the persisted setting below — identical to the old inline
// sequence in main.js.
const _loaded = parseRaw(localStorage.getItem(KEY), 15);
export const state = _loaded
  ? { blocks: _loaded.blocks, recentLabels: _loaded.recentLabels, settings: _loaded.settings }
  : { blocks: [], recentLabels: [], settings: defaultSettings(15) };

/* ---------- multi-tab-safe persistence ----------
   Two open instances (PWA window + browser tab) share timelog.v1. A blind
   "write the whole in-memory state" is last-write-wins: the stale tab's next
   save would silently drop everything the other tab logged. Instead:

     * save() re-reads localStorage SYNCHRONOUSLY first. If another tab wrote
       since our last sync, its payload is three-way-merged (base = our last
       synced snapshot, mine = in-memory, theirs = storage) at slot level
       BEFORE we write. Because every writer does this, each write descends
       from the previous one — writes form a linear history and nothing is
       ever clobbered, even when the async `storage` event hasn't arrived yet.
     * initCrossTabSync() additionally listens for the `storage` event so the
       OTHER tab's UI updates promptly (re-render via the callback).
     * Deletions only propagate when `theirs` is provably newer than our base
       (savedAt stamp); otherwise the merge is purely additive — the failure
       mode is a resurrected block, never a lost one. */

let _synced = localStorage.getItem(KEY);   // raw string we last read/wrote
let _base = snapshot(state);               // parsed content at that point
let _baseSavedAt = _loaded?.savedAt || 0;

function snapshot(s) {
  return JSON.parse(JSON.stringify({ blocks: s.blocks, recentLabels: s.recentLabels, settings: s.settings }));
}
function contentStr(s) {
  return JSON.stringify({ blocks: s.blocks, recentLabels: s.recentLabels, settings: s.settings });
}
function absorb(m) { state.blocks = m.blocks; state.recentLabels = m.recentLabels; state.settings = m.settings; }

/** Pure three-way merge of two divergent states over a common base.
    Blocks are keyed by their slot start time; per slot, a side that changed
    it (vs base) wins, with `mine` winning a direct conflict (we are the
    active writer). A slot missing from `theirs` is only treated as a
    deletion when trustTheirsDeletes is set — otherwise kept (additive). */
export function mergeStates(base, mine, theirs, trustTheirsDeletes = true) {
  const mapOf = (blocks) => { const m = new Map(); for (const b of blocks || []) m.set(new Date(b.start).getTime(), b); return m; };
  const bm = mapOf(base.blocks), mm = mapOf(mine.blocks), tm = mapOf(theirs.blocks);
  const eq = (a, b) => a === b ||
    (!!a && !!b && a.label === b.label && new Date(a.end).getTime() === new Date(b.end).getTime());
  const blocks = [];
  for (const k of new Set([...bm.keys(), ...mm.keys(), ...tm.keys()])) {
    const b = bm.get(k), m = mm.get(k), t = tm.get(k);
    let pick;
    if (!eq(m, b)) pick = m;                       // we changed/added/deleted it → ours wins
    else if (tm.has(k)) pick = t;                   // theirs (changed or unchanged)
    else pick = trustTheirsDeletes ? undefined : m; // absent in theirs: delete vs. keep
    if (pick) blocks.push(pick);
  }
  blocks.sort((a, b) => new Date(a.start) - new Date(b.start));
  // recentLabels: the side that changed them leads, the other fills up.
  const mineChanged = JSON.stringify(mine.recentLabels) !== JSON.stringify(base.recentLabels);
  const lead = (mineChanged ? mine.recentLabels : theirs.recentLabels) || [];
  const tail = (mineChanged ? theirs.recentLabels : mine.recentLabels) || [];
  const recentLabels = [...new Set([...lead, ...tail])].slice(0, 12);
  // settings: per key — our local change wins, otherwise theirs.
  const settings = {};
  const bs = base.settings || {}, ms = mine.settings || {}, ts = theirs.settings || {};
  for (const k of new Set([...Object.keys(bs), ...Object.keys(ms), ...Object.keys(ts)])) {
    settings[k] = (ms[k] !== bs[k]) ? ms[k] : (k in ts ? ts[k] : ms[k]);
  }
  return { blocks, recentLabels, settings };
}

let _onSaveError = null;
/** Register the one handler invoked when persisting fails (e.g. quota
    exceeded) — main.js wires this to a visible "export your data!" toast. */
export function onSaveError(cb) { _onSaveError = cb; }

/** Persist the current state object to localStorage — multi-tab-safe (merges
    any concurrent write from another tab first) and guarded: a failed write
    (QuotaExceededError, private mode, …) never throws into the caller; it
    reports through onSaveError instead so the UI can warn the user. */
export function save() {
  try {
    const raw = localStorage.getItem(KEY);
    if (raw !== null && raw !== _synced) {
      const theirs = parseRaw(raw, state.settings.intervalMin || 15);
      if (theirs) {
        absorb(mergeStates(_base, state, theirs,
          Number.isFinite(theirs.savedAt) && theirs.savedAt >= _baseSavedAt));
      }
    }
    const savedAt = Date.now();
    const out = JSON.stringify({
      blocks: state.blocks, recentLabels: state.recentLabels, settings: state.settings, savedAt,
    });
    localStorage.setItem(KEY, out);
    _synced = out; _base = snapshot(state); _baseSavedAt = savedAt;
  } catch (e) {
    if (_onSaveError) _onSaveError(e);
  }
}

/** Pull another tab's write into memory (merge, never clobber). Reads the
    CURRENT storage content (not the event payload — queued stale events thus
    collapse into one up-to-date sync). Returns true iff memory changed. */
export function syncFromStorage() {
  const raw = localStorage.getItem(KEY);
  if (raw === null || raw === _synced) return false;   // our own write was last
  const theirs = parseRaw(raw, state.settings.intervalMin || 15);
  if (!theirs) return false;
  const before = contentStr(state);
  absorb(mergeStates(_base, state, theirs,
    Number.isFinite(theirs.savedAt) && theirs.savedAt >= _baseSavedAt));
  _synced = raw;
  _base = snapshot(theirs);
  _baseSavedAt = Math.max(_baseSavedAt, theirs.savedAt || 0);
  // If the merge kept local data `theirs` lacks (exotic write race), write the
  // union back so both tabs converge on it.
  if (contentStr(state) !== contentStr(theirs)) save();
  return contentStr(state) !== before;
}

/** Wire the cross-tab `storage` listener (browser only — called from main.js
    so this module stays loadable in Node tests). onExternalChange fires after
    another tab's write has been merged into memory; re-render there. */
export function initCrossTabSync(onExternalChange) {
  window.addEventListener('storage', (e) => {
    if (e.key !== KEY) return;
    if (syncFromStorage() && onExternalChange) onExternalChange();
  });
}

/* ---------- reassigned runtime values (accessor-gated) ---------- */

// Active block size in minutes, from settings. Reconciled with the persisted
// value, falling back to 15 when the stored interval isn't a known option.
let _slotMin = INTERVALS.includes(state.settings.intervalMin) ? state.settings.intervalMin : 15;
export function getSlotMin() { return _slotMin; }
export function setSlotMin(v) { _slotMin = v; }

// Reminder (ping) cadence — decoupled from the block size. While coupled the
// reminder simply follows the block size, so getReminderMin() returns the
// active slotMin and the stored reminderMin is irrelevant; once decoupled the
// stored value is used as-is. setReminderMin/setReminderCoupled persist into
// state.settings (the multi-tab merge picks them up like any other setting).
export function getReminderCoupled() { return state.settings.reminderCoupled !== false; }
export function setReminderCoupled(b) { state.settings.reminderCoupled = !!b; }
export function getReminderMin() {
  if (getReminderCoupled()) return getSlotMin();
  const r = state.settings.reminderMin;
  return INTERVALS.includes(r) ? r : getSlotMin();
}
export function setReminderMin(v) { state.settings.reminderMin = v; }

// Left-most visible day in the calendar.
let _anchor = startOfDay(new Date());
export function getAnchor() { return _anchor; }
export function setAnchor(d) { _anchor = d; }

/** Columns for the current viewport: 3 on desktop, 1 on phones. */
export function colsForViewport() {
  return window.matchMedia('(max-width:760px)').matches ? 1 : 3;
}
let _dayCols = colsForViewport();
export function getDayCols() { return _dayCols; }
export function setDayCols(v) { _dayCols = v; }

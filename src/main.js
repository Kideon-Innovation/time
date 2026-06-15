// App entry module (ES module, loaded via <script type="module"> in
// index.html). This is the former inline <script> from index.html, moved
// under src/ so Vite bundles + hashes it and so it can import the extracted
// pure helpers (src/time.js). Behaviour is unchanged: a type="module" script
// is deferred, so it still runs after the DOM is parsed (same as the old
// end-of-body classic script). XLSX is a UMD global loaded earlier in <head>.
//
// Modules are always strict mode, so the explicit "use strict" is dropped.
import './style.css';
import {
  startOfDay, addDays, iso, hhmm,
  blockDurMin, fmtDur,
  floorSlot as floorSlotMin, nextBoundary as nextBoundaryMin,
} from './time.js';
import {
  state, save, onSaveError, initCrossTabSync,
  INTERVALS, PX_PER_MIN, DOW,
  getSlotMin, setSlotMin, getAnchor, setAnchor,
  getReminderMin, setReminderMin, getReminderCoupled, setReminderCoupled,
  colsForViewport, getDayCols, setDayCols,
} from './state.js';
import {
  uid, bumpRecent, clearRange,
  beat, gapSlots, morningMode,
} from './blocks.js';
import { $ } from './ui/dom.js';
import { toast, notify, beep } from './ui/notify.js';
import {
  initCalendar, render, renderHeaderClock,
  updateNowLine, updateCountdown, scrollToNow,
} from './ui/calendar.js';
import {
  openScrim, closeScrim, close, openScrims,
  openEdit, openLogNow, openMorningPing, openPing,
} from './ui/dialogs.js';
import { attachDrag } from './ui/drag.js';
import { rowsToBlocks } from './importRows.js';

// Thin wrappers so the rest of the app keeps calling floorSlot(d) /
// nextBoundary(d) unchanged; they supply the active block size to the pure,
// parameterised implementations in time.js.
const floorSlot = (d) => floorSlotMin(d, getSlotMin());
// Cadence (WHEN to ping) follows the reminder interval, which is decoupled from
// the block size (getReminderMin() returns the block size while coupled). The
// gap computation in blocks.js stays on the block size (getSlotMin()).
const nextReminder = (d) => nextBoundaryMin(d, getReminderMin());

/* ============================================================
   KIDEON time — passive time tracker.  MIT License.
   State + config live in state.js; pure time helpers in time.js.
   ============================================================ */

/* ============================================================
   UI LAYER (all extracted under src/ui/)
   The calendar render cluster (render / renderHeaderClock / renderRange /
   renderLegend / renderCalendar + scrollToNow / updateNowLine / updateCountdown)
   lives in src/ui/calendar.js; the ping/catch-up/range-entry/edit dialogs plus
   the scrim/modal a11y machinery in src/ui/dialogs.js; the toast/notify/beep
   feedback primitives in src/ui/notify.js; and drag-to-select in src/ui/drag.js.
   They are all imported above. renderCalendar needs two behaviours from the
   higher UI layer — openEdit (dialogs.js) and attachDrag (drag.js) — which are
   handed to it once via initCalendar (see the import block + boot()); this DI
   keeps calendar.js free of an import cycle (calendar -> drag -> dialogs ->
   calendar). dialogs.js pulls `toast` from notify.js directly. What remains in
   main.js is boot/wiring, the export/import flows, and the timer/heartbeat loop.
   ============================================================ */

/* ============================================================
   EXPORT (xlsx via SheetJS)
   ============================================================ */
function exportRows(){
  let from=$("expFrom").value?new Date($("expFrom").value+"T00:00:00"):null;
  let to=$("expTo").value?new Date($("expTo").value+"T23:59:59"):null;
  return [...state.blocks].sort((a,b)=>new Date(a.start)-new Date(b.start))
    .filter(b=>{ const s=new Date(b.start); return (!from||s>=from)&&(!to||s<=to); });
}
// Shared wording for an inverted export range, used by both the inline #expCount
// hint and the Export/DATEV toasts so the user sees ONE consistent message.
const INVERTED_RANGE_MSG="„Von“ liegt nach „Bis“ — bitte Zeitraum prüfen.";
function rangeInverted(){
  const fromV=$("expFrom").value, toV=$("expTo").value;
  return !!(fromV && toV && fromV>toV);
}
function updateExpCount(){
  const el=$("expCount");
  if(rangeInverted()){
    // Inverted range silently yields 0 — point the user at the swapped dates.
    el.textContent=INVERTED_RANGE_MSG;
    el.classList.add("warn");
    return;
  }
  el.classList.remove("warn");
  el.textContent=exportRows().length+" Blöcke im Zeitraum";
}
function doExport(){
  if(typeof XLSX==="undefined"){ toast("Excel-Funktion gerade nicht verfügbar"); return; }
  // An inverted range yields 0 rows for a non-obvious reason; surface the same
  // hint the inline counter shows instead of the generic "Nichts zu exportieren".
  if(rangeInverted()){ updateExpCount(); toast(INVERTED_RANGE_MSG); return; }
  const rows=exportRows();
  if(!rows.length){ toast("Nichts zu exportieren"); return; }
  const data=[["Datum","Wochentag","Start","Ende","Dauer (min)","Tätigkeit"]];
  rows.forEach(b=>{ const s=new Date(b.start),e=new Date(b.end);
    data.push([s.toLocaleDateString("de-DE"),DOW[s.getDay()],hhmm(s),hhmm(e),blockDurMin(b),b.label]); });
  const ws=XLSX.utils.aoa_to_sheet(data);
  ws["!cols"]=[{wch:12},{wch:10},{wch:8},{wch:8},{wch:12},{wch:34}];
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"KIDEON time");
  const tag=(rows[0].start.slice(0,10))+"_bis_"+(rows[rows.length-1].start.slice(0,10));
  XLSX.writeFile(wb,"kideon_time_"+tag+".xlsx");
  markExported();
  close("exportScrim"); toast("Exportiert: "+rows.length+" Blöcke");
}

/* ============================================================
   DATEV-LOHN-EXPORT (Datei) — Bewegungsdaten als Semikolon-CSV
   für DATEV Lohn und Gehalt. Stunden pro Kalendertag verdichtet.
   Spalten: Personalnummer;Datum;Lohnart;Stunden (CRLF, ASCII-rein).
   Personalnummer + Lohnart sind Pflicht und liegen separat im
   localStorage (DATEV_KEY), getrennt vom Zeit-Datenmodell.
   ============================================================ */
const DATEV_KEY="timelog.datev.v1";
function loadDatevCfg(){ try{ return JSON.parse(localStorage.getItem(DATEV_KEY))||{}; }catch(e){ return {}; } }
function saveDatevCfg(cfg){ try{ localStorage.setItem(DATEV_KEY, JSON.stringify(cfg)); }catch(e){} }
function fmtDateDE(d){ const z=n=>String(n).padStart(2,"0"); return z(d.getDate())+"."+z(d.getMonth()+1)+"."+d.getFullYear(); }
function fmtHours(min){ return (min/60).toFixed(2).replace(".",","); }
// exportRows() ist bereits nach start sortiert → Tagesreihenfolge bleibt erhalten.
function datevLohnRows(){
  const byDay=new Map();
  // Bucket each block under its START day. A midnight-crossing block (e.g. the
  // 23:45→00:00 last slot) counts entirely under the day it began on — by design,
  // consistent with exportRows()'s start-day sorting. Total minutes are unaffected.
  exportRows().forEach(b=>{ const key=b.start.slice(0,10);
    byDay.set(key,(byDay.get(key)||0)+blockDurMin(b)); });
  return [...byDay.entries()].map(([key,min])=>
    ({datum:fmtDateDE(new Date(key+"T00:00:00")), min}));
}
function doExportDatev(){
  const pnr=$("datevPnr").value.trim(), la=$("datevLa").value.trim();
  saveDatevCfg({pnr,la});
  if(!pnr||!la){ $("datevDetails").open=true; toast("Personalnummer und Lohnart angeben (vom Lohnbüro)"); return; }
  // Digits only (QA N3): a stray ";" (or any other char) in these fields would
  // shift the CSV columns and DATEV Lohn und Gehalt would misread the file.
  if(!/^\d+$/.test(pnr)||!/^\d+$/.test(la)){ $("datevDetails").open=true; toast("Personalnummer und Lohnart: nur Ziffern (vom Lohnbüro)"); return; }  // Same inverted-range hint as the Excel export, kept consistent with #expCount.
  if(rangeInverted()){ updateExpCount(); toast(INVERTED_RANGE_MSG); return; }
  const days=datevLohnRows();
  if(!days.length){ toast("Nichts zu exportieren"); return; }
  const lines=["Personalnummer;Datum;Lohnart;Stunden"];
  days.forEach(d=>lines.push(pnr+";"+d.datum+";"+la+";"+fmtHours(d.min)));
  const csv=lines.join("\r\n")+"\r\n";
  const blob=new Blob([csv],{type:"text/csv;charset=utf-8"});
  const url=URL.createObjectURL(blob), a=document.createElement("a");
  const rows=exportRows();
  const tag=rows[0].start.slice(0,10)+"_bis_"+rows[rows.length-1].start.slice(0,10);
  a.href=url; a.download="datev_lohn_"+tag+".csv";
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),0);
  markExported();
  close("exportScrim"); toast("DATEV-Lohn exportiert: "+days.length+" Tage");
}

/* ============================================================
   IMPORT (xlsx via SheetJS) — round-trips the export format.
   Reads Datum / Start / Ende / Tätigkeit; Wochentag + Dauer ignored.
   Conflict policy: slot-overwrite (imported block wins on equal start).
   The pure row→block parsing/validation (German text cells AND native Excel
   date/time cells, midnight-crossing rules, slot snapping) lives in
   src/importRows.js so it is unit-testable; this is just the I/O shell.
   ============================================================ */
function importXlsx(file){
  if(typeof XLSX==="undefined"){ toast("Excel-Funktion gerade nicht verfügbar"); return; }
  const reader=new FileReader();
  reader.onerror=()=>toast("Datei konnte nicht gelesen werden");
  reader.onload=ev=>{
    let rows;
    try{
      const wb=XLSX.read(ev.target.result,{type:"array",cellDates:true});
      const ws=wb.Sheets[wb.SheetNames[0]];
      // raw:true keeps native date/time cells as Date objects (parseDE/parseHM
      // handle those); text cells (our own export) stay as strings.
      rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true});
    }catch(err){ toast("Keine gültige Excel-Datei"); return; }
    applyImport(rows);
  };
  reader.readAsArrayBuffer(file);
}
function applyImport(rows){
  const res=rowsToBlocks(rows,getSlotMin());
  if(res.error==="empty"){ toast("Datei ist leer"); return; }
  if(res.error==="columns"){ toast("Spalten Datum/Start/Ende/Tätigkeit fehlen"); return; }
  const {ok,bad,blocks}=res;
  if(!ok){ toast("Keine gültigen Zeilen gefunden"); return; }
  for(const b of blocks){
    // Replace whatever the imported block covers across its FULL span, the same
    // way in-app drag does (fillRange→clearRange). A plain start-only overwrite
    // left finer pre-existing blocks behind when a coarse row (e.g. 09:00–11:00)
    // landed over them, producing overlapping/stacked blocks and a double-counted
    // day total. clearRange clips straddling blocks at the edges.
    clearRange(new Date(b.start), new Date(b.end));
    state.blocks.push({id:uid(),start:b.start,end:b.end,label:b.label});
    bumpRecent(b.label);
  }
  save(); render();
  toast("Importiert: "+ok+" Blöcke"+(bad?(" ("+bad+" übersprungen)"):""));
}

/* ============================================================
   TIMER
   ============================================================ */
// A direct-manipulation gesture (move/resize/range-select) marks the body while
// the pointer is held. A render() during that window rebuilds the calendar DOM
// out from under the pointer, dropping the captured block and silently aborting
// the drag — so any timer-driven render/ping must wait for the pointer to lift.
function gestureInFlight(){ return document.body.classList.contains("dragging"); }
// Single self-perpetuating timer aimed at the next reminder boundary. Tracked
// in _tickT so a cadence change (block/reminder setting, or a cross-tab sync)
// can re-arm it to the new boundary without leaking a parallel timer chain.
let _tickT=null;
function scheduleTick(){
  if(_tickT) clearTimeout(_tickT);
  const ms=nextReminder(new Date()).getTime()-Date.now()+200;
  _tickT=setTimeout(()=>{ _tickT=null; onBoundary(); scheduleTick(); }, Math.max(ms,500));
}
/* Single entry point for all three ping triggers (boundary tick, tab return,
   initial load). Morgen-Modus (see morningMode in blocks.js): after a night
   without logging there is nothing worth back-filling — gapSlots() already
   reports no gaps, and instead of the retro catch-up we ask only about the
   CURRENT slot in present tense. `loud` adds beep+notification (boundary
   tick only, matching the previous behaviour of the three call sites). */
function pingOpenSlots(loud){
  if($("intro").classList.contains("show")) return;
  // Never (re)build the ping while ANY dialog is up. If the open dialog IS the
  // ping itself, rebuilding would wipe text the user is typing (M3); if it's
  // another dialog (edit/export/…), the ping would stack on top of it and
  // break focus + Esc (M4). The gaps aren't lost: a deferred boundary tick
  // (see onBoundary) re-pings right after the dialog closes, and the next
  // tab return / boundary covers the rest.
  if(openScrims().length) return;
  if(morningMode()){
    if(loud){ if(state.settings.soundOn) beep(); notify(1); }
    openMorningPing();
    return;
  }
  const gaps=gapSlots();
  if(!gaps.length) return;
  if(loud){ if(state.settings.soundOn) beep(); notify(gaps.length); }
  openPing(true);
}
// Guard so at most ONE deferred boundary retry is ever in flight. Without it
// every boundary that elapses while a dialog is open (i.e. the whole time the
// user is away from the tab) spawned its OWN 250ms retry loop; after a long
// absence dozens of loops were polling at once, and on dialog-close they
// cascaded — each dismissal instantly re-opened another ping ("modal spam").
let _boundaryRetryPending=false;
function onBoundary(){
  // Defer the whole boundary tick (render + catch-up ping) until any in-progress
  // drag releases (a render would tear the gesture down mid-flight) AND until
  // any open dialog closes — so the tick neither rebuilds an open catch-up
  // (M3, typed text) nor stacks a ping over another modal (M4), and the ping
  // fires promptly once the dialog is dismissed.
  if(gestureInFlight() || openScrims().length){
    if(_boundaryRetryPending) return;   // a single retry is already queued
    _boundaryRetryPending=true;
    const retry=()=>{
      if(gestureInFlight() || openScrims().length){ setTimeout(retry,250); return; }
      _boundaryRetryPending=false;
      onBoundary();
    };
    setTimeout(retry,250);
    return;
  }
  pingOpenSlots(true);
  refreshExportReminder();   // re-check within the existing cadence tick (covers day rollover)
  render();
}
// tickClock glues the per-second header refresh; its three constituents
// (renderHeaderClock / updateCountdown / updateNowLine) live in ui/calendar.js.
function tickClock(){ renderHeaderClock(); updateCountdown(); updateNowLine(); }

/* ============================================================
   EXCEL-EXPORT REMINDER
   Nothing is backed up to a server — if the user forgets to export and the
   browser data gets wiped, everything is gone. So we nudge them to export.

   Anchor for "days since last export":
     - last successful export timestamp (timelog.lastExport.v1), if any;
     - else the start of the OLDEST logged block (their oldest unsaved data),
       so a brand-new user with no data is never nagged, and we only nag once
       there is actually data older than the threshold that has never been saved.
   Overdue = more than EXPORT_REMINDER_DAYS days since that anchor.
   ============================================================ */
const LAST_EXPORT_KEY = "timelog.lastExport.v1";
const EXPORT_REMINDER_DAYS = 3;

function markExported(){ try{ localStorage.setItem(LAST_EXPORT_KEY, new Date().toISOString()); }catch(e){} refreshExportReminder(); }
function lastExportTime(){
  try{ const v=localStorage.getItem(LAST_EXPORT_KEY); if(v){ const t=new Date(v).getTime(); if(!isNaN(t)) return t; } }catch(e){}
  return null;
}
function oldestBlockTime(){
  let min=null;
  for(const b of state.blocks){ const t=new Date(b.start).getTime();
    if(!isNaN(t) && (min===null || t<min)) min=t; }
  return min;
}
/* central calc: whole days since the export anchor, or null when there is
   simply nothing to protect (no export ever AND no data). */
function daysSinceExport(){
  const anchorT = lastExportTime() ?? oldestBlockTime();
  if(anchorT===null) return null;
  return Math.floor((Date.now()-anchorT)/86400000);
}
function exportOverdue(){
  const d=daysSinceExport();
  return d!==null && d>EXPORT_REMINDER_DAYS;
}
function todayKey(){ return iso(startOfDay(new Date())).slice(0,10); }

/* In-app nudge: at most once per day, dismissible, re-appears next day if still
   overdue. Notification fires once per day too (same day-gate), best-effort. */
function exportNudgeWants(){
  return exportOverdue()
      && state.settings.exportReminderDay !== todayKey()
      && !$("intro").classList.contains("show");
}
function refreshExportReminder(){
  const show = exportNudgeWants();
  $("exportNudge").hidden = !show;
  if(show){
    const d=daysSinceExport();
    $("exportNudgeTxt").innerHTML =
      "Seit <b>"+d+" Tagen</b> kein Excel-Export — lad deine Daten als Excel runter, "
      + "damit nichts verloren geht. Sie liegen nur auf diesem Gerät.";
    notifyExportOverdue(d);
  }
  // Only ONE banner at a time: the data-loss nudge wins, so re-sync the
  // notifications nudge (it hides itself while the export nudge is up).
  refreshNotifyNudge();
}
function notifyExportOverdue(days){
  if(!state.settings.notifyOn || !("Notification" in window) || Notification.permission!=="granted") return;
  // gate to once per day via the same dismiss-day mechanism is too aggressive
  // (dismiss would also kill the banner); use a dedicated per-day notify flag.
  if(state.settings.exportNotifyDay===todayKey()) return;
  try{
    new Notification("KIDEON time — Daten sichern",{
      body:"Seit "+days+" Tagen kein Excel-Export. Lad deine Daten runter, damit nichts verloren geht.",
      tag:"timelog-export", renotify:false });
    state.settings.exportNotifyDay=todayKey(); save();
  }catch(e){}
}
$("exportNudgeDismiss").onclick=()=>{ state.settings.exportReminderDay=todayKey(); save(); refreshExportReminder(); };
$("exportNudgeGo").onclick=()=>{ updateExpCount(); openScrim("exportScrim"); };

/* ============================================================
   WIRING
   ============================================================ */
$("prevBtn").onclick=()=>{ setAnchor(addDays(getAnchor(),-1)); render(); };
$("nextBtn").onclick=()=>{ setAnchor(addDays(getAnchor(),1)); render(); };
$("todayBtn").onclick=()=>{ setAnchor(startOfDay(new Date())); render(); scrollToNow(); };
$("datePick").onchange=e=>{ if(e.target.value){ setAnchor(startOfDay(new Date(e.target.value+"T00:00:00"))); render(); } };
$("logNowBtn").onclick=()=>openLogNow();
$("exportBtn").onclick=()=>{ updateExpCount(); openScrim("exportScrim"); };
$("expCancel").onclick=()=>close("exportScrim");
$("expAll").onclick=()=>{ $("expFrom").value=""; $("expTo").value=""; updateExpCount(); };
$("expGo").onclick=doExport;
$("expDatev").onclick=doExportDatev;
(function initDatevInputs(){ const c=loadDatevCfg(); if(c.pnr) $("datevPnr").value=c.pnr; if(c.la) $("datevLa").value=c.la; })();
$("expFrom").onchange=updateExpCount; $("expTo").onchange=updateExpCount;

/* overflow menu */
const menuEl=$("menu"), menuBtn=$("menuBtn");
function closeMenu(){ menuEl.hidden=true; menuBtn.setAttribute("aria-expanded","false"); menuEl.style.right=""; }
function openMenu(){
  menuEl.hidden=false; menuBtn.setAttribute("aria-expanded","true");
  // Keep the dropdown fully on-screen. On phones the menu button sits near the
  // left edge, so the default right:0 panel spills off the left; nudge it back
  // inside the viewport (handles either edge).
  menuEl.style.right="";
  const r=menuEl.getBoundingClientRect(), gap=8;
  if(r.left<gap) menuEl.style.right=(r.left-gap)+"px";
  else if(r.right>window.innerWidth-gap) menuEl.style.right=(r.right-(window.innerWidth-gap))+"px";
}
menuBtn.onclick=e=>{ e.stopPropagation(); menuEl.hidden?openMenu():closeMenu(); };
menuEl.addEventListener("click",e=>{ if(e.target.closest(".menu-item")) closeMenu(); });
document.addEventListener("mousedown",e=>{
  if(!menuEl.hidden && !menuBtn.parentElement.contains(e.target)) closeMenu(); });

/* import */
$("importBtn").onclick=()=>$("importFile").click();
$("importFile").onchange=e=>{ const f=e.target.files[0]; if(f) importXlsx(f); e.target.value=""; };

/* first-run intro */
function showIntro(){ $("intro").classList.add("show"); $("intro").setAttribute("aria-hidden","false"); $("intro").scrollTop=0;
  // Mirror the modal pattern: pull the app out of the a11y tree + tab order so
  // focus can't leak into the controls hidden behind the intro overlay.
  const app=$("app"); app.setAttribute("aria-hidden","true"); try{ app.inert=true; }catch(e){} }
function hideIntro(){ $("intro").classList.remove("show"); $("intro").setAttribute("aria-hidden","true");
  const app=$("app"); app.removeAttribute("aria-hidden"); try{ app.inert=false; }catch(e){} }

function dismissIntro(){ state.settings.introSeen=true; save(); hideIntro(); refreshNotifyNudge(); refreshExportReminder(); }
$("introStart").onclick=dismissIntro;
$("introClose").onclick=dismissIntro;
// Click the backdrop (outside the card) to dismiss — the card itself doesn't bubble.
$("intro").addEventListener("mousedown",e=>{ if(e.target===$("intro")) dismissIntro(); });
$("aboutBtn").onclick=()=>showIntro();

/* ============================================================
   PWA: install flow, service worker, responsive re-render
   ============================================================ */
const isIOS = /iphone|ipad|ipod/i.test(navigator.userAgent) ||
              (navigator.platform==="MacIntel" && navigator.maxTouchPoints>1);
// Live check — display-mode can flip while the page is open (e.g. user installs
// then opens the standalone window), so don't cache it.
function isStandalone(){
  return window.matchMedia("(display-mode: standalone)").matches ||
         window.navigator.standalone===true;
}
let deferredPrompt=null;
let relatedAppInstalled=false;   // set async via getInstalledRelatedApps()

/* Single source of truth for the install controls' visibility.
   Install is reachable from two places — the hamburger menu and the intro /
   "Was ist das?" card — and both are hidden once the app is installed (detected
   via standalone display-mode, navigator.standalone, OR a matching installed
   related app). As long as the app is NOT installed they're shown: where we
   captured a beforeinstallprompt (Chrome/Edge) the click fires the native
   prompt, otherwise (iOS, Firefox, Safari, …) it opens step-by-step install
   help — so the instructions are reachable everywhere. */
function installed(){ return isStandalone() || relatedAppInstalled; }
function refreshInstallBtn(){
  const inst=installed();
  $("installBtn").hidden=inst;
  $("introInstallBtn").hidden=inst;
}

window.addEventListener("beforeinstallprompt",e=>{ e.preventDefault(); deferredPrompt=e; refreshInstallBtn(); });
window.addEventListener("appinstalled",()=>{ deferredPrompt=null; relatedAppInstalled=true;
  refreshInstallBtn(); toast("Installiert — viel Spaß 🎉"); });
// If the user installs and the window switches to standalone while open, re-evaluate.
window.matchMedia("(display-mode: standalone)").addEventListener?.("change", refreshInstallBtn);

async function triggerInstall(){
  if(deferredPrompt){
    deferredPrompt.prompt();
    const {outcome}=await deferredPrompt.userChoice;
    deferredPrompt=null;
    if(outcome==="accepted") relatedAppInstalled=true;
    refreshInstallBtn();
    return;
  }
  openInstallHelp();                       // iOS / browsers without the prompt API
}
$("installBtn").onclick=triggerInstall;
$("introInstallBtn").onclick=triggerInstall;
$("installClose").onclick=()=>close("installScrim");

/* Some browsers (Android Chrome with a `related_applications`/scope match, and
   notably catches the "already installed" case where beforeinstallprompt never
   fires) expose installed PWAs here. Best-effort, async, then re-evaluate. */
if(navigator.getInstalledRelatedApps){
  navigator.getInstalledRelatedApps()
    .then(apps=>{ if(apps && apps.length){ relatedAppInstalled=true; refreshInstallBtn(); } })
    .catch(()=>{});
}

function openInstallHelp(){
  const ua=navigator.userAgent, isFirefox=/firefox/i.test(ua);
  // Safari on macOS: no address-bar install icon — installs via Share → "Zum Dock
  // hinzufügen" (Safari 17+/Sonoma). Chrome/Edge on Mac still get the default
  // branch (they DO show the address-bar icon). isIOS already claims iPads that
  // masquerade as Mac, so this matches only genuine desktop Safari.
  const isMacSafari = /Macintosh/.test(ua) && /Safari/.test(ua) &&
                      !/Chrome|Chromium|CriOS|Edg|OPR/.test(ua) && !isIOS;
  // The lead is set per-platform too: "Home-Bildschirm" is an iOS/Android term —
  // on the Mac the app lands in the Dock, on other desktops it's just a window —
  // so a single shared lead would be wrong on at least one platform.
  let steps, title, lead;
  if(isIOS){
    title="Auf iPhone / iPad installieren";
    lead="So legst du KIDEON time wie eine native App auf deinen Home-Bildschirm — läuft danach offline, im eigenen Fenster, ohne Browser-Leiste.";
    steps=['Tippe unten in Safari auf <b>Teilen</b> <span class="shareglyph">⬆</span>',
           'Wähle <b>„Zum Home-Bildschirm"</b>',
           'Oben rechts auf <b>Hinzufügen</b> — fertig.'];
  } else if(isMacSafari){
    title="Auf dem Mac installieren";
    lead="So legst du KIDEON time wie eine native App ins Dock — läuft danach offline, im eigenen Fenster, ohne Browser-Leiste.";
    steps=['Klick in Safari oben rechts auf <b>Teilen</b> <span class="shareglyph">⬆</span> (oder Menü <b>Ablage</b>)',
           'Wähle <b>„Zum Dock hinzufügen"</b>',
           'Bestätigen — läuft danach im eigenen Fenster. (Safari 17 / macOS Sonoma)'];
  } else if(isFirefox){
    title="Installieren";
    lead="So installierst du KIDEON time als eigenständige App — läuft danach offline, im eigenen Fenster, ohne Browser-Leiste.";
    steps=['Öffne das Browser-<b>Menü</b> (⋮ oben rechts)',
           'Wähle <b>„App installieren"</b> / „Zum Startbildschirm hinzufügen"',
           'Bestätigen — läuft danach im eigenen Fenster.'];
  } else {
    title="Installieren";
    lead="So installierst du KIDEON time als eigenständige App — läuft danach offline, im eigenen Fenster, ohne Browser-Leiste.";
    steps=['Klick in der Adressleiste auf das <b>Installieren-Symbol</b> (↗ bzw. Monitor mit Pfeil)',
           'Alternativ Browser-<b>Menü</b> (⋮) → <b>„App installieren"</b>',
           'Bestätigen — KIDEON time läuft danach offline im eigenen Fenster.'];
  }
  $("installTitle").textContent=title;
  $("installLead").textContent=lead;
  $("installBody").innerHTML='<ul class="ioslist">'+steps.map(s=>"<li>"+s+"</li>").join("")+'</ul>';
  openScrim("installScrim");
}

// re-render when crossing the mobile/desktop breakpoint
let resizeT=null;
window.addEventListener("resize",()=>{ clearTimeout(resizeT); resizeT=setTimeout(()=>{
  const c=colsForViewport(); if(c!==getDayCols()){ setDayCols(c); render(); }
},150); });

/* Service worker: offline app shell + auto-update.
   The Workbox SW is generated and registered by vite-plugin-pwa
   (registerType:'autoUpdate', injectRegister:'auto'). On a new deploy the
   fresh SW installs, skipWaiting + clientsClaim take control immediately and
   the page reloads to the new version — no manual hard-refresh. The old
   hand-written sw.js + manual registration were removed in favor of this. */

function applyTheme(){
  const dark=state.settings.theme==="dark";
  document.documentElement.setAttribute("data-theme", dark?"dark":"light");
  // Label by the action the button performs, so SR users hear what activating it does.
  const btn=$("themeBtn");
  btn.textContent=dark?"☀️ Helles Design":"🌙 Dunkles Design";
  btn.setAttribute("aria-label", dark?"Zu hellem Design wechseln":"Zu dunklem Design wechseln");
  btn.removeAttribute("aria-pressed");
  // tint the OS status bar / window chrome to match the app
  const meta=$("metaTheme"); if(meta) meta.setAttribute("content", dark?"#0f2137":"#faf9f7");
}
$("themeBtn").onclick=()=>{ state.settings.theme=(state.settings.theme==="light")?"dark":"light"; save(); applyTheme(); };

// "Blockgröße = dein Abrechnungstakt" education: a transient toast shown ONCE
// per settings change (i.e. the first time the user changes the block size in
// this session), not a blocking modal and not persisted across sessions.
let _blockEduShown=false;
function initInterval(){
  const sel=$("intervalSel");
  sel.innerHTML=INTERVALS.map(m=>`<option value="${m}">${m} min</option>`).join("");
  sel.value=String(getSlotMin());
  sel.onchange=()=>{
    setSlotMin(parseInt(sel.value,10));
    state.settings.intervalMin=getSlotMin();
    // While coupled, the reminder cadence follows the block size — drag it
    // along and reflect the inherited value in the (disabled) reminder select.
    if(getReminderCoupled()){ setReminderMin(getSlotMin()); $("reminderSel").value=String(getSlotMin()); }
    save();
    render(); updateCountdown();
    scheduleTick();   // cadence may have changed (coupled) → re-arm the timer
    if(!_blockEduShown){
      _blockEduShown=true;
      toast("Blockgröße = dein Abrechnungstakt. Erinnerungen stellst du separat ein.");
    } else {
      toast("Blockgröße: "+getSlotMin()+" min");
    }
  };
}

// Reminder (Erinnerung) controls: a "An Blockgröße koppeln" checkbox and an
// interval select that is disabled (and shows the inherited block size) while
// coupled. Mirrors initInterval(): init from state, wire handlers, persist.
function initReminder(){
  const chk=$("reminderCoupleChk"), sel=$("reminderSel");
  sel.innerHTML=INTERVALS.map(m=>`<option value="${m}">${m} min</option>`).join("");
  syncReminderControls();
  chk.onchange=()=>{
    setReminderCoupled(chk.checked);
    if(chk.checked) setReminderMin(getSlotMin());   // re-couple → snap to block size
    save();
    syncReminderControls();
    scheduleTick();   // cadence changed → re-arm the timer
  };
  sel.onchange=()=>{
    setReminderMin(parseInt(sel.value,10));
    save();
    scheduleTick();
    toast("Erinnerung: "+getReminderMin()+" min");
  };
}
// Reflect coupled/decoupled state into the two controls (disabled + shown value).
function syncReminderControls(){
  const chk=$("reminderCoupleChk"), sel=$("reminderSel");
  const coupled=getReminderCoupled();
  chk.checked=coupled;
  sel.disabled=coupled;
  sel.value=String(getReminderMin());   // shows inherited block size while coupled
}

function refreshSoundBtn(){
  $("soundBtn").textContent=state.settings.soundOn?"🔇 Ton ausschalten":"🔊 Ton anschalten";
  $("soundBtn").classList.toggle("ghost",!state.settings.soundOn);
  $("soundBtn").setAttribute("aria-pressed", state.settings.soundOn?"true":"false");
}
$("soundBtn").onclick=()=>{ state.settings.soundOn=!state.settings.soundOn; save();
  refreshSoundBtn();
  if(state.settings.soundOn) beep(); refreshNotifyNudge(); toast("Ton "+(state.settings.soundOn?"eingeschaltet":"ausgeschaltet")); };

$("notifyBtn").onclick=async()=>{
  if(!("Notification" in window)){ toast("Dieser Browser kann keine Erinnerungen anzeigen"); return; }
  // Real on/off toggle: if currently on, turn off; otherwise ask + turn on.
  if(notifyGranted()){
    state.settings.notifyOn=false; save(); refreshNotifyBtn();
    toast("Erinnerungen ausgeschaltet"); return;
  }
  let p=Notification.permission;
  if(p!=="granted") p=await Notification.requestPermission();
  state.settings.notifyOn=(p==="granted"); save(); refreshNotifyBtn();
  toast(p==="denied" ? "Im Browser blockiert — bitte dort erlauben"
                     : (state.settings.notifyOn?"Erinnerungen eingeschaltet":"Erinnerungen ausgeschaltet"));
};
function refreshNotifyBtn(){
  const on=state.settings.notifyOn && ("Notification" in window) && Notification.permission==="granted";
  $("notifyBtn").textContent=on?"🔕 Erinnerungen ausschalten":"🔔 Erinnerungen anschalten";
  $("notifyBtn").classList.toggle("ghost",!on);
  $("notifyBtn").setAttribute("aria-pressed", on?"true":"false");
  refreshNotifyNudge();
}

/* ---------- notifications-off nudge ----------
   Reminds the user KIDEON time can only ping them when OS notifications are on.
   Shown once until dismissed; vanishes on its own once notifications are granted. */
function notifyGranted(){
  return state.settings.notifyOn && ("Notification" in window) && Notification.permission==="granted";
}
function refreshNotifyNudge(){
  const supported = "Notification" in window;
  // Only ONE banner at a time — the export/data-loss nudge takes priority, so
  // hold this one back whenever that one is (or would be) showing.
  const show = supported && !notifyGranted() && !state.settings.notifyNudgeDismissed
            && !$("intro").classList.contains("show")
            && !exportNudgeWants();
  $("notifyNudge").hidden = !show;
  if(show){
    // sound is the fallback ping channel — mention it if that's off too
    $("nudgeHint").textContent = state.settings.soundOn
      ? "Jetzt aktivieren — den Ping-Takt stellst du im Menü unter Blockgröße ein."
      : "Auch der Ton ist aus. Jetzt aktivieren — den Takt stellst du im Menü unter Blockgröße ein.";
  }
}
$("nudgeDismiss").onclick=()=>{ state.settings.notifyNudgeDismissed=true; save(); refreshNotifyNudge(); };
$("nudgeEnable").onclick=async()=>{
  if("Notification" in window){
    let p=Notification.permission;
    if(p!=="granted") p=await Notification.requestPermission();
    state.settings.notifyOn=(p==="granted");
  }
  if(!state.settings.soundOn){ state.settings.soundOn=true; refreshSoundBtn(); beep(); }
  save(); refreshNotifyBtn();
  if(notifyGranted()) toast("Erinnerungen eingeschaltet — du wirst jetzt erinnert");
  else { openMenu(); toast("Den Ping-Takt stellst du hier unter Blockgröße ein"); }
};

// close scrim on backdrop click / Esc
// Require the press to BOTH start and end on the backdrop itself
// (pointerdown→click on the scrim). A touch tap that opens a dialog — e.g.
// tapping an empty calendar slot — synthesises a trailing mouse/click on the
// freshly shown backdrop; gating on a real pointerdown that lands on the scrim
// ignores those ghost events so the dialog no longer slams shut the instant it
// opens on touch.
document.querySelectorAll(".scrim").forEach(s=>{
  let downOnSelf=false;
  s.addEventListener("pointerdown",e=>{ downOnSelf = e.target===s; });
  s.addEventListener("click",e=>{ if(downOnSelf && e.target===s) closeScrim(s.id); downOnSelf=false; });
});
document.addEventListener("keydown",e=>{ if(e.key==="Escape"){ openScrims().forEach(s=>closeScrim(s.id)); closeMenu(); if($("intro").classList.contains("show")) dismissIntro(); } });

/* save() failed (quota exceeded, private mode, …) — the data did NOT persist.
   Surface it loudly with a direct path to the Excel export so nothing is lost;
   every further failing save re-raises the toast (persistent enough by design). */
onSaveError(()=>toast("Speichern fehlgeschlagen — bitte Daten als Excel exportieren!",
  { action:"Exportieren", onAction:()=>{ updateExpCount(); openScrim("exportScrim"); } }));

/* another tab/PWA window wrote timelog.v1 — state.js merged it into memory;
   refresh everything that renders from state. Deferred while a drag gesture is
   in flight (same reason as onBoundary: render() would tear the gesture down). */
function onExternalState(){
  if(gestureInFlight()){ setTimeout(onExternalState,250); return; }
  if(INTERVALS.includes(state.settings.intervalMin) && state.settings.intervalMin!==getSlotMin()){
    setSlotMin(state.settings.intervalMin); $("intervalSel").value=String(getSlotMin());
  }
  // Another tab may have (de)coupled the reminder or picked a new cadence —
  // reflect it in the controls and re-arm the ping timer to the new cadence.
  syncReminderControls(); scheduleTick();
  applyTheme(); refreshSoundBtn(); refreshNotifyBtn(); render();
}
initCrossTabSync(onExternalState);

// returning to tab → catch up
document.addEventListener("visibilitychange",()=>{ if(!document.hidden){ recordBeat(); refreshExportReminder(); render(); pingOpenSlots(false); } });
window.addEventListener("focus",recordBeat);

/* heartbeat: stamp the current slot while the tab is visible.
   If a brand-new slot lit up, repaint so the rail stays current. */
function recordBeat(){ if(document.hidden||gestureInFlight()) return; if(beat()) render(); }

/* ---------- boot ---------- */
function boot(){
  // Hand the calendar renderer the two event/modal behaviours it needs: opening
  // the edit dialog (now in dialogs.js) and attaching drag-to-select (now in
  // drag.js, imported above and passed through).
  initCalendar({ openEdit, attachDrag });
  applyTheme();
  if(!state.settings.introSeen) showIntro();
  initInterval();
  initReminder();
  refreshSoundBtn();
  refreshNotifyBtn();
  refreshExportReminder();            // remind to export if overdue (once/day)
  recordBeat();                       // mark this slot alive on load
  render();
  scrollToNow();
  tickClock();
  setInterval(tickClock,1000);
  setInterval(recordBeat,60000);      // heartbeat every ~60s while visible
  scheduleTick();

  // Set the install control's initial state (shows on iOS / once a prompt is
  // captured; hidden + "✓ Installiert" pill when already installed).
  refreshInstallBtn();

  // manifest shortcuts: ?action=log | export
  const action=new URLSearchParams(location.search).get("action");
  if(action==="log"){ setTimeout(openLogNow,350); }
  else if(action==="export"){ setTimeout(()=>{ updateExpCount(); openScrim("exportScrim"); },350); }
  else { setTimeout(()=>pingOpenSlots(false), 900); }  // initial catch-up
}
boot();

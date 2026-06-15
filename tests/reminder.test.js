// Unit assertions for the decoupled reminder interval (Erinnerung) feature:
//   * reminderMin / reminderCoupled load + default-fill (reminderMin ←
//     intervalMin, reminderCoupled ← true) so existing users keep today's
//     behaviour (cadence = block size) with no migration.
//   * coupled: getReminderMin() follows the block size; decoupled: independent.
//   * 3-way merge / multi-tab keeps both new fields consistent.
//   * Cadence (nextBoundary) is driven by reminderMin while gapSlots() keeps
//     using slotMin, and the current (running) slot is never a gap — also when
//     the reminder is SMALLER than the block size (the "no spam" guarantee).
//
// Same stub pattern as blocks.test.js / storage.test.js: install minimal
// browser globals BEFORE the dynamic import so state.js loads in Node unchanged.

import { test, before, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import { iso, nextBoundary } from '../src/time.js';

function makeLocalStorage() {
  const m = new Map();
  return {
    getItem: (k) => (m.has(k) ? m.get(k) : null),
    setItem: (k, v) => { m.set(k, String(v)); },
    removeItem: (k) => { m.delete(k); },
    clear: () => m.clear(),
  };
}
globalThis.localStorage = makeLocalStorage();
globalThis.window = { matchMedia: () => ({ matches: false }) };
if (typeof globalThis.performance === 'undefined') {
  globalThis.performance = { now: () => Date.now() };
}

let S; let B;
before(async () => { S = await import('../src/state.js'); B = await import('../src/blocks.js'); });

const at = (h, m = 0) => new Date(2026, 5, 8, h, m, 0); // 8 Jun 2026, local
const blk = (start, end, label) => ({ start: iso(start), end: iso(end), label });
const st = (blocks, recentLabels = [], settings = {}) => ({ blocks, recentLabels, settings });

beforeEach(() => {
  S.state.blocks = [];
  S.state.recentLabels = [];
  S.setSlotMin(15);
  S.state.settings.intervalMin = 15;
  S.state.settings.reminderMin = 15;
  S.state.settings.reminderCoupled = true;
});

/* ---------------- defaults + accessors ---------------- */

test('defaults: coupled reminder follows the block size', () => {
  S.setSlotMin(30); S.state.settings.intervalMin = 30;
  assert.equal(S.getReminderCoupled(), true);
  assert.equal(S.getReminderMin(), 30); // follows slotMin, ignores stored reminderMin
  S.setSlotMin(6); S.state.settings.intervalMin = 6;
  assert.equal(S.getReminderMin(), 6);
});

test('default-fill on load: missing reminderMin ← intervalMin, reminderCoupled ← true', () => {
  // a legacy payload with no reminder fields at all
  localStorage.setItem(S.KEY, JSON.stringify({
    blocks: [], recentLabels: [], settings: { intervalMin: 20, theme: 'light' }, savedAt: Date.now(),
  }));
  // syncFromStorage merges the foreign payload through parseRaw's default-fill
  S.syncFromStorage();
  assert.equal(S.state.settings.reminderMin, 20);
  assert.equal(S.state.settings.reminderCoupled, true);
});

test('decoupled: reminder is independent of the block size', () => {
  S.setSlotMin(15); S.state.settings.intervalMin = 15;
  S.setReminderCoupled(false);
  S.setReminderMin(60);
  assert.equal(S.getReminderCoupled(), false);
  assert.equal(S.getReminderMin(), 60);          // own value, not slotMin
  S.setSlotMin(6); S.state.settings.intervalMin = 6;
  assert.equal(S.getReminderMin(), 60);          // unaffected by block-size change
});

test('decoupled with an invalid stored reminderMin falls back to slotMin', () => {
  S.setSlotMin(15);
  S.setReminderCoupled(false);
  S.state.settings.reminderMin = 7; // not on the ladder
  assert.equal(S.getReminderMin(), 15);
});

/* ---------------- 3-way merge / multi-tab ---------------- */

test('mergeStates keeps reminderMin/reminderCoupled per key (local change wins)', () => {
  const base = st([], [], { intervalMin: 15, reminderMin: 15, reminderCoupled: true });
  const mine = st([], [], { intervalMin: 15, reminderMin: 60, reminderCoupled: false }); // we decoupled
  const theirs = st([], [], { intervalMin: 15, reminderMin: 15, reminderCoupled: true });
  const m = S.mergeStates(base, mine, theirs);
  assert.equal(m.settings.reminderCoupled, false);
  assert.equal(m.settings.reminderMin, 60);
});

test('mergeStates: a foreign reminder change propagates when we did not touch it', () => {
  const base = st([], [], { reminderMin: 15, reminderCoupled: true });
  const mine = st([], [], { reminderMin: 15, reminderCoupled: true });
  const theirs = st([], [], { reminderMin: 30, reminderCoupled: false }); // other tab decoupled
  const m = S.mergeStates(base, mine, theirs);
  assert.equal(m.settings.reminderCoupled, false);
  assert.equal(m.settings.reminderMin, 30);
});

/* ---------------- cadence uses reminderMin, gaps use slotMin ---------------- */

test('cadence boundary uses reminderMin (decoupled, larger than block)', () => {
  S.setSlotMin(15);
  S.setReminderCoupled(false);
  S.setReminderMin(60);
  // 09:07 → next reminder boundary at 10:00 (60-min cadence), NOT 09:15 (block)
  const b = nextBoundary(at(9, 7), S.getReminderMin());
  assert.equal(b.getTime(), at(10).getTime());
});

test('cadence boundary uses reminderMin (decoupled, smaller than block)', () => {
  S.setSlotMin(60);
  S.setReminderCoupled(false);
  S.setReminderMin(15);
  // 09:07 → next reminder boundary at 09:15 (15-min cadence), NOT 10:00 (block)
  const b = nextBoundary(at(9, 7), S.getReminderMin());
  assert.equal(b.getTime(), at(9, 15).getTime());
});

test('gapSlots still tiles in block-size (slotMin) units, independent of reminderMin', () => {
  S.setSlotMin(60);
  S.setReminderCoupled(false);
  S.setReminderMin(15);
  // one logged block ending 08:00; "now" 10:30 → gaps are the 08:00 and 09:00
  // block-sized slots (the 10:00 slot is the running one → excluded).
  S.state.blocks = [blk(at(7), at(8), 'A')];
  const gaps = B.gapSlots(at(10, 30));
  assert.deepEqual(gaps.map((d) => d.getTime()), [at(8).getTime(), at(9).getTime()]);
});

test('no-spam: running slot is never a gap even when reminder < block size', () => {
  S.setSlotMin(60);
  S.setReminderCoupled(false);
  S.setReminderMin(15);
  S.state.blocks = [blk(at(8), at(9), 'A')];
  // now = 09:15 — inside the 09:00–10:00 block-slot. That running slot must not
  // be reported, no matter how often the (15-min) reminder fires.
  const gaps = B.gapSlots(at(9, 15));
  assert.ok(!gaps.some((d) => d.getTime() === at(9).getTime()),
    'the running 09:00 block-slot must never be a gap');
  const gaps2 = B.gapSlots(at(9, 45)); // still the same running slot
  assert.ok(!gaps2.some((d) => d.getTime() === at(9).getTime()));
});

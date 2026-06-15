import { test, expect } from '@playwright/test';

// E2E for the decoupled reminder interval (Erinnerung) UI in the menu:
//   * couple/uncouple toggles #reminderSel's disabled state
//   * decoupled lets you pick a differing value that persists across reload
//   * a one-time toast appears when the block size changes
//   * sub-block cadence (reminder < block) nags only about EARLIER empty
//     block-slots, never the running one (the "no spam" guarantee)

async function seed(page, settings = {}) {
  await page.goto('./');
  await page.evaluate((settings) => {
    localStorage.setItem('timelog.v1', JSON.stringify({
      blocks: [],
      recentLabels: [],
      settings: Object.assign({
        theme: 'light', introSeen: true, soundOn: true, notifyOn: false,
        intervalMin: 15, reminderMin: 15, reminderCoupled: true,
        notifyNudgeDismissed: false,
      }, settings),
    }));
    localStorage.removeItem('timelog.lastExport.v1');
  }, settings);
  await page.reload();
  await page.waitForTimeout(1100);
}

async function closeModal(page) {
  for (let i = 0; i < 6; i++) {
    if (!(await page.locator('.scrim.show').count())) return;
    await page.keyboard.press('Escape');
    await page.waitForTimeout(150);
  }
}

test('reminder is coupled by default: #reminderSel disabled and shows the block size', async ({ page }) => {
  await seed(page);
  await closeModal(page);
  await page.click('#menuBtn');
  await expect(page.locator('#reminderCoupleChk')).toBeChecked();
  await expect(page.locator('#reminderSel')).toBeDisabled();
  await expect(page.locator('#reminderSel')).toHaveValue('15');
  // permanent education line under the block-size dropdown
  await expect(page.locator('.menu-hint')).toContainText('Abrechnungstakt');
});

test('uncoupling enables #reminderSel; a differing value persists across reload', async ({ page }) => {
  await seed(page);
  await closeModal(page);
  await page.click('#menuBtn');
  await page.uncheck('#reminderCoupleChk');
  await expect(page.locator('#reminderSel')).toBeEnabled();
  await page.selectOption('#reminderSel', '60');   // reminder 60 ≠ block 15
  await expect(page.locator('#reminderSel')).toHaveValue('60');

  await page.reload();
  await page.waitForTimeout(1100);
  await closeModal(page);
  await page.click('#menuBtn');
  await expect(page.locator('#reminderCoupleChk')).not.toBeChecked();
  await expect(page.locator('#reminderSel')).toBeEnabled();
  await expect(page.locator('#reminderSel')).toHaveValue('60');
  await expect(page.locator('#intervalSel')).toHaveValue('15'); // block size untouched
});

test('coupled: changing the block size drags the reminder along (display + value)', async ({ page }) => {
  await seed(page);
  await closeModal(page);
  await page.click('#menuBtn');
  await page.selectOption('#intervalSel', '30');
  await expect(page.locator('#reminderSel')).toHaveValue('30'); // inherited value updated
  const stored = await page.evaluate(() => JSON.parse(localStorage.getItem('timelog.v1')).settings);
  expect(stored.intervalMin).toBe(30);
  expect(stored.reminderMin).toBe(30);     // dragged along
  expect(stored.reminderCoupled).toBe(true);
});

test('one-time toast appears when changing the block size, then reverts to the plain toast', async ({ page }) => {
  await seed(page);
  await closeModal(page);
  await page.click('#menuBtn');
  const toast = page.locator('#toast');
  await page.selectOption('#intervalSel', '30');
  await expect(toast).toContainText('Abrechnungstakt');
  await expect(toast).toContainText('separat');
  // second change in the same session → plain toast, education not repeated
  await page.selectOption('#intervalSel', '6');
  await expect(toast).toContainText('Blockgröße: 6 min');
  await expect(toast).not.toContainText('Abrechnungstakt');
});

test('sub-block reminder nags only about earlier empty slots, never the running one', async ({ page }) => {
  // Block size 60, reminder 15 (sub-block). Seed one block ending two block-slots
  // ago so there is exactly ONE earlier empty 60-min hole; land "now" mid running
  // slot. The catch-up ping must list that earlier hole but never a row covering
  // the currently running slot — the "no spam" guarantee.
  const pad = (n) => String(n).padStart(2, '0');
  await page.goto('./');
  const labels = await page.evaluate(() => {
    const HOUR = 60 * 60000;
    const floorHr = (t) => { const d = new Date(t); d.setMinutes(0, 0, 0); return d; };
    const runStart = floorHr(Date.now()).getTime();   // start of the running 60-min slot
    const mk = (start, label) => ({
      start: new Date(start).toISOString(),
      end: new Date(start + HOUR).toISOString(), label,
    });
    localStorage.setItem('timelog.v1', JSON.stringify({
      blocks: [mk(runStart - 2 * HOUR, 'Frueher')],   // one block two slots back
      recentLabels: ['Frueher'],
      settings: { theme: 'light', introSeen: true, soundOn: false, notifyOn: false,
        intervalMin: 60, reminderMin: 15, reminderCoupled: false,
        notifyNudgeDismissed: false },
    }));
    localStorage.removeItem('timelog.lastExport.v1');
    const runH = new Date(runStart).getHours();
    const gapH = new Date(runStart - HOUR).getHours();
    return { runH, gapH };
  });
  await page.reload();
  await page.waitForTimeout(1300);

  // The ping dialog auto-opens (initial catch-up). With a 2h look-back cap over
  // 60-min slots and a block ending one slot ago, exactly one earlier empty slot
  // remains → the single-slot ping, whose subtitle is the gap's "HH:MM – HH:MM".
  // (Either path proves the invariant; we read whichever rendered.)
  await expect(page.locator('#pingScrim.scrim.show')).toBeVisible();
  const sub = (await page.locator('#pingSub').textContent()) || '';
  const rows = (await page.locator('#pingBody .gaprow .tg').allTextContents()).join(' | ');
  const shown = `${sub} | ${rows}`;
  // The earlier empty 60-min slot is offered to fill…
  expect(shown).toContain(`${pad(labels.gapH)}:00`);
  // …but the dialog never targets the currently running slot — its hour as a
  // RANGE START ("HH:00 –" / "HH:00–") must be absent. No spam about live work.
  expect(shown).not.toMatch(new RegExp(`${pad(labels.runH)}:00\\s*–`));
});

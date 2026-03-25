// @ts-check
const { test: base } = require('@playwright/test');
require('dotenv').config();

const MONTHS = {
  '1': 'January', '2': 'February', '3': 'March', '4': 'April',
  '5': 'May', '6': 'June', '7': 'July', '8': 'August',
  '9': 'September', '10': 'October', '11': 'November', '12': 'December',
};

/**
 * Shared fixture that provides an authenticated Odoo page.
 * Every test that uses `test` from this file starts already logged in.
 */
const test = base.extend({
  page: async ({ page }, use) => {
    const baseURL  = process.env.ODOO_BASE_URL;
    const user     = process.env.ODOO_ADMIN_USER;
    const password = process.env.ODOO_ADMIN_PASSWORD;

    if (!baseURL || !user || !password) {
      throw new Error(
        'Missing env vars. Ensure ODOO_BASE_URL, ODOO_ADMIN_USER and ' +
        'ODOO_ADMIN_PASSWORD are set in your .env file.'
      );
    }

    // ── Login ──────────────────────────────────────────────────────────────
    await page.goto('/web/login');
    await page.locator('#login').fill(user);
    await page.locator('#password').fill(password);
    await page.getByRole('button', { name: 'Log in' }).click();
    // Wait for the Odoo backend home screen
    await page.waitForURL(/\/web/, { timeout: 30000 });
    await page.waitForTimeout(1500);

    await use(page);
  },
});

/** Navigate to OT Sheets list directly via action XML id. */
async function goToOtSheets(page) {
  await page.goto('/web#action=hr_ot_sheet.action_hr_ot_sheet');
  await page.waitForSelector('.o_list_view', { timeout: 15000 });
}

/**
 * Create a new OT Sheet with the given month (number string "1"–"12") and year.
 * Leaves the page on the saved form.
 */
async function createOtSheet(page, month, year) {
  await page.getByRole('button', { name: /new/i }).first().click();
  await page.waitForSelector('.o_field_widget[name="month"] select', { timeout: 10000 });
  // Select by visible label text
  await page.locator('.o_field_widget[name="month"] select').selectOption({ label: MONTHS[month] });
  await page.locator('.o_field_widget[name="year"] input').fill(String(year));
  await page.keyboard.press('Tab');
  // Save to generate the sequence reference
  await page.locator('.o_form_button_save').click();
  await page.waitForTimeout(1000);
}

module.exports = { test, expect: base.expect, goToOtSheets, createOtSheet };

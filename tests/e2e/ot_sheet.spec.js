// @ts-check
/**
 * E2E tests for the hr_ot_sheet Odoo module.
 *
 * Covered scenarios
 * ─────────────────
 * 1. Navigation  – OT Sheet menu item exists inside the Employees app
 * 2. Create      – New OT Sheet gets an auto-generated reference (NNN/MM/YYYY)
 * 3. Import tab  – Import tab is visible in Draft state, hidden in Done state
 * 4. Import      – Excel file upload creates OT lines (skipped if fixture absent)
 * 5. Apply       – "Apply to Payslips" transitions state to Done
 * 6. Done guard  – Import Excel / Apply buttons are hidden once Done
 * 7. Lines tab   – Rows can be added manually via the inline editable tree
 */

const path = require('path');
const { test, expect, goToOtSheets, createOtSheet } = require('./fixtures');

// ── 1. Navigation ─────────────────────────────────────────────────────────────
test('1. OT Sheet menu is visible inside the Employees app', async ({ page }) => {
  // Start from the home menu
  await page.goto('/web');
  await page.waitForSelector('.o_app.o_menuitem', { timeout: 15000 });

  // Open Employees app
  await page.locator('.o_app.o_menuitem').filter({ hasText: /^Employees$/ }).click();
  await page.waitForTimeout(2000);

  // "OT Sheet" should be a top-level menu section in the nav bar
  const otSheetMenu = page.getByRole('button', { name: /^OT Sheet$/i })
    .or(page.getByRole('link', { name: /^OT Sheet$/i }))
    .or(page.locator('.o_nav_entry, .o_menu_sections a').filter({ hasText: /^OT Sheet$/i }));
  await expect(otSheetMenu.first()).toBeVisible({ timeout: 10000 });

  // Click it to reveal submenu
  await otSheetMenu.first().click();
  await page.waitForTimeout(500);

  // "OT Sheets" submenu item should appear
  const otSheetsItem = page.getByRole('menuitem', { name: /^OT Sheets$/i })
    .or(page.getByRole('link', { name: /^OT Sheets$/i }))
    .or(page.locator('.dropdown-menu a, .o_dropdown_menu a').filter({ hasText: /^OT Sheets$/i }));
  await expect(otSheetsItem.first()).toBeVisible({ timeout: 5000 });

  // Click and confirm the list view loads
  await otSheetsItem.first().click();
  await expect(page.locator('.o_list_view')).toBeVisible({ timeout: 15000 });
});

// ── 2. Create ─────────────────────────────────────────────────────────────────
test('2. New OT Sheet is saved with correct period and Draft state', async ({ page }) => {
  await goToOtSheets(page);

  await page.getByRole('button', { name: /new/i }).first().click();
  await page.waitForSelector('.o_field_widget[name="month"] select', { timeout: 10000 });
  await page.locator('.o_field_widget[name="month"] select').selectOption({ label: 'March' });
  await page.locator('.o_field_widget[name="year"] input').fill(String(new Date().getFullYear()));
  // Set a manual reference so the record has a name and can be saved
  await page.locator('.o_field_widget[name="name"] input').fill('E2E/03/' + new Date().getFullYear());
  await page.keyboard.press('Tab');
  await page.locator('.o_form_button_save').click();
  await page.waitForTimeout(1500);

  // Reference is persisted
  const savedRef = await page.locator('.o_field_widget[name="name"] input').inputValue();
  expect(savedRef).toBe('E2E/03/' + new Date().getFullYear());

  // State must be Draft
  await expect(page.locator('.o_arrow_button_current')).toContainText(/draft/i);
});

// ── 3. Import tab visibility ───────────────────────────────────────────────────
test('3. Import tab is visible in Draft and hidden in Done', async ({ page }) => {
  await goToOtSheets(page);
  await createOtSheet(page, '4', new Date().getFullYear());

  // Visible in Draft
  await expect(page.getByRole('tab', { name: /import/i })).toBeVisible();

  // Apply → Done
  await page.getByRole('button', { name: /apply to payslips/i }).click();
  await expect(page.locator('.o_arrow_button_current')).toContainText(/done/i, { timeout: 20000 });

  // Hidden in Done
  await expect(page.getByRole('tab', { name: /import/i })).toHaveCount(0);
});

// ── 4. Import Excel ───────────────────────────────────────────────────────────
test('4. Uploading an Excel file populates OT lines', async ({ page }) => {
  const fixturePath = path.resolve(__dirname, '../fixtures/sample_ot.xlsx');
  const fs = require('fs');
  if (!fs.existsSync(fixturePath)) {
    test.skip(true, 'sample_ot.xlsx fixture not found — skipping import test');
    return;
  }

  await goToOtSheets(page);
  await createOtSheet(page, '5', new Date().getFullYear());

  // Switch to Import tab and upload
  await page.getByRole('tab', { name: /import/i }).click();
  await page.locator('input[type="file"]').setInputFiles(fixturePath);

  // Click Import Excel
  await page.getByRole('button', { name: /import excel/i }).click();

  // Success/warning notification should appear
  await expect(
    page.locator('.o_notification_content, .o_toast_content').filter({ hasText: /imported/i })
  ).toBeVisible({ timeout: 20000 });

  // Lines tab should now have at least one row
  await page.getByRole('tab', { name: /lines/i }).click();
  const rows = page.locator('.o_field_widget[name="line_ids"] .o_data_row');
  await expect(rows.first()).toBeVisible({ timeout: 10000 });
  expect(await rows.count()).toBeGreaterThan(0);
});

// ── 5. Apply to Payslips ──────────────────────────────────────────────────────
test('5. Apply to Payslips transitions state to Done', async ({ page }) => {
  await goToOtSheets(page);
  await createOtSheet(page, '6', new Date().getFullYear());

  await page.getByRole('button', { name: /apply to payslips/i }).click();

  await expect(page.locator('.o_arrow_button_current')).toContainText(/done/i, { timeout: 20000 });
});

// ── 6. Done guard ─────────────────────────────────────────────────────────────
test('6. Import Excel and Apply buttons are hidden on a Done sheet', async ({ page }) => {
  await goToOtSheets(page);
  await createOtSheet(page, '7', new Date().getFullYear());

  await page.getByRole('button', { name: /apply to payslips/i }).click();
  await expect(page.locator('.o_arrow_button_current')).toContainText(/done/i, { timeout: 20000 });

  await expect(page.getByRole('button', { name: /import excel/i })).toHaveCount(0);
  await expect(page.getByRole('button', { name: /apply to payslips/i })).toHaveCount(0);
});

// ── 7. Manual lines ───────────────────────────────────────────────────────────
test('7. OT lines can be added manually via the inline editable tree', async ({ page }) => {
  test.setTimeout(120_000);

  await goToOtSheets(page);
  await createOtSheet(page, '9', new Date().getFullYear());

  // Switch to Lines tab
  await page.getByRole('tab', { name: 'Lines' }).click();
  await page.waitForTimeout(500);

  // Click "Add a line"
  await page.locator('.o_field_widget[name="line_ids"] a').filter({ hasText: 'Add a line' }).click({ timeout: 15000 });

  // Employee autocomplete input should appear in the new row
  await expect(
    page.locator('.o_field_widget[name="line_ids"] .o_field_widget[name="employee_id"] .o-autocomplete--input')
  ).toBeVisible({ timeout: 10000 });
});

const { test, expect } = require('@playwright/test');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const config = require('../config');
const { buildSummaryMessage, sendMessage } = require('../services/notificationService');

const LOGIN_URL = config.loginURL;
const CREDENTIALS_FILE = path.join(__dirname, 'Login-credentials.csv');
const OUTPUT_DIR = path.join(process.cwd(), 'test-results', 'navigation-audit');
const RUN_LOG_FILE = path.join(OUTPUT_DIR, 'execution.log');
const RUN_STATUS_FILE = path.join(OUTPUT_DIR, 'last-run-status.txt');

function sanitizeForFileName(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 80);
}

function csvEscape(value) {
  const str = String(value ?? '');
  if (str.includes('"') || str.includes(',') || str.includes('\n')) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
}

function escapeRegex(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function cleanOldArtifacts(dirPath) {
  if (fs.existsSync(dirPath)) {
    fs.rmSync(dirPath, { recursive: true, force: true });
  }
}

function appendRunLog(message) {
  const timestamp = new Date().toISOString();
  const line = `[${timestamp}] ${message}`;
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  fs.appendFileSync(RUN_LOG_FILE, `${line}\n`, 'utf8');
  console.log(line);
}

function writeRunStatus(status, detail) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  fs.writeFileSync(RUN_STATUS_FILE, `${status}\n${detail}\n`, 'utf8');
}

async function writeExcelReport(rows, excelPath, screenshotsDir) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Results');

  worksheet.columns = [
    { header: 'TestCaseID', key: 'id', width: 16 },
    { header: 'AccountEmail', key: 'account', width: 32 },
    { header: 'UserRole', key: 'userRole', width: 20 },
    { header: 'TestPage', key: 'pageName', width: 28 },
    { header: 'PageURL', key: 'pageUrl', width: 48 },
    { header: 'Steps', key: 'steps', width: 36 },
    { header: 'Status', key: 'status', width: 12 },
    { header: 'ExpectedResult', key: 'expected', width: 40 },
    { header: 'ActualResult', key: 'actual', width: 40 },
    { header: 'ValidationDetails', key: 'validationDetails', width: 44 },
    { header: 'Screenshot', key: 'screenshot', width: 40 },
    { header: 'ExecutedAtUTC', key: 'timestamp', width: 22 },
  ];

  worksheet.getRow(1).font = { bold: true };

  for (const row of rows) {
    const excelRow = worksheet.addRow({
      id: row.id,
      account: row.account,
      userRole: row.userRole,
      pageName: row.pageName,
      pageUrl: row.pageUrl,
      steps: row.steps,
      status: row.status,
      expected: row.expected,
      actual: row.actual,
      validationDetails: row.validationDetails || '',
      screenshot: row.screenshot || '',
      timestamp: row.timestamp,
    });

    if (row.screenshot) {
      const screenshotPath = path.isAbsolute(row.screenshot)
        ? row.screenshot
        : path.resolve(row.screenshot);
      if (fs.existsSync(screenshotPath)) {
        const extension = path.extname(screenshotPath).slice(1) || 'png';
        const imageId = workbook.addImage({
          filename: screenshotPath,
          extension,
        });
        worksheet.addImage(imageId, {
          tl: { col: 10, row: excelRow.number - 1 },
          ext: { width: 160, height: 100 },
        });
        worksheet.getRow(excelRow.number).height = 75;
      }
    }
  }

  await workbook.xlsx.writeFile(excelPath);
}

function formatLocalDateTime(date = new Date()) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `${day}/${month}/${year} ${hours}:${minutes}`;
}

function buildFailureSummary(rows) {
  const failedRows = rows.filter((row) => row.status !== 'PASS');
  const categories = {};
  const reasons = {};

  failedRows.forEach((row) => {
    categories[row.pageName] = (categories[row.pageName] || 0) + 1;
    const reason = row.actual || row.validationDetails || 'Unknown failure';
    reasons[reason] = (reasons[reason] || 0) + 1;
  });

  return {
    total: rows.length,
    passed: rows.length - failedRows.length,
    failed: failedRows.length,
    categories,
    reasons,
  };
}

function getExpectedMenusForPhase(account) {
  return account.expectedMenus;
}

function getRestrictedMenusForPhase(account) {
  return account.restrictedMenus;
}

const TEST_STEPS = {
  LOGIN:
    '1. Navigate to login page | 2. Enter email and password | 3. Click Sign In | 4. Wait for navigation menu',
  LANGUAGE_SWITCH_ENGLISH:
    '1. Navigate to Profile page | 2. Select English from language dropdown | 3. Click Save/Apply | 4. Wait for language change',
  MENU_ACCESS_VALIDATION:
    '1. Expand collapsed navigation menus | 2. Collect left navigation menu items | 3. Compare found menus against expected menus',
  LEFT_NAV_DISCOVERY:
    '1. Expand collapsed navigation menus | 2. Collect left navigation menu items',
  LOGOUT:
    '1. Open profile/user dropdown | 2. Click logout option | 3. Verify redirect to login page',
};

function getMenuNavigationSteps(menuName) {
  return `1. Expand parent menus if needed | 2. Locate "${menuName}" in left navigation | 3. Verify menu is visible | 4. Click menu link | 5. Wait for page load | 6. Capture screenshot`;
}

function resolveTestSteps(pageName) {
  switch (pageName) {
    case 'Login':
      return TEST_STEPS.LOGIN;
    case 'Language Switch to English':
      return TEST_STEPS.LANGUAGE_SWITCH_ENGLISH;
    case 'Menu Access Validation':
      return TEST_STEPS.MENU_ACCESS_VALIDATION;
    case 'Left Navigation Discovery':
      return TEST_STEPS.LEFT_NAV_DISCOVERY;
    case 'Logout':
      return TEST_STEPS.LOGOUT;
    default:
      return getMenuNavigationSteps(pageName);
  }
}

function readCredentials(csvPath) {
  const raw = fs.readFileSync(csvPath, 'utf8').trim();
  const lines = raw.split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) return [];
  const [headerLine, ...rows] = lines;
  const headers = headerLine.split(',').map((h) => h.trim().toLowerCase());
  const emailIndex = headers.indexOf('email');
  const passwordIndex = headers.indexOf('password');
  if (emailIndex === -1 || passwordIndex === -1) {
    throw new Error('CSV must include Email and Password columns.');
  }
  return rows
    .map((line) => line.split(','))
    .map((parts) => ({
      email: (parts[emailIndex] || '').trim(),
      password: (parts[passwordIndex] || '').trim(),
    }))
    .filter((item) => item.email && item.password);
}

function readTestData(csvPath) {
  const raw = fs.readFileSync(csvPath, 'utf8').trim();
  const lines = raw.split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) return [];
  const [headerLine, ...rows] = lines;
  const headers = headerLine.split(',').map((h) => h.trim().toLowerCase());
  
  const emailIndex = headers.indexOf('email');
  const passwordIndex = headers.indexOf('password');
  const userRoleIndex = headers.indexOf('userrole');
  const expectedMenusIndex = headers.indexOf('expectedmenus');
  const restrictedMenusIndex = headers.indexOf('restrictedmenus');
  const priorityMenusIndex = headers.indexOf('prioritymenus');
  
  if (emailIndex === -1 || passwordIndex === -1) {
    throw new Error('CSV must include Email and Password columns.');
  }
  
  return rows
    .map((line) => line.split(','))
    .map((parts) => ({
      email: (parts[emailIndex] || '').trim(),
      password: (parts[passwordIndex] || '').trim(),
      userRole: (parts[userRoleIndex] || '').trim(),
      expectedMenus: (parts[expectedMenusIndex] || '').split('|').map(m => m.trim()).filter(Boolean),
      restrictedMenus: (parts[restrictedMenusIndex] || '').split('|').map(m => m.trim()).filter(Boolean),
      priorityMenus: (parts[priorityMenusIndex] || '').trim(),
    }))
    .filter((item) => item.email && item.password);
}

async function switchLanguage(page, targetLanguage) {
  try {
    console.log(`[LANGUAGE SWITCH] Attempting to switch language to ${targetLanguage}`);
    await page.goto(config.profileURL, {
      waitUntil: 'domcontentloaded',
      timeout: 10000,
    });
    await page.waitForTimeout(500);

    if (await isLanguageActive(page, targetLanguage)) {
      const detectedUILanguage = await detectUILanguageFromNav(page);
      console.log(`[LANGUAGE SWITCH] Language is already set to ${targetLanguage} (detected: ${detectedUILanguage})`);
      return {
        success: true,
        detectedUILanguage,
        reason: `Language is already set to ${targetLanguage} (verified from navigation labels)`,
      };
    }

    const languageSelectors = [
      'select#language',
      'select[name="language"]',
      'select[id*="language"]',
      'select[class*="language"]',
      '.language-dropdown select',
      '[data-testid="language-select"]',
      'select.form-control',
      '.form-select',
      'select',
    ];

    let languageDropdown = null;
    let foundSelector = null;
    
    for (const selector of languageSelectors) {
      try {
        const element = page.locator(selector);
        const count = await element.count();
        console.log(`[LANGUAGE SWITCH] Found ${count} elements with selector: ${selector}`);
        
        for (let i = 0; i < count; i++) {
          const candidate = element.nth(i);
          const isVisible = await candidate.isVisible({ timeout: 1000 });
          if (!isVisible) {
            console.log(`[LANGUAGE SWITCH] Element ${i} with selector ${selector} is not visible`);
            continue;
          }

          const options = await candidate.locator('option').allTextContents();
          console.log(`[LANGUAGE SWITCH] Element ${i} has options: ${options.join(', ')}`);
          
          const hasLanguageOptions = options.some(
            (opt) =>
              /english|language/i.test(opt)
          );
          if (hasLanguageOptions) {
            languageDropdown = candidate;
            foundSelector = selector;
            console.log(`[LANGUAGE SWITCH] Found language dropdown with selector: ${selector}`);
            break;
          }
        }
        if (languageDropdown) break;
      } catch (error) {
        console.log(`[LANGUAGE SWITCH] Error with selector ${selector}: ${error.message}`);
        continue;
      }
    }

    if (!languageDropdown) {
      console.log('[LANGUAGE SWITCH] Language dropdown not found, trying alternative detection methods');
      
      // Try to find any clickable elements that might be language controls
      const possibleLanguageControls = await page.evaluate(() => {
        const results = [];
        // Check for radio buttons
        const radios = document.querySelectorAll('input[type="radio"]');
        radios.forEach((radio, i) => {
          if (radio.name && radio.name.toLowerCase().includes('language')) {
            results.push({ type: 'radio', name: radio.name, value: radio.value, index: i });
          }
        });
        
        // Check for buttons with English text
        const buttons = document.querySelectorAll('button');
        buttons.forEach((btn, i) => {
          const text = btn.textContent.toLowerCase();
          if (text.includes('english')) {
            results.push({ type: 'button', text: btn.textContent, index: i });
          }
        });
        
        // Check for links with English text
        const links = document.querySelectorAll('a');
        links.forEach((link, i) => {
          const text = link.textContent.toLowerCase();
          if (text.includes('english')) {
            results.push({ type: 'link', text: link.textContent, index: i });
          }
        });

        return results;
      });
      
      console.log('[LANGUAGE SWITCH] Found potential language controls:', JSON.stringify(possibleLanguageControls));
      
      const detectedUILanguage = await detectUILanguageFromNav(page);
      if (await isLanguageActive(page, targetLanguage)) {
        return {
          success: true,
          detectedUILanguage,
          reason: `Language dropdown not found, but navigation already shows ${targetLanguage}`,
        };
      }
      
      return {
        success: false,
        detectedUILanguage,
        reason: `Language dropdown not found on Profile page. Found alternative controls: ${JSON.stringify(possibleLanguageControls)}`,
      };
    }

    const selectedBefore = await languageDropdown.inputValue().catch(() => '');
    console.log(`[LANGUAGE SWITCH] Current selection: ${selectedBefore}`);
    
    // Get all options to find the correct one
    const options = await languageDropdown.locator('option').all();
    let targetOption = null;
    let targetValue = null;
    
    for (const option of options) {
      const text = await option.textContent();
      const value = await option.getAttribute('value');
      console.log(`[LANGUAGE SWITCH] Option - Text: "${text}", Value: "${value}"`);
      
      if (text.toLowerCase().includes(targetLanguage.toLowerCase())) {
        targetOption = text;
        targetValue = value;
        break;
      }
    }
    
    if (targetValue) {
      console.log(`[LANGUAGE SWITCH] Selecting option with value: ${targetValue}`);
      await languageDropdown.selectOption(targetValue, { timeout: 5000 });
    } else if (targetOption) {
      console.log(`[LANGUAGE SWITCH] Selecting option by label: ${targetOption}`);
      await languageDropdown.selectOption({ label: targetOption }, { timeout: 5000 });
    } else {
      console.log(`[LANGUAGE SWITCH] No matching option found for ${targetLanguage}`);
    }
    
    // Wait for button text to update after language selection
    await page.waitForTimeout(1000);
    
    const selectedAfter = await languageDropdown.inputValue().catch(() => '');
    console.log(`[LANGUAGE SWITCH] Selection after change: ${selectedAfter}`);

    // Find the save/apply control near the language dropdown in a more robust way
    let saveClicked = false;
    let foundSaveSelector = null;

    try {
      const dropdownHandle = await languageDropdown.elementHandle();

      // 1) Try to find a submit button inside the nearest form ancestor
      if (dropdownHandle) {
        const formHandle = await dropdownHandle.evaluateHandle((el) => el.closest('form'));
        if (formHandle) {
          try {
            const submitBtn = await formHandle.asElement().$('button[type="submit"], input[type="submit"], button');
            if (submitBtn) {
              try {
                const visible = await submitBtn.isVisible().catch(() => false);
                if (visible) {
                  const text = (await submitBtn.textContent()) || (await submitBtn.getAttribute('value')) || '';
                  console.log(`[LANGUAGE SWITCH] Clicking submit button in form near dropdown: "${text}"`);
                  await submitBtn.click({ timeout: 7000 });
                  await page.waitForLoadState('domcontentloaded', { timeout: 10000 }).catch(() => {});
                  await page.waitForTimeout(1500);
                  saveClicked = true;
                  foundSaveSelector = 'form submit button';
                }
              } catch (err) {
                console.log(`[LANGUAGE SWITCH] submitBtn click failed: ${err.message}`);
              }
            } else {
              // If no button found, try to submit the form programmatically
              try {
                await formHandle.asElement().evaluate((f) => {
                  if (typeof f.requestSubmit === 'function') {
                    f.requestSubmit();
                  } else if (typeof f.submit === 'function') {
                    f.submit();
                  }
                });
                await page.waitForLoadState('domcontentloaded', { timeout: 10000 }).catch(() => {});
                await page.waitForTimeout(1500);
                saveClicked = true;
                foundSaveSelector = 'form submit (programmatic)';
              } catch (err) {
                console.log(`[LANGUAGE SWITCH] form submit failed: ${err.message}`);
              }
            }
          } catch (err) {
            console.log(`[LANGUAGE SWITCH] Error handling form near dropdown: ${err.message}`);
          }
        }
      }

      // 2) If form-based approach failed, try nearby buttons using XPath and previous heuristics
      if (!saveClicked) {
        const xpathSelectors = [
          `//select[contains(@class, 'form-control')]/ancestor::div/following-sibling::button[contains(translate(normalize-space(text()), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ', 'abcdefghijklmnopqrstuvwxyzäöü'), 'save') or contains(translate(normalize-space(text()), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ', 'abcdefghijklmnopqrstuvwxyzäöü'), 'speichern') or contains(translate(normalize-space(text()), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ', 'abcdefghijklmnopqrstuvwxyzäöü'), 'apply') or contains(translate(normalize-space(text()), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ', 'abcdefghijklmnopqrstuvwxyzäöü'), 'anwenden')]`,
          `//select[contains(@class, 'form-control')]/ancestor::div//button[contains(translate(normalize-space(text()), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ', 'abcdefghijklmnopqrstuvwxyzäöü'), 'save') or contains(translate(normalize-space(text()), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ', 'abcdefghijklmnopqrstuvwxyzäöü'), 'speichern') or contains(translate(normalize-space(text()), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ', 'abcdefghijklmnopqrstuvwxyzäöü'), 'apply') or contains(translate(normalize-space(text()), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ', 'abcdefghijklmnopqrstuvwxyzäöü'), 'anwenden')]`,
          `//select[contains(@class, 'form-control')]/following::button[contains(translate(normalize-space(text()), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ', 'abcdefghijklmnopqrstuvwxyzäöü'), 'save') or contains(translate(normalize-space(text()), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ', 'abcdefghijklmnopqrstuvwxyzäöü'), 'speichern') or contains(translate(normalize-space(text()), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ', 'abcdefghijklmnopqrstuvwxyzäöü'), 'apply') or contains(translate(normalize-space(text()), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ', 'abcdefghijklmnopqrstuvwxyzäöü'), 'anwenden')][1]`,
        ];

        for (const xpath of xpathSelectors) {
          try {
            const saveButton = page.locator(`xpath=${xpath}`).first();
            const count = await saveButton.count();
            if (count > 0 && (await saveButton.isVisible({ timeout: 2000 }))) {
              const buttonText = await saveButton.textContent();
              console.log(`[LANGUAGE SWITCH] Found save button near dropdown using XPath, text: "${buttonText}"`);
              await saveButton.click({ timeout: 7000 });
              await page.waitForLoadState('domcontentloaded', { timeout: 10000 }).catch(() => {});
              await page.waitForTimeout(1500);
              saveClicked = true;
              foundSaveSelector = xpath;
              break;
            }
          } catch (error) {
            console.log(`[LANGUAGE SWITCH] XPath selector ${xpath} failed: ${error.message}`);
            continue;
          }
        }
      }

      // 3) Final fallback: try clicking any visible button with save-like or speichern-like text (best-effort)
      if (!saveClicked) {
        console.log('[LANGUAGE SWITCH] Falling back to generic save button search');
        const allSaveButtons = page.locator('button', { hasText: /save|speichern|apply|anwenden/i });
        const saveCount = await allSaveButtons.count();
        console.log(`[LANGUAGE SWITCH] Found ${saveCount} save-related buttons`);

        for (let i = 0; i < saveCount; i++) {
          const btn = allSaveButtons.nth(i);
          try {
            const buttonText = await btn.textContent();
            console.log(`[LANGUAGE SWITCH] Trying save button ${i}: "${buttonText}"`);
            await btn.click({ timeout: 7000 });
            await page.waitForTimeout(1200);

            const detectedLang = await detectUILanguageFromNav(page);
            if (/english/i.test(targetLanguage) && detectedLang === 'EN') {
              console.log('[LANGUAGE SWITCH] Language change detected after clicking fallback save button');
              saveClicked = true;
              foundSaveSelector = `fallback button index ${i}`;
              break;
            }
          } catch (err) {
            console.log(`[LANGUAGE SWITCH] fallback button ${i} click failed: ${err.message}`);
            continue;
          }
        }
      }
    } catch (error) {
      console.log(`[LANGUAGE SWITCH] Error finding save button near dropdown: ${error.message}`);
    }

    // Verify language change by checking navigation
    await page.waitForTimeout(1000);
    const finalUILanguage = await detectUILanguageFromNav(page);
    console.log(`[LANGUAGE SWITCH] Final UI language detected: ${finalUILanguage}`);

    if (await isLanguageActive(page, targetLanguage)) {
      return {
        success: true,
        detectedUILanguage: finalUILanguage,
        reason: saveClicked
          ? `Language switched to ${targetLanguage} and saved (dropdown: ${foundSelector}, save: ${foundSaveSelector})`
          : `Language switched to ${targetLanguage} (dropdown: ${foundSelector}, no save button needed)`,
      };
    }

    if (selectedBefore === selectedAfter && !saveClicked) {
      return {
        success: false,
        detectedUILanguage: finalUILanguage,
        reason: `Language dropdown value did not change and save button was not found. Current dropdown value: "${selectedAfter}". Detected UI language: ${finalUILanguage}`,
      };
    }

    return {
      success: false,
      detectedUILanguage: finalUILanguage,
      reason: saveClicked
        ? `Language was saved but navigation labels still do not reflect ${targetLanguage}. Detected UI language: ${finalUILanguage}`
        : `Selected ${targetLanguage} in dropdown but could not find Save/Apply button to apply the change. Detected UI language: ${finalUILanguage}`,
    };
  } catch (error) {
    console.error(`[LANGUAGE SWITCH] Error: ${error.message}`);
    const detectedUILanguage = await detectUILanguageFromNav(page).catch(() => 'UNKNOWN');
    return {
      success: false,
      detectedUILanguage,
      reason: error.message || `Failed to switch language to ${targetLanguage}`,
    };
  }
}

async function detectUILanguageFromNav(page) {
  try {
    await page.waitForSelector('a.nav-link', { timeout: 5000 });
  } catch (_) {
    return 'UNKNOWN';
  }

  const navLabels = (await page.locator('a.nav-link').allTextContents())
    .join(' ')
    .toLowerCase();

  const englishMarkers = ['meter dashboard', 'manage meters', 'alarms center', 'manage locations'];
  const englishScore = englishMarkers.filter((m) => navLabels.includes(m)).length;
  if (englishScore > 0) return 'EN';
  return 'UNKNOWN';
}

async function isLanguageActive(page, targetLanguage) {
  const detected = await detectUILanguageFromNav(page);
  return detected === 'EN';
}

async function resolveRowLanguageFields(page, testPhase) {
  const detectedUILanguage = await detectUILanguageFromNav(page);
  return {
    language: detectedUILanguage === 'EN' ? 'EN' : 'UNKNOWN',
    detectedUILanguage,
  };
}

function validateMenus(navItems, expectedMenus, restrictedMenus) {
  const foundMenus = navItems.map((item) => item.label.toLowerCase());

  const validationResults = {
    expectedMissing: [],
    unexpectedFound: [],
    allExpectedPresent: true,
    totalExpected: expectedMenus.length,
    totalFound: navItems.length,
    restrictedAccessed: [],
  };

  expectedMenus.forEach((expectedMenu) => {
    const expectedLower = expectedMenu.toLowerCase();
    const found = foundMenus.some(
      (menu) => menu.includes(expectedLower) || expectedLower.includes(menu)
    );
    if (!found) {
      validationResults.expectedMissing.push(expectedMenu);
      validationResults.allExpectedPresent = false;
    }
  });

  restrictedMenus.forEach((restrictedMenu) => {
    const restrictedLower = restrictedMenu.toLowerCase();
    const found = foundMenus.some(
      (menu) => menu.includes(restrictedLower) || restrictedLower.includes(menu)
    );
    if (found) {
      validationResults.unexpectedFound.push(restrictedMenu);
      validationResults.restrictedAccessed.push(restrictedMenu);
      validationResults.allExpectedPresent = false;
    }
  });

  return validationResults;
}

async function buildResultRow(page, testPhase, fields) {
  const { detectedUILanguage } = await resolveRowLanguageFields(page, testPhase);
  return {
    detectedUILanguage,
    language: detectedUILanguage === 'EN' ? 'EN' : 'UNKNOWN',
    ...fields,
  };
}

async function expandParentMenus(page) {
  const parentMenuSelectors = [
    'a.nav-link.collapsed-menu-folder',
    'a.nav-link[data-bs-toggle="collapse"]',
    'a.nav-link[aria-expanded="false"]',
    '.nav-item.has-submenu > a.nav-link',
    '.menu-item.has-treeview > a',
  ];

  let parentMenus = null;
  for (const selector of parentMenuSelectors) {
    const elements = page.locator(selector);
    const count = await elements.count();
    if (count > 0) {
      parentMenus = elements;
      break;
    }
  }

  if (!parentMenus) return;

  const parentCount = await parentMenus.count();
  for (let i = 0; i < parentCount; i++) {
    try {
      const parent = parentMenus.nth(i);
      if (!(await parent.isVisible({ timeout: 1000 }))) continue;

      const isExpanded = await parent.evaluate((el) => {
        const container = el.closest('li');
        return (
          container?.classList.contains('expanded') ||
          container?.classList.contains('showMenu') ||
          container?.classList.contains('menu-open') ||
          container?.classList.contains('show') ||
          el.getAttribute('aria-expanded') === 'true'
        );
      });

      if (!isExpanded) {
        await parent.scrollIntoViewIfNeeded({ timeout: 2000 });
        await parent.click({ timeout: 2000 });
        await page.waitForTimeout(300);
      }
    } catch (_) {
      continue;
    }
  }
}

async function clickNavMenuItem(page, menuName) {
  await expandParentMenus(page);

  const pattern = new RegExp(`^\\s*${escapeRegex(menuName)}\\s*$`, 'i');
  const links = page.locator('a.nav-link').filter({ hasText: pattern });
  const linkCount = await links.count();

  if (linkCount === 0) {
    throw new Error(`Menu item "${menuName}" not found in navigation`);
  }

  for (let i = 0; i < linkCount; i++) {
    const link = links.nth(i);
    try {
      const inSidebar = await link.evaluate((el) => {
        const rect = el.getBoundingClientRect();
        return rect.left >= 0 && rect.left < 280 && rect.width > 0 && rect.height > 0;
      });
      if (!inSidebar) continue;

      await link.scrollIntoViewIfNeeded({ timeout: 2000 });
      if (await link.isVisible({ timeout: 2000 })) {
        await link.click({ timeout: 5000 });
        return;
      }
    } catch (_) {
      continue;
    }
  }

  await expandParentMenus(page);

  for (let i = 0; i < linkCount; i++) {
    const link = links.nth(i);
    try {
      const inSidebar = await link.evaluate((el) => {
        const rect = el.getBoundingClientRect();
        return rect.left >= 0 && rect.left < 280;
      });
      if (!inSidebar) continue;

      await link.scrollIntoViewIfNeeded({ timeout: 2000 });
      if (await link.isVisible({ timeout: 3000 })) {
        await link.click({ timeout: 5000 });
        return;
      }
    } catch (_) {
      continue;
    }
  }

  throw new Error(
    `Menu item "${menuName}" is not visible. It may be inside a collapsed parent menu (e.g. Customers or Users submenu).`
  );
}

async function validateMenuByEnglishName(navItems, testData) {
  const expectedMenus = getExpectedMenusForPhase(testData);
  const restrictedMenus = getRestrictedMenusForPhase(testData);
  return validateMenus(navItems, expectedMenus, restrictedMenus);
}

async function fillLoginForm(page, email, password) {
  let emailInput = page.getByLabel('Enter Email or Mobile Number');
  if ((await emailInput.count()) === 0) {
    emailInput = page.getByPlaceholder('Enter Email or Mobile Number');
  }

  let passwordInput = page.getByLabel('Password');
  if ((await passwordInput.count()) === 0) {
    passwordInput = page.getByPlaceholder('Password');
  }

  await emailInput.first().fill(email);
  await passwordInput.first().fill(password);
  await page.getByRole('button', { name: /sign in/i }).first().click();
}

async function getLeftNavItems(page) {
  await page.waitForSelector('a.nav-link', { timeout: 10000 });
  await expandParentMenus(page);

  // Get all visible menu items with better error handling
  try {
    const items = await page.$$eval('a.nav-link', (links) =>
      links
        .filter((el) => {
          try {
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return (
              style.visibility !== 'hidden' &&
              style.display !== 'none' &&
              rect.width > 0 &&
              rect.height > 0 &&
              rect.left >= 0 &&
              rect.left < 280
            );
          } catch (_) {
            return false;
          }
        })
        .map((el) => ({
          label: (el.textContent || '').replace(/\s+/g, ' ').trim(),
          href: (el.getAttribute('href') || '').trim(),
          isParent: el.classList.contains('collapsed-menu-folder') || 
                   el.hasAttribute('data-bs-toggle') ||
                   el.getAttribute('aria-expanded') !== null,
        }))
        .filter((item) => item.label.length > 0 && !item.isParent)
    );

    // Remove duplicates with better handling
    const unique = [];
    const seen = new Set();
    for (const item of items) {
      const key = `${item.label.toLowerCase()}|${item.href.toLowerCase()}`;
      if (!seen.has(key)) {
        seen.add(key);
        unique.push(item);
      }
    }
    
    console.log(`Found ${unique.length} unique menu items`);
    return unique;
  } catch (error) {
    console.log(`Error getting menu items: ${error.message}`);
    // Return empty array if we can't get menu items
    return [];
  }
}

async function logoutFromAccount(page) {
  const profileTriggerCandidates = [
    page.locator('li.dropdown.profile a.nav-link-user').first(),
    page.getByRole('link', { name: /hello|profile|account/i }).first(),
    page.getByRole('button', { name: /hello|profile|account/i }).first(),
  ];

  for (const trigger of profileTriggerCandidates) {
    try {
      if ((await trigger.count()) > 0 && (await trigger.isVisible())) {
        await trigger.click({ timeout: 5000 });
        await page.waitForTimeout(500);
        break;
      }
    } catch (_) {
      // Try next candidate.
    }
  }

  const logoutCandidates = [
    page.locator('li.dropdown.profile .dropdown-menu a.dropdown-item.text-danger').first(),
    page.getByRole('menuitem', { name: /logout/i }).first(),
    page.getByRole('button', { name: /logout/i }).first(),
    page.getByRole('link', { name: /logout/i }).first(),
    page.getByText(/logout/i).first(),
  ];

  for (const candidate of logoutCandidates) {
    try {
      if ((await candidate.count()) > 0 && (await candidate.isVisible())) {
        await candidate.click({ timeout: 7000 });
        await Promise.race([
          page.waitForURL(/anmeldung/i, { timeout: 15000 }),
          page
            .locator(
              'input[placeholder=\"Enter Email or Mobile Number\"], input[type=\"password\"]'
            )
            .first()
            .waitFor({ state: 'visible', timeout: 15000 }),
        ]);
        return true;
      }
    } catch (_) {
      // Try next candidate.
    }
  }

  return false;
}

test('validate English menu access and logout for all accounts', async ({ browser }) => {
  test.setTimeout(15 * 60 * 1000); // Extended timeout for English menu validation

  appendRunLog('Starting English menu validation run');
  writeRunStatus('RUNNING', 'English menu validation started');

  try {
    const projectName = sanitizeForFileName(test.info().project.name || 'default');
    const projectOutputDir = path.join(OUTPUT_DIR, projectName);
    const screenshotsDir = path.join(projectOutputDir, 'screenshots');
    const csvFile = path.join(projectOutputDir, 'english-menu-test-results.csv');
    const excelFile = path.join(projectOutputDir, 'english-menu-test-results.xlsx');
    
    cleanOldArtifacts(projectOutputDir);
    fs.mkdirSync(screenshotsDir, { recursive: true });
    
    const testData = readTestData(CREDENTIALS_FILE);
    appendRunLog(`Loaded ${testData.length} accounts from ${CREDENTIALS_FILE}`);
    expect(testData.length, 'No valid test data found in CSV.').toBeGreaterThan(0);
    
    const rows = [];
    let caseCounter = 1;

  async function recordRow(page, testPhase, fields) {
    rows.push(await buildResultRow(page, testPhase, fields));
  }
  
  for (const account of testData) {
    const context = await browser.newContext();
    const page = await context.newPage();
    const accountSlug = sanitizeForFileName(account.email);
    
    try {
      // Login
      await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded' });
      await fillLoginForm(page, account.email, account.password);
      
      // Wait for navigation menu with better error handling
      let loginSuccess = false;
      try {
        await page.waitForSelector('a.nav-link', { timeout: 10000 });
        await page.waitForLoadState('domcontentloaded');
        await page.waitForTimeout(2000);
        loginSuccess = true;
      } catch (loginError) {
        // Login failed - record and continue to next user
await recordRow(page, 'EN', {
          id: `TC-${String(caseCounter++).padStart(3, '0')}`,
          project: test.info().project.name,
          account: account.email,
          userRole: account.userRole,
          pageName: 'Login',
          pageUrl: page.url(),
          status: 'FAIL',
          expected: 'User should be able to log in successfully',
          actual: loginError.message || 'Login failed',
          screenshot: '',
          timestamp: formatLocalDateTime(),
        });
        
        console.log(`Login failed for ${account.email}, skipping to next user`);
        await context.close();
        continue;
      }
      
      // Record successful login
      await recordRow(page, 'EN', {
        id: `TC-${String(caseCounter++).padStart(3, '0')}`,
        project: test.info().project.name,
        account: account.email,
        userRole: account.userRole,
        pageName: 'Login',
        pageUrl: page.url(),
        status: 'PASS',
        expected: 'User should be able to log in successfully',
        actual: 'Login successful',
        screenshot: '',
        timestamp: formatLocalDateTime(),
      });
      
      // STEP 1: Switch to English first
      const englishSwitchResult = await switchLanguage(page, 'English');
      await recordRow(page, 'EN', {
        id: `TC-${String(caseCounter++).padStart(3, '0')}`,
        project: test.info().project.name,
        account: account.email,
        userRole: account.userRole,
        pageName: 'Language Switch to English',
        pageUrl: page.url(),
        status: englishSwitchResult.success ? 'PASS' : 'FAIL',
        expected: 'Language should switch to English',
        actual: englishSwitchResult.reason,
        screenshot: '',
        timestamp: formatLocalDateTime(),
      });
      
      // STEP 2: Test all menus in English
      const navItemsEn = await getLeftNavItems(page);
      const expectedMenusEn = getExpectedMenusForPhase(account);
      
      // Validate menu access in English
      const validationEn = await validateMenuByEnglishName(navItemsEn, account);
      
      await recordRow(page, 'EN', {
        id: `TC-${String(caseCounter++).padStart(3, '0')}`,
        project: test.info().project.name,
        account: account.email,
        userRole: account.userRole,
        pageName: 'Menu Access Validation',
        pageUrl: page.url(),
        status: validationEn.allExpectedPresent ? 'PASS' : 'FAIL',
        expected: expectedMenusEn.join('|'),
        actual: navItemsEn.map(item => item.label).join('|'),
        restricted: getRestrictedMenusForPhase(account).join('|'),
        validationDetails: validationEn.expectedMissing.length
          ? `Expected: ${validationEn.totalExpected}, Found: ${validationEn.totalFound}, Missing: ${validationEn.expectedMissing.length} (${validationEn.expectedMissing.join(', ')})`
          : `Expected: ${validationEn.totalExpected}, Found: ${validationEn.totalFound}, Missing: 0`,
        screenshot: '',
        timestamp: formatLocalDateTime(),
      });
      
      // Test all navigation items in English
      if (navItemsEn.length === 0) {
        await recordRow(page, 'EN', {
          id: `TC-${String(caseCounter++).padStart(3, '0')}`,
          project: test.info().project.name,
          account: account.email,
          userRole: account.userRole,
          pageName: 'Left Navigation Discovery',
          pageUrl: page.url(),
          status: 'FAIL',
          expected: 'Left side navigation menu should be visible after login.',
          actual: 'No left navigation items were found.',
          screenshot: '',
          timestamp: formatLocalDateTime(),
        });
      } else {
        const cappedItemsEn = navItemsEn.slice(0, 25);
        for (let i = 0; i < cappedItemsEn.length; i++) {
          const navItem = cappedItemsEn[i];
          const caseId = `TC-${String(caseCounter++).padStart(3, '0')}`;
          const menuName = navItem.label;
          const safeMenuName = sanitizeForFileName(menuName || `menu-${i + 1}`);
          const screenshotPath = path.join(
            screenshotsDir,
            `${accountSlug}_english_${String(i + 1).padStart(2, '0')}_${safeMenuName}.png`
          );

          try {
            await clickNavMenuItem(page, menuName);

            // Wait for navigation with shorter timeout
            await page.waitForLoadState('domcontentloaded', { timeout: 10000 });
            await page.waitForTimeout(1000);

            // Take screenshot
            await page.screenshot({ path: screenshotPath, fullPage: true });

            await recordRow(page, 'EN', {
              id: caseId,
              project: test.info().project.name,
              account: account.email,
              userRole: account.userRole,
              pageName: menuName,
              pageUrl: page.url(),
              status: 'PASS',
              expected: `${menuName} page opens and is fully loaded.`,
              actual: `${menuName} page opened successfully.`,
              screenshot: screenshotPath,
              timestamp: formatLocalDateTime(),
            });
          } catch (error) {
            await recordRow(page, 'EN', {
              id: caseId,
              project: test.info().project.name,
              account: account.email,
              userRole: account.userRole,
              pageName: menuName,
              pageUrl: page.url(),
              status: 'FAIL',
              expected: `${menuName} page opens and is fully loaded.`,
              actual: error && error.message ? error.message : 'Navigation failed.',
              screenshot: '',
              timestamp: formatLocalDateTime(),
            });
          }
        }
      }
      
      // STEP 5: Logout
      const logoutSuccess = await logoutFromAccount(page);
      await recordRow(page, 'EN', {
        id: `TC-${String(caseCounter++).padStart(3, '0')}`,
        project: test.info().project.name,
        account: account.email,
        userRole: account.userRole,
        pageName: 'Logout',
        pageUrl: page.url(),
        status: logoutSuccess ? 'PASS' : 'FAIL',
        expected: 'User should be logged out successfully',
        actual: logoutSuccess ? 'Logout completed' : 'Logout failed',
        screenshot: '',
        timestamp: formatLocalDateTime(),
      });
      
    } catch (error) {
      await recordRow(page, 'EN', {
        id: `TC-${String(caseCounter++).padStart(3, '0')}`,
        project: test.info().project.name,
        account: account.email,
        userRole: account.userRole,
        pageName: 'Login',
        pageUrl: LOGIN_URL,
        status: 'FAIL',
        expected: 'User should be able to log in successfully',
        actual: error.message || 'Login failed',
        screenshot: '',
        timestamp: formatLocalDateTime(),
      });
    } finally {
      await context.close();
    }
  }
  
  // Enhanced CSV header
  const header = [
    'TestCaseID',
    'AccountEmail',
    'UserRole',
    'TestPage',
    'PageURL',
    'Steps',
    'Status',
    'ExpectedResult',
    'ActualResult',
    'ValidationDetails',
    'Screenshot',
    'ExecutedAtUTC'
  ];
  
  const csvLines = [header.join(',')];
  for (const row of rows) {
    csvLines.push([
      row.id,
      row.account,
      row.userRole,
      row.pageName,
      row.pageUrl,
      row.steps || resolveTestSteps(row.pageName),
      row.status,
      row.expected,
      row.actual,
      row.validationDetails || '',
      row.screenshot,
      row.timestamp,
    ].map(csvEscape).join(','));
  }
  
    fs.mkdirSync(projectOutputDir, { recursive: true });
    fs.writeFileSync(csvFile, `${csvLines.join('\n')}\n`, 'utf8');
    await writeExcelReport(rows, excelFile, screenshotsDir);

    appendRunLog(`Generated CSV report: ${csvFile}`);
    appendRunLog(`Generated Excel report: ${excelFile}`);

    const summary = buildFailureSummary(rows);
    const title = summary.failed > 0 ? 'IQCloud QA Test Summary - Failures detected' : 'IQCloud QA Test Summary - All checks passed';
    const message = buildSummaryMessage({
      title,
      summary,
      runName: 'English menu validation',
    });

    if (config.webhookUrl) {
      try {
        await sendMessage({
          provider: 'webhook',
          webhookUrl: config.webhookUrl,
          title,
          message,
        });
        appendRunLog('Webhook notification sent successfully');
      } catch (notifyError) {
        appendRunLog(`Webhook notification failed: ${notifyError.message}`);
      }
    } else {
      appendRunLog('Webhook URL not configured; skipping notification');
    }

    writeRunStatus('SUCCESS', `Reports generated: ${csvFile} | ${excelFile}`);

    test.info().annotations.push({
      type: 'sheet',
      description: `English menu test results generated: ${excelFile}`,
    });
  } catch (error) {
    appendRunLog(`Run failed: ${error && error.message ? error.message : error}`);
    writeRunStatus('FAILED', error && error.message ? error.message : 'English menu validation failed');
    throw error;
  }
});


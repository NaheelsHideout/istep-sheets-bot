import { chromium } from 'playwright';
import { google } from 'googleapis';
import * as XLSX from 'xlsx';
import fs from 'fs';

const {
  ISTEP_LOGIN_URL,
  ISTEP_REPORT_URL,
  ISTEP_FILES_URL,
  ISTEP_USERNAME,
  ISTEP_PASSWORD,
  GOOGLE_SHEET_ID,
  GOOGLE_SERVICE_ACCOUNT_JSON
} = process.env;

const RAW_SHEET_NAME = 'Raw_iStep';
const SUMMARY_SHEET_NAME = 'Summary';

for (const [k, v] of Object.entries({
  ISTEP_LOGIN_URL,
  ISTEP_REPORT_URL,
  ISTEP_FILES_URL,
  ISTEP_USERNAME,
  ISTEP_PASSWORD,
  GOOGLE_SHEET_ID,
  GOOGLE_SERVICE_ACCOUNT_JSON
})) {
  if (!v) throw new Error(`Missing required secret/env: ${k}`);
}

function clean(value) {
  return String(value ?? '')
    .replace(/\s+/g, ' ')
    .replace(/\u00a0/g, ' ')
    .trim();
}

function scoreToNumber(value) {
  const n = Number(clean(value).replace('%', ''));
  return Number.isFinite(n) ? n : '';
}

function parseDate(value) {
  const s = clean(value);
  if (!s || s === '-') return null;
  const d = new Date(s.replace(' ', 'T'));
  return Number.isNaN(d.getTime()) ? null : d;
}

function detectType(subject) {
  const s = clean(subject).toLowerCase();

  if (isExcludedSubject(subject)) return 'Excluded';

  if (
    s.includes('new brand setup') ||
    s.includes('new brand') ||
    s.includes('brand setup')
  ) return 'Build';

  if (
    s.includes('outlet catalogue update request') ||
    s.includes('catalogue update') ||
    s.includes('catalog update') ||
    s.includes('replace menu') ||
    s.includes('menu replace') ||
    s.includes('add new item') ||
    s.includes('add new items') ||
    s.includes('move item') ||
    s.includes('move items') ||
    s.includes('location update') ||
    s.includes('remove items')
  ) return 'Update';

  return 'Other';
}

function isExcludedSubject(subject) {
  const s = clean(subject).toLowerCase();

  return (
    s.includes('new brand setup (shops)') ||
    s.includes('new outlet for existing brand (shops)') ||
    s.includes('outlet catalogue update request- shops') ||
    s.includes('outlet catalogue update request - shops')
  );
}

function normalizeMarket(value) {
  const s = clean(value).toLowerCase();

  if (
    s.includes('uae') ||
    s.includes('dubai') ||
    s.includes('abu dhabi') ||
    s.includes('sharjah') ||
    s.includes('al ain')
  ) return 'UAE';

  if (
    s.includes('jor') ||
    s.includes('jordan') ||
    s.includes('amman') ||
    s.includes('irbid') ||
    s.includes('zarqa')
  ) return 'JOR';

  return clean(value);
}

function avg(values) {
  const nums = values.filter(v => typeof v === 'number' && !Number.isNaN(v));
  if (!nums.length) return null;
  return nums.reduce((a, b) => a + b, 0) / nums.length;
}

function fmtPct(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return '-';
  return `${Number(value).toFixed(2).replace(/\.00$/, '')}%`;
}

function uniqueTickets(rows) {
  const map = new Map();

  for (const r of rows) {
    if (!r.reference) continue;

    if (!map.has(r.reference)) {
      map.set(r.reference, { ...r });
    } else {
      const e = map.get(r.reference);
      e.sentBack = e.sentBack || r.sentBack ? 1 : 0;

      if (!e.date && r.date) e.date = r.date;
      if (!e.ticketScore && r.ticketScore) e.ticketScore = r.ticketScore;
      if (!e.market && r.market) e.market = r.market;
      if (!e.city && r.city) e.city = r.city;

      map.set(r.reference, e);
    }
  }

  return [...map.values()];
}

async function login(page) {
  console.log('Opening iStep login page...');
  await page.goto(ISTEP_LOGIN_URL, { waitUntil: 'networkidle' });

  console.log('Filling login form...');

  const usernameSelectors = [
    'input[name="email"]',
    'input[name="username"]',
    'input[type="email"]',
    'input[type="text"]'
  ];

  const passwordSelectors = [
    'input[name="password"]',
    'input[type="password"]'
  ];

  let filledUser = false;
  for (const selector of usernameSelectors) {
    const el = page.locator(selector).first();
    if (await el.count()) {
      await el.fill(ISTEP_USERNAME);
      filledUser = true;
      break;
    }
  }

  if (!filledUser) throw new Error('Could not find username input.');

  let filledPass = false;
  for (const selector of passwordSelectors) {
    const el = page.locator(selector).first();
    if (await el.count()) {
      await el.fill(ISTEP_PASSWORD);
      filledPass = true;
      break;
    }
  }

  if (!filledPass) throw new Error('Could not find password input.');

  const loginButtons = [
    'button[type="submit"]',
    'input[type="submit"]',
    'button:has-text("Login")',
    'button:has-text("Log in")',
    'button:has-text("Sign in")'
  ];

  let clicked = false;
  for (const selector of loginButtons) {
    const el = page.locator(selector).first();
    if (await el.count()) {
      await el.click({ force: true });
      clicked = true;
      break;
    }
  }

  if (!clicked) await page.keyboard.press('Enter');

  await page.waitForTimeout(5000);
  await page.waitForLoadState('networkidle').catch(() => {});
  console.log('Login completed.');
}

async function selectReportAndExport(page) {
  console.log('Opening report page...');
  await page.goto(ISTEP_REPORT_URL, { waitUntil: 'networkidle' });
  await page.waitForTimeout(4000);

  console.log('Selecting Evaluation...');
  const evaluationRadio = page.locator('input[type="radio"]').first();
  if (await evaluationRadio.count()) {
    await evaluationRadio.check({ force: true }).catch(async () => {
      await evaluationRadio.click({ force: true });
    });
  } else {
    const evalText = page.locator('text=Evaluation').first();
    if (await evalText.count()) await evalText.click({ force: true });
  }

  await page.waitForTimeout(1000);

  console.log('Selecting Overall Evaluation Tickets...');

  let selected = false;
  const selects = page.locator('select');

  for (let i = 0; i < await selects.count(); i++) {
    const s = selects.nth(i);
    const options = await s.locator('option').allTextContents().catch(() => []);
    if (options.some(o => clean(o).includes('Overall Evaluation Tickets'))) {
      await s.selectOption({ label: 'Overall Evaluation Tickets' });
      selected = true;
      break;
    }
  }

  if (!selected) {
    const dropdowns = [
      'text=--Select--',
      '.select2-selection',
      '.form-select',
      '.form-control'
    ];

    for (const selector of dropdowns) {
      const d = page.locator(selector).first();
      if (await d.count()) {
        await d.click({ force: true }).catch(() => {});
        await page.waitForTimeout(1000);

        const option = page.locator('text=Overall Evaluation Tickets').first();
        if (await option.count()) {
          await option.click({ force: true });
          selected = true;
          break;
        }
      }
    }
  }

  if (!selected) {
    console.log('Report dropdown not selected automatically. Continuing.');
  }

  await page.waitForTimeout(1500);

  console.log('Clicking Apply...');
  const apply = page.locator('button:has-text("Apply"), input[value="Apply"], text=Apply').first();
  if (await apply.count()) {
    await apply.click({ force: true });
  } else {
    throw new Error('Apply button not found.');
  }

  await page.waitForTimeout(8000);

  console.log('Clicking Export...');
  const exportBtn = page.locator('button:has-text("Export"), a:has-text("Export"), text=Export').first();

  if (!(await exportBtn.count())) {
    throw new Error('Export button not found.');
  }

  await exportBtn.click({ force: true });

  console.log('Export clicked. Waiting for file to be generated...');
  await page.waitForTimeout(15000);
}

async function downloadLatestExport(page) {
  console.log('Opening File Management...');
  await page.goto(ISTEP_FILES_URL, { waitUntil: 'networkidle' });

  const expectedName = 'Overall_Evaluation_Tickets';
  const maxAttempts = 18; // 18 x 20 seconds = 6 minutes

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    console.log(`Checking export file attempt ${attempt}/${maxAttempts}...`);

    await page.waitForTimeout(3000);

    const rows = page.locator('table tbody tr');
    const count = await rows.count();

    for (let i = 0; i < count; i++) {
      const row = rows.nth(i);
      const txt = clean(await row.innerText().catch(() => ''));

      if (txt.includes(expectedName)) {
        console.log('Found latest Overall Evaluation export.');

        const downloadIcon = row.locator('a, button, i').last();

        const downloadPromise = page.waitForEvent('download', { timeout: 60000 });
        await downloadIcon.click({ force: true });
        const download = await downloadPromise;

        const path = './istep-export.xlsx';
        await download.saveAs(path);

        console.log(`Downloaded export to ${path}`);
        return path;
      }
    }

    console.log('File not ready yet. Refreshing...');
    const refreshBtn = page.locator('button:has(i), button, a').filter({ hasText: '' }).last();

    await page.reload({ waitUntil: 'networkidle' }).catch(() => {});
    await page.waitForTimeout(17000);
  }

  throw new Error('Latest Overall Evaluation export did not appear after waiting.');
}

function parseExportFile(filePath) {
  console.log('Parsing Excel export...');

  if (!fs.existsSync(filePath)) {
    throw new Error(`Export file not found: ${filePath}`);
  }

  const workbook = XLSX.readFile(filePath);
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];

  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ''
  });

  if (!matrix.length) throw new Error('Excel file is empty.');

  const headerRowIndex = matrix.findIndex(row =>
    row.map(clean).includes('Reference ID') &&
    row.map(clean).includes('Subject') &&
    row.map(clean).includes('Ticket Score')
  );

  if (headerRowIndex === -1) {
    throw new Error('Could not find header row in Excel export.');
  }

  const headers = matrix[headerRowIndex].map(clean);
  const dataRows = matrix.slice(headerRowIndex + 1);

  const idx = name => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());

  const indexes = {
    reference: idx('Reference ID'),
    subject: idx('Subject'),
    ticketScore: idx('Ticket Score'),
    catName: idx('Catalogue Name'),
    catUserId: idx('Catalogue User Id'),
    catSentBack: idx('Catalogue Sent Back To Catalog'),
    catCity: idx('Catalogue City'),
    catMarket: idx('Catalogue Market'),
    catScore: idx('Catalogue Score'),
    catDate: idx('Catalogue Date & Time'),
    studioName: idx('Studio Name'),
    studioUserId: idx('Studio User Id'),
    studioSentBack: idx('Studio Sent Back To Catalog'),
    studioCity: idx('Studio City'),
    studioMarket: idx('Studio Market'),
    studioScore: idx('Studio Score'),
    studioFormName: idx('Studio Form Name'),
    studioDate: idx('Studio Date & Time')
  };

  if (indexes.reference === -1 || indexes.subject === -1 || indexes.ticketScore === -1) {
    console.log(headers.slice(0, 30));
    throw new Error('Required export columns missing.');
  }

  let lastReference = '';
  let lastSubject = '';
  let lastTicketScore = '';

  const get = (row, i) => i >= 0 ? clean(row[i]) : '';

  const rows = [];

  for (const row of dataRows) {
    let reference = get(row, indexes.reference);
    let subject = get(row, indexes.subject);
    let ticketScoreRaw = get(row, indexes.ticketScore);

    if (!reference && lastReference) reference = lastReference;
    if (!subject && lastSubject) subject = lastSubject;
    if (!ticketScoreRaw && lastTicketScore) ticketScoreRaw = lastTicketScore;

    if (!reference || !reference.startsWith('TH-')) continue;

    lastReference = reference;
    lastSubject = subject;
    lastTicketScore = ticketScoreRaw;

    if (isExcludedSubject(subject)) continue;

    const catMarket = get(row, indexes.catMarket);
    const studioMarket = get(row, indexes.studioMarket);
    const catCity = get(row, indexes.catCity);
    const studioCity = get(row, indexes.studioCity);

    const catSentBack = Number(get(row, indexes.catSentBack) || 0) || 0;
    const studioSentBack = Number(get(row, indexes.studioSentBack) || 0) || 0;

    const catDate = get(row, indexes.catDate);
    const studioDate = get(row, indexes.studioDate);

    rows.push({
      reference,
      subject,
      ticketScore: scoreToNumber(ticketScoreRaw),
      type: detectType(subject),
      market: normalizeMarket(catMarket || studioMarket || subject),
      city: catCity || studioCity,
      date: catDate || studioDate,
      dateObj: parseDate(catDate || studioDate),
      sentBack: catSentBack > 0 || studioSentBack > 0 ? 1 : 0,

      catName: get(row, indexes.catName),
      catUserId: get(row, indexes.catUserId),
      catSentBack,
      catCity,
      catMarket,
      catScore: scoreToNumber(get(row, indexes.catScore)),
      catDate,

      studioName: get(row, indexes.studioName),
      studioUserId: get(row, indexes.studioUserId),
      studioSentBack,
      studioCity,
      studioMarket,
      studioScore: scoreToNumber(get(row, indexes.studioScore)),
      studioFormName: get(row, indexes.studioFormName),
      studioDate
    });
  }

  const unique = uniqueTickets(rows);

  console.log(`Parsed ${rows.length} rows, ${unique.length} unique tickets after exclusions.`);
  return unique;
}

function buildSummary(rows) {
  const now = new Date();

  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const weekStart = new Date(today);
  weekStart.setDate(today.getDate() - 6);

  const mtdStart = new Date(today.getFullYear(), today.getMonth(), 1);
  const end = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);

  const inRange = (r, start, finish) => {
    const d = r.dateObj || parseDate(r.date);
    return d && d >= start && d <= finish;
  };

  const weekRows = rows.filter(r => inRange(r, weekStart, end));
  const mtdRows = rows.filter(r => inRange(r, mtdStart, end));

  const metrics = periodRows => {
    const total = periodRows.length;
    const sentBack = periodRows.filter(r => r.sentBack).length;

    return {
      sentBack,
      sentBackRate: total ? sentBack / total * 100 : null,
      totalScore: avg(periodRows.map(r => r.ticketScore)),

      buildOverall: avg(periodRows.filter(r => r.type === 'Build').map(r => r.ticketScore)),
      buildUAE: avg(periodRows.filter(r => r.type === 'Build' && r.market === 'UAE').map(r => r.ticketScore)),
      buildJOR: avg(periodRows.filter(r => r.type === 'Build' && r.market === 'JOR').map(r => r.ticketScore)),

      updateOverall: avg(periodRows.filter(r => r.type === 'Update').map(r => r.ticketScore)),
      updateUAE: avg(periodRows.filter(r => r.type === 'Update' && r.market === 'UAE').map(r => r.ticketScore)),
      updateJOR: avg(periodRows.filter(r => r.type === 'Update' && r.market === 'JOR').map(r => r.ticketScore))
    };
  };

  const week = metrics(weekRows);
  const mtd = metrics(mtdRows);

  return [
    ['Sent back to catalog', `${mtd.sentBack} Tickets - ${fmtPct(mtd.sentBackRate)}`],
    ['Quality score for Week to date', fmtPct(week.totalScore)],
    ['MTD Total Score', fmtPct(mtd.totalScore)],
    ['MTD Build Score', `Overall: ${fmtPct(mtd.buildOverall)} | UAE: ${fmtPct(mtd.buildUAE)} | JOR: ${fmtPct(mtd.buildJOR)}`],
    ['MTD Update Score', `Overall: ${fmtPct(mtd.updateOverall)} | UAE: ${fmtPct(mtd.updateUAE)} | JOR: ${fmtPct(mtd.updateJOR)}`]
  ];
}

async function getSheetsClient() {
  let credentialsText = GOOGLE_SERVICE_ACCOUNT_JSON;

  if (!credentialsText.trim().startsWith('{')) {
    credentialsText = Buffer.from(credentialsText, 'base64').toString('utf8');
  }

  const credentials = JSON.parse(credentialsText);

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });

  return google.sheets({ version: 'v4', auth });
}

async function updateGoogleSheet(rows) {
  console.log('Updating Google Sheet...');

  const sheets = await getSheetsClient();
  const updatedAt = new Date().toISOString();

  const rawHeaders = [
    'Reference ID',
    'Subject',
    'Ticket Score',
    'Type',
    'Market',
    'City',
    'Date',
    'Sent Back',
    'Catalogue Name',
    'Catalogue User Id',
    'Catalogue Sent Back To Catalog',
    'Catalogue City',
    'Catalogue Market',
    'Catalogue Score',
    'Catalogue Date & Time',
    'Studio Name',
    'Studio User Id',
    'Studio Sent Back To Catalog',
    'Studio City',
    'Studio Market',
    'Studio Score',
    'Studio Form Name',
    'Studio Date & Time',
    'Updated At'
  ];

  const rawValues = [
    rawHeaders,
    ...rows.map(r => [
      r.reference,
      r.subject,
      r.ticketScore,
      r.type,
      r.market,
      r.city,
      r.date,
      r.sentBack,
      r.catName,
      r.catUserId,
      r.catSentBack,
      r.catCity,
      r.catMarket,
      r.catScore,
      r.catDate,
      r.studioName,
      r.studioUserId,
      r.studioSentBack,
      r.studioCity,
      r.studioMarket,
      r.studioScore,
      r.studioFormName,
      r.studioDate,
      updatedAt
    ])
  ];

  const summaryValues = buildSummary(rows);

  await sheets.spreadsheets.values.clear({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: `${RAW_SHEET_NAME}!A:Z`
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: `${RAW_SHEET_NAME}!A1`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: rawValues }
  });

  await sheets.spreadsheets.values.clear({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: `${SUMMARY_SHEET_NAME}!A:B`
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: `${SUMMARY_SHEET_NAME}!A1`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: summaryValues }
  });

  console.log(`Updated Raw_iStep and Summary with ${rows.length} tickets.`);
}

async function main() {
  const browser = await chromium.launch({ headless: true });

  const page = await browser.newPage({
    viewport: { width: 1600, height: 1000 },
    acceptDownloads: true
  });

  try {
    await login(page);
    await selectReportAndExport(page);
    const filePath = await downloadLatestExport(page);
    const rows = parseExportFile(filePath);

    if (!rows.length) throw new Error('No rows after parsing export.');

    await updateGoogleSheet(rows);
  } catch (error) {
    console.error('Bot failed:', error.message);

    await page.screenshot({
      path: 'istep-error.png',
      fullPage: true
    }).catch(() => {});

    throw error;
  } finally {
    await browser.close();
  }
}

main();

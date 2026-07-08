import { chromium } from 'playwright';
import { google } from 'googleapis';
import XLSX from 'xlsx';
import fs from 'fs';

process.env.TZ = 'Asia/Amman';

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

for (const [key, value] of Object.entries({
  ISTEP_LOGIN_URL,
  ISTEP_REPORT_URL,
  ISTEP_FILES_URL,
  ISTEP_USERNAME,
  ISTEP_PASSWORD,
  GOOGLE_SHEET_ID,
  GOOGLE_SERVICE_ACCOUNT_JSON
})) {
  if (!value) throw new Error(`Missing required secret/env: ${key}`);
}

function clean(value) {
  return String(value ?? '')
    .replace(/\s+/g, ' ')
    .replace(/\u00a0/g, ' ')
    .trim();
}

function pad2(n) {
  return String(n).padStart(2, '0');
}

function scoreToNumber(value) {
  const n = Number(clean(value).replace('%', ''));
  return Number.isFinite(n) ? n : '';
}

function parseDate(value) {
  const s = clean(value);
  if (!s || s === '-') return null;

  if (/^\d+(\.\d+)?$/.test(s)) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(excelEpoch.getTime() + Number(s) * 24 * 60 * 60 * 1000);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  const d = new Date(s.replace(' ', 'T'));
  return Number.isNaN(d.getTime()) ? null : d;
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

function detectType(subject) {
  const s = clean(subject).toLowerCase();

  if (isExcludedSubject(subject)) return 'Excluded';

  if (
    s.includes('new brand setup') ||
    s.includes('new brand') ||
    s.includes('brand setup')
  ) {
    return 'Build';
  }

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
  ) {
    return 'Update';
  }

  return 'Other';
}

function normalizeMarket(value) {
  const s = clean(value).toLowerCase();

  if (
    s.includes('uae') ||
    s.includes('dubai') ||
    s.includes('abu dhabi') ||
    s.includes('sharjah') ||
    s.includes('al ain')
  ) {
    return 'UAE';
  }

  if (
    s.includes('jor') ||
    s.includes('jordan') ||
    s.includes('amman') ||
    s.includes('irbid') ||
    s.includes('zarqa')
  ) {
    return 'JOR';
  }

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

  for (const row of rows) {
    if (!row.reference) continue;

    if (!map.has(row.reference)) {
      map.set(row.reference, { ...row });
    } else {
      const existing = map.get(row.reference);

      existing.sentBack = existing.sentBack || row.sentBack ? 1 : 0;

      if (!existing.date && row.date) existing.date = row.date;
      if (!existing.dateObj && row.dateObj) existing.dateObj = row.dateObj;
      if (!existing.ticketScore && row.ticketScore) existing.ticketScore = row.ticketScore;
      if (!existing.market && row.market) existing.market = row.market;
      if (!existing.city && row.city) existing.city = row.city;

      map.set(row.reference, existing);
    }
  }

  return [...map.values()];
}

async function clickFirstAvailable(page, selectors, label) {
  for (const selector of selectors) {
    const locator = page.locator(selector).first();

    if (await locator.count()) {
      try {
        if (await locator.isVisible().catch(() => false)) {
          await locator.click({ force: true });
          console.log(`Clicked ${label} using selector: ${selector}`);
          return true;
        }
      } catch (err) {
        console.log(`Could not click ${label} using selector ${selector}: ${err.message}`);
      }
    }
  }

  return false;
}

function getMtdRangeForIstep() {
  const now = new Date();
  const from = new Date(now.getFullYear(), now.getMonth(), 1);
  const to = now;

  return {
    fromText: `${pad2(from.getMonth() + 1)}/${pad2(from.getDate())}/${from.getFullYear()}`,
    toText: `${pad2(to.getMonth() + 1)}/${pad2(to.getDate())}/${to.getFullYear()}`
  };
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
    const input = page.locator(selector).first();

    if (await input.count()) {
      await input.fill(ISTEP_USERNAME);
      filledUser = true;
      break;
    }
  }

  if (!filledUser) throw new Error('Could not find username input.');

  let filledPass = false;

  for (const selector of passwordSelectors) {
    const input = page.locator(selector).first();

    if (await input.count()) {
      await input.fill(ISTEP_PASSWORD);
      filledPass = true;
      break;
    }
  }

  if (!filledPass) throw new Error('Could not find password input.');

  const clickedLogin = await clickFirstAvailable(
    page,
    [
      'button[type="submit"]',
      'input[type="submit"]',
      'button:has-text("Login")',
      'button:has-text("Log in")',
      'button:has-text("Sign in")'
    ],
    'login'
  );

  if (!clickedLogin) await page.keyboard.press('Enter');

  await page.waitForTimeout(5000);
  await page.waitForLoadState('networkidle').catch(() => {});

  console.log('Login completed.');
}

async function setMtdDateRange(page) {
  const { fromText, toText } = getMtdRangeForIstep();

  console.log(`Setting iStep date range to MTD: ${fromText} → ${toText}`);

  let customSelected = false;

  const selects = page.locator('select');
  const selectCount = await selects.count();

  for (let i = 0; i < selectCount; i++) {
    const select = selects.nth(i);

    const options = await select.locator('option').evaluateAll(opts =>
      opts.map(o => ({
        label: (o.textContent || '').trim(),
        value: o.getAttribute('value') || ''
      }))
    ).catch(() => []);

    const hasLast24 = options.some(o => o.label.toLowerCase().includes('last 24 hours'));
    const customOption = options.find(o =>
      o.label.toLowerCase().includes('custom')
    );

    if (hasLast24 && customOption) {
      console.log(`Selecting date range option: ${customOption.label}`);

      try {
        await select.selectOption(customOption.value || { label: customOption.label });
      } catch {
        await select.selectOption({ label: customOption.label });
      }

      await select.dispatchEvent('change').catch(() => {});
      customSelected = true;
      break;
    }
  }

  if (!customSelected) {
    console.log('Could not select Custom Date from native select. Trying visible custom date option...');

    customSelected = await clickFirstAvailable(
      page,
      [
        'button:has-text("Custom Date")',
        'button:has-text("Custom")',
        'a:has-text("Custom Date")',
        'a:has-text("Custom")',
        'text=Custom Date',
        'text=Custom'
      ],
      'Custom Date'
    );
  }

  if (!customSelected) {
    throw new Error('Could not select Custom Date range.');
  }

  await page.waitForTimeout(2000);

  const allDateInputs = page.locator(
    'input[placeholder="mm/dd/yyyy"], input[type="date"], .modal input[type="text"], input[name*="from"], input[name*="to"], input[id*="from"], input[id*="to"]'
  );

  const visibleInputs = [];
  const inputCount = await allDateInputs.count();

  for (let i = 0; i < inputCount; i++) {
    const input = allDateInputs.nth(i);

    const visible = await input.isVisible().catch(() => false);
    const enabled = await input.isEnabled().catch(() => false);

    if (visible && enabled) {
      visibleInputs.push(input);
    }
  }

  if (visibleInputs.length < 2) {
    throw new Error(`Could not find the two visible custom date inputs. Found ${visibleInputs.length}.`);
  }

  const fromInput = visibleInputs[0];
  const toInput = visibleInputs[1];

  console.log('Filling MTD from/to date inputs...');

  await fromInput.click({ force: true });
  await fromInput.press(process.platform === 'darwin' ? 'Meta+A' : 'Control+A').catch(() => {});
  await fromInput.fill(fromText);

  await toInput.click({ force: true });
  await toInput.press(process.platform === 'darwin' ? 'Meta+A' : 'Control+A').catch(() => {});
  await toInput.fill(toText);

  await page.waitForTimeout(1000);

  const submitted = await clickFirstAvailable(
    page,
    [
      'button:has-text("Submit")',
      'input[value="Submit"]',
      'button:has-text("Apply")',
      'input[value="Apply"]',
      'text=Submit'
    ],
    'date Submit'
  );

  if (!submitted) {
    throw new Error('Could not click Submit on custom date modal.');
  }

  await page.waitForTimeout(4000);
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
    const evaluationText = page.locator('text=Evaluation').first();

    if (await evaluationText.count()) {
      await evaluationText.click({ force: true });
    }
  }

  await page.waitForTimeout(1200);

  console.log('Selecting Overall Evaluation Tickets...');

  let selected = false;
  const selects = page.locator('select');
  const selectCount = await selects.count();

  for (let i = 0; i < selectCount; i++) {
    const select = selects.nth(i);
    const options = await select.locator('option').allTextContents().catch(() => []);

    if (options.some(option => clean(option).includes('Overall Evaluation Tickets'))) {
      await select.selectOption({ label: 'Overall Evaluation Tickets' });
      selected = true;
      console.log('Selected Overall Evaluation Tickets using normal select.');
      break;
    }
  }

  if (!selected) {
    const dropdownSelectors = [
      'text=--Select--',
      '.select2-selection',
      '.form-select',
      '.form-control'
    ];

    for (const selector of dropdownSelectors) {
      const dropdown = page.locator(selector).first();

      if (await dropdown.count()) {
        await dropdown.click({ force: true }).catch(() => {});
        await page.waitForTimeout(1000);

        const option = page.locator('text=Overall Evaluation Tickets').first();

        if (await option.count()) {
          await option.click({ force: true });
          selected = true;
          console.log('Selected Overall Evaluation Tickets using dropdown click.');
          break;
        }
      }
    }
  }

  if (!selected) {
    console.log('Could not auto-select Overall Evaluation Tickets. Continuing anyway...');
  }

  await page.waitForTimeout(1500);

  await setMtdDateRange(page);

  console.log('Clicking Apply...');

  const clickedApply = await clickFirstAvailable(
    page,
    [
      'button:has-text("Apply")',
      'input[value="Apply"]',
      'text=Apply'
    ],
    'Apply'
  );

  if (!clickedApply) throw new Error('Apply button not found.');

  console.log('Waiting for report table/export button...');
  await page.waitForTimeout(10000);
  await page.waitForLoadState('networkidle').catch(() => {});

  console.log('Clicking Export...');

  const clickedExport = await clickFirstAvailable(
    page,
    [
      'button:has-text("Export")',
      'a:has-text("Export")',
      'text=Export'
    ],
    'Export'
  );

  if (!clickedExport) throw new Error('Export button not found.');

  console.log('Export clicked. Waiting for file to generate...');
  await page.waitForTimeout(20000);
}

async function downloadLatestExport(page) {
  console.log('Opening File Management...');
  await page.goto(ISTEP_FILES_URL, { waitUntil: 'networkidle' });

  const expectedName = 'Overall_Evaluation_Tickets';
  const maxAttempts = 18;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    console.log(`Checking export file attempt ${attempt}/${maxAttempts}...`);

    await page.waitForTimeout(4000);

    const rows = page.locator('table tbody tr');
    const count = await rows.count();

    for (let i = 0; i < count; i++) {
      const row = rows.nth(i);
      const rowText = clean(await row.innerText().catch(() => ''));

      if (rowText.includes(expectedName)) {
        console.log(`Found export row: ${rowText}`);

        const downloadCandidates = [
          row.locator('a').last(),
          row.locator('button').last(),
          row.locator('i').last()
        ];

        for (const candidate of downloadCandidates) {
          if (await candidate.count()) {
            const downloadPromise = page.waitForEvent('download', { timeout: 60000 });
            await candidate.click({ force: true });

            const download = await downloadPromise;
            const path = './istep-export.xlsx';

            await download.saveAs(path);

            console.log(`Downloaded export to ${path}`);
            return path;
          }
        }

        throw new Error('Found export row but could not click download icon.');
      }
    }

    console.log('File not ready yet. Refreshing File Management...');
    await page.reload({ waitUntil: 'networkidle' }).catch(() => {});
    await page.waitForTimeout(16000);
  }

  throw new Error('Latest Overall Evaluation export did not appear after waiting.');
}

function parseExportFile(filePath) {
  console.log('Parsing Excel export...');

  if (!fs.existsSync(filePath)) {
    throw new Error(`Export file not found: ${filePath}`);
  }

  const fileBuffer = fs.readFileSync(filePath);
  const workbook = XLSX.read(fileBuffer, { type: 'buffer', cellDates: false });
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];

  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: '',
    raw: false
  });

  if (!matrix.length) throw new Error('Excel file is empty.');

  const headerRowIndex = matrix.findIndex(row => {
    const cleaned = row.map(clean);
    return (
      cleaned.includes('Reference ID') &&
      cleaned.includes('Subject') &&
      cleaned.includes('Ticket Score')
    );
  });

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
    console.log('Headers found:', headers.slice(0, 40));
    throw new Error('Required export columns missing.');
  }

  let lastReference = '';
  let lastSubject = '';
  let lastTicketScore = '';

  const get = (row, index) => {
    if (index < 0) return '';
    return clean(row[index]);
  };

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

    const date = catDate || studioDate;

    rows.push({
      reference,
      subject,
      ticketScore: scoreToNumber(ticketScoreRaw),
      type: detectType(subject),
      market: normalizeMarket(catMarket || studioMarket || subject),
      city: catCity || studioCity,
      date,
      dateObj: parseDate(date),
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
  weekStart.setDate(today.getDate() - today.getDay());

  const mtdStart = new Date(today.getFullYear(), today.getMonth(), 1);
  const end = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);

  console.log(`Summary WTD range: ${weekStart.toISOString()} → ${end.toISOString()}`);
  console.log(`Summary MTD range: ${mtdStart.toISOString()} → ${end.toISOString()}`);

  const inRange = (row, start, finish) => {
    const d = row.dateObj || parseDate(row.date);
    return d && d >= start && d <= finish;
  };

  const weekRows = rows.filter(row => inRange(row, weekStart, end));
  const mtdRows = rows.filter(row => inRange(row, mtdStart, end));

  const metrics = periodRows => {
    const total = periodRows.length;
    const sentBack = periodRows.filter(row => row.sentBack).length;

    return {
      total,
      sentBack,
      sentBackRate: total ? sentBack / total * 100 : null,
      totalScore: avg(periodRows.map(row => row.ticketScore)),

      buildOverall: avg(periodRows.filter(row => row.type === 'Build').map(row => row.ticketScore)),
      buildUAE: avg(periodRows.filter(row => row.type === 'Build' && row.market === 'UAE').map(row => row.ticketScore)),
      buildJOR: avg(periodRows.filter(row => row.type === 'Build' && row.market === 'JOR').map(row => row.ticketScore)),

      updateOverall: avg(periodRows.filter(row => row.type === 'Update').map(row => row.ticketScore)),
      updateUAE: avg(periodRows.filter(row => row.type === 'Update' && row.market === 'UAE').map(row => row.ticketScore)),
      updateJOR: avg(periodRows.filter(row => row.type === 'Update' && row.market === 'JOR').map(row => row.ticketScore))
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
    ...rows.map(row => [
      row.reference,
      row.subject,
      row.ticketScore,
      row.type,
      row.market,
      row.city,
      row.date,
      row.sentBack,
      row.catName,
      row.catUserId,
      row.catSentBack,
      row.catCity,
      row.catMarket,
      row.catScore,
      row.catDate,
      row.studioName,
      row.studioUserId,
      row.studioSentBack,
      row.studioCity,
      row.studioMarket,
      row.studioScore,
      row.studioFormName,
      row.studioDate,
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
  const browser = await chromium.launch({
    headless: true
  });

  const page = await browser.newPage({
    viewport: {
      width: 1600,
      height: 1000
    },
    acceptDownloads: true
  });

  try {
    await login(page);
    await selectReportAndExport(page);

    const filePath = await downloadLatestExport(page);
    const rows = parseExportFile(filePath);

    if (!rows.length) {
      throw new Error('No rows after parsing export.');
    }

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

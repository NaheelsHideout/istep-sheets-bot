import { chromium } from 'playwright';
import { google } from 'googleapis';

const {
  ISTEP_LOGIN_URL,
  ISTEP_REPORT_URL,
  ISTEP_USERNAME,
  ISTEP_PASSWORD,
  GOOGLE_SHEET_ID,
  GOOGLE_SERVICE_ACCOUNT_JSON
} = process.env;

const RAW_SHEET_NAME = 'Raw_iStep';

function required(name, value) {
  if (!value) throw new Error(`Missing required secret/env: ${name}`);
}

required('ISTEP_LOGIN_URL', ISTEP_LOGIN_URL);
required('ISTEP_REPORT_URL', ISTEP_REPORT_URL);
required('ISTEP_USERNAME', ISTEP_USERNAME);
required('ISTEP_PASSWORD', ISTEP_PASSWORD);
required('GOOGLE_SHEET_ID', GOOGLE_SHEET_ID);
required('GOOGLE_SERVICE_ACCOUNT_JSON', GOOGLE_SERVICE_ACCOUNT_JSON);

function clean(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .replace(/\u00a0/g, ' ')
    .trim();
}

function scoreToNumber(value) {
  const n = Number(String(value || '').replace('%', '').trim());
  return Number.isFinite(n) ? n : '';
}

function detectType(subject) {
  const s = String(subject || '').toLowerCase();

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
  const s = String(value || '').toLowerCase();

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

  let usernameFilled = false;

  for (const selector of usernameSelectors) {
    const locator = page.locator(selector).first();

    if (await locator.count()) {
      await locator.fill(ISTEP_USERNAME);
      usernameFilled = true;
      break;
    }
  }

  if (!usernameFilled) {
    throw new Error('Could not find username/email input on iStep login page.');
  }

  let passwordFilled = false;

  for (const selector of passwordSelectors) {
    const locator = page.locator(selector).first();

    if (await locator.count()) {
      await locator.fill(ISTEP_PASSWORD);
      passwordFilled = true;
      break;
    }
  }

  if (!passwordFilled) {
    throw new Error('Could not find password input on iStep login page.');
  }

  const loginButtons = [
    'button[type="submit"]',
    'input[type="submit"]',
    'button:has-text("Login")',
    'button:has-text("Sign in")',
    'button:has-text("Log in")'
  ];

  let clicked = false;

  for (const selector of loginButtons) {
    const locator = page.locator(selector).first();

    if (await locator.count()) {
      await locator.click({ force: true });
      clicked = true;
      break;
    }
  }

  if (!clicked) {
    await page.keyboard.press('Enter');
  }

  await page.waitForTimeout(5000);
  await page.waitForLoadState('networkidle').catch(() => {});

  console.log('Login step completed.');
}

async function openReport(page) {
  console.log('Opening iStep report page...');
  await page.goto(ISTEP_REPORT_URL, { waitUntil: 'networkidle' });
  await page.waitForTimeout(4000);

  console.log('Selecting Evaluation type...');

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

  await page.waitForTimeout(1500);

  console.log('Selecting Overall Evaluation Tickets...');

  const selects = page.locator('select');
  const selectCount = await selects.count();

  let selectedReport = false;

  for (let i = 0; i < selectCount; i++) {
    const select = selects.nth(i);
    const options = await select.locator('option').allTextContents().catch(() => []);

    if (options.some(o => clean(o).includes('Overall Evaluation Tickets'))) {
      await select.selectOption({ label: 'Overall Evaluation Tickets' });
      selectedReport = true;
      console.log('Selected report using normal select.');
      break;
    }
  }

  if (!selectedReport) {
    const dropdownSelectors = [
      'select',
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
          selectedReport = true;
          console.log('Selected report using dropdown click.');
          break;
        }
      }
    }
  }

  if (!selectedReport) {
    console.log('Could not auto-select Overall Evaluation Tickets. Continuing anyway...');
  }

  await page.waitForTimeout(1500);

  console.log('Clicking Apply...');

  const applySelectors = [
    'button:has-text("Apply")',
    'input[value="Apply"]',
    'text=Apply'
  ];

  let clickedApply = false;

  for (const selector of applySelectors) {
    const applyButton = page.locator(selector).first();

    if (await applyButton.count()) {
      await applyButton.click({ force: true });
      clickedApply = true;
      break;
    }
  }

  if (!clickedApply) {
    throw new Error('Could not find Apply button.');
  }

  console.log('Waiting for report table to load...');

  await page.waitForTimeout(10000);
  await page.waitForLoadState('networkidle').catch(() => {});

  const table = page.locator('table#example, table').first();
  await table.waitFor({ timeout: 60000 });

  console.log('Report table found.');
}

async function extractTable(page) {
  console.log('Extracting table...');

  const result = await page.evaluate(() => {
    function cellText(cell) {
      return (cell.innerText || '')
        .replace(/\s+/g, ' ')
        .replace(/\u00a0/g, ' ')
        .trim();
    }

    const table = document.querySelector('table#example') || document.querySelector('table');

    if (!table) {
      throw new Error('No table found.');
    }

    const headerCells = Array.from(table.querySelectorAll('thead th')).map(cellText);

    const bodyRows = Array.from(table.querySelectorAll('tbody tr'));
    const grid = [];
    const rowspans = {};

    bodyRows.forEach((tr) => {
      const row = [];
      let colIndex = 0;

      while (rowspans[colIndex] && rowspans[colIndex].remaining > 0) {
        row[colIndex] = rowspans[colIndex].value;
        rowspans[colIndex].remaining -= 1;

        if (rowspans[colIndex].remaining === 0) {
          delete rowspans[colIndex];
        }

        colIndex++;
      }

      const cells = Array.from(tr.querySelectorAll('td'));

      cells.forEach(td => {
        while (row[colIndex] !== undefined) colIndex++;

        const value = cellText(td);
        const rowspan = Number(td.getAttribute('rowspan') || '1');

        row[colIndex] = value;

        if (rowspan > 1) {
          rowspans[colIndex] = {
            value,
            remaining: rowspan - 1
          };
        }

        colIndex++;
      });

      grid.push(row);
    });

    return {
      headers: headerCells,
      rows: grid
    };
  });

  const headers = result.headers.map(clean);
  const rows = result.rows;

  console.log('Headers found:', headers.slice(0, 25));

  const indexOf = (name) => {
    const wanted = name.toLowerCase();
    return headers.findIndex(h => h.toLowerCase() === wanted);
  };

  const idx = {
    reference: indexOf('Reference ID'),
    subject: indexOf('Subject'),
    ticketScore: indexOf('Ticket Score'),

    catName: indexOf('Catalogue Name'),
    catUserId: indexOf('Catalogue User Id'),
    catSentBack: indexOf('Catalogue Sent Back To Catalog'),
    catCity: indexOf('Catalogue City'),
    catMarket: indexOf('Catalogue Market'),
    catScore: indexOf('Catalogue Score'),
    catDate: indexOf('Catalogue Date & Time'),

    studioName: indexOf('Studio Name'),
    studioUserId: indexOf('Studio User Id'),
    studioSentBack: indexOf('Studio Sent Back To Catalog'),
    studioCity: indexOf('Studio City'),
    studioMarket: indexOf('Studio Market'),
    studioScore: indexOf('Studio Score'),
    studioFormName: indexOf('Studio Form Name'),
    studioDate: indexOf('Studio Date & Time')
  };

  if (idx.reference === -1 || idx.subject === -1 || idx.ticketScore === -1) {
    throw new Error('Could not find required columns: Reference ID, Subject, Ticket Score.');
  }

  const get = (row, index) => {
    if (index === -1 || index === undefined) return '';
    return clean(row[index]);
  };

  const output = rows
    .map(r => {
      const reference = get(r, idx.reference);
      const subject = get(r, idx.subject);

      const catMarket = get(r, idx.catMarket);
      const studioMarket = get(r, idx.studioMarket);
      const catCity = get(r, idx.catCity);
      const studioCity = get(r, idx.studioCity);

      const market = normalizeMarket(catMarket || studioMarket || subject);
      const type = detectType(subject);

      const catSentBack = Number(get(r, idx.catSentBack) || 0) || 0;
      const studioSentBack = Number(get(r, idx.studioSentBack) || 0) || 0;

      const catDate = get(r, idx.catDate);
      const studioDate = get(r, idx.studioDate);

      return {
        reference,
        subject,
        ticketScore: scoreToNumber(get(r, idx.ticketScore)),
        type,
        market,
        city: catCity || studioCity,
        date: catDate || studioDate,
        sentBack: catSentBack > 0 || studioSentBack > 0 ? 1 : 0,

        catName: get(r, idx.catName),
        catUserId: get(r, idx.catUserId),
        catSentBack,
        catCity,
        catMarket,
        catScore: scoreToNumber(get(r, idx.catScore)),
        catDate,

        studioName: get(r, idx.studioName),
        studioUserId: get(r, idx.studioUserId),
        studioSentBack,
        studioCity,
        studioMarket,
        studioScore: scoreToNumber(get(r, idx.studioScore)),
        studioFormName: get(r, idx.studioFormName),
        studioDate
      };
    })
    .filter(r => r.reference && r.reference.startsWith('TH-'));

  console.log(`Extracted ${output.length} table rows.`);

  return output;
}

function uniqueTickets(rows) {
  const map = new Map();

  for (const row of rows) {
    if (!map.has(row.reference)) {
      map.set(row.reference, { ...row });
    } else {
      const existing = map.get(row.reference);

      existing.sentBack = existing.sentBack || row.sentBack ? 1 : 0;

      if (!existing.date && row.date) existing.date = row.date;
      if (!existing.ticketScore && row.ticketScore) existing.ticketScore = row.ticketScore;
      if (!existing.market && row.market) existing.market = row.market;
      if (!existing.city && row.city) existing.city = row.city;

      map.set(row.reference, existing);
    }
  }

  return Array.from(map.values());
}

async function updateGoogleSheet(rows) {
  console.log('Updating Google Sheet...');

  let credentialsText = GOOGLE_SERVICE_ACCOUNT_JSON;

  if (!credentialsText.trim().startsWith('{')) {
    credentialsText = Buffer.from(credentialsText, 'base64').toString('utf8');
  }

  const credentials = JSON.parse(credentialsText);

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });

  const sheets = google.sheets({
    version: 'v4',
    auth
  });

  const headers = [
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

  const now = new Date().toISOString();
  const unique = uniqueTickets(rows);

  const values = [
    headers,
    ...unique.map(r => [
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
      now
    ])
  ];

  await sheets.spreadsheets.values.clear({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: `${RAW_SHEET_NAME}!A:Z`
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: `${RAW_SHEET_NAME}!A1`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values
    }
  });

  console.log(`Google Sheet updated with ${unique.length} unique tickets.`);
}

async function main() {
  const browser = await chromium.launch({
    headless: true
  });

  const page = await browser.newPage({
    viewport: {
      width: 1440,
      height: 1000
    }
  });

  try {
    await login(page);
    await openReport(page);

    const rows = await extractTable(page);

    if (!rows.length) {
      throw new Error('No iStep rows extracted.');
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

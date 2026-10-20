import { google } from 'googleapis';
import { mapMasterData } from './workbook.js';

const SHEET_SCOPE = 'https://www.googleapis.com/auth/spreadsheets';
const normalize = (value) => String(value ?? '').trim().toLowerCase();

function parseDateFormula(value) {
  const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return match ? `=DATE(${Number(match[1])},${Number(match[2])},${Number(match[3])})` : '';
}

function columnLetters(index) {
  let letters = '';
  while (index > 0) {
    const remainder = (index - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    index = Math.floor((index - 1) / 26);
  }
  return letters;
}

function mappedClient(headers, values) {
  const record = Object.fromEntries(headers.map((header, index) => [header, values[index] ?? '']));
  const client = mapMasterData(record);
  for (const field of ['id', 'name', 'email', 'phone', 'translator', 'frequency', 'outcome', 'paymentStatus', 'notes']) {
    client[field] = String(client[field] ?? '');
  }
  client.daysRemaining = String(client.daysRemaining ?? '');
  return client;
}

export function createGoogleSheetStore({ spreadsheetId, sheetName = 'MasterData', serviceAccountEmail, privateKey }, api = google) {
  const quotedSheet = `'${sheetName.replaceAll("'", "''")}'`;
  const range = `${quotedSheet}!A:P`;
  let queuedWrite = Promise.resolve();

  function getSheetsClient() {
    if (!spreadsheetId || !serviceAccountEmail || !privateKey) {
      throw new Error('Google Sheets requires GOOGLE_SHEET_ID, GOOGLE_SERVICE_ACCOUNT_EMAIL, and GOOGLE_PRIVATE_KEY.');
    }
    const auth = new api.auth.JWT({
      email: serviceAccountEmail,
      key: privateKey.replace(/\\n/g, '\n'),
      scopes: [SHEET_SCOPE],
    });
    return api.sheets({ version: 'v4', auth });
  }

  async function getRows() {
    const sheets = getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'SERIAL_NUMBER',
    });
    const values = response.data.values || [];
    if (!values.length) throw new Error(`The ${sheetName} tab is empty or unavailable.`);
    return { sheets, headers: values[0].map((header) => String(header ?? '')), rows: values.slice(1) };
  }

  async function list() {
    const { headers, rows } = await getRows();
    return rows
      .map((row) => mappedClient(headers, row))
      .filter((client) => client.name)
      .sort((left, right) => left.name.localeCompare(right.name));
  }

  async function appendClient(input) {
    const { sheets, headers, rows } = await getRows();
    const email = String(input.email ?? '').trim();
    const clients = rows.map((row) => mappedClient(headers, row));
    if (email && clients.some((client) => normalize(client.email) === normalize(email))) {
      throw new Error('A client with this email already exists.');
    }

    const nextId = clients.reduce((largest, client) => {
      const match = client.id.match(/^CL-(\d+)$/i);
      return match ? Math.max(largest, Number(match[1])) : largest;
    }, 0) + 1;
    const clientId = `CL-${String(nextId).padStart(3, '0')}`;
    const rowNumber = rows.length + 2;
    const row = Array(headers.length).fill('');
    const indexes = new Map(headers.map((header, index) => [normalize(header), index]));
    const put = (header, value, isFormula = false) => {
      const index = indexes.get(normalize(header));
      if (index === undefined) return;
      row[index] = !isFormula && typeof value === 'string' && /^[=+\-@]/.test(value) ? `'${value}` : value;
    };
    const putFormula = (header, value) => put(header, value, true);
    const dateFormula = parseDateFormula;

    put('ID No.', clientId);
    put('Client Name', String(input.name ?? '').trim());
    put('Email', email);
    put('PHONE', String(input.phone ?? '').trim());
    put('Translator Needed', input.translator || 'No');
    putFormula('Date Onboarded', dateFormula(input.onboardingDate) || '=TODAY()');
    put('Meeting Frequency', input.frequency || '');
    putFormula('Date of Interview', dateFormula(input.interviewDate));
    putFormula('Date Contacted', dateFormula(input.lastContactDate));
    put('Outcome', input.outcome || '');
    put('Outcome ', input.outcome || '');
    putFormula('Next date of Contact', dateFormula(input.nextContactDate));
    putFormula('Last Date With Attorney', dateFormula(input.lastAttorneyDate));
    put('Payment Status', input.paymentStatus || 'Pending');
    putFormula('Todays Date', '=TODAY()');
    putFormula('Days Remaining', `=IF(K${rowNumber}="","",K${rowNumber}-N${rowNumber})`);
    put('NOTES', input.notes || '');

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [row] },
    });

    return {
      id: clientId,
      name: String(input.name ?? '').trim(),
      email,
      phone: String(input.phone ?? '').trim(),
      translator: input.translator || 'No',
      onboardingDate: input.onboardingDate || new Date().toISOString().slice(0, 10),
      frequency: input.frequency || '',
      interviewDate: input.interviewDate || '',
      lastContactDate: input.lastContactDate || '',
      outcome: input.outcome || '',
      nextContactDate: input.nextContactDate || '',
      lastAttorneyDate: input.lastAttorneyDate || '',
      paymentStatus: input.paymentStatus || 'Pending',
      notes: input.notes || '',
      daysRemaining: input.nextContactDate
        ? Math.round((new Date(`${input.nextContactDate}T00:00:00Z`) - new Date(`${new Date().toISOString().slice(0, 10)}T00:00:00Z`)) / 86400000)
        : '',
    };
  }

  async function completeFollowUp(clientId) {
    const { sheets, headers, rows } = await getRows();
    const idIndex = headers.findIndex((header) => normalize(header) === normalize('ID No.'));
    const nextContactIndex = headers.findIndex((header) => normalize(header) === normalize('Next date of Contact'));
    const rowIndex = rows.findIndex((row) => String(row[idIndex] ?? '') === clientId);
    if (rowIndex < 0) throw new Error('Client not found.');
    const row = rows[rowIndex];
    const scheduledDate = String(row[nextContactIndex] ?? '').trim();
    if (!scheduledDate) throw new Error('This client has no scheduled follow-up to complete.');
    if (normalize(scheduledDate) === 'done') throw new Error('This follow-up is already marked done.');

    const rowNumber = rowIndex + 2;
    const cell = `${columnLetters(nextContactIndex + 1)}${rowNumber}`;
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${quotedSheet}!${cell}`,
      valueInputOption: 'RAW',
      requestBody: { values: [['Done']] },
    });
    row[nextContactIndex] = 'Done';
    return mappedClient(headers, row);
  }

  return {
    list,
    add(input) {
      const operation = queuedWrite.then(() => appendClient(input));
      queuedWrite = operation.catch(() => {});
      return operation;
    },
    completeFollowUp(clientId) {
      const operation = queuedWrite.then(() => completeFollowUp(clientId));
      queuedWrite = operation.catch(() => {});
      return operation;
    },
  };
}

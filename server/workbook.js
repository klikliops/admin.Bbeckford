import ExcelJS from 'exceljs';
import { readFile, rename, unlink, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import xmlDom from '@xmldom/xmldom';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

const { DOMParser, XMLSerializer } = xmlDom;
const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const DOC_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const HEADER_ROW = 1;
const MASTER_SHEET = 'MasterData';
let writeQueue = Promise.resolve();

const normalize = (value) => String(value ?? '').trim().toLowerCase();
const cellText = (cell) => {
  if (!cell || cell.value == null) return '';
  if (cell.value instanceof Date) return cell.value.toISOString().slice(0, 10);
  if (typeof cell.value === 'object' && 'result' in cell.value) {
    const result = cell.value.result;
    if (result instanceof Date) return result.toISOString().slice(0, 10);
    return String(result ?? '').trim();
  }
  return String(cell.text ?? cell.value).trim();
};

function rowsAsObjects(sheet) {
  if (!sheet) return [];
  const headers = new Map();
  sheet.getRow(HEADER_ROW).eachCell({ includeEmpty: false }, (cell, column) => {
    const header = cellText(cell);
    if (header) headers.set(column, header);
  });

  const results = [];
  for (let rowNumber = HEADER_ROW + 1; rowNumber <= sheet.actualRowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const record = {};
    for (const [column, header] of headers) {
      record[header] = cellText(row.getCell(column));
    }
    if (Object.values(record).some(Boolean)) results.push(record);
  }
  return results;
}

function normalizeDate(value) {
  if (!value) return '';
  const normalized = String(value).trim();
  const serial = Number(normalized);
  if (/^\d{5}(?:\.\d+)?$/.test(normalized) && serial > 20000 && serial < 80000) {
    return new Date(Date.UTC(1899, 11, 30) + serial * 86400000).toISOString().slice(0, 10);
  }
  return normalized;
}

export function mapMasterData(row) {
  return {
    id: row['ID No.'] || '',
    name: row['Client Name'] || '',
    email: row.Email || '',
    phone: row.PHONE || '',
    translator: row['Translator Needed'] || '',
    onboardingDate: normalizeDate(row['Date Onboarded']),
    frequency: row['Meeting Frequency'] || '',
    interviewDate: normalizeDate(row['Date of Interview']),
    lastContactDate: normalizeDate(row['Date Contacted']),
    outcome: row['Outcome'] || row['Outcome '] || '',
    nextContactDate: normalizeDate(row['Next date of Contact']),
    lastAttorneyDate: normalizeDate(row['Last Date With Attorney']),
    paymentStatus: row['Payment Status'] || '',
    todayDate: normalizeDate(row['Todays Date']),
    daysRemaining: row['Days Remaining '] || row['Days Remaining'] || '',
    notes: row.NOTES || row.Notes || '',
  };
}

async function loadWorkbook(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  return workbook;
}

export async function readClients(filePath) {
  const workbook = await loadWorkbook(filePath);
  const clients = rowsAsObjects(workbook.getWorksheet(MASTER_SHEET))
    .map(mapMasterData)
    .filter((client) => client.name);
  return clients.sort((left, right) => left.name.localeCompare(right.name));
}

function columnName(index) {
  let name = '';
  while (index > 0) {
    const remainder = (index - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    index = Math.floor((index - 1) / 26);
  }
  return name;
}

function findColumn(sheet, ...labels) {
  const wanted = new Set(labels.map(normalize));
  let result = 0;
  sheet.getRow(HEADER_ROW).eachCell({ includeEmpty: false }, (cell, column) => {
    if (wanted.has(normalize(cellText(cell)))) result = column;
  });
  return result;
}

function nextClientId(sheet) {
  let largest = 0;
  for (let rowNumber = HEADER_ROW + 1; rowNumber <= sheet.actualRowCount; rowNumber += 1) {
    const value = cellText(sheet.getRow(rowNumber).getCell(findColumn(sheet, 'ID No.')));
    const match = value.match(/^CL-(\d+)$/i);
    if (match) largest = Math.max(largest, Number(match[1]));
  }
  return `CL-${String(largest + 1).padStart(3, '0')}`;
}

function excelSerial(value) {
  return Math.floor((Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate()) - Date.UTC(1899, 11, 30)) / 86400000);
}

function dateFromInput(value) {
  if (!value) return null;
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return match ? new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3]))) : null;
}

function makeCell(document, column, rowNumber, value, style) {
  const cell = document.createElementNS(MAIN_NS, 'c');
  cell.setAttribute('r', `${columnName(column)}${rowNumber}`);
  if (style) cell.setAttribute('s', style);
  if (value instanceof Date) {
    const stored = document.createElementNS(MAIN_NS, 'v');
    stored.appendChild(document.createTextNode(String(excelSerial(value))));
    cell.appendChild(stored);
    return cell;
  }
  if (value && typeof value === 'object' && 'formula' in value) {
    const formula = document.createElementNS(MAIN_NS, 'f');
    formula.appendChild(document.createTextNode(value.formula));
    const result = document.createElementNS(MAIN_NS, 'v');
    result.appendChild(document.createTextNode(String(value.result ?? '')));
    cell.appendChild(formula);
    cell.appendChild(result);
    return cell;
  }
  cell.setAttribute('t', 'inlineStr');
  const inline = document.createElementNS(MAIN_NS, 'is');
  const text = document.createElementNS(MAIN_NS, 't');
  text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
  text.appendChild(document.createTextNode(String(value)));
  inline.appendChild(text);
  cell.appendChild(inline);
  return cell;
}

function appendWorksheetRow(zip, worksheetPath, rowValues) {
  const document = new DOMParser().parseFromString(strFromU8(zip[worksheetPath]), 'application/xml');
  const sheetData = document.getElementsByTagNameNS(MAIN_NS, 'sheetData').item(0);
  if (!sheetData) throw new Error(`The worksheet data section is missing in ${worksheetPath}.`);
  const rowNumber = rowValues.rowNumber;
  const values = rowValues.values;
  const rows = sheetData.getElementsByTagNameNS(MAIN_NS, 'row');
  let row;
  let previousRow;
  let nextRow;
  for (let index = 0; index < rows.length; index += 1) {
    const currentRow = rows.item(index);
    const currentNumber = Number(currentRow.getAttribute('r'));
    if (currentNumber === rowNumber) row = currentRow;
    if (currentNumber === rowNumber - 1) previousRow = currentRow;
    if (!nextRow && currentNumber > rowNumber) nextRow = currentRow;
  }
  if (!row) {
    row = document.createElementNS(MAIN_NS, 'row');
    row.setAttribute('r', String(rowNumber));
    sheetData.insertBefore(row, nextRow);
  }

  for (const [column, value] of values) {
    if (value !== '' && value != null) {
      const reference = `${columnName(column)}${rowNumber - 1}`;
      const currentReference = `${columnName(column)}${rowNumber}`;
      const existingCells = row.getElementsByTagNameNS(MAIN_NS, 'c');
      let existingCell;
      let priorCell;
      for (let index = 0; index < existingCells.length; index += 1) {
        if (existingCells.item(index).getAttribute('r') === currentReference) existingCell = existingCells.item(index);
      }
      const priorCells = previousRow?.getElementsByTagNameNS(MAIN_NS, 'c');
      if (priorCells) {
        for (let index = 0; index < priorCells.length; index += 1) {
          if (priorCells.item(index).getAttribute('r') === reference) priorCell = priorCells.item(index);
        }
      }
      const style = existingCell?.getAttribute('s') || priorCell?.getAttribute('s');
      const replacement = makeCell(document, column, rowNumber, value, style);
      if (existingCell) row.replaceChild(replacement, existingCell);
      else row.appendChild(replacement);
    }
  }
  const dimension = document.getElementsByTagNameNS(MAIN_NS, 'dimension').item(0);
  if (dimension) {
    const ref = dimension.getAttribute('ref');
    const match = ref.match(/(\$?[A-Z]+\$?)(\d+)$/);
    if (match && Number(match[2]) < rowNumber) dimension.setAttribute('ref', `${match[1]}${rowNumber}`);
  }
  const validations = document.getElementsByTagNameNS(MAIN_NS, 'dataValidation');
  for (let index = 0; index < validations.length; index += 1) {
    const validation = validations.item(index);
    const ranges = validation.getAttribute('sqref').split(/\s+/).filter(Boolean);
    const updatedRanges = ranges.map((range) => range.replace(/(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?)(\d+)$/, (match, prefix, endRow) =>
      Number(endRow) === rowNumber - 1 ? `${prefix}${rowNumber}` : match,
    ));
    const formula = validation.getElementsByTagNameNS(MAIN_NS, 'formula1').item(0)?.textContent || '';
    if (validation.getAttribute('type') === 'custom' && formula.includes('DATEVALUE')) {
      for (const column of [6, 8, 9, 11, 12]) {
        const reference = `${columnName(column)}${rowNumber}`;
        if (!updatedRanges.some((range) => range.includes(reference))) updatedRanges.push(reference);
      }
    }
    validation.setAttribute('sqref', updatedRanges.join(' '));
  }
  zip[worksheetPath] = strToU8(new XMLSerializer().serializeToString(document));
}

function makeEmptyCell(document, column, rowNumber, style) {
  const cell = document.createElementNS(MAIN_NS, 'c');
  cell.setAttribute('r', `${columnName(column)}${rowNumber}`);
  if (style) cell.setAttribute('s', style);
  return cell;
}

function updateWorksheetCells(zip, worksheetPath, rowNumber, values) {
  const document = new DOMParser().parseFromString(strFromU8(zip[worksheetPath]), 'application/xml');
  const sheetData = document.getElementsByTagNameNS(MAIN_NS, 'sheetData').item(0);
  const rows = sheetData?.getElementsByTagNameNS(MAIN_NS, 'row');
  let row;
  let previousRow;
  for (let index = 0; rows && index < rows.length; index += 1) {
    const currentRow = rows.item(index);
    const currentNumber = Number(currentRow.getAttribute('r'));
    if (currentNumber === rowNumber) row = currentRow;
    if (currentNumber === rowNumber - 1) previousRow = currentRow;
  }
  if (!row) throw new Error(`Client row ${rowNumber} was not found in MasterData.`);

  for (const [column, value] of values) {
    const reference = `${columnName(column)}${rowNumber}`;
    const existingCells = row.getElementsByTagNameNS(MAIN_NS, 'c');
    let existingCell;
    let style;
    for (let index = 0; index < existingCells.length; index += 1) {
      if (existingCells.item(index).getAttribute('r') === reference) existingCell = existingCells.item(index);
    }
    style = existingCell?.getAttribute('s');
    if (!style) {
      const previousCells = previousRow?.getElementsByTagNameNS(MAIN_NS, 'c');
      for (let index = 0; previousCells && index < previousCells.length; index += 1) {
        if (previousCells.item(index).getAttribute('r') === `${columnName(column)}${rowNumber - 1}`) {
          style = previousCells.item(index).getAttribute('s');
        }
      }
    }

    const replacement = value === ''
      ? makeEmptyCell(document, column, rowNumber, style)
      : makeCell(document, column, rowNumber, value, style);
    if (existingCell) {
      row.replaceChild(replacement, existingCell);
      continue;
    }

    const cells = row.getElementsByTagNameNS(MAIN_NS, 'c');
    let nextCell;
    for (let index = 0; index < cells.length; index += 1) {
      const currentColumn = Number(cells.item(index).getAttribute('r').match(/^[A-Z]+/)?.[0]
        .split('').reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0));
      if (currentColumn > column) {
        nextCell = cells.item(index);
        break;
      }
    }
    row.insertBefore(replacement, nextCell);
  }
  zip[worksheetPath] = strToU8(new XMLSerializer().serializeToString(document));
}

async function replaceWorkbookFile(filePath, archive) {
  const temporaryPath = join(dirname(filePath), `.case-management-${process.pid}-${Date.now()}.xlsx`);
  await writeFile(temporaryPath, zipSync(archive, { level: 6 }));
  try {
    await rename(temporaryPath, filePath);
  } catch (error) {
    await unlink(temporaryPath).catch(() => {});
    if (error.code === 'EPERM' || error.code === 'EACCES') {
      throw new Error('The workbook is locked by Excel or another process. Close it, then try again.');
    }
    throw error;
  }
}

function worksheetPaths(zip) {
  const parser = new DOMParser();
  const workbook = parser.parseFromString(strFromU8(zip['xl/workbook.xml']), 'application/xml');
  const relationships = parser.parseFromString(strFromU8(zip['xl/_rels/workbook.xml.rels']), 'application/xml');
  const targets = new Map();
  const relationshipNodes = relationships.getElementsByTagNameNS(PACKAGE_REL_NS, 'Relationship');
  for (let index = 0; index < relationshipNodes.length; index += 1) {
    const relation = relationshipNodes.item(index);
    targets.set(relation.getAttribute('Id'), relation.getAttribute('Target'));
  }
  const result = new Map();
  const sheets = workbook.getElementsByTagNameNS(MAIN_NS, 'sheet');
  for (let index = 0; index < sheets.length; index += 1) {
    const sheet = sheets.item(index);
    const target = targets.get(sheet.getAttributeNS(DOC_REL_NS, 'id'));
    if (!target) continue;
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
    result.set(sheet.getAttribute('name'), path.replace(/^xl\/\.\//, 'xl/'));
  }
  return result;
}

function appendByHeaders(sheet, values) {
  return Object.entries(values)
    .map(([header, value]) => [findColumn(sheet, header), value])
    .filter(([column, value]) => column && value !== '' && value != null);
}

async function addClientToWorkbook(filePath, input) {
  const workbook = await loadWorkbook(filePath);
  const master = workbook.getWorksheet(MASTER_SHEET);
  if (!master) throw new Error('The workbook must contain a MasterData sheet.');

  const email = String(input.email ?? '').trim();
  if (email && (await readClients(filePath)).some((client) => normalize(client.email) === normalize(email))) {
    throw new Error('A client with this email already exists.');
  }

  const clientId = nextClientId(master);
  const now = new Date();
  const createdOn = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
  const rowNumber = master.actualRowCount + 1;
  const nextContact = dateFromInput(input.nextContactDate);
  const daysUntilNextContact = nextContact ? Math.round((nextContact - createdOn) / 86400000) : '';
  const masterRow = appendByHeaders(master, {
    'ID No.': clientId,
    'Client Name': String(input.name ?? '').trim(),
    Email: email,
    PHONE: String(input.phone ?? '').trim(),
    'Translator Needed': input.translator,
    'Date Onboarded': dateFromInput(input.onboardingDate) || createdOn,
    'Meeting Frequency': input.frequency,
    'Date of Interview': dateFromInput(input.interviewDate),
    'Date Contacted': dateFromInput(input.lastContactDate),
    'Outcome ': input.outcome,
    'Next date of Contact': nextContact,
    'Last Date With Attorney': dateFromInput(input.lastAttorneyDate),
    'Payment Status': input.paymentStatus,
    NOTES: input.notes,
    'Todays Date': { formula: 'TODAY()', result: excelSerial(createdOn) },
    'Days Remaining ': { formula: `K${rowNumber}-N${rowNumber}`, result: daysUntilNextContact },
  });
  const archive = unzipSync(new Uint8Array(await readFile(filePath)));
  const paths = worksheetPaths(archive);
  const masterPath = paths.get(MASTER_SHEET);
  if (!masterPath) throw new Error('Could not find the MasterData worksheet in the workbook.');

  appendWorksheetRow(archive, masterPath, { rowNumber, values: masterRow });
  const temporaryPath = join(dirname(filePath), `.case-management-${process.pid}.xlsx`);
  await writeFile(temporaryPath, zipSync(archive, { level: 6 }));
  try {
    await rename(temporaryPath, filePath);
  } catch (error) {
    await unlink(temporaryPath).catch(() => {});
    if (error.code === 'EPERM' || error.code === 'EACCES') {
      throw new Error('The workbook is locked by Excel or another process. Close it, then submit again.');
    }
    throw error;
  }

  return {
    id: clientId,
    name: String(input.name ?? '').trim(),
    email,
    phone: String(input.phone ?? '').trim(),
    translator: input.translator || '',
    onboardingDate: (dateFromInput(input.onboardingDate) || createdOn).toISOString().slice(0, 10),
    frequency: input.frequency || '',
    interviewDate: input.interviewDate || '',
    lastContactDate: input.lastContactDate || '',
    outcome: input.outcome || '',
    nextContactDate: input.nextContactDate || '',
    lastAttorneyDate: input.lastAttorneyDate || '',
    paymentStatus: input.paymentStatus || '',
    notes: input.notes || '',
    daysRemaining: daysUntilNextContact,
  };
}

async function updateClientInWorkbook(filePath, clientId, input) {
  const workbook = await loadWorkbook(filePath);
  const master = workbook.getWorksheet(MASTER_SHEET);
  if (!master) throw new Error('The workbook must contain a MasterData sheet.');
  const idColumn = findColumn(master, 'ID No.');
  let rowNumber = 0;
  for (let currentRow = HEADER_ROW + 1; currentRow <= master.actualRowCount; currentRow += 1) {
    if (cellText(master.getRow(currentRow).getCell(idColumn)) === clientId) rowNumber = currentRow;
  }
  if (!rowNumber) throw new Error('Client not found.');

  const email = String(input.email ?? '').trim();
  if (email && (await readClients(filePath)).some((client) => client.id !== clientId && normalize(client.email) === normalize(email))) {
    throw new Error('A client with this email already exists.');
  }
  const nextContact = String(input.nextContactDate ?? '').trim();
  const nextContactValue = normalize(nextContact) === 'done' ? 'done' : dateFromInput(nextContact) || '';
  if (nextContact && !nextContactValue) throw new Error('Next contact must be a date or done.');

  const values = [
    ['Client Name', String(input.name ?? '').trim()],
    ['Email', email],
    ['PHONE', String(input.phone ?? '').trim()],
    ['Translator Needed', String(input.translator ?? '').trim()],
    ['Date Onboarded', dateFromInput(input.onboardingDate) || ''],
    ['Meeting Frequency', String(input.frequency ?? '').trim()],
    ['Date of Interview', dateFromInput(input.interviewDate) || ''],
    ['Date Contacted', dateFromInput(input.lastContactDate) || ''],
    ['Outcome ', String(input.outcome ?? '').trim()],
    ['Next date of Contact', nextContactValue],
    ['Last Date With Attorney', dateFromInput(input.lastAttorneyDate) || ''],
    ['Payment Status', String(input.paymentStatus ?? '').trim()],
    ['NOTES', String(input.notes ?? '').trim()],
  ].map(([header, value]) => [findColumn(master, header), value]).filter(([column]) => column);

  const archive = unzipSync(new Uint8Array(await readFile(filePath)));
  const masterPath = worksheetPaths(archive).get(MASTER_SHEET);
  if (!masterPath) throw new Error('Could not find the MasterData worksheet in the workbook.');
  updateWorksheetCells(archive, masterPath, rowNumber, values);
  await replaceWorkbookFile(filePath, archive);
  return (await readClients(filePath)).find((client) => client.id === clientId);
}

async function completeFollowUpInWorkbook(filePath, clientId) {
  const workbook = await loadWorkbook(filePath);
  const master = workbook.getWorksheet(MASTER_SHEET);
  if (!master) throw new Error('The workbook must contain a MasterData sheet.');
  const idColumn = findColumn(master, 'ID No.');
  const nextContactColumn = findColumn(master, 'Next date of Contact');
  let rowNumber = 0;
  for (let currentRow = HEADER_ROW + 1; currentRow <= master.actualRowCount; currentRow += 1) {
    if (cellText(master.getRow(currentRow).getCell(idColumn)) === clientId) rowNumber = currentRow;
  }
  if (!rowNumber) throw new Error('Client not found.');
  const scheduledDate = cellText(master.getRow(rowNumber).getCell(nextContactColumn));
  if (!scheduledDate) {
    throw new Error('This client has no scheduled follow-up to complete.');
  }
  if (normalize(scheduledDate) === 'done') throw new Error('This follow-up is already marked done.');

  const archive = unzipSync(new Uint8Array(await readFile(filePath)));
  const masterPath = worksheetPaths(archive).get(MASTER_SHEET);
  if (!masterPath) throw new Error('Could not find the MasterData worksheet in the workbook.');
  updateWorksheetCells(archive, masterPath, rowNumber, [[nextContactColumn, 'done']]);
  await replaceWorkbookFile(filePath, archive);
  return (await readClients(filePath)).find((client) => client.id === clientId);
}

export function createWorkbookStore(filePath) {
  return {
    list: () => readClients(filePath),
    add: (input) => {
      const operation = writeQueue.then(() => addClientToWorkbook(filePath, input));
      writeQueue = operation.catch(() => {});
      return operation;
    },
    completeFollowUp(clientId) {
      const operation = writeQueue.then(() => completeFollowUpInWorkbook(filePath, clientId));
      writeQueue = operation.catch(() => {});
      return operation;
    },
    update(clientId, input) {
      const operation = writeQueue.then(() => updateClientInWorkbook(filePath, clientId, input));
      writeQueue = operation.catch(() => {});
      return operation;
    },
  };
}

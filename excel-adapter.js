(function (global) {
  'use strict';

  const SHEET_NAME = 'NG_Record';
  const TABLE_NAME = 'NGRecords';
  const HEADERS = ['Record Date', 'Record Time', 'Part No.', 'Job No.', 'Cause', 'NG Qty'];

  function isAvailable() {
    const office = global.Office;
    return Boolean(office && office.context && office.HostType && office.context.host === office.HostType.Excel && global.Excel && typeof global.Excel.run === 'function');
  }

  function requireExcel() {
    if (!isAvailable()) throw new Error('Excel workbook connection is not available. Open NG Record from Excel Add-ins and try again.');
  }

  async function getWorksheet(context, create) {
    let sheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAME);
    sheet.load('name');
    await context.sync();
    if (sheet.isNullObject) {
      if (!create) return null;
      sheet = context.workbook.worksheets.add(SHEET_NAME);
      await context.sync();
    }
    return sheet;
  }

  async function getTable(context, create) {
    const sheet = await getWorksheet(context, create);
    if (!sheet) return null;

    let table = sheet.tables.getItemOrNullObject(TABLE_NAME);
    table.load('name');
    await context.sync();
    if (table.isNullObject) {
      if (!create) return null;

      const usedRange = sheet.getUsedRange();
      usedRange.load('values');
      await context.sync();
      const usedValues = usedRange.values || [];
      const hasExistingContent = usedValues.flat().some(value => value !== null && value !== undefined && String(value).trim() !== '');
      if (hasExistingContent) {
        throw new Error('The NG_Record worksheet already contains data. Move or rename that data before creating the NGRecords table.');
      }

      sheet.getRange('A1:F1').values = [HEADERS];
      table = sheet.tables.add('A1:F1', true);
      table.name = TABLE_NAME;
      await context.sync();
      table = sheet.tables.getItemOrNullObject(TABLE_NAME);
      table.load('name');
      await context.sync();
    }

    const headerRange = table.getHeaderRowRange();
    headerRange.load('values');
    await context.sync();
    const actualHeaders = (headerRange.values && headerRange.values[0] || []).map(value => String(value));
    if (actualHeaders.length !== HEADERS.length || HEADERS.some((header, index) => actualHeaders[index] !== header)) {
      throw new Error('The NGRecords table has incompatible headers. Keep its columns in this order: Record Date, Record Time, Part No., Job No., Cause, NG Qty.');
    }
    return table;
  }

  async function appendRecord(record) {
    requireExcel();
    return global.Excel.run(async context => {
      const table = await getTable(context, true);
      [2, 3, 4].forEach(index => {
        table.columns.getItemAt(index).getRange().numberFormat = [['@']];
      });
      table.rows.add(-1, [[record.date, record.time, record.part, record.job, record.cause || '', Number(record.qty)]]);
      await context.sync();
    });
  }

  async function readRecords() {
    requireExcel();
    return global.Excel.run(async context => {
      const table = await getTable(context, false);
      if (!table) return [];
      const range = table.getRange();
      range.load('values');
      await context.sync();
      const body = (range.values || []).slice(1);
      return body.filter(row => row.some(value => value !== null && value !== undefined && value !== '')).map(row => ({
        date: String(row[0] || ''),
        time: String(row[1] || ''),
        part: String(row[2] || ''),
        job: String(row[3] || ''),
        cause: String(row[4] || ''),
        qty: Number(row[5]) || 0,
      }));
    });
  }

  const adapter = Object.freeze({ isAvailable, appendRecord, readRecords });
  global.NGExcel = adapter;
  if (typeof module !== 'undefined' && module.exports) module.exports = adapter;
})(globalThis);

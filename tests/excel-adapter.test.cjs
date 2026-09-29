const test = require('node:test');
const assert = require('node:assert/strict');

const HEADERS = ['Record Date', 'Record Time', 'Part No.', 'Job No.', 'Cause', 'NG Qty'];

function createExcelMock({ host = 'Excel', sheetName = null, headers = null, rows = [], cellValues = null } = {}) {
  const state = { sheets: new Map(), tables: new Map(), formats: [] };
  const makeRange = values => ({
    values,
    load() {},
    set numberFormat(value) { state.formats.push(value); },
    format: {
      set numberFormat(value) { state.formats.push(value); },
    },
  });
  const makeTable = (name, tableHeaders, tableRows = []) => {
    const table = {
      get name() { return name; },
      set name(value) { state.tables.delete(name); name = value; state.tables.set(name, table); },
      load() {},
      getRange: () => makeRange([tableHeaders, ...tableRows]),
      getHeaderRowRange: () => makeRange([tableHeaders]),
      columns: { getItemAt: index => ({ getRange: () => makeRange([]) }) },
      rows: {
        add: (_index, values) => {
          if (values[0].some((value, index) => typeof value === 'string' && value.startsWith('=') && !state.formats.some(format => format === '@' || (Array.isArray(format) && format.some(row => row.includes('@')))))) {
            throw new Error('Formula-like text was written without text formatting');
          }
          tableRows.push(...values);
        },
      },
    };
    state.tables.set(name, table);
    return table;
  };
  const makeSheet = name => {
    const cells = {};
    if (cellValues) cells['A1:F1'] = cellValues;
    const tables = {
      getItemOrNullObject: tableName => state.tables.get(tableName) || { isNullObject: true, load() {} },
      add: () => makeTable('Table1', cells['A1:F1']?.[0] || [], []),
    };
    const sheet = {
      name,
      isNullObject: false,
      load() {},
      getRange: address => ({
        get values() { return cells[address]; },
        set values(value) { cells[address] = value; },
      }),
      getUsedRange: () => makeRange(Object.values(cells).flat()),
      tables,
    };
    state.sheets.set(name, sheet);
    if (headers) makeTable('NGRecords', headers, rows);
    return sheet;
  };
  if (sheetName) makeSheet(sheetName);
  const workbook = {
    worksheets: {
      getItemOrNullObject: name => state.sheets.get(name) || { isNullObject: true, load() {} },
      add: name => makeSheet(name),
    },
  };
  const Office = { HostType: { Excel: 'Excel' }, context: { host } };
  const Excel = { run: callback => callback({ workbook, sync: async () => {} }) };
  global.Office = Office;
  global.Excel = Excel;
  delete require.cache[require.resolve('../excel-adapter.js')];
  return { adapter: require('../excel-adapter.js'), state, clearOffice: () => { delete global.Office; delete global.Excel; } };
}

test('detects Excel host only when Office and Excel.run are available', () => {
  const excelHost = createExcelMock();
  assert.equal(excelHost.adapter.isAvailable(), true);
  excelHost.clearOffice();

  const wordHost = createExcelMock({ host: 'Word' });
  assert.equal(wordHost.adapter.isAvailable(), false);
  wordHost.clearOffice();

  delete global.Office;
  delete global.Excel;
  const adapter = require('../excel-adapter.js');
  assert.equal(adapter.isAvailable(), false);
});

test('creates the NG worksheet and table then appends one record', async () => {
  const mock = createExcelMock();
  await mock.adapter.appendRecord({ date: '2026-09-29', time: '12:30:00', part: 'PART-1', job: 'JOB-1', cause: 'ไฟดับ', qty: 10 });
  const table = mock.state.tables.get('NGRecords');
  assert.ok(mock.state.sheets.has('NG_Record'));
  assert.deepEqual(table.getHeaderRowRange().values, [HEADERS]);
  assert.deepEqual(table.getRange().values[1], ['2026-09-29', '12:30:00', 'PART-1', 'JOB-1', 'ไฟดับ', 10]);
  mock.clearOffice();
});

test('reads records from an existing table and returns no records for a header-only table', async () => {
  const populated = createExcelMock({ sheetName: 'NG_Record', headers: HEADERS, rows: [['2026-09-29', '12:30:00', 'P', 'J', 'เครื่องเสีย', 10]] });
  assert.deepEqual(await populated.adapter.readRecords(), [{ date: '2026-09-29', time: '12:30:00', part: 'P', job: 'J', cause: 'เครื่องเสีย', qty: 10 }]);
  populated.clearOffice();

  const empty = createExcelMock({ sheetName: 'NG_Record', headers: HEADERS });
  assert.deepEqual(await empty.adapter.readRecords(), []);
  empty.clearOffice();
});

test('rejects an incompatible NGRecords table without changing its rows', async () => {
  const before = [['old date', 'old part']];
  const mock = createExcelMock({ sheetName: 'NG_Record', headers: ['Old Date', 'Old Part'], rows: before });
  await assert.rejects(mock.adapter.appendRecord({ date: '2026-09-29', time: '12:30:00', part: 'P', job: 'J', cause: '', qty: 1 }), /headers/i);
  assert.deepEqual(mock.state.tables.get('NGRecords').getRange().values[1], before[0]);
  mock.clearOffice();
});

test('refuses to initialize a nonempty NG_Record worksheet', async () => {
  const mock = createExcelMock({ sheetName: 'NG_Record', cellValues: [['unrelated data']] });
  await assert.rejects(mock.adapter.appendRecord({ date: '2026-09-29', time: '12:30:00', part: 'P', job: 'J', cause: '', qty: 1 }), /already contains data/i);
  assert.equal(mock.state.tables.has('NGRecords'), false);
  mock.clearOffice();
});

test('preserves barcode identifiers beginning with zero or an equals sign as text', async () => {
  const mock = createExcelMock();
  await mock.adapter.appendRecord({ date: '2026-09-29', time: '12:30:00', part: '00123', job: '=1+1', cause: 'เครื่องเสีย', qty: 2 });
  assert.deepEqual(mock.state.tables.get('NGRecords').getRange().values[1], ['2026-09-29', '12:30:00', '00123', '=1+1', 'เครื่องเสีย', 2]);
  assert.ok(mock.state.formats.some(format => format === '@' || (Array.isArray(format) && format.some(row => row.includes('@')))));
  mock.clearOffice();
});

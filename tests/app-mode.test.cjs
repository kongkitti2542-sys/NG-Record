const test = require('node:test');
const assert = require('node:assert/strict');

function loadStore() {
  delete require.cache[require.resolve('../record-store.js')];
  return require('../record-store.js');
}

test('Excel mode delegates record reads and writes only to the workbook adapter', async () => {
  const calls = [];
  const record = { date: '2026-09-29', time: '12:30:00', part: 'P', job: 'J', cause: '', qty: 10 };
  const excel = {
    isAvailable: () => true,
    appendRecord: async value => calls.push(['excel:add', value]),
    readRecords: async () => { calls.push(['excel:all']); return [record]; },
  };
  const indexedDb = {
    add: async () => calls.push(['browser:add']),
    all: async () => { calls.push(['browser:all']); return []; },
  };
  const store = loadStore().create({ excel, isExcelHost: true, indexedDb });
  assert.equal(store.mode, 'excel');
  await store.add(record);
  assert.deepEqual(await store.all(), [record]);
  assert.deepEqual(calls, [['excel:add', record], ['excel:all']]);
});

test('standalone mode delegates reads and writes only to IndexedDB', async () => {
  const calls = [];
  const record = { date: '2026-09-29', part: 'P' };
  const excel = {
    isAvailable: () => false,
    appendRecord: async () => calls.push(['excel:add']),
    readRecords: async () => { calls.push(['excel:all']); return []; },
  };
  const indexedDb = {
    add: async value => calls.push(['browser:add', value]),
    all: async () => { calls.push(['browser:all']); return [record]; },
  };
  const store = loadStore().create({ excel, isExcelHost: false, indexedDb });
  assert.equal(store.mode, 'browser');
  await store.add(record);
  assert.deepEqual(await store.all(), [record]);
  assert.deepEqual(calls, [['browser:add', record], ['browser:all']]);
});

test('Excel write failures propagate without writing to IndexedDB', async () => {
  let browserWrites = 0;
  const failure = new Error('Workbook is read-only');
  const store = loadStore().create({
    isExcelHost: true,
    excel: { isAvailable: () => true, appendRecord: async () => { throw failure; }, readRecords: async () => [] },
    indexedDb: { add: async () => { browserWrites += 1; }, all: async () => [] },
  });
  await assert.rejects(store.add({ part: 'P' }), /read-only/);
  assert.equal(browserWrites, 0);
});

test('Excel mode refuses to start without an available Excel adapter', () => {
  const indexedDb = { add: async () => {}, all: async () => [] };
  assert.throws(() => loadStore().create({ isExcelHost: true, excel: { isAvailable: () => false }, indexedDb }), /Excel workbook connection is not available/i);
});

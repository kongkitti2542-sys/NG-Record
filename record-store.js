(function (global) {
  'use strict';

  function create(options) {
    const { excel, isExcelHost, indexedDb } = options;
    if (isExcelHost) {
      if (!excel || typeof excel.isAvailable !== 'function' || !excel.isAvailable()) {
        throw new Error('Excel workbook connection is not available. Reopen NG Record from Excel Add-ins.');
      }
      return Object.freeze({
        mode: 'excel',
        add: record => excel.appendRecord(record),
        all: () => excel.readRecords(),
      });
    }

    if (!indexedDb || typeof indexedDb.add !== 'function' || typeof indexedDb.all !== 'function') {
      throw new Error('Browser-local storage is not available.');
    }
    return Object.freeze({
      mode: 'browser',
      add: record => indexedDb.add(record),
      all: () => indexedDb.all(),
    });
  }

  const api = Object.freeze({ create });
  global.NGRecordStore = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(globalThis);

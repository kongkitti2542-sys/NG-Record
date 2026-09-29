(() => {
  'use strict';

  const $ = selector => document.querySelector(selector);
  const today = () => {
    const date = new Date();
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
  };
  const excelAddinRequested = new URLSearchParams(window.location.search).get('addin') === '1';
  let databasePromise;
  let storeUnavailable = false;

  function openDatabase() {
    if (!databasePromise) {
      databasePromise = new Promise((resolve, reject) => {
        const request = indexedDB.open('ng-record-db', 1);
        request.onupgradeneeded = () => {
          if (!request.result.objectStoreNames.contains('records')) {
            request.result.createObjectStore('records', { keyPath: 'id', autoIncrement: true });
          }
        };
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
    }
    return databasePromise;
  }

  async function browserAll() {
    const database = await openDatabase();
    return new Promise((resolve, reject) => {
      const request = database.transaction('records').objectStore('records').getAll();
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async function browserAdd(value) {
    const database = await openDatabase();
    return new Promise((resolve, reject) => {
      const transaction = database.transaction('records', 'readwrite');
      transaction.objectStore('records').add(value);
      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error);
    });
  }

  function setStorageStatus(message, isError = false) {
    const status = $('#storageStatus');
    status.textContent = message;
    status.className = `storage-status${isError ? ' error' : ''}`;
  }

  function loadOfficeJs() {
    if (window.Office && typeof window.Office.onReady === 'function') return Promise.resolve();
    return new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
      script.onload = resolve;
      script.onerror = () => reject(new Error('โหลด Office.js ไม่สำเร็จ ตรวจสอบการเชื่อมต่ออินเทอร์เน็ตแล้วลองอีกครั้ง'));
      document.head.append(script);
    });
  }

  async function initializeStore() {
    let isExcelHost = false;
    if (excelAddinRequested) {
      try {
        await loadOfficeJs();
        const info = await window.Office.onReady();
        isExcelHost = info.host === window.Office.HostType.Excel && window.NGExcel.isAvailable();
        if (!isExcelHost) throw new Error('ไม่พบ workbook ที่เปิดอยู่ ให้เปิด NG Record จากเมนู Add-ins ใน Excel');
      } catch (error) {
        storeUnavailable = true;
        setStorageStatus(`เชื่อมต่อ Excel ไม่สำเร็จ: ${error.message}`, true);
        $('#save').disabled = true;
        return null;
      }
    }

    try {
      const store = window.NGRecordStore.create({
        excel: window.NGExcel,
        isExcelHost,
        indexedDb: { add: browserAdd, all: browserAll },
      });
      if (store.mode === 'excel') {
        setStorageStatus('Excel เชื่อมต่อกับ workbook ปัจจุบัน · บันทึกลงตาราง NGRecords');
        $('#entrySub').textContent = 'รายการนี้จะถูกบันทึกลง workbook ที่เปิดอยู่';
      } else {
        setStorageStatus('เบราว์เซอร์ · ข้อมูลเก็บไว้ในเบราว์เซอร์เครื่องนี้');
        $('#entrySub').textContent = 'สแกนหรือพิมพ์ข้อมูล แล้วบันทึกลงเบราว์เซอร์เครื่องนี้';
      }
      return store;
    } catch (error) {
      storeUnavailable = true;
      setStorageStatus(error.message, true);
      $('#save').disabled = true;
      return null;
    }
  }

  const storePromise = initializeStore();
  const notify = message => {
    const toast = $('#toast');
    toast.textContent = message;
    toast.classList.add('show');
    setTimeout(() => toast.classList.remove('show'), 2500);
  };
  const showNotice = (message, kind) => {
    const notice = $('#notice');
    notice.textContent = message;
    notice.className = `notice show ${kind}`;
  };
  const setToday = () => { $('#date').value = today(); };
  const localTime = () => new Intl.DateTimeFormat('en-GB', {
    hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23',
  }).format(new Date());
  const formatTime = value => {
    if (/^\d{2}:\d{2}:\d{2}$/.test(String(value))) return String(value);
    const date = new Date(value);
    return Number.isNaN(date.valueOf()) ? String(value || '') : date.toLocaleTimeString('th-TH', {
      hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false,
    });
  };
  const dateTime = record => `${record.date} ${formatTime(record.time)}`;
  setToday();

  document.querySelectorAll('.tab').forEach(button => {
    button.onclick = () => {
      document.querySelectorAll('.tab,.panel').forEach(element => element.classList.remove('active'));
      button.classList.add('active');
      $(`#${button.dataset.tab}`).classList.add('active');
      if (button.dataset.tab === 'history') render();
    };
  });

  $('#entryForm').addEventListener('submit', async event => {
    event.preventDefault();
    const button = $('#save');
    const quantity = $('#qty').valueAsNumber;
    const part = $('#part').value.trim();
    const job = $('#job').value.trim();
    if (!part || !job || !Number.isInteger(quantity) || quantity <= 0 || !$('#date').value) {
      showNotice('กรุณาตรวจสอบข้อมูลให้ครบ และ NG Qty ต้องเป็นจำนวนเต็มมากกว่า 0', 'error');
      return;
    }

    button.disabled = true;
    button.textContent = 'กำลังบันทึก…';
    try {
      const store = await storePromise;
      if (!store) throw new Error('ที่เก็บข้อมูลยังไม่พร้อม กรุณาตรวจสอบสถานะการเชื่อมต่อ');
      const timestamp = store.mode === 'excel' ? localTime() : new Date().toISOString();
      await store.add({
        part,
        job,
        cause: $('#cause').value.trim(),
        qty: quantity,
        date: $('#date').value,
        time: timestamp,
      });
      showNotice('บันทึกสำเร็จ พร้อมบันทึกรายการถัดไป', 'ok');
      $('#part').value = '';
      $('#job').value = '';
      $('#cause').value = '';
      $('#qty').value = '';
      setToday();
      $('#part').focus();
    } catch (error) {
      showNotice(`บันทึกไม่สำเร็จ: ${error.message || 'ตรวจสอบที่เก็บข้อมูลแล้วลองอีกครั้ง'}`, 'error');
    } finally {
      button.disabled = storeUnavailable;
      button.textContent = 'บันทึกรายการ';
    }
  });

  async function filteredRecords(store) {
    const from = $('#from').value;
    const to = $('#to').value;
    const partFilter = $('#partFilter').value.trim().toLowerCase();
    const jobFilter = $('#jobFilter').value.trim().toLowerCase();
    const records = await store.all();
    return records
      .filter(record => (!from || record.date >= from)
        && (!to || record.date <= to)
        && (!partFilter || String(record.part || '').toLowerCase().includes(partFilter))
        && (!jobFilter || String(record.job || '').toLowerCase().includes(jobFilter)))
      .sort((left, right) => `${right.date}T${formatTime(right.time)}`.localeCompare(`${left.date}T${formatTime(left.time)}`));
  }

  function renderEmpty(message) {
    const body = $('#rows');
    body.replaceChildren();
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 5;
    cell.className = 'empty';
    cell.textContent = message;
    row.append(cell);
    body.append(row);
  }

  async function render() {
    const body = $('#rows');
    try {
      const store = await storePromise;
      if (!store) {
        $('#count').textContent = '0';
        $('#total').textContent = '0';
        renderEmpty('ยังเชื่อมต่อกับ Excel ไม่ได้');
        return;
      }
      const records = await filteredRecords(store);
      body.replaceChildren();
      $('#count').textContent = records.length.toLocaleString();
      $('#total').textContent = records.reduce((sum, record) => sum + (Number(record.qty) || 0), 0).toLocaleString();
      if (!records.length) {
        renderEmpty('ไม่พบรายการ');
        return;
      }
      for (const record of records) {
        const row = document.createElement('tr');
        [dateTime(record), record.part, record.job, record.cause || '', Number(record.qty).toLocaleString()].forEach((value, index) => {
          const cell = document.createElement('td');
          cell.textContent = value;
          if (index === 4) cell.className = 'qty';
          row.append(cell);
        });
        body.append(row);
      }
    } catch (error) {
      $('#count').textContent = '0';
      $('#total').textContent = '0';
      renderEmpty(`อ่านประวัติไม่สำเร็จ: ${error.message || 'ตรวจสอบ workbook แล้วลองอีกครั้ง'}`);
    }
  }

  ['from', 'to', 'partFilter', 'jobFilter'].forEach(id => $('#'+id).addEventListener('input', render));
  $('#clear').onclick = () => {
    ['from', 'to', 'partFilter', 'jobFilter'].forEach(id => { $('#'+id).value = ''; });
    render();
  };

  function filename(extension) { return `NG_Record_${today()}.${extension}`; }
  function download(blob, name) {
    const anchor = document.createElement('a');
    anchor.href = URL.createObjectURL(blob);
    anchor.download = name;
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(anchor.href), 1000);
  }
  const columns = ['วันที่เวลา', 'Part No.', 'Job No.', 'สาเหตุการเกิด', 'NG Qty'];
  const csvCell = value => `"${String(value).replaceAll('"', '""')}"`;
  async function currentExportRecords() {
    const store = await storePromise;
    if (!store) throw new Error('ที่เก็บข้อมูลยังไม่พร้อม');
    return filteredRecords(store);
  }
  $('#csv').onclick = async () => {
    try {
      const records = await currentExportRecords();
      const lines = [columns, ...records.map(record => [dateTime(record), record.part, record.job, record.cause || '', record.qty])]
        .map(row => row.map(csvCell).join(','));
      download(new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }), filename('csv'));
      notify('ส่งออก CSV แล้ว');
    } catch (error) {
      notify(`ส่งออกไม่สำเร็จ: ${error.message}`);
    }
  };

  function xmlEscape(value) {
    return String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;');
  }
  function crc32(bytes) {
    let checksum = 0xffffffff;
    for (const byte of bytes) {
      checksum ^= byte;
      for (let bit = 0; bit < 8; bit += 1) checksum = (checksum >>> 1) ^ ((checksum & 1) ? 0xedb88320 : 0);
    }
    return (checksum ^ 0xffffffff) >>> 0;
  }
  function zip(files) {
    const encoder = new TextEncoder();
    const parts = [];
    const centralDirectory = [];
    let offset = 0;
    for (const [name, content] of files) {
      const nameBytes = encoder.encode(name);
      const data = encoder.encode(content);
      const checksum = crc32(data);
      const localHeader = new Uint8Array(30 + nameBytes.length);
      const localView = new DataView(localHeader.buffer);
      localView.setUint32(0, 0x04034b50, true);
      localView.setUint16(4, 20, true);
      localView.setUint32(14, checksum, true);
      localView.setUint32(18, data.length, true);
      localView.setUint32(22, data.length, true);
      localView.setUint16(26, nameBytes.length, true);
      localHeader.set(nameBytes, 30);
      parts.push(localHeader, data);

      const directoryEntry = new Uint8Array(46 + nameBytes.length);
      const directoryView = new DataView(directoryEntry.buffer);
      directoryView.setUint32(0, 0x02014b50, true);
      directoryView.setUint16(4, 20, true);
      directoryView.setUint16(6, 20, true);
      directoryView.setUint32(16, checksum, true);
      directoryView.setUint32(20, data.length, true);
      directoryView.setUint32(24, data.length, true);
      directoryView.setUint16(28, nameBytes.length, true);
      directoryView.setUint32(42, offset, true);
      directoryEntry.set(nameBytes, 46);
      centralDirectory.push(directoryEntry);
      offset += localHeader.length + data.length;
    }
    const directorySize = centralDirectory.reduce((sum, entry) => sum + entry.length, 0);
    const end = new Uint8Array(22);
    const endView = new DataView(end.buffer);
    endView.setUint32(0, 0x06054b50, true);
    endView.setUint16(8, centralDirectory.length, true);
    endView.setUint16(10, centralDirectory.length, true);
    endView.setUint32(12, directorySize, true);
    endView.setUint32(16, offset, true);
    return new Blob([...parts, ...centralDirectory, end], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  $('#xlsx').onclick = async () => {
    try {
      const records = await currentExportRecords();
      const data = [columns, ...records.map(record => [dateTime(record), record.part, record.job, record.cause || '', Number(record.qty)])];
      const cell = (value, row, column) => {
        const reference = `${String.fromCharCode(65 + column)}${row}`;
        return typeof value === 'number'
          ? `<c r="${reference}"><v>${value}</v></c>`
          : `<c r="${reference}" t="inlineStr"><is><t>${xmlEscape(value)}</t></is></c>`;
      };
      const rows = data.map((values, index) => `<row r="${index + 1}">${values.map((value, column) => cell(value, index + 1, column)).join('')}</row>`).join('');
      const files = [
        ['[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'],
        ['_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
        ['xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="NG Record" sheetId="1" r:id="rId1"/></sheets></workbook>'],
        ['xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'],
        ['xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`],
      ];
      download(zip(files), filename('xlsx'));
      notify('ส่งออก Excel แล้ว');
    } catch (error) {
      notify(`ส่งออกไม่สำเร็จ: ${error.message}`);
    }
  };
})();

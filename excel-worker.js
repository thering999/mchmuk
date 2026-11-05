// Web Worker: parse XLSX/CSV in background thread so UI doesn't freeze
importScripts('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');

self.onmessage = function(e) {
    const { buffer, sheetName } = e.data;
    try {
        const data = new Uint8Array(buffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellNF: false, cellText: false });
        const sheetNames = workbook.SheetNames;
        const target = sheetName || (sheetNames.includes('DATA') ? 'DATA' : sheetNames[0]);
        const sheet = workbook.Sheets[target];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        const rowsText = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
        self.postMessage({ sheetNames, targetSheet: target, rows, rowsText });
    } catch(err) {
        self.postMessage({ error: err.message });
    }
};

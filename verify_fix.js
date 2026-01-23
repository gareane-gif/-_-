
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Mock browser functions needed by script.js-like logic
function normalizeId2(val) {
    if (val == null) return '';
    if (typeof val === 'number') return String(Math.floor(val));
    let s = String(val).trim();
    s = s.replace(/[\u0660-\u0669]/g, c => String(c.charCodeAt(0) - 0x0660));
    s = s.replace(/[\u06F0-\u06F9]/g, c => String(c.charCodeAt(0) - 0x06F0));
    s = s.replace(/\s+/g, '');
    s = s.replace(/[^a-zA-Z0-9]/g, '');
    if (/^\d+$/.test(s)) s = s.replace(/^0+/, '');
    return s.toUpperCase();
}

function isIdMatch(target, candidate) {
    const normTarget = normalizeId2(target);
    const normCandidate = normalizeId2(candidate);
    if (!normTarget || !normCandidate) return false;
    if (normTarget === normCandidate) return true;
    const numTarget = normTarget.replace(/[^0-9]/g, '');
    const numCandidate = normCandidate.replace(/[^0-9]/g, '');
    if (numTarget && numCandidate && numTarget === numCandidate && numTarget.length >= 4) {
        return true;
    }
    return false;
}

const xlsDir = 'xls';
const tests = [
    { file: 'قسم الحاسوب25-26.xls', id: 'CO231004' },
    { file: 'قسم الكهرباء25-26.xls', id: 'EL251001' },
    { file: 'قسم المحاسبة25-26.xls', id: 'AC242041' },
    { file: 'قسم الميكانيكا 25-26.xls', id: 'ME242002' }
];

tests.forEach(test => {
    console.log(`\nTesting ${test.file} for ID ${test.id}...`);
    const filePath = path.join(xlsDir, test.file);
    if (!fs.existsSync(filePath)) {
        console.log("  File not found!");
        return;
    }
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const range = XLSX.utils.decode_range(sheet['!ref']);

    let found = false;
    for (let r = range.s.r; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
            const cell = sheet[XLSX.utils.encode_cell({ r: r, c: c })];
            const val = cell ? cell.v : '';
            if (isIdMatch(test.id, val)) {
                console.log(`  FOUND ID ${test.id} at Row ${r + 1}, Col ${c + 1}: '${val}'`);
                found = true;
                break;
            }
        }
        if (found) break;
    }
    if (!found) console.log(`  NOT FOUND: ${test.id}`);
});

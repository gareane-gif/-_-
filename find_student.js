#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const idToFind = process.argv[2];
if (!idToFind) {
  console.error('Usage: node find_student.js <studentId>');
  process.exit(1);
}

function normalizeId(val) {
  if (val == null) return '';
  let s = String(val).trim();
  // remove non-digits and leading zeros
  s = s.replace(/[^0-9]/g, '');
  s = s.replace(/^0+/, '');
  return s;
}

const target = normalizeId(idToFind);
const dir = path.join(__dirname, 'xls');
if (!fs.existsSync(dir)) {
  console.error('Directory not found:', dir);
  process.exit(1);
}

const files = fs.readdirSync(dir).filter(f => f.toLowerCase().endsWith('.xls') || f.toLowerCase().endsWith('.xlsx'));
if (files.length === 0) {
  console.log('No Excel files found in', dir);
  process.exit(0);
}

let foundAny = false;
for (const file of files) {
  const filePath = path.join(dir, file);
  let workbook;
  try {
    workbook = XLSX.readFile(filePath, { cellDates: false });
  } catch (err) {
    console.error('Failed to read', file, err.message);
    continue;
  }

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

    // Find header row (look in first 10 rows for Arabic header keywords)
    let headerRowIndex = -1;
    for (let i = 0; i < Math.min(10, json.length); i++) {
      const row = json[i] || [];
      if (row.some(cell => {
        const s = (cell || '').toString();
        return /قيد|القيد|الرقم|الطالب|الاسم|تسلسل/i.test(s);
      })) {
        headerRowIndex = i;
        break;
      }
    }
    if (headerRowIndex === -1) headerRowIndex = 0;

    const headers = (json[headerRowIndex] || []).map(h => (h || '').toString().trim());

    // Try to find ID column by header keywords
    let idColIndex = headers.findIndex(h => /قيد|القيد|الرقم|الطالب|id|student|تسلسل/i.test(h));

    // If not found, try to detect column with numeric-looking values matching target
    if (idColIndex === -1) {
      for (let r = headerRowIndex + 1; r < Math.min(json.length, headerRowIndex + 200); r++) {
        const row = json[r] || [];
        for (let c = 0; c < Math.min(row.length, 40); c++) {
          const val = normalizeId(row[c]);
          if (val && val === target) {
            idColIndex = c;
            break;
          }
        }
        if (idColIndex !== -1) break;
      }
    }

    if (idColIndex === -1) continue;

    for (let r = headerRowIndex + 1; r < json.length; r++) {
      const row = json[r] || [];
      const cell = row[idColIndex];
      if (normalizeId(cell) === target) {
        // collect full row and component rows
        const rowData = [];
        const workRowData = [];
        const finalRowData = [];
        const totalRowData = [];

        for (let c = 0; c < row.length; c++) {
          rowData.push(row[c]);
          workRowData.push((json[r + 1] || [])[c]);
          finalRowData.push((json[r + 3] || [])[c]);
          totalRowData.push((json[r + 4] || [])[c]);
        }

        console.log('\n--- MATCH FOUND ---');
        console.log('File:', file);
        console.log('Sheet:', sheetName);
        console.log('Header row index (1-based):', headerRowIndex + 1);
        console.log('Data row index (1-based):', r + 1);
        console.log('ID column index (1-based):', idColIndex + 1);
        console.log('ID cell value:', cell);
        console.log('Row values:');

        // Print detailed report for each subject
        console.log(headers.map((h, i) => {
          if (!h) return null;
          const w = workRowData[i] || '-';
          const f = finalRowData[i] || '-';
          const t = totalRowData[i] || '-';
          // basic filter for non-subject columns
          if (/القيد|الطالب|الرقم|تسلسل|ملاحظ/.test(h)) return `${h}: ${rowData[i]}`;

          return `${h}: أعمال=${w} | اختبار=${f} | كلية=${t}`;
        }).filter(Boolean).join('\n'));

        foundAny = true;
      }
    }
  }
}

if (!foundAny) console.log('No matches found for', idToFind);
else console.log('\nSearch complete.');

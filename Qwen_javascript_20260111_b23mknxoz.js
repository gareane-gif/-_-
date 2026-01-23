document.getElementById('searchBtn').addEventListener('click', function () {
  const fileInput = document.getElementById('file');
  const studentId = document.getElementById('studentId').value.trim();
  const resultDiv = document.getElementById('result');

  if (!fileInput.files.length) {
    alert('يرجى اختيار ملف Excel أولاً.');
    return;
  }

  if (!studentId) {
    alert('يرجى إدخال الرقم الدراسي.');
    return;
  }

  const file = fileInput.files[0];
  // validate numeric form of studentId (support Arabic-Indic digits)
  if (normalizeId2(studentId) === '') {
    alert('يرجى إدخال رقم القيد بالأرقام (مثال: 232029). تأكد أنك لم تستخدم أرقام عربية مثل ٢٣٢٠٢٩.');
    return;
  }

  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      let foundStudent = null;
      let sheetName = null;
      const scannedSheets = []; // for diagnostics
      console.log('Searching for student ID (raw):', studentId);
      console.log('Searching for student ID (normalized):', normalizeId2(studentId));

      // create an on-page debug panel so you don't need DevTools
      let debugDiv = document.getElementById('debug');
      if (!debugDiv) {
        debugDiv = document.createElement('pre');
        debugDiv.id = 'debug';
        debugDiv.style = 'background:#f8f9fa;border:1px solid #ddd;padding:8px;max-height:200px;overflow:auto;font-size:12px;direction:ltr;text-align:left;margin-top:8px;';
        const container = document.getElementById('result').parentNode;
        container.insertBefore(debugDiv, document.getElementById('result'));
      }
      function logDebug() {
        try {
          console.log.apply(console, arguments);
        } catch (e) {}
        try {
          const parts = Array.from(arguments).map(a => (typeof a === 'object' ? JSON.stringify(a) : String(a)));
          debugDiv.textContent += parts.join(' ') + '\n';
        } catch (e) {}
      }
      logDebug('Searching for student ID (raw):', studentId);
      logDebug('Searching for student ID (normalized):', normalizeId2(studentId));

      workbook.SheetNames.forEach(name => {
        if (foundStudent) return;
        const sheet = workbook.Sheets[name];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        console.log('Scanning sheet:', name, 'rows:', json.length);
        scannedSheets.push({ name, rows: json.length });
        logDebug('Scanning sheet:', name, 'rows:', json.length);

        let headers = [];
        let startRow = 0;
      // Look for a header row by scanning the first few rows and checking any cell for header keywords
      for (let i = 0; i < Math.min(json.length, 20); i++) {
        const row = json[i] || [];
        if (row.some(cell => {
          const s = (cell || '').toString();
          return /اسم|الطالب|القيد|الرقم|تسلسل/i.test(s);
        })) {
          headers = row.map(h => h?.toString().trim() || '');
          startRow = i + 1;
          console.log('Header detected in sheet', name, 'at row', i+1, 'headers:', headers);
          logDebug('Header detected in sheet', name, 'at row', i+1, 'headers:', headers);
          break;
        }
      }

      const idColIndex = headers.findIndex(h => /قيد|القيد|الرقم|الطالب|id|student/i.test(h || ''));
      const nameColIndex = headers.findIndex(h => /اسم|الطالب|name/i.test(h || ''));
      console.log('Sheet', name, 'idColIndex:', idColIndex, 'nameColIndex:', nameColIndex, 'startRow:', startRow);
      logDebug('Sheet', name, 'idColIndex:', idColIndex, 'nameColIndex:', nameColIndex, 'startRow:', startRow);

      // If we found an ID column, search using it
      if (idColIndex !== -1) {
        for (let r = startRow; r < json.length; r++) {
          const row = json[r];
          if (!row || row.length === 0) continue;
          const cellValue = row[idColIndex];
          const id = normalizeId2(cellValue);
          if (id) {
            console.log('Checking row', r+1, 'ID cell raw:', cellValue, 'normalized:', id);
            logDebug('Checking row', r+1, 'ID cell raw:', cellValue, 'normalized:', id);
          }
          if (id === normalizeId2(studentId)) {
            console.log('Match in sheet', name, 'row', r+1);
            logDebug('Match in sheet', name, 'row', r+1, 'ID raw:', cellValue, 'normalized:', id);
            foundStudent = {
              name: (nameColIndex !== -1 ? row[nameColIndex] : 'غير معروف') || 'غير معروف',
              row: row,
              headers: headers
            };
            sheetName = name;
            break;
          }
        }
      }

      // Fallback: scan all cells in the sheet for a matching ID (handles cases where headers are missing or ID column not found)
      if (!foundStudent) {
        for (let r = startRow; r < json.length; r++) {
          const row = json[r] || [];
          for (let c = 0; c < row.length; c++) {
            const normalized = normalizeId2(row[c]);
            if (normalized) {
              console.log('Scanning cell', name, 'r', r+1, 'c', c+1, 'raw:', row[c], 'normalized:', normalized);
              logDebug('Scanning cell', name, 'r', r+1, 'c', c+1, 'raw:', row[c], 'normalized:', normalized);
            }
            if (normalized === normalizeId2(studentId)) {
              console.log('Fallback match in sheet', name, 'at row', r+1, 'col', c+1);
              logDebug('Fallback match in sheet', name, 'at row', r+1, 'col', c+1, 'raw:', row[c], 'normalized:', normalized);
              // try to pick a name from known columns or nearby cells
              let studentName = 'غير معروف';
              if (nameColIndex !== -1 && row[nameColIndex]) studentName = row[nameColIndex];
              else if (c + 1 < row.length && typeof row[c + 1] === 'string' && row[c + 1].trim()) studentName = row[c + 1];
              else if (c - 1 >= 0 && typeof row[c - 1] === 'string' && row[c - 1].trim()) studentName = row[c - 1];
              foundStudent = {
                name: studentName,
                row: row,
                headers: headers
              };
              sheetName = name;
              break;
            }
          }
          if (foundStudent) break;
        }
      }

      if (foundStudent) {
        console.log('Student found in sheet', sheetName, foundStudent);
        displayResult(foundStudent, sheetName);
      } else {
        console.warn('No student found for', studentId, 'scannedSheets:', scannedSheets);
        logDebug('No student found for', studentId, 'scannedSheets:', scannedSheets);
        resultDiv.innerHTML = `<p style="color:red; text-align:center;">لم يتم العثور على طالب برقم القيد: ${studentId}</p><p style="text-align:center; font-size:smaller;">الملفات الممسوحة: ${scannedSheets.map(s=>s.name).join(', ') || 'لا توجد'}</p>`;
        resultDiv.classList.add('show');
      }
    }
  };

  reader.readAsArrayBuffer(file);

// Robust ID normalizer that handles Western digits, Arabic-Indic (٠-٩), Persian (۰-۹), and fullwidth digits
function normalizeId2(val) {
  if (val == null) return '';
  if (typeof val === 'number') return String(Math.floor(val));
  let s = String(val).trim();
  // map Arabic-Indic digits (U+0660 - U+0669)
  s = s.replace(/[\u0660-\u0669]/g, c => String(c.charCodeAt(0) - 0x0660));
  // map Eastern Arabic-Indic / Persian digits (U+06F0 - U+06F9)
  s = s.replace(/[\u06F0-\u06F9]/g, c => String(c.charCodeAt(0) - 0x06F0));
  // map fullwidth digits (U+FF10 - U+FF19)
  s = s.replace(/[\uFF10-\uFF19]/g, c => String(c.charCodeAt(0) - 0xFF10));
  // remove non-digits and leading zeros
  s = s.replace(/[^0-9]/g, '');
  s = s.replace(/^0+/, '');
  return s;
}

});

function normalizeId(val) {
  if (val == null) return '';
  if (typeof val === 'number') return String(Math.floor(val));
  let s = String(val).trim();
  // Replace Arabic-Indic digits (٠-٩), Eastern Arabic-Indic / Persian digits (۰-۹), and fullwidth digits to ASCII digits
  s = s.replace(/[
-
]/g, '');
  s = s.replace(/[
-
]/g, '');
  s = s.replace(/[
-
]/g, '');
  s = s.replace(/[60- 669]/g, c => String(c.charCodeAt(0) - 0x0660));
  s = s.replace(/[ 6F0- 6F9]/g, c => String(c.charCodeAt(0) - 0x06F0));
  s = s.replace(/[ FF10- FF19]/g, c => String(c.charCodeAt(0) - 0xFF10));
  // Remove non-digits and leading zeros
  s = s.replace(/[^0-9]/g, '');
  s = s.replace(/^0+/, '');
  return s;
}

function displayResult(student, sheetName) {
  const { name, row, headers } = student;
  // ensure name is not the generic 'غير معروف' if possible
  const studentName = (name && name !== 'غير معروف') ? name : (row[2] || row[1] || row[0] || 'غير معروف');
  let tableHTML = `
    <h2>نتيجة الطالب: ${studentName}</h2>
    <p><strong>القسم:</strong> ${extractDepartment(sheetName)}</p>
    <table class="result-table">
      <thead><tr>
        <th>المادة</th>
        <th>الدرجة</th>
        <th>التقدير</th>
      </tr></thead>
      <tbody>
  `;

  const skipKeywords = ['المجموع', 'معدل', 'نتيجة', 'ملاحظ', 'التقدي', 'عدد', 'أعمال', 'امتحان', 'تقييم', 'الطالب', 'الرقم', 'الاسم', 'تسلسل'];

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    if (!header) continue;

    const shouldSkip = skipKeywords.some(kw => header.includes(kw));
    if (shouldSkip) continue;

    const grade = row[i] != null ? row[i] : '-';
    let estimate = '-';
    if (typeof grade === 'number') {
      if (grade >= 90) estimate = 'ممتاز';
      else if (grade >= 80) estimate = 'جيد جداً';
      else if (grade >= 70) estimate = 'جيد';
      else if (grade >= 60) estimate = 'مقبول';
      else estimate = 'ضعيف';
    }

    tableHTML += `
      <tr>
        <td>${header}</td>
        <td>${grade}</td>
        <td class="${estimate === 'ضعيف' ? 'fail' : estimate === 'ممتاز' ? 'success' : ''}">${estimate}</td>
      </tr>
    `;
  }

  tableHTML += `</tbody></table>`;
  document.getElementById('result').innerHTML = tableHTML;
  document.getElementById('result').classList.add('show');
}

function extractDepartment(sheetName) {
  if (sheetName.includes('حاسوب')) return 'علوم الحاسوب';
  if (sheetName.includes('طاقة') || sheetName.includes('طاقات')) return 'الطاقات المتجددة';
  if (sheetName.includes('كهرباء')) return 'الهندسة الكهربائية';
  if (sheetName.includes('مساحة')) return 'المساحة';
  if (sheetName.includes('محاسبة')) return 'المحاسبة';
  return 'غير محدد';
}
// Robust normalizer: handles Western digits, Arabic-Indic (٠-٩), Persian (۰-۹), and fullwidth digits
function normalizeId2(val) {
  if (val == null) return '';
  if (typeof val === 'number') return String(Math.floor(val));
  let s = String(val).trim();
  // FIX: معالجة الأرقام التي تحتوي على كسور عشرية مخزنة كنص (مثل "252001.0")
  if (s.includes('.')) {
    const parts = s.split('.');
    // إذا كان الجزء العشري عبارة عن أصفار فقط، نتجاهله
    if (parts[1] && /^0+$/.test(parts[1])) s = parts[0];
  }
  // Normalize Arabic-Indic/Persian digits
  s = s.replace(/[\u0660-\u0669]/g, c => String(c.charCodeAt(0) - 0x0660));
  s = s.replace(/[\u06F0-\u06F9]/g, c => String(c.charCodeAt(0) - 0x06F0));
  // Remove spaces and special chars
  s = s.replace(/\s+/g, '');
  s = s.replace(/[^a-zA-Z0-9]/g, '');
  // Always remove leading zeros (to match Python logic)
  s = s.replace(/^0+/, '');
  return s.toUpperCase();
}

function isIdMatch(target, candidate) {
  const normTarget = normalizeId2(target);
  const normCandidate = normalizeId2(candidate);
  if (!normTarget || !normCandidate) return false;
  if (normTarget === normCandidate) return true;

  const alphaTarget = normTarget.replace(/[0-9]/g, '');
  const alphaCandidate = normCandidate.replace(/[0-9]/g, '');
  
  const numTarget = normTarget.replace(/[^0-9]/g, '');
  const numCandidate = normCandidate.replace(/[^0-9]/g, '');

  // If both have letters, the letters MUST match
  if (alphaTarget && alphaCandidate && alphaTarget !== alphaCandidate) {
    return false;
  }
  
  // If numeric parts match and are significant (e.g. >= 4 digits), consider it a match
  if (numTarget && numCandidate && numTarget === numCandidate && numTarget.length >= 4) {
    return true;
  }
  return false;
}

function normalizeArabicText(text) {
  if (!text) return "";
  let s = String(text);
  s = s.replace(/[\u0660-\u0669]/g, c => String(c.charCodeAt(0) - 0x0660));
  s = s.replace(/[\u06F0-\u06F9]/g, c => String(c.charCodeAt(0) - 0x06F0));
  s = s.replace(/[\u064B-\u065F]/g, ""); // Remove Harakat
  s = s.replace(/[أإآ]/g, "ا");      // Standardize Alef
  s = s.replace(/ة/g, "ه");          // Standardize Teh Marbuta
  s = s.replace(/[ىي]/g, "ي");       // Standardize Yeh/Alef Maksura
  s = s.replace(/[ـ\s]/g, "");       // Remove Tatweel and spaces
  s = s.replace(/[^\u0621-\u064Aa-zA-Z0-9]/g, "");
  return s.trim();
}

function extractSemesterNumber(text) {
  const norm = normalizeArabicText(text);
  if (norm.includes("اول") || norm.includes("1")) return 1;
  if (norm.includes("ثاني") || norm.includes("2")) return 2;
  if (norm.includes("ثالث") || norm.includes("3")) return 3;
  if (norm.includes("رابع") || norm.includes("4")) return 4;
  if (norm.includes("خامس") || norm.includes("5")) return 5;
  if (norm.includes("سادس") || norm.includes("6")) return 6;
  if (norm.includes("سابع") || norm.includes("7")) return 7;
  if (norm.includes("ثامن") || norm.includes("8")) return 8;
  return null;
}

window.__SCRIPT_LOADED = true;

function pickName(rowData, idColIdx, nameColIdx) {
  const tried = [];
  if (typeof nameColIdx === 'number' && nameColIdx >= 0) tried.push(nameColIdx);
  if (typeof idColIdx === 'number' && idColIdx >= 0) {
    if (idColIdx - 1 >= 0) tried.push(idColIdx - 1);
    if (idColIdx - 2 >= 0) tried.push(idColIdx - 2);
  }
  tried.push(2);
  for (const ci of tried) {
    if (ci >= 0 && ci < rowData.length) {
      const v = rowData[ci];
      if (typeof v === 'string' && v.trim()) return v.trim();
      if (typeof v === 'number' && String(v).trim()) return String(v).trim();
    }
  }
  for (let i = 0; i < rowData.length; i++) {
    const v = rowData[i];
    if (typeof v === 'string' && /[A-Za-z\u0600-\u06FF]/.test(v)) {
      return v.trim();
    }
    if (typeof v === 'number') {
      const s = String(v).trim();
      if (s) return s;
    }
  }
  return 'غير معروف';
}

window.__currentUser = null;
window.__workbookCache = new Map(); 
console.log("System Loaded: v20260123_ULTRA_STABLE");

const SERVER_FILES = [
  // المحاولة بالأرقام (الحل الأضمن لتجنب مشاكل اللغة العربية في الروابط)
  'xls/1.xls', 'xls/2.xls', 'xls/3.xls', 'xls/4.xls', 'xls/5.xls', 'xls/6.xls',
  'xls/1.xlsx', 'xls/2.xlsx', 'xls/3.xlsx', 'xls/4.xlsx', 'xls/5.xlsx', 'xls/6.xlsx',
  // المحاولة بالأسماء الأصلية (كاحتياط)
  'xls/قسم الحاسوب25-26.xls', 'xls/قسم الطاقة25-26.xls', 'xls/قسم الكهرباء25-26.xls',
  'xls/قسم المحاسبة25-26.xls', 'xls/قسم المساحة25-26.xls', 'xls/قسم الميكانيكا 25-26.xls'
];

function clearCache() {
  window.__workbookCache.clear();
}

function checkLogin() {
  const user = document.getElementById('username').value.trim();
  const pass = document.getElementById('password').value.trim();
  const errorEl = document.getElementById('loginError');
  if (user.toLowerCase() === 'admin' && pass === 'admin123') {
    alert('مرحباً بك أيها المدير (Admin). سيتم إظهار خانة رفع الملفات الآن.');
    loginAs('admin');
  } else if (user && pass === '123456') {
    alert('تم الدخول كطالب برقم قيد: ' + user);
    loginAs('student', user);
  } else {
    errorEl.style.display = 'block';
    alert('خطأ: اسم مستخدم أو كلمة سر غير صحيحة.');
  }
}

function loginAs(role, id = null) {
  window.__currentUser = { role, id };
  document.getElementById('loginOverlay').style.display = 'none';
  document.getElementById('mainApp').style.display = 'block';
  const adminSections = document.getElementById('adminSections');
  const searchParts = document.getElementById('searchParts');
  const studentLinkSection = document.getElementById('studentLinkSection');
  const logoutBtn = document.getElementById('logoutBtn');
  
  if (role === 'admin') {
    if (adminSections) adminSections.style.setProperty('display', 'block', 'important');
    if (searchParts) searchParts.style.display = 'block';
    if (studentLinkSection) studentLinkSection.style.display = 'block';
    if (logoutBtn) logoutBtn.style.display = 'flex';
  } else if (role === 'public_student') {
    if (adminSections) adminSections.style.setProperty('display', 'none', 'important');
    if (searchParts) searchParts.style.display = 'block';
    if (studentLinkSection) studentLinkSection.style.display = 'none';
    if (logoutBtn) logoutBtn.style.display = 'none'; // Students using the link don't see logout
  } else {
    // Normal student login (fixed ID)
    if (adminSections) adminSections.style.setProperty('display', 'none', 'important');
    if (searchParts) searchParts.style.display = 'none';
    if (studentLinkSection) studentLinkSection.style.display = 'block';
    if (logoutBtn) logoutBtn.style.display = 'flex';
    if (id) doSearch(id);
  }
}

function logout() {
  window.__currentUser = null;
  document.getElementById('loginOverlay').style.display = 'flex';
  document.getElementById('mainApp').style.display = 'none';
  document.getElementById('result').innerHTML = '';
  document.getElementById('result').classList.remove('show');
  document.getElementById('username').value = '';
  document.getElementById('password').value = '';
  document.getElementById('loginError').style.display = 'none';
  const uploadStatus = document.getElementById('uploadStatus');
  if (uploadStatus) {
    uploadStatus.textContent = '';
    uploadStatus.style.display = 'none';
  }
  const fileInput = document.getElementById('file');
  if (fileInput) fileInput.value = ''; // Reset file input on logout
}

if (!window.__searchListenerAttached) {
  window.__searchListenerAttached = true;
  document.getElementById('loginBtn').addEventListener('click', checkLogin);
  document.getElementById('logoutBtn').addEventListener('click', logout);
  const fileInput = document.getElementById('file');
  const uploadStatus = document.getElementById('uploadStatus');
  if (fileInput) {
    fileInput.addEventListener('change', function() {
      clearCache();
      const count = this.files.length;
      if (count > 0) {
        if (uploadStatus) {
          uploadStatus.textContent = `✅ تم رفع ${count} ملفات بنجاح. المنظومة جاهزة للبحث الآن.`;
          uploadStatus.style.display = 'block';
        }
        alert(`تم رفع ${count} ملفات بنجاح. المنظومة جاهزة للبحث الآن.`);
      } else {
        if (uploadStatus) uploadStatus.style.display = 'none';
      }
    });
  }
  document.getElementById('password').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') checkLogin();
  });
  document.getElementById('searchBtn').addEventListener('click', function () {
    const studentId = document.getElementById('studentId').value.trim();
    if (!studentId) {
      alert('يرجى إدخال رقم دراسي.');
      return;
    }
    doSearch(studentId);
  });
  document.getElementById('genLinkBtn').addEventListener('click', function () {
    if (window.location.protocol === 'file:') {
      alert('⚠️ تنبيه هام:\nأنت تعمل على ملف محلي (file://).\n\nالرابط الذي سيتم توليده لن يعمل عند الطلاب إلا إذا تم رفع المشروع على استضافة ويب أو خادم محلي.\n\nإذا أرسلت هذا الرابط لطالب، لن يتمكن متصفحه من الوصول لملفات النتائج الموجودة على جهازك.');
    }
    
    const url = window.location.href.split('?')[0].split('#')[0] + '?mode=student';
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(url).then(() => {
        alert('تم نسخ رابط المنظومة بنجاح. يمكنك إرساله لزملائك.');
      }).catch(err => {
        prompt('انسخ الرابط التالي لمشاركته:', url);
      });
    } else {
      prompt('انسخ الرابط التالي لمشاركته:', url);
    }
  });
}

// Auto-login for student link
window.addEventListener('DOMContentLoaded', () => {
  const urlParams = new URLSearchParams(window.location.search);
  if (urlParams.get('mode') === 'student') {
    loginAs('public_student');
  }
});

async function doSearch(studentId) {
  if (window.__searchInProgress) return;
  if (typeof XLSX === 'undefined') {
    alert('خطأ: لم يتم تحميل مكتبة المعالجة (SheetJS).\nتأكد من اتصالك بالإنترنت أو قم بتحميل المكتبة محلياً.');
    return;
  }

  const fileInput = document.getElementById('file');
  const resultDiv = document.getElementById('result');
  
  let filesToProcess = [];
  let isServerFetch = false;

  if (fileInput && fileInput.files.length > 0) {
    filesToProcess = Array.from(fileInput.files);
  } else {
    // If no files selected, use the predefined server files
    filesToProcess = SERVER_FILES.map(path => ({ name: path, isServer: true }));
    isServerFetch = true;
  }

  window.__searchInProgress = true;
  const searchBtn = document.getElementById('searchBtn');
  const originalBtnText = searchBtn.textContent;
  searchBtn.disabled = true;
  searchBtn.textContent = 'جاري البحث...';
  
  if (isServerFetch) {
    resultDiv.innerHTML = `<p style="text-align:center; color: var(--accent-color);">جاري تحميل النتائج من الخادم... يرجى الانتظار</p>`;
    resultDiv.classList.add('show');
  }

  let foundStudent = null;
  let sheetNameText = null;
  let foundWorkbook = null;
  let scannedFilesCount = 0;
  let fetchErrors = [];

  try {
    for (let fileInfo of filesToProcess) {
      let workbook;
      const cacheKey = fileInfo.isServer ? fileInfo.name : (fileInfo.name + fileInfo.size + fileInfo.lastModified);
      
      if (window.__workbookCache.has(cacheKey)) {
        workbook = window.__workbookCache.get(cacheKey);
      } else {
        if (fileInfo.isServer) {
          // Check for file:// protocol which blocks fetch
          if (window.location.protocol === 'file:') {
            console.warn('Fetch skipped: Browser blocks local file access (CORS). Use a server or upload files.');
            continue;
          }
          // Fetch from server
          try {
            const fileNameOnly = fileInfo.name.split('/').pop();
            // جرب المسارات النسبية أولاً لتجنب مشاكل النطاق (Base URL)
            const folderPrefixes = ['xls/', './xls/', 'XLS/', '']; 
            
            let response = null;
            let lastStatus = 0;
            let lastUrlTried = "";

            // محاولة عدة صيغ لاسم الملف لضمان التوافق مع الخادم (NFC/NFD/Encoded)
            const nameVariants = [
              fileNameOnly,
              encodeURIComponent(fileNameOnly),
              fileNameOnly.normalize('NFC'),
              encodeURIComponent(fileNameOnly.normalize('NFC')),
              fileNameOnly.normalize('NFD')
            ];
            const uniqueNames = [...new Set(nameVariants)];

            outerLoop: for (const prefix of folderPrefixes) {
              for (const variant of uniqueNames) {
                const p = prefix + variant;
                lastUrlTried = p;
                try {
                  // إضافة cache: 'no-store' لضمان عدم تحميل صفحة 404 مخزنة
                  const r = await fetch(p, { cache: 'no-store' });
                  lastStatus = r.status;
                  if (r.ok) {
                    const cType = r.headers.get("content-type");
                    if (cType && cType.includes("text/html")) continue;
                    
                    response = r;
                    break outerLoop;
                  }
                } catch (e) { continue; }
              }
            }

            if (!response || !response.ok) {
              fetchErrors.push({ file: fileNameOnly, status: lastStatus, url: lastUrlTried });
              continue; 
            }
            
            const data = await response.arrayBuffer();
            workbook = XLSX.read(new Uint8Array(data), { type: 'array', raw: true });
            window.__workbookCache.set(cacheKey, workbook);
          } catch (fetchErr) {
            fetchErrors.push({ file: fileInfo.name, error: fetchErr.message });
            continue;
          }
        } else {
          // Read from input
          workbook = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
              try {
                const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', raw: true });
                window.__workbookCache.set(cacheKey, wb);
                resolve(wb);
              } catch (err) {
                reject(err);
              }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(fileInfo);
          });
        }
      }

      if (workbook) {
        scannedFilesCount++;
        for (let sName of workbook.SheetNames) {
          const sheet = workbook.Sheets[sName];
        const refStr = sheet['!ref'];
        if (!refStr) continue;
        const range = XLSX.utils.decode_range(refStr);

        let potentialHeaderRows = [];
        const maxHScan = Math.min(60, range.e.r - range.s.r + 1);
        let firstHeaderIdx = -1;
        for (let r = range.s.r; r < range.s.r + maxHScan; r++) {
          const rowData = [];
          for (let c = range.s.c; c <= range.e.c; c++) {
            const cell = sheet[XLSX.utils.encode_cell({ r: r, c: c })];
            rowData.push(cell ? (cell.v || '').toString().trim() : '');
          }
          if (rowData.some(cell => /قيد|الرقم|الطالب|اسم/i.test(cell)) ||
            rowData.some(cell => /وحدات|ساعات|مجموع|معدل|تقييم/i.test(cell))) {
            firstHeaderIdx = r;
            potentialHeaderRows.push({ rowIndex: r, data: rowData });
            break; // Found the START of the header block, now collect following rows carefully
          }
        }

        if (firstHeaderIdx !== -1) {
          const mainH = potentialHeaderRows[0];
          let idCIdx = -1, nameCIdx = -1;
          mainH.data.forEach((h, i) => {
            if (/قيد|الرقم|id|student/i.test(h)) idCIdx = i;
            if (/اسم|name/i.test(h)) nameCIdx = i;
          });

          // Capture up to 5 following rows as potential headers (stopping at students)
          for (let off = 1; off <= 5; off++) {
            const rowIdx = firstHeaderIdx + off;
            if (rowIdx >= range.e.r) break;
            const rowData = [];
            for (let c = range.s.c; c <= range.e.c; c++) {
              const cell = sheet[XLSX.utils.encode_cell({ r: rowIdx, c: c })];
              rowData.push(cell ? (cell.v || '').toString().trim() : '');
            }
            // Stop if ID-like pattern in the suspected ID column
            const idVal = idCIdx !== -1 ? normalizeId2(rowData[idCIdx]) : '';
            if (idVal && idVal.length >= 4 && /^[A-Z]*\d+$/i.test(idVal)) break;
            // Also stop if the row looks like a student name row (has text but no header-like keywords)
            if (nameCIdx !== -1 && rowData[nameCIdx] && rowData[nameCIdx].length > 4 && 
                !rowData.some(cell => /وحدات|ساعات|مجموع|معدل|تقييم|فصل|مواد/i.test(cell))) break;
            
            potentialHeaderRows.push({ rowIndex: rowIdx, data: rowData });
          }

          const startSearchRow = mainH.rowIndex + 1;
          for (let r = startSearchRow; r <= range.e.r; r++) {
            let rowFound = false;
            // First check the predicted ID column if available
            if (idCIdx !== -1) {
              const cell = sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + idCIdx })];
              if (cell && isIdMatch(studentId, cell.v)) rowFound = true;
            }
            // Fallback: search the entire row if not found in idCIdx
            if (!rowFound) {
              for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = sheet[XLSX.utils.encode_cell({ r: r, c: c })];
                if (cell && isIdMatch(studentId, cell.v)) {
                  rowFound = true;
                  break;
                }
              }
            }

            if (rowFound) {
              foundStudent = extractStudentData(sheet, r, range, potentialHeaderRows, idCIdx, nameCIdx);
              sheetNameText = sName;
              foundWorkbook = workbook;
              break;
            }
          }
        }
        if (foundStudent) break;
      }
      if (foundStudent) break;
    }
  }

  if (foundStudent) {
      displayResult(foundStudent, sheetNameText, foundWorkbook);
    } else {
      if (scannedFilesCount === 0) {
        if (window.location.protocol === 'file:') {
          resultDiv.innerHTML = `
            <div style="color:#721c24; background-color:#f8d7da; border:1px solid #f5c6cb; padding:20px; border-radius:5px; text-align:center;">
              <h3>⚠️ تعذر الوصول للملفات</h3>
              <p>أنت تحاول استخدام البحث التلقائي ولكنك تفتح الصفحة كملف محلي (file://).</p>
              <p>تمنع المتصفحات هذا الإجراء لأسباب أمنية.</p>
              <hr style="border-top:1px solid #f5c6cb; margin:10px 0;">
              <p><strong>الحل:</strong> يجب تشغيل المنظومة عبر خادم (مثل: <code>python -m http.server</code>) أو رفعها على استضافة ويب.</p>
            </div>`;
        } else {
          let diagInfo = fetchErrors.map(e => `File: ${e.file} | Status: ${e.status || 'Error'} | URL: ${e.url || e.error}`).join('<br>');
          resultDiv.innerHTML = `
            <div style="color:red; text-align:center; padding: 20px; border: 1px solid #ffccd5; background: #fff5f6; border-radius: 8px;">
              <p style="font-weight:bold; margin-bottom:10px;">تعذر الوصول إلى ملفات النتائج على الخادم.</p>
              <p style="font-size: 0.85rem; color: #666;">يرجى التأكد من رفع مجلد (xls) الذي يحتوي على ملفات الإكسل بجانب ملف index.html.</p>
              <div style="margin-top:15px; padding:10px; background:#fff; border:1px solid #ddd; font-family:monospace; font-size:0.7rem; text-align:left; direction:ltr; overflow-x:auto;">
                <b>Diagnostic Info (All attempts failed):</b><br>
                ${diagInfo || 'No files reached.'}
              </div>
            </div>
          `;
        }
      } else {
        resultDiv.innerHTML = `<p style="color:red; text-align:center;">لم يتم العثور على طالب برقم القيد: ${studentId}</p>
                               <p style="text-align:center; font-size:smaller; color:#666;">(تم البحث في ${scannedFilesCount} ملفات)</p>`;
      }
      resultDiv.classList.add('show');
    }
  } catch (err) {
    console.error("Error during search:", err);
    resultDiv.innerHTML = `
      <div style="color:red; text-align:center; padding: 20px; border: 1px solid #ffccd5; background: #fff5f6; border-radius: 8px;">
        <p style="font-weight:bold; margin-bottom:10px;">حدث خطأ أثناء معالجة الملفات.</p>
        <p style="font-size: 0.85rem; color: #666;">تفاصيل الخطأ: ${err.message || 'خطأ غير معروف'}</p>
        <p style="font-size: 0.85rem; color: #666; margin-top: 5px;">يرجى التأكد من أن ملف الإكسل غير محمي بكلمة سر ومن صحة البيانات.</p>
      </div>
    `;
    resultDiv.classList.add('show');
  }

  window.__searchInProgress = false;
  searchBtn.disabled = false;
  searchBtn.textContent = originalBtnText;
}

function extractStudentData(sheet, r, range, headerRows, idColIdx, nameColIdx) {
  const data = {
    name: '', row: [], workRow: [], finalRow: [], totalRow: [], gpaRow: [],
    headerRows: headerRows, studentRowIdx: r,
    blockMessage: ''
  };

  // 1. Try to find a "Notes" column from headers
  let notesColIdx = -1;
  headerRows.forEach(hRow => {
    hRow.data.forEach((h, i) => {
      if (/ملاحظ|تنبيه|حجب|قرار/i.test(String(h))) notesColIdx = i;
    });
  });

  // 2. Scan for block message in:
  // - Column Z (index 25) as requested
  // - Identified Notes column
  // - Any column containing blocking keywords (fallback)
  
  // تحسين كلمات الحجب لتكون أكثر دقة وتجنب الحجب الخطأ بسبب أسماء المواد
  const blockingKeywords = /مراجعة\s*(القسم|المسجل|الادارة|الإدارة|المالية)|حجب|إيقاف|موقوف|الشؤون\s*المالية|إيقاف\s*القيد/i;
  
  for (let offset = 0; offset <= 6; offset++) {
    const checkR = r + offset;
    // السماح بفحص صفوف إضافية قليلاً حتى لو تجاوزت النطاق الرسمي للجدول
    if (checkR > range.e.r + 5) break;
    
    // توسيع نطاق البحث ليشمل العمود Z (رقم 25) حتى لو كان الجدول ينتهي قبله
    const maxCol = Math.max(range.e.c, 26);
    
    for (let c = range.s.c; c <= maxCol; c++) {
      // Prioritize Column Z, Notes column, or any column with keywords
      const isPriorityCol = (c === 25 || c === notesColIdx);
      const cell = sheet[XLSX.utils.encode_cell({ r: checkR, c: c })];
      if (cell && cell.v && String(cell.v).trim()) {
        const val = String(cell.v).trim();
        
        // If it's a priority column OR contains blocking keywords, and is NOT a header
        if ((isPriorityCol || blockingKeywords.test(val)) && !/ملاحظ|نتيجة|تقدير|معدل|فصلي|المادة/i.test(val)) {
          data.blockMessage = val;
          return data; // Stop early if found
        }
      }
    }
  }

  // Label detection logic similar to Python
  let workRowIdx = null, finalRowIdx = null, totalRowIdx = null, unitsRowIdx = null;
  const labelReWork = /اعمال|أعمال/i;
  const labelReFinal = /امتحان|الأمتحان|اختبار/i;
  const labelReTotal = /المجموع|مجموع/i;
  const labelReUnits = /عدد\s*الوحدات/i;

  // Scan 8 rows for labels in columns 0 to 10 (labels can move)
  for (let rr = r; rr < Math.min(range.e.r + 1, r + 8); rr++) {
    for (let cc = range.s.c; cc <= Math.min(range.e.c, range.s.c + 10); cc++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: rr, c: cc })];
      if (!cell || !cell.v) continue;
      const lbl = String(cell.v).trim();
      if (workRowIdx === null && labelReWork.test(lbl)) workRowIdx = rr;
      if (finalRowIdx === null && labelReFinal.test(lbl)) finalRowIdx = rr;
      if (totalRowIdx === null && labelReTotal.test(lbl)) totalRowIdx = rr;
      if (unitsRowIdx === null && labelReUnits.test(lbl)) unitsRowIdx = rr;
    }
  }

  // Fallback to fixed offsets +1, +2, +3
  if (workRowIdx === null && finalRowIdx === null && totalRowIdx === null) {
    workRowIdx = r + 1;
    finalRowIdx = r + 2;
    totalRowIdx = r + 3;
    if (unitsRowIdx === null) unitsRowIdx = r;
  }
  if (unitsRowIdx === null) unitsRowIdx = r;

  for (let c = range.s.c; c <= range.e.c; c++) {
    const getVal = (ri) => {
      if (ri == null || ri > range.e.r) return '';
      const cell = sheet[XLSX.utils.encode_cell({ r: ri, c: c })];
      return cell ? cell.v : '';
    };
    data.row.push(getVal(unitsRowIdx));
    data.workRow.push(getVal(workRowIdx));
    data.finalRow.push(getVal(finalRowIdx));
    data.totalRow.push(getVal(totalRowIdx));
    data.gpaRow.push(getVal(totalRowIdx + 1)); // Heuristic for GPA/Evaluation row
  }

  // For pickName, we still want the ID row
  const studentRowData = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: r, c: c })];
    studentRowData.push(cell ? cell.v : '');
  }
  data.name = pickName(studentRowData, idColIdx, nameColIdx);
  return data;
}

function displayResult(student, sheetName, workbook) {
  const { name, row, workRow, finalRow, totalRow, gpaRow, headerRows, studentRowIdx, blockMessage } = student;
  const resultDiv = document.getElementById('result');

  // If there's a blocking message in Column Z, show only that
  if (blockMessage) {
    resultDiv.innerHTML = `
      <div style="background: #fff3cd; color: #856404; border: 1px solid #ffeeba; padding: 20px; border-radius: 8px; text-align: center; margin-top: 20px; font-weight: bold; font-size: 1.2rem;">
        <h2 style="margin-top: 0; color: #856404;">تنبيه بخصوص نتيجة الطالب: ${name}</h2>
        <p>${blockMessage}</p>
      </div>
    `;
    resultDiv.classList.add('show');
    return;
  }

  const sheet = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const studentSemester = extractSemesterNumber(sheetName);

  // Determine Grouping Info for headers
  const groupingRow = headerRows.find(h => h && h.data && h.data.some(v => v && String(v).includes("مواد")));
  const subjectNameRow = headerRows.length > 0 ? headerRows[headerRows.length - 1] : null;

  if (!subjectNameRow || !subjectNameRow.data) {
    resultDiv.innerHTML = `<p style="color:red; text-align:center;">تعذر تحديد بنية الجدول في هذا الملف (${sheetName}). يرجى التأكد من وجود عناوين للمواد.</p>`;
    resultDiv.classList.add('show');
    return;
  }

  // Find GPA Column
  let gpaColIdx = -1;
  for (const hRow of headerRows) {
    hRow.data.forEach((h, i) => {
      const v = (h || '').toString();
      if (/معدل|فصلي|GPA/i.test(v)) gpaColIdx = i;
    });
    if (gpaColIdx !== -1) break;
  }

  if (gpaColIdx === -1) {
    for (let i = gpaRow.length - 1; i >= Math.max(0, gpaRow.length - 10); i--) {
      const v = parseFloat(gpaRow[i]);
      if (!isNaN(v) && v > 0 && v <= 100) {
        gpaColIdx = i;
        break;
      }
    }
  }

  // Calculate Column-Specific "Required" Status
  const columnInfo = subjectNameRow.data.map((h, i) => {
    let groupText = "";
    if (groupingRow) {
      // Find the merged/spanning group text for this column
      for (let c = i; c >= 0; c--) {
        if (groupingRow.data[c]) { groupText = groupingRow.data[c]; break; }
      }
    }

    const normGroup = normalizeArabicText(groupText);
    const isBacklogGroup = normGroup.includes("مطالب") || normGroup.includes("بواقي") || normGroup.includes("تحميل");
    const groupSemester = extractSemesterNumber(groupText);

    let isRequired = isBacklogGroup;
    if (studentSemester && groupSemester && groupSemester !== studentSemester) isRequired = true;

    // Explicit keywords in subjects or student block
    const requireRegex = /طالب|حمل|بقي|باقي|رسب|عاده|مكمل|دور|غايب|غائب|مقصور/;
    const colCells = [row[i], workRow[i], finalRow[i], totalRow[i], gpaRow[i]];
    if (colCells.some(c => c && requireRegex.test(normalizeArabicText(String(c))))) isRequired = true;
    if (h && requireRegex.test(normalizeArabicText(h))) isRequired = true;

    return { header: h, isRequired, isGpa: (i === gpaColIdx) };
  });

  let gpaValue = (gpaColIdx !== -1 && gpaRow[gpaColIdx]) ? gpaRow[gpaColIdx] : null;
  if (gpaValue != null && !isNaN(parseFloat(gpaValue))) gpaValue = parseFloat(gpaValue).toFixed(2);

  const semesterNames = {
    1: 'الأول',
    2: 'الثاني',
    3: 'الثالث',
    4: 'الرابع',
    5: 'الخامس',
    6: 'السادس',
    7: 'السابع',
    8: 'الثامن'
  };
  const semesterText = studentSemester ? ` - الفصل ${semesterNames[studentSemester] || studentSemester}` : '';

  let tableHTML = `
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; border-bottom: 2px solid #eee; padding-bottom: 15px;">
      <h2 style="margin: 0; color: var(--primary-color); font-family: 'Amiri', serif;">نتيجة طالب: ${name}</h2>
      ${gpaValue ? `<div style="font-size: 1.1rem; font-weight: bold; background: #f8f9fa; padding: 8px 15px; border-radius: 8px; border: 1px solid #ddd; color: var(--primary-color);">المعدل الفصلي: <span style="font-size: 1.3rem; margin-right: 5px;">${gpaValue}</span></div>` : ''}
    </div>
    <p style="font-size: 1.1rem; margin-bottom: 15px;"><strong>القسم:</strong> ${extractDepartment(sheetName)}${semesterText}</p>
    <table class="result-table">
      <thead><tr>
        <th>المادة</th>
        <th>أعمال الفصل</th>
        <th>الامتحان النهائي</th>
        <th>المجموع</th>
        <th>التقدير</th>
      </tr></thead>
      <tbody>
  `;

  const seenHeaders = {};
  columnInfo.forEach((info, i) => {
    // 1. Pick the best header name and detect if it's a metadata column
    let header = "";
    let isColumnMetadata = false;
    const metadataRegex = /رقم\s*تسلسل|رقم\s*القيد|اسم\s*الطالب|تقييم|ملاحظ|تسلسل|مجموع|وحدات|ساعات|نتيجة|تقدير|معدل|فصلي|عام/i;

    headerRows.forEach(hRow => {
      let rawV = (hRow.data[i] || "").trim();
      if (rawV && metadataRegex.test(rawV)) isColumnMetadata = true;

      let v = rawV;
      // Forward-fill check for merged cells/headers
      if (!v) {
        for (let prev = i - 1; prev >= 0; prev--) {
          if (hRow.data[prev]) { v = hRow.data[prev]; break; }
        }
      }
      if (!v) return;

      const nv = normalizeArabicText(v);
      const isMetadata = metadataRegex.test(v);
      const isGrouping = nv.includes("مواد") || nv.includes("فصل");
      const isNumeric = /^\d+(\.\d+)?$/.test(v); // Numeric strings are usually units, not names
      
      if (!isMetadata && !isGrouping && !isNumeric) header = v;
    });

    if (isColumnMetadata) return; // Skip columns that are clearly metadata in any header row
    if (info.isGpa) return;
    if (!header || metadataRegex.test(header)) return;

    const finalNorm = normalizeArabicText(header);
    if (finalNorm.includes("مواد") || finalNorm.includes("فصل")) return;

    // Handle duplicates
    let displayHeader = header;
    if (seenHeaders[header]) {
      displayHeader = `${header} (${i})`;
    }
    seenHeaders[header] = true;

    // 2. Visibility Logic
    const units = parseFloat(row[i]) || 0;
    const workVal = (workRow[i] != null) ? String(workRow[i]).trim() : '';
    const finalVal = (finalRow[i] != null) ? String(finalRow[i]).trim() : '';
    const totalVal = (totalRow[i] != null) ? String(totalRow[i]).trim() : '';

    const hasGrade = (workVal && workVal !== "0" && workVal !== "0.0" && workVal !== "0.00") ||
      (finalVal && finalVal !== "0" && finalVal !== "0.0" && finalVal !== "0.00") ||
      (totalVal && totalVal !== "0" && totalVal !== "0.0" && totalVal !== "0.00");

    // Hide if units are 0 and there are no valid grades, unless explicitly required (backlog)
    // FIX: إظهار المواد المطالب بها (Backlog) حتى لو كانت الدرجة 0 أو لم تُرصد بعد
    const shouldShow = (units > 0) || hasGrade || info.isRequired;
    if (!shouldShow) return;

    if (info.isRequired) displayHeader += ' (مطالب)';

    let estimate = '-';
    const gradeNum = parseFloat(totalVal);
    if (!isNaN(gradeNum)) {
      if (gradeNum >= 85) estimate = 'ممتاز';
      else if (gradeNum >= 75) estimate = 'جيد جداً';
      else if (gradeNum >= 65) estimate = 'جيد';
      else if (gradeNum >= 50) estimate = 'مقبول';
      else estimate = 'ضعيف';
    }

    tableHTML += `
      <tr>
        <td>${displayHeader}</td>
        <td>${workVal}</td>
        <td>${finalVal}</td>
        <td style="font-weight:bold;">${totalVal}</td>
        <td class="${estimate === 'ضعيف' ? 'fail' : estimate === 'ممتاز' ? 'success' : ''}">${estimate}</td>
      </tr>
    `;
  });

  tableHTML += `</tbody></table>`;

  document.getElementById('result').innerHTML = `
    <div class="action-buttons"><button class="print-btn" onclick="window.print()">طباعة النتيجة</button></div>
    <div style="flex-grow: 1;">${tableHTML}</div>
    <div class="signature-section">
      <div class="sig-block"><span class="sig-title">المسجل العام للمعهد</span><div class="sig-line"></div><b>التوقيع</b></div>
    </div>
  `;
  document.getElementById('result').classList.add('show');
}

function extractDepartment(sheetName) {
  const norm = normalizeArabicText(sheetName);
  if (norm.includes('حاسوب')) return 'علوم الحاسوب';
  if (norm.includes('طاقه') || norm.includes('طاقات')) return 'الطاقات المتجددة';
  if (norm.includes('كهرباء')) return 'الهندسة الكهربائية';
  if (norm.includes('مساحه')) return 'المساحة';
  if (norm.includes('محاسبه')) return 'المحاسبة';
  if (norm.includes('ميكانيكا')) return 'الهندسة الميكانيكية';
  return 'غير محدد';
}

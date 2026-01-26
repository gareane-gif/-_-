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
console.log("System Loaded: v20260123_FINAL_ULTRA_FIX_WITH_BLOCKING_IMPROVEMENTS");

const SERVER_DEPT_NAMES = [
  'accounting', 
  'computer', 
  'electric', 
  'energy', 
  'mechanical', 
  'surveying'
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
  const printAllSection = document.getElementById('printAllSection');
  const academicYearSection = document.querySelector('.academic-year-section');
  const logoutBtn = document.getElementById('logoutBtn');
  
  if (role === 'admin') {
    if (adminSections) adminSections.style.setProperty('display', 'block', 'important');
    if (searchParts) searchParts.style.display = 'block';
    if (studentLinkSection) studentLinkSection.style.display = 'block';
    if (printAllSection) printAllSection.style.display = 'block';
    if (academicYearSection) academicYearSection.style.display = 'block';
    if (logoutBtn) logoutBtn.style.display = 'flex';
    
    // Load academic year settings
    loadAcademicYearSettings();
  } else if (role === 'public_student') {
    if (adminSections) adminSections.style.setProperty('display', 'none', 'important');
    if (searchParts) searchParts.style.display = 'block';
    if (studentLinkSection) studentLinkSection.style.display = 'none';
    if (printAllSection) printAllSection.style.display = 'none';
    if (academicYearSection) academicYearSection.style.display = 'none';
    if (logoutBtn) logoutBtn.style.display = 'none'; // Students using the link don't see logout
  } else {
    // Normal student login (fixed ID)
    if (adminSections) adminSections.style.setProperty('display', 'none', 'important');
    if (searchParts) searchParts.style.display = 'none';
    if (studentLinkSection) studentLinkSection.style.display = 'block';
    if (printAllSection) printAllSection.style.display = 'none';
    if (academicYearSection) academicYearSection.style.display = 'none';
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
          uploadStatus.style.color = "#0056b3";
          uploadStatus.textContent = `⏳ جاري معالجة ${count} ملفات...`;
          uploadStatus.style.display = 'block';
        }
        
        // محاكاة بسيطة ليشعر المستخدم بالعملية
        setTimeout(() => {
          if (uploadStatus) {
            uploadStatus.style.color = "#28a745";
            uploadStatus.textContent = `✅ تم تجهيز ${count} ملفات بنجاح. يمكنك البحث الآن.`;
          }
        }, 500);
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
        prompt('انسخ الرابط التالي لمشاركاته:', url);
      });
    } else {
      prompt('انسخ الرابط التالي لمشاركاته:', url);
    }
  });
  
  // Add event listeners for print options
  document.getElementById('printAllBtn').addEventListener('click', function () {
    showPrintOptionsDialog();
  });
  
  // Add event listener for print confirmation
  document.addEventListener('click', function(e) {
    if (e.target.id === 'confirmPrintBtn') {
      const includeBlocked = document.getElementById('includeBlockedCheckbox').checked;
      closePrintOptionsDialog();
      printAllDepartmentsResults(includeBlocked);
    } else if (e.target.id === 'cancelPrintBtn' || e.target.classList.contains('print-options-overlay')) {
      closePrintOptionsDialog();
    }
  });
  
  // Academic year settings event listener
  const saveAcademicYearBtn = document.getElementById('saveAcademicYearBtn');
  if (saveAcademicYearBtn) {
    saveAcademicYearBtn.addEventListener('click', saveAcademicYearSettings);
  }
  
  // Add a test data button for demonstration
  if (document.getElementById('adminSections')) {
    const testDataBtn = document.createElement('button');
    testDataBtn.innerHTML = '📊 إنشاء بيانات تجريبية';
    testDataBtn.style.cssText = 'margin: 8px 0 0; width: 100%; height: 44px; background-color: #17a2b8 !important; color: white; display: flex; align-items: center; justify-content: center; gap: 8px;';
    testDataBtn.onclick = createTestData;
    
    const testDiv = document.createElement('div');
    testDiv.id = 'testDataSection';
    testDiv.style.display = 'none';
    testDiv.style.flex = '1';
    testDiv.style.minWidth = '250px';
    testDiv.innerHTML = '<label>بيانات تجريبية:</label>';
    testDiv.appendChild(testDataBtn);
    
    document.querySelector('.input-section > div').appendChild(testDiv);
    
    // Show test section for admin
    if (window.__currentUser && window.__currentUser.role === 'admin') {
      testDiv.style.display = 'block';
    }
  }
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
    // محاولة البحث في السيرفر بهدوء دون إظهار أخطاء معقدة
    filesToProcess = SERVER_DEPT_NAMES.map(name => ({ name: name, isServer: true }));
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
          if (window.location.protocol === 'file:') {
            console.warn('Fetch skipped: Browser blocks local file access (CORS). Use a server or upload files.');
            continue;
          }
          
          try {
            const baseName = fileInfo.name;
            let response = null;
            let lastStatus = 0;
            let lastUrlTried = "";

            // توليد قائمة بكل الاحتمالات الممكنة للرابط (تغطية شاملة لحالة الأحرف)
            const possibleFolders = ['xls/', 'XLS/', './xls/', ''];
            const possibleNames = [baseName, baseName.toLowerCase(), baseName.toUpperCase(), baseName.charAt(0).toUpperCase() + baseName.slice(1)];
            const possibleExts = ['.xlsx', '.xls', '.XLSX', '.XLS', ''];
            const uniqueUrls = new Set();
            
            for (const f of possibleFolders) {
              for (const n of possibleNames) {
                for (const e of possibleExts) {
                  uniqueUrls.add(f + n + e);
                }
              }
            }

            // محاولة كل الاحتمالات حتى النجاح
            outerLoop: for (const targetUrl of uniqueUrls) {
              lastUrlTried = targetUrl;
              try {
                // استخدام رابط مطلق بناءً على موقع الصفحة لضمان الدقة في GitHub Pages
                const absoluteUrl = new URL(targetUrl, window.location.href).href;
                const r = await fetch(absoluteUrl, { cache: 'no-store' });
                lastStatus = r.status;
                if (r.ok) {
                  const cType = r.headers.get("content-type");
                  if (cType && cType.includes("text/html")) continue;
                  response = r;
                  break outerLoop;
                }
              } catch (e) { continue; }
            }

            if (!response || !response.ok) {
              fetchErrors.push({ file: baseName, status: lastStatus, url: lastUrlTried });
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
            console.log("Found header at row:", r + 1, "Data:", rowData);
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
              console.log("Student found at Excel row:", r + 1);
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
        if (isServerFetch) {
          resultDiv.innerHTML = `
            <div style="text-align:center; padding: 20px; border: 1px solid #ddd; background: #f9f9f9; border-radius: 8px;">
              <p style="font-weight:bold; color: #555;">تعذر العثور على ملفات النتائج على السيرفر، أو لم يتم رفع ملفات من الجهاز.</p>
              <p style="font-size: 0.9rem; color: #888; margin-top:10px;"><b>للمدير:</b> يرجى التأكد من رفع مجلد <code>xls</code> على GitHub وبداخله الملفات.</p>
              <p style="font-size: 0.9rem; color: #888;"><b>للطالب:</b> يرجى مراجعة إدارة المعهد للتأكد من جاهزية النتائج إلكترونياً.</p>
              <div style="margin-top:15px; font-size:0.7rem; color:#bbb; cursor:pointer;" onclick="this.nextElementSibling.style.display='block'">+ عرض تفاصيل الخطأ التقني</div>
              <div style="display:none; text-align:left; direction:ltr; font-family:monospace; font-size:0.65rem; background:#eee; padding:5px; margin-top:5px;">
                Tried paths for first file: ${fetchErrors.length > 0 ? fetchErrors[0].url : 'None'}<br>
                Status: ${fetchErrors.length > 0 ? fetchErrors[0].status : 'No attempts'}
              </div>
            </div>
          `;
        }
      } else {
        resultDiv.innerHTML = `<p style="color:red; text-align:center;">لم يتم العثور على طالب برقم القيد: ${studentId}</p>
                               <p style="text-align:center; font-size:smaller; color:#666;">(تم البحث في ${scannedFilesCount} ملفات تم رفعها)</p>`;
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

  // 1. البحث الذكي عن عمود الملاحظات من الصفوف 5-9 (أي من 4 إلى 8 في النظام صفر-based)
  let notesColIdx = -1;
  
  // أولاً: البحث في الصفوف 5-9 (Excel rows 5-9) للعثور على "ملاحظ" أو "ملاحظات" أو "ملاحظة"
  for (let rowIdx = 4; rowIdx <= 8 && notesColIdx === -1; rowIdx++) {
    if (rowIdx <= range.e.r) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = sheet[XLSX.utils.encode_cell({ r: rowIdx, c: c })];
        if (cell && cell.v && /ملاحظ|ملاحظات|ملاحظة|تنبيه|حجب|قرار/i.test(normalizeArabicText(String(cell.v)))) {
          notesColIdx = c;
          console.log("Found notes column at index:", c, "in row:", rowIdx, "with value:", cell.v);
          break; // وجدناه في هذا الصف، لا نحتاج لفحص باقي الأعمدة
        }
      }
    }
  }
  
  // ثانياً: إذا لم يجد في الصفوف 5-9، يبحث في كل صفوف الهيدر
  if (notesColIdx === -1) {
    headerRows.forEach(hRow => {
      hRow.data.forEach((h, i) => {
        if (notesColIdx === -1 && /ملاحظ|ملاحظات|ملاحظة|تنبيه|حجب|قرار/i.test(normalizeArabicText(String(h)))) {
          notesColIdx = i;
          console.log("Found notes column in header at index:", i, "with value:", h);
        }
      });
    });
  }

  // 2. إذا وجد عمود الملاحظات، اقرأ القيمة مباشرة من صف الطالب
  if (notesColIdx !== -1) {
    console.log("Checking for note in column:", notesColIdx, "starting from row:", r, "(Excel row:", r+1, ")");
    for (let offset = 0; offset <= 3; offset++) {
      const checkR = r + offset;
      if (checkR > range.e.r) break;
      
      const cell = sheet[XLSX.utils.encode_cell({ r: checkR, c: notesColIdx })];
      if (cell && cell.v && String(cell.v).trim()) {
        const val = String(cell.v).trim();
        console.log("Cell value at row", checkR, "col", notesColIdx, ":", val);
        
        // تحسين: تحقق مما إذا كانت القيمة تحتوي على كلمات حجب أو كلمات عادية (غير رأس عمود)
        if (/حجب|منع|إيقاف|إيقاف مؤقت|حظر|موقوف|ممنوع/i.test(val)) {
          console.log("Found block message:", val);
          data.blockMessage = val;
          return data; // توقف فوراً إذا وجدت كلمة حجب
        } else if (!/ملاحظ|تنبيه|نتيجة|تقدير|معدل|فصلي|المادة/i.test(val)) {
          // إذا لم تكن كلمة حجب، لكنها ليست أيضًا عنوان عمود، اعتبرها ملاحظة عامة
          console.log("Found general note:", val);
          data.blockMessage = val;
          return data; // توقف فوراً إذا وجد ملاحظة
        }
      }
    }
  }

  // 3. Fallback: البحث في العمود Z (رقم 25) كما كان سابقاً (للتوافقية مع الملفات القديمة)
  console.log("Checking fallback Z column (index 25) starting from row:", r, "(Excel row:", r+1, ")");
  for (let offset = 0; offset <= 6; offset++) {
    const checkR = r + offset;
    if (checkR > range.e.r + 5) break;
    
    const cellZ = sheet[XLSX.utils.encode_cell({ r: checkR, c: 25 })];
    if (cellZ && cellZ.v && String(cellZ.v).trim()) {
      const valZ = String(cellZ.v).trim();
      console.log("Cell value at row", checkR, "col Z (25):", valZ);
      
      // تحسين: تحقق مما إذا كانت القيمة تحتوي على كلمات حجب أو كلمات عادية (غير رأس عمود)
      if (/حجب|منع|إيقاف|إيقاف مؤقت|حظر|موقوف|ممنوع/i.test(valZ)) {
        console.log("Found block message in Z column:", valZ);
        data.blockMessage = valZ;
        return data; // توقف فوراً إذا وجدت كلمة حجب في Z
      } else if (!/ملاحظ|نتيجة|تقدير|معدل|فصلي|المادة/i.test(valZ)) {
        // إذا لم تكن كلمة حجب، لكنها ليست أيضًا عنوان عمود، اعتبرها ملاحظة عامة
        console.log("Found general note in Z column:", valZ);
        data.blockMessage = valZ;
        return data; // توقف فوراً إذا وجد ملاحظة في Z
      }
    }
  }

  // Label detection logic (units, work, final, total)
  let workRowIdx = null, finalRowIdx = null, totalRowIdx = null, unitsRowIdx = null;
  const labelReWork = /اعمال|أعمال/i;
  const labelReFinal = /امتحان|الأمتحان|اختبار/i;
  const labelReTotal = /المجموع|مجموع/i;
  const labelReUnits = /عدد\s*الوحدات|وحدات|ساعات/i;

  for (let rr = Math.max(0, r - 5); rr < Math.min(range.e.r + 1, r + 15); rr++) {
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

  // Smart fallback for units row: look for a row with numeric values (1-10) in subject columns
  if (unitsRowIdx === null) {
    for (let rr = Math.max(0, r - 5); rr < Math.min(range.e.r + 1, r + 5); rr++) {
      let numericCount = 0;
      for (let cc = range.s.c + 5; cc <= range.e.c; cc++) {
        const cell = sheet[XLSX.utils.encode_cell({ r: rr, c: cc })];
        if (cell && typeof cell.v === 'number' && cell.v > 0 && cell.v < 10) numericCount++;
      }
      if (numericCount > 3) { unitsRowIdx = rr; break; }
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
  
  // Get academic year from control panel settings
  const academicYearSettings = localStorage.getItem('academicYearSettings');
  let academicYear = 'خريف 2025-2026'; // Default value
  if (academicYearSettings) {
    try {
      const settings = JSON.parse(academicYearSettings);
      if (settings.academicYearText) {
        academicYear = settings.academicYearText;
      }
    } catch (e) {
      console.error('Error parsing academic year settings:', e);
    }
  }

  let tableHTML = `
    <div style="position: relative; padding-top: 120px;">
      <div style="position: absolute; top: 20px; left: 20px; width: 80px; height: 80px; opacity: 0.3; z-index: 1; pointer-events: none;">
        <img src="278143110_656758089110417_4349344937401225995_n.png" alt="شعار المعهد" style="width: 100%; height: 100%; object-fit: contain;" />
      </div>
      <div style="text-align: center; margin-bottom: 20px; margin-top: 20px;">
        <h2 style="margin: 0; color: var(--primary-color); font-family: 'Amiri', serif; font-size: 1.5rem;">المعهد العالي للعلوم والتقنية الجفرة بسوكنة</h2>
      </div>
      <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; border-bottom: 2px solid #eee; padding-bottom: 15px;">
        <h2 style="margin: 0; color: var(--primary-color); font-family: 'Amiri', serif;">نتيجة الطالب: ${name}</h2>
        ${gpaValue ? `<div style="font-size: 1.1rem; font-weight: bold; background: #f8f9fa; padding: 8px 15px; border-radius: 8px; border: 1px solid #ddd; color: var(--primary-color);">المعدل الفصلي: <span style="font-size: 1.3rem; margin-right: 5px;">${gpaValue}</span></div>` : ''}
      </div>
      <p style="font-size: 1.1rem; margin-bottom: 15px;"><strong>القسم:</strong> ${extractDepartment(sheetName)}${semesterText}</p>
      <p style="font-size: 1.0rem; margin-bottom: 15px;"><strong>الفصل الدراسي:</strong> ${academicYear}</p>
    </div>
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

    // 2. Visibility Logic: عرض المادة إذا كان لها درجة أو عدد وحدات أكبر من 0
    const rawUnit = row[i];
    const units = parseFloat(rawUnit) || 0;
    const workVal = (workRow[i] != null) ? String(workRow[i]).trim() : '';
    const finalVal = (finalRow[i] != null) ? String(finalRow[i]).trim() : '';
    const totalVal = (totalRow[i] != null) ? String(totalRow[i]).trim() : '';

    const hasGrade = (workVal && workVal !== "0") || (finalVal && finalVal !== "0") || (totalVal && totalVal !== "0");
    
    // إذا لم تكن هناك درجة ولم تكن هناك وحدات، نتجاهل المادة
    if (units === 0 && !hasGrade) return;

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
    <div class="action-buttons"><button class="print-btn" onclick="printResult()">طباعة النتيجة</button></div>
    <div style="position: relative; padding-top: 60px;">
      <div style="flex-grow: 1;">
        <div class="table-container">
          ${tableHTML}
        </div>
      </div>
    </div>
  `;
  document.getElementById('result').classList.add('show');
}

function printResult() {
  // Get the current result content
  const resultDiv = document.getElementById('result');
  
  // Create a simplified content with only essential elements
  const tempDiv = document.createElement('div');
  tempDiv.innerHTML = resultDiv.innerHTML;
  
  // Remove the action buttons (print button)
  const actionButtons = tempDiv.querySelector('.action-buttons');
  if (actionButtons) actionButtons.remove();
  
  // Extract the table container
  const tableContainer = tempDiv.querySelector('.table-container');
  const tableElement = tableContainer ? tableContainer.querySelector('.result-table') : tempDiv.querySelector('.result-table');
  
  // Build content without any institutional branding that might be in the original
  let contentHtml = '';
  
  // Add table
  if (tableElement) {
    // If we have a table container, use it; otherwise use just the table
    if (tableContainer) {
      contentHtml += tableContainer.outerHTML;
    } else {
      contentHtml += `<div class="table-container">${tableElement.outerHTML}</div>`;
    }
  }
  
  // Add signature section
  contentHtml += `
    <div class="signature-section">
      <div class="registrar-text">المسجل العام للمعهد</div>
      <div class="signature-line"></div>
      <b>التوقيع</b>
    </div>
  `;
  
  // Open the print.html file in a new window
  const printWindow = window.open('', '_blank', 'width=800,height=600');
  
  // Write the HTML content directly to the window to avoid URL in print
  printWindow.document.write(`
    <!DOCTYPE html>
    <html lang="ar" dir="rtl">
    <head>
      <meta charset="UTF-8">
      <title>طباعة نتيجة الطالب</title>
      <style>
        @page {
          size: A4 portrait;
          margin: 0.5cm; /* Reduced margin to prevent two-page issue */
        }
        body { 
          font-family: 'Amiri', serif; 
          margin: 0.5cm; 
          background: white;
          color: black;
        }
        .page-container {
          position: relative;
          min-height: 800px;
          padding-top: 20px; /* Reduced padding since no header */
        }
        .table-container {
          overflow-x: auto;
          width: 100%;
          margin: 20px 0;
        }
        
        .result-table {
          width: 100%;
          border-collapse: collapse;
          margin: 0;
          font-size: 14px;
          min-width: 100%;
        }
        .result-table th, .result-table td {
          border: 1px solid #333;
          padding: 8px;
          text-align: center;
        }
        .result-table th {
          background-color: #003366 !important;
          color: white !important;
          font-weight: bold;
          -webkit-print-color-adjust: exact !important;
          print-color-adjust: exact !important;
        }
        .result-table tr:nth-child(even) {
          background-color: #f9f9f9;
          -webkit-print-color-adjust: exact !important;
          print-color-adjust: exact !important;
        }
        .signature-section {
          margin-top: 60px;
          text-align: center;
        }
        .registrar-text {
          font-weight: bold;
          margin-bottom: 50px;
          font-family: 'Amiri', serif;
          font-size: 1.1rem;
        }
        .signature-line {
          width: 200px;
          margin: 20px auto;
          border-top: 1px solid #000;
        }
        .signature-section {
          margin-top: 30px; /* Reduced margin to prevent two-page issue */
          text-align: center;
        }
      </style>
    </head>
    <body>
      <div class="page-container">
        <div class="content" id="content">
          ${contentHtml}
        </div>
      </div>
    </body>
    </html>
  `);
  
  printWindow.document.close();
  printWindow.focus();
  
  // Print after a short delay to ensure content is loaded
  setTimeout(() => {
    printWindow.print();
    printWindow.close();
  }, 500);
}



// Print Options Dialog Functions
function showPrintOptionsDialog() {
  // Create overlay
  const overlay = document.createElement('div');
  overlay.className = 'print-options-overlay';
  overlay.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: rgba(0, 0, 0, 0.5);
    display: flex;
    justify-content: center;
    align-items: center;
    z-index: 10000;
  `;
  
  // Create dialog
  const dialog = document.createElement('div');
  dialog.style.cssText = `
    background: white;
    padding: 30px;
    border-radius: 12px;
    width: 90%;
    max-width: 500px;
    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
    text-align: center;
  `;
  
  dialog.innerHTML = `
    <h3 style="color: var(--primary-color); margin-bottom: 20px; font-family: 'Amiri', serif;">خيارات الطباعة</h3>
    <p style="margin-bottom: 25px; color: #555;">اختر خيارات الطباعة لنتائج جميع الأقسام:</p>
    
    <div style="margin-bottom: 25px; text-align: right;">
      <label style="display: flex; align-items: center; gap: 10px; margin-bottom: 15px; cursor: pointer;">
        <input type="checkbox" id="includeBlockedCheckbox" checked style="transform: scale(1.3);">
        <span>تضمين النتائج المحجوبة (_blocked results_)</span>
      </label>
      <p style="font-size: 0.9rem; color: #777; margin: 5px 30px 0 0; text-align: right;">
        إذا تم إلغاء التحديد، سيتم استثناء النتائج التي تحتوي على رسائل حجب أو ملاحظات خاصة
      </p>
    </div>
    
    <div style="display: flex; gap: 15px; justify-content: center; margin-top: 30px;">
      <button id="confirmPrintBtn" style="
        background-color: var(--accent-color) !important;
        color: white;
        border: none;
        padding: 12px 25px;
        border-radius: 6px;
        cursor: pointer;
        font-weight: bold;
        font-size: 1rem;
      ">طباعة النتائج</button>
      
      <button id="cancelPrintBtn" style="
        background-color: #6c757d !important;
        color: white;
        border: none;
        padding: 12px 25px;
        border-radius: 6px;
        cursor: pointer;
        font-weight: bold;
        font-size: 1rem;
      ">إلغاء</button>
    </div>
  `;
  
  overlay.appendChild(dialog);
  document.body.appendChild(overlay);
}

function closePrintOptionsDialog() {
  const overlay = document.querySelector('.print-options-overlay');
  if (overlay) {
    overlay.remove();
  }
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

// Print all departments results with cover sheets
async function printAllDepartmentsResults(includeBlocked = true) {
  if (typeof XLSX === 'undefined') {
    alert('خطأ: لم يتم تحميل مكتبة المعالجة (SheetJS).\nتأكد من اتصالك بالإنترنت أو قم بتحميل المكتبة محلياً.');
    return;
  }

  const fileInput = document.getElementById('file');
  const resultDiv = document.getElementById('result');
  
  let filesToProcess = [];
  let isServerFetch = false;
  let useMockData = false;

  if (fileInput && fileInput.files.length > 0) {
    filesToProcess = Array.from(fileInput.files);
  } else if (window.__workbookCache.has('mock_data')) {
    // Use mock data if available
    useMockData = true;
    filesToProcess = [{ name: 'mock_data', isMock: true }];
  } else {
    // محاولة البحث في السيرفر بهدوء دون إظهار أخطاء معقدة
    filesToProcess = SERVER_DEPT_NAMES.map(name => ({ name: name, isServer: true }));
    isServerFetch = true;
  }

  if (filesToProcess.length === 0) {
    alert('لا توجد ملفات نتائج لطباعة نتائج الأقسام. يرجى رفع ملفات Excel أولاً.');
    return;
  }

  resultDiv.innerHTML = '<p style="text-align:center; color: var(--accent-color);">جاري تجهيز طباعة نتائج جميع الأقسام... يرجى الانتظار</p>';
  resultDiv.classList.add('show');

  // Hide non-print elements temporarily
  const originalDisplay = {};
  const hideSelectors = ['.upload-section', '.input-section', '.action-buttons', '.debug-div', '.logout-btn'];
  hideSelectors.forEach(selector => {
    const elements = document.querySelectorAll(selector);
    elements.forEach(el => {
      originalDisplay[el.id || el.className] = el.style.display;
      el.style.display = 'none';
    });
  });

  try {
    let allResultsHTML = '';
    let processedCount = 0;
    
    for (let fileInfo of filesToProcess) {
      let workbook;
      const cacheKey = fileInfo.isServer ? fileInfo.name : (fileInfo.name + fileInfo.size + fileInfo.lastModified);
      
      if (window.__workbookCache.has(cacheKey)) {
        workbook = window.__workbookCache.get(cacheKey);
      } else {
        if (fileInfo.isMock) {
          workbook = window.__workbookCache.get('mock_data');
          console.log("Using mock data workbook");
        } else if (fileInfo.isServer) {
          if (window.location.protocol === 'file:') {
            console.warn('Fetch skipped: Browser blocks local file access (CORS). Use a server or upload files.');
            continue;
          }
          
          try {
            const baseName = fileInfo.name;
            let response = null;
            let lastStatus = 0;
            let lastUrlTried = "";

            // توليد قائمة بكل الاحتمالات الممكنة للرابط (تغطية شاملة لحالة الأحرف)
            const possibleFolders = ['xls/', 'XLS/', './xls/', ''];
            const possibleNames = [baseName, baseName.toLowerCase(), baseName.toUpperCase(), baseName.charAt(0).toUpperCase() + baseName.slice(1)];
            const possibleExts = ['.xlsx', '.xls', '.XLSX', '.XLS', ''];
            const uniqueUrls = new Set();
            
            for (const f of possibleFolders) {
              for (const n of possibleNames) {
                for (const e of possibleExts) {
                  uniqueUrls.add(f + n + e);
                }
              }
            }

            // محاولة كل الاحتمالات حتى النجاح
            outerLoop: for (const targetUrl of uniqueUrls) {
              lastUrlTried = targetUrl;
              try {
                // استخدام رابط مطلق بناءً على موقع الصفحة لضمان الدقة في GitHub Pages
                const absoluteUrl = new URL(targetUrl, window.location.href).href;
                const r = await fetch(absoluteUrl, { cache: 'no-store' });
                lastStatus = r.status;
                if (r.ok) {
                  const cType = r.headers.get("content-type");
                  if (cType && cType.includes("text/html")) continue;
                  response = r;
                  break outerLoop;
                }
              } catch (e) { continue; }
            }

            if (!response || !response.ok) {
              continue; 
            }
            
            const data = await response.arrayBuffer();
            workbook = XLSX.read(new Uint8Array(data), { type: 'array', raw: true });
            window.__workbookCache.set(cacheKey, workbook);
          } catch (fetchErr) {
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
        processedCount++;
        
        // Generate cover page for this department
        const deptName = extractDepartment(workbook.SheetNames[0]);
        console.log("Processing department:", deptName, "with", workbook.SheetNames.length, "sheets");
        
        // Get academic year from control panel settings
        const academicYearSettings = localStorage.getItem('academicYearSettings');
        let academicYearText = 'خريف 2025-2026'; // Default value
        if (academicYearSettings) {
          try {
            const settings = JSON.parse(academicYearSettings);
            if (settings.academicYearText) {
              academicYearText = settings.academicYearText;
            }
          } catch (e) {
            console.error('Error parsing academic year settings:', e);
          }
        }
        
        allResultsHTML += '<div class="department-cover" style="position: relative; padding-top: 120px; page-break-before: always; text-align: center; padding: 50px 20px; background-color: white;">' +
                         '<div style="position: absolute; top: 20px; left: 20px; width: 80px; height: 80px; opacity: 0.3; z-index: 1; pointer-events: none;">' +
                         '<img src="278143110_656758089110417_4349344937401225995_n.png" alt="شعار المعهد" style="width: 100%; height: 100%; object-fit: contain;" />' +
                         '</div>' +
                         '<div style="text-align: center; margin-bottom: 20px; margin-top: 20px;">' +
                         '<h2 style="margin: 0; color: var(--primary-color); font-family: \'Amiri\', serif; font-size: 1.5rem;">المعهد العالي للعلوم والتقنية الجفرة بسوكنة</h2>' +
                         '</div>' +
                         '<h1 style="color: var(--primary-color); margin-bottom: 30px; font-size: 2rem;">' + deptName + '</h1>' +
                         '<h2 style="color: var(--accent-color); margin-bottom: 50px; font-size: 1.5rem;">نتائج الطلاب</h2>' +
                         '<div style="font-size: 1.2rem; margin-top: 50px;">' +
                           '<p>السنة الدراسية: ' + academicYearText + '</p>' +
                           '<p>عدد الجداول: ' + workbook.SheetNames.length + '</p>' +
                         '</div>' +
                       '</div>';
        
        // Process each sheet in the workbook
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
            console.log("Main header data:", mainH.data);
            let idCIdx = -1, nameCIdx = -1;
            mainH.data.forEach((h, i) => {
              if (/قيد|الرقم|id|student/i.test(h)) idCIdx = i;
              if (/اسم|name/i.test(h)) nameCIdx = i;
            });
            console.log("ID column index:", idCIdx, "Name column index:", nameCIdx);

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
              // Check if this row contains a valid student ID
              let hasValidId = false;
              let idValue = '';
              
              // First check the predicted ID column if available
              if (idCIdx !== -1) {
                const cell = sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + idCIdx })];
                if (cell && cell.v) {
                  idValue = cell.v;
                  const normalizedId = normalizeId2(idValue);
                  // Check if it looks like a valid student ID (has letters and numbers)
                  if (normalizedId && normalizedId.length >= 4 && /^[A-Z]*\d+[A-Z]*\d*$/.test(normalizedId)) {
                    hasValidId = true;
                  }
                }
              }
              
              // Fallback: search the entire row for any cell that looks like an ID
              if (!hasValidId) {
                for (let c = range.s.c; c <= range.e.c; c++) {
                  const cell = sheet[XLSX.utils.encode_cell({ r: r, c: c })];
                  if (cell && cell.v) {
                    const normalizedId = normalizeId2(cell.v);
                    if (normalizedId && normalizedId.length >= 4 && /^[A-Z]*\d+[A-Z]*\d*$/.test(normalizedId)) {
                      hasValidId = true;
                      idValue = cell.v;
                      break;
                    }
                  }
                }
              }

              if (hasValidId) {
                console.log("Found student at row:", r + 1, "ID:", idValue);
                const foundStudent = extractStudentData(sheet, r, range, potentialHeaderRows, idCIdx, nameCIdx);
                console.log("Extracted student data:", foundStudent);
                
                // Check if student has blocked message and handle according to includeBlocked setting
                if (foundStudent.blockMessage && !includeBlocked) {
                  console.log("Skipping blocked student:", foundStudent.name);
                  continue; // Skip this student if blocked results should be excluded
                }
                
                const studentResult = generateStudentResultHTML(foundStudent, sName, workbook, includeBlocked);
                console.log("Generated student result HTML length:", studentResult.length);
                allResultsHTML += '<div class="student-result" style="page-break-before: always; padding: 20px; background-color: white;">' +
                                 studentResult +
                                 '</div>';
              } else {
                console.log("No valid ID found at row:", r + 1);
              }
            }
          }
        }
      }
    }
    
    if (processedCount > 0) {
      resultDiv.innerHTML = allResultsHTML;
      
      // Ensure content is visible for printing
      resultDiv.classList.add('show');
      
      // Force a reflow to ensure content is rendered
      void resultDiv.offsetWidth;
      
      // Trigger print after a brief delay to ensure content is rendered
      setTimeout(() => {
        window.print();
        
        // Restore hidden elements after print dialog is closed
        setTimeout(() => {
          hideSelectors.forEach(selector => {
            const elements = document.querySelectorAll(selector);
            elements.forEach(el => {
              const idOrClass = el.id || el.className;
              if (originalDisplay[idOrClass] !== undefined) {
                el.style.display = originalDisplay[idOrClass];
              }
            });
          });
        }, 1000);
      }, 500);
    } else {
      resultDiv.innerHTML = '<p style="color:red; text-align:center;">لم يتم العثور على ملفات نتائج لطباعة نتائج الأقسام.</p>';
      resultDiv.classList.add('show');
    }
  } catch (err) {
    console.error("Error during print all departments:", err);
    resultDiv.innerHTML = '<p style="color:red; text-align:center;">حدث خطأ أثناء تجهيز نتائج الأقسام: ' + err.message + '</p>';
  }
}

// Helper function to generate student result HTML
function generateStudentResultHTML(student, sheetName, workbook, includeBlocked = true) {
  const { name, row, workRow, finalRow, totalRow, gpaRow, headerRows, studentRowIdx, blockMessage } = student;

  // If there's a blocking message in Column Z, handle based on includeBlocked setting
  if (blockMessage) {
    if (includeBlocked) {
      // Show blocked message with registrar info
      return '<div style="position: relative; padding-top: 120px;">' +
             '<div style="position: absolute; top: 20px; left: 20px; width: 80px; height: 80px; opacity: 0.3; z-index: 1; pointer-events: none;">' +
             '<img src="278143110_656758089110417_4349344937401225995_n.png" alt="شعار المعهد" style="width: 100%; height: 100%; object-fit: contain;" />' +
             '</div>' +
             '<div style="text-align: center; margin-bottom: 20px; margin-top: 20px;">' +
             '<h2 style="margin: 0; color: var(--primary-color); font-family: \'Amiri\', serif; font-size: 1.5rem;">المعهد العالي للعلوم والتقنية الجفرة بسوكنة</h2>' +
             '</div>' +
             '<div style="background: #fff3cd; color: #856404; border: 1px solid #ffeeba; padding: 20px; border-radius: 8px; text-align: center; margin-top: 20px; font-weight: bold; font-size: 1.2rem;">' +
             '<h2 style="margin-top: 0; color: #856404;">تنبيه بخصوص نتيجة الطالب: ' + name + '</h2>' +
             '<p>' + blockMessage + '</p>' +
             '</div>' +
             '<div style="margin-top: 40px; text-align: center;">' +
             '<div style="display: flex; justify-content: space-around; margin-top: 50px;">' +
               '<div style="text-align: center; min-width: 200px;">' +
                 '<span style="display: block; font-weight: bold; margin-bottom: 50px; font-family: \'Amiri\', serif;">المسجل العام للمعهد</span>' +
                 '<div style="border-top: 1px solid #000; width: 150px; margin: 0 auto;"></div>' +
                 '<b style="display: block; margin-top: 10px;">التوقيع</b>' +
               '</div>' +
             '</div>' +
             '</div>' +
             '</div>';
    } else {
      // This shouldn't happen since blocked students are filtered out when includeBlocked=false
      return '<div style="display: none;"></div>';
    }
  }

  const sheet = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const studentSemester = extractSemesterNumber(sheetName);

  // Determine Grouping Info for headers
  const groupingRow = headerRows.find(h => h && h.data && h.data.some(v => v && String(v).includes("مواد")));
  const subjectNameRow = headerRows.length > 0 ? headerRows[headerRows.length - 1] : null;

  if (!subjectNameRow || !subjectNameRow.data) {
    return '<p style="color:red; text-align:center;">تعذر تحديد بنية الجدول في هذا الملف (' + sheetName + '). يرجى التأكد من وجود عناوين للمواد.</p>';
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
      // Find the merged-spanning group text for this column
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
  const semesterText = studentSemester ? ' - الفصل ' + (semesterNames[studentSemester] || studentSemester) : '';
  
  // Get academic year from control panel settings
  const academicYearSettings = localStorage.getItem('academicYearSettings');
  let academicYear = 'خريف 2025-2026'; // Default value
  if (academicYearSettings) {
    try {
      const settings = JSON.parse(academicYearSettings);
      if (settings.academicYearText) {
        academicYear = settings.academicYearText;
      }
    } catch (e) {
      console.error('Error parsing academic year settings:', e);
    }
  }

  let tableHTML = '<div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; border-bottom: 2px solid #eee; padding-bottom: 15px;">' +
               '<h2 style="margin: 0; color: var(--primary-color); font-family: \'Amiri\', serif;">نتيجة الطالب: ' + name + '</h2>' +
               (gpaValue ? '<div style="font-size: 1.1rem; font-weight: bold; background: #f8f9fa; padding: 8px 15px; border-radius: 8px; border: 1px solid #ddd; color: var(--primary-color);">المعدل الفصلي: <span style="font-size: 1.3rem; margin-right: 5px;">' + gpaValue + '</span></div>' : '') +
               '</div>' +
               '<p style="font-size: 1.1rem; margin-bottom: 15px;"><strong>القسم:</strong> ' + extractDepartment(sheetName) + semesterText + '</p>' +
               '<p style="font-size: 1.0rem; margin-bottom: 15px;"><strong>الفصل الدراسي:</strong> ' + academicYear + '</p>' +
               '<table class="result-table">' +
                 '<thead><tr>' +
                   '<th>المادة</th>' +
                   '<th>أعمال الفصل</th>' +
                   '<th>الامتحان النهائي</th>' +
                   '<th>المجموع</th>' +
                   '<th>التقدير</th>' +
                 '</tr></thead>' +
                 '<tbody>';

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
      displayHeader = header + ' (' + i + ')';
    }
    seenHeaders[header] = true;

    // 2. Visibility Logic: عرض المادة إذا كان لها درجة أو عدد وحدات أكبر من 0
    const rawUnit = row[i];
    const units = parseFloat(rawUnit) || 0;
    const workVal = (workRow[i] != null) ? String(workRow[i]).trim() : '';
    const finalVal = (finalRow[i] != null) ? String(finalRow[i]).trim() : '';
    const totalVal = (totalRow[i] != null) ? String(totalRow[i]).trim() : '';

    const hasGrade = (workVal && workVal !== "0") || (finalVal && finalVal !== "0") || (totalVal && totalVal !== "0");
    
    // إذا لم تكن هناك درجة ولم تكن هناك وحدات، نتجاهل المادة
    if (units === 0 && !hasGrade) return;

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

    tableHTML += '<tr>' +
               '<td>' + displayHeader + '</td>' +
               '<td>' + workVal + '</td>' +
               '<td>' + finalVal + '</td>' +
               '<td style="font-weight:bold;">' + totalVal + '</td>' +
               '<td class="' + (estimate === 'ضعيف' ? 'fail' : estimate === 'ممتاز' ? 'success' : '') + '">' + estimate + '</td>' +
               '</tr>';
  });

  tableHTML += '</tbody></table>';
  
  return '<div style="flex-grow: 1; position: relative; padding-top: 120px;">' +
         '<div style="position: absolute; top: 20px; left: 20px; width: 80px; height: 80px; opacity: 0.3; z-index: 1; pointer-events: none;">' +
         '<img src="278143110_656758089110417_4349344937401225995_n.png" alt="شعار المعهد" style="width: 100%; height: 100%; object-fit: contain;" />' +
         '</div>' +
         '<div style="text-align: center; margin-bottom: 20px; margin-top: 20px;">' +
         '<h2 style="margin: 0; color: var(--primary-color); font-family: \'Amiri\', serif; font-size: 1.5rem;">المعهد العالي للعلوم والتقنية الجفرة بسوكنة</h2>' +
         '</div>' +
         tableHTML +
         '<div style="margin-top: 40px; text-align: center;">' +
         '<div style="display: flex; justify-content: space-around; margin-top: 50px;">' +
           '<div style="text-align: center; min-width: 200px;">' +
             '<span style="display: block; font-weight: bold; margin-bottom: 50px; font-family: \'Amiri\', serif;">المسجل العام للمعهد</span>' +
             '<div style="border-top: 1px solid #000; width: 150px; margin: 0 auto;"></div>' +
             '<b style="display: block; margin-top: 10px;">التوقيع</b>' +
           '</div>' +
         '</div>' +
         '</div>' +
         '</div>';
}

// Function to create test data for demonstration
function createTestData() {
  // Create a mock workbook with sample data
  const mockWorkbook = {
    SheetNames: ['حاسوب_الثالث', 'كهرباء_الثاني'],
    Sheets: {
      'حاسوب_الثالث': createMockSheet('AC'),
      'كهرباء_الثاني': createMockSheet('EL')
    }
  };
  
  // Store in cache
  window.__workbookCache.set('mock_data', mockWorkbook);
  
  alert('✅ تم إنشاء بيانات تجريبية. الآن يمكنك الضغط على زر "طباعة نتائج جميع الأقسام" لاختبار الوظيفة.');
}

function createMockSheet(prefix) {
  // Create a mock sheet structure
  const sheet = {};
  
  // Add header rows
  sheet['A1'] = { v: 'رقم القيد' };
  sheet['B1'] = { v: 'اسم الطالب' };
  sheet['C1'] = { v: 'المواد' };
  sheet['D1'] = { v: 'الوحدات' };
  sheet['E1'] = { v: 'أعمال الفصل' };
  sheet['F1'] = { v: 'الامتحان النهائي' };
  sheet['G1'] = { v: 'المجموع' };
  sheet['H1'] = { v: 'التقدير' };
  sheet['I1'] = { v: 'المعدل الفصلي' };
  
  // Add student data rows
  for (let i = 0; i < 5; i++) {
    const row = i + 2;
    sheet[`A${row}`] = { v: prefix + '23' + (100 + i) };
    sheet[`B${row}`] = { v: `طالب ${i + 1}` };
    sheet[`C${row}`] = { v: 'مادة 1' };
    sheet[`D${row}`] = { v: 3 };
    sheet[`E${row}`] = { v: Math.floor(Math.random() * 40) + 10 };
    sheet[`F${row}`] = { v: Math.floor(Math.random() * 60) + 40 };
    sheet[`G${row}`] = { v: Math.floor(Math.random() * 50) + 50 };
    sheet[`H${row}`] = { v: 'جيد' };
    sheet[`I${row}`] = { v: (Math.random() * 30 + 70).toFixed(2) };
  }
  
  // Define sheet range
  sheet['!ref'] = 'A1:I' + (5 + 1);
  
  return sheet;
}

// Academic Year Management Functions

// Load saved academic year settings from localStorage
function loadAcademicYearSettings() {
  const savedSettings = localStorage.getItem('academicYearSettings');
  if (savedSettings) {
    try {
      const settings = JSON.parse(savedSettings);
      
      // Update form fields
      const academicYearInput = document.getElementById('academicYearInput');
      
      if (academicYearInput && settings.academicYearText) {
        academicYearInput.value = settings.academicYearText;
      }
      
      // Update display
      updateAcademicYearDisplay(settings.academicYearText);
    } catch (e) {
      console.error('Error loading academic year settings:', e);
    }
  }
}

// Save academic year settings to localStorage
function saveAcademicYearSettings() {
  const academicYearInput = document.getElementById('academicYearInput');
  const statusDiv = document.getElementById('academicYearStatus');
  
  if (!academicYearInput) return;
  
  const academicYearText = academicYearInput.value.trim();
  
  if (!academicYearText) {
    showAlert('يرجى إدخال نص السنة الدراسية', 'error');
    return;
  }
  
  const settings = {
    academicYearText: academicYearText,
    lastUpdated: new Date().toISOString()
  };
  
  try {
    localStorage.setItem('academicYearSettings', JSON.stringify(settings));
    updateAcademicYearDisplay(academicYearText);
    showAlert('✅ تم حفظ نص السنة الدراسية بنجاح', 'success');
    
    // Update any currently displayed results
    const resultDiv = document.getElementById('result');
    if (resultDiv && resultDiv.innerHTML) {
      updatePrintResultsAcademicYear(academicYearText);
    }
  } catch (e) {
    console.error('Error saving academic year settings:', e);
    showAlert('خطأ في حفظ الإعدادات. تأكد من مساحة التخزين الكافية.', 'error');
  }
}

// Update the academic year display in the header
function updateAcademicYearDisplay(text) {
  // Update the academic year input field
  const academicYearInput = document.getElementById('academicYearInput');
  if (academicYearInput) {
    academicYearInput.value = text;
  }
  
  // Also update in print results if they exist
  updatePrintResultsAcademicYear(text);
}

// Update academic year in printed results
function updatePrintResultsAcademicYear(text) {
  const resultDiv = document.getElementById('result');
  if (resultDiv && resultDiv.innerHTML) {
    // Update any academic year references in the results
    const academicYearElements = resultDiv.querySelectorAll('p');
    academicYearElements.forEach(element => {
      if (element.textContent.includes('الفصل الدراسي:')) {
        element.innerHTML = `<strong>الفصل الدراسي:</strong> ${text}`;
      }
    });
  }
  
  // Also update any existing department cover pages in batch print results
  const departmentCovers = document.querySelectorAll('.department-cover');
  departmentCovers.forEach(cover => {
    const paragraphs = cover.querySelectorAll('p');
    paragraphs.forEach(p => {
      if (p.textContent.includes('السنة الدراسية:')) {
        p.innerHTML = `السنة الدراسية: ${text}`;
      }
    });
  });
}

// Show alert/status message
function showAlert(message, type = 'info') {
  const statusDiv = document.getElementById('academicYearStatus');
  if (statusDiv) {
    statusDiv.textContent = message;
    statusDiv.style.display = 'block';
    
    // Set color based on type
    switch(type) {
      case 'success':
        statusDiv.style.color = '#28a745';
        break;
      case 'error':
        statusDiv.style.color = '#dc3545';
        break;
      case 'warning':
        statusDiv.style.color = '#ffc107';
        break;
      default:
        statusDiv.style.color = '#17a2b8';
    }
    
    // Auto-hide after 5 seconds
    setTimeout(() => {
      statusDiv.style.display = 'none';
    }, 5000);
  }
}



// Initialize academic year on page load
window.addEventListener('DOMContentLoaded', function() { 
  // Load settings if user is already logged in
  if (window.__currentUser && window.__currentUser.role == 'admin') {
    loadAcademicYearSettings();
  }
});



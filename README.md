# نتائج_الطلاب — تشغيل المشروع

هذا المستودع يحتوي على أدوات بسيطة للبحث في ملفات Excel (قوائم نتائج طلاب). الهدف: العثور على صف الطالب حسب رقم القيد واستخراج صف النتيجة.

محتويات مهمة:
- `find_student.py` — بحث نصي/مطبّع عن رقم القيد عبر ملفات `xls/`.
- `search_any.py` — بحث تجريدي عن سلاسل نصية داخل جميع الملفات في `xls/`.
- `dump_sheet.py` — أداة مساعدة لطباعة أول صفوف وعيّنة من القيم لملف محدد داخل `xls/`.
- `find_student.js` و`index.html` + `script.js` — نسخة Node وواجهة مستخدم تعتمد على SheetJS (client-side).
 - `search_accounting.py`, `dump_accounting.py` — أدوات مساعدة خاصة بورقة المحاسبة.

التغييرات التي أجريتها لتحسين قابلية التشغيل:
- سكربتات البايثون الآن تستخدم `pandas` لقراءة أوراق العمل، وتختار المحرك (`xlrd` أو `openpyxl`) اعتمادًا على الامتداد.
- أضفت `requirements.txt` يحتوي على `pandas`, `openpyxl`, و`xlrd==1.2.0` (محتاج لقراءة `.xls`).
- أبقيت سلوك البحث كما كان (كشف رؤوس عربية، تطبيع الأرقام، تقارير 1-based).

تشغيل محلي (بايثون):
1. انشئ بيئة افتراضية (مستحسن) ثم ثبت الاعتماديات:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1   # PowerShell
pip install -r requirements.txt
```

2. مثال استخدام:
```powershell
python find_student.py 232029
python search_any.py "اسم"
python dump_sheet.py
```

أمثلة جديدة:
```powershell
python find_student.py AC252002 --file "قسم المحاسبة25-26.xls" --format json
python find_student.py AC252002 --format csv --out result.csv
python search_accounting.py AC252002
```

ملاحظات حول Node وواجهة الويب:
- `find_student.js` يحتاج الحزمة `xlsx` (SheetJS). أضفت `package.json` صغير لسهولة التثبيت.
- واجهة الويب (`index.html` + `script.js`) تستخدم CDN لـ SheetJS؛ افتح الصفحة في المتصفح واستخدم زر اختيار الملف لتحميل ملف Excel محليًا. إذا كان المتصفح غير متصل بالإنترنت، ثبت نسخة محلية من SheetJS أو شغّل صفحة عبر خادم محلي.

تشغيل Node (اختياري):
```powershell
npm install
node find_student.js 232029
```

تشغيل واجهة الويب محليًا (موصى به لتجنب سياسات CORS عند تحميل ملفات عبر الشبكة):
```powershell
# من داخل المجلد المشروع
python -m http.server 8000
# ثم افتح http://localhost:8000/index.html في المتصفح
```

ماذا أفعل بعد ذلك؟
- إذا تفضّل، أستطيع استبدال استخدام `xlrd` بالكامل أو تحويل ملفات `.xls` إلى `.xlsx` تلقائيًا.
- إذا ظهر خطأ عند التشغيل، انسخ رسالة الخطأ هنا وسأحلّها فورًا.

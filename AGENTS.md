# Amazon Free Shipping to Israel Alert — Win App

> ⚠️ קובץ זה זהה ל-`CLAUDE.md` / `AGENTS.md` באותה תיקייה. כל עדכון חייב להיכתב **בשני הקבצים**.

# 🗄️ ארכיון — הפרויקט הופסק

**אפליקציית ה-Windows ירדה מהאוויר. השירות הוא web-only.**
הקוד החי נמצא בשני פרויקטים אחרים:

- **`Amazon Free Shipping to Israel Alert SaaS/`** — הבאקאנד והדשבורד (`app.amzfreeil.com`, Railway)
- **`amzfreeil-www/`** — האתר השיווקי והבלוג (`www.amzfreeil.com`, Vercel)

התיקייה הזו נשמרת להיסטוריה בלבד. **לא לפתח כאן, לא לבנות installer, לא לשחרר גרסה.**
כל בקשה שנוגעת ל"האפליקציה" צריכה להיות מנותבת ל-SaaS או ל-www.

---

## מה היה כאן (תיעוד היסטורי)

### שרשרת ההרצה במחשב המשתמש
```
AmazonIsraelFreeShipAlert.exe   ← launcher של PyInstaller (בלי Python מוטמע)
  └─ python.exe gui.py          ← Python 3.13 של המערכת, כ-subprocess
       ├─ checker.py            ← Playwright async, Chromium headed
       ├─ notifier.py           ← Gmail SMTP
       ├─ scheduler.py          ← APScheduler בתוך thread של gui.py
       ├─ config.py / config.json
       └─ state.py / state.json
```

ה-launcher הוא exe זעיר שמריץ את ה-Python האמיתי כ-subprocess, כדי שהרחבות ה-C (greenlet, PIL) ייטענו ע"י אותו מפרש שבו `pip` התקין אותן.

**ניקוי הסביבה ב-launcher:** ה-Setup הוא עצמו PyInstaller bundle שמגדיר `TCL_LIBRARY`/`TK_LIBRARY` לתיקיית `_MEI<n>` שלו. אם המשתנים דולפים ל-`python.exe gui.py`, Tkinter לא מוצא את `init.tcl` — ה-launcher מנקה אותם לפני ה-spawn.

**זיהוי משלוח חינם:** בלוק המשלוח מכיל את שלוש המחרוזות `"free delivery"` · `"to israel"` · `"eligible orders"`.

### גרסאות ו-build
`version.py` היה מקור האמת; `python build_installer.py` ייצר את
`AmazonIsraelFreeShipAlert_Setup_v{VERSION}.exe` (~17.5MB) עם ה-launcher מוטמע בפנים.
`*.exe` ו-`install.py` הם תוצרי בנייה ולא נדחפו ל-git.

### החלטות תכנון ששרדו
- `--only-binary :all:` ל-pip — greenlet חייב wheel מקומפל.
- `_find_python()` מחפש קודם ב-`~/AppData/Local/Programs/Python/Python3*/python.exe` ומדלג על stubs של `WindowsApps`; PATH רק כ-fallback.
- קוד יציאה 0 אינו שגיאה — הוא אומר "האפליקציה כבר רצה ב-tray".
- `PLAYWRIGHT_BROWSERS_PATH` נקבע במפורש ל-`%LOCALAPPDATA%\ms-playwright`.

## שפה מועדפת
עברית — כל התגובות והמסמכים בעברית.

## ⛔ אסור בלי אישור מפורש
- **אין לפתח, לבנות או לשחרר מכאן.** הפרויקט מוקפא.
- **אין להפנות לסקיל `sr-ptd-skill` ולא לנתיב `C:/projects/Skills/Dev_doc_for_skills`** — שניהם לא קיימים. הדרישה הזו הוסרה מהקובץ.
- אין להחזיר מסרים על "אפליקציית Windows" לשום נכס חי (`llms.txt`, `about.html`, `guide.html`) — זה מזין מנועי AI במידע שגוי.

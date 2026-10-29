# שלח ל-Planner – תוסף Outlook של Elcom IT

תוסף שהופך מייל למשימה ב-Planner, עם בחירת תוכנית, Bucket, שיוך ותאריך יעד.
עובד באאוטלוק החדש, בקלאסי, ב-Outlook on the web וב-Outlook ל-iOS/Android.

**מה התוסף עושה:**
- מציג את כל התוכניות בקבוצות M365 שאתה חבר בהן.
- **מועדפים** ⭐ ו**זכירת התוכנית וה-Bucket האחרונים**. ההגדרות מסונכרנות בין המחשב לטלפון.
- כותרת המשימה = נושא המייל (אפשר לערוך).
- התיאור כולל שולח, תאריך ותוכן המייל, וקישור למייל המקורי.
- שיוך לעצמך כברירת מחדל, או לכל חבר בקבוצה.

---

## שלב 1 – הגדרות
פתח את `config.json` ומלא שני ערכים:

```json
"clientId": "ה-Application (client) ID מ-Entra",
"githubUser": "שם המשתמש שלך ב-GitHub",
```

## שלב 2 – העלאה ל-GitHub
1. ב-GitHub צור repository חדש בשם **`mail-to-planner`**, מסוג **Public**.
   GitHub Pages חינמי רק ל-repo ציבורי. אין בקוד שום סוד, וה-Client ID אינו סוד.
2. העלה את כל הקבצים מהתיקייה (בלי `node_modules` ובלי `dist`):
   ```bash
   git init
   git add .
   git commit -m "Initial"
   git branch -M main
   git remote add origin https://github.com/elibz-elcom/mail-to-planner.git
   git push -u origin main
   ```
3. ב-repo, תחת **Settings → Pages → Build and deployment → Source**, בחר **GitHub Actions**.
4. בלשונית **Actions** תראה את ה-Deploy רץ. אחרי כדקה האתר זמין בכתובת:
   `https://elibz-elcom.github.io/mail-to-planner/`

## שלב 3 – הגדרת האפליקציה ב-Entra
בתוך האפליקציה ב-**App registrations**:

**Authentication → Add a platform → Single-page application.** הוסף שני Redirect URIs:
```
brk-multihub://elibz-elcom.github.io
https://elibz-elcom.github.io/mail-to-planner/redirect.html
```
- הראשון משמש להתחברות האוטומטית בתוך אאוטלוק (NAA).
- השני הוא גיבוי לגרסאות אאוטלוק ישנות.

ב-**API permissions → Add a permission → Microsoft Graph → Delegated**, הוסף:
| הרשאה | למה |
|---|---|
| `User.Read` | זיהוי המשתמש |
| `Tasks.ReadWrite` | יצירת משימות וקריאת תוכניות ו-Buckets |
| `GroupMember.Read.All` | רשימת הקבוצות והחברים לשיוך |
| `Mail.Read` | קישור למייל המקורי |

ולסיום לחץ **Grant admin consent for Elcom**.

## שלב 4 – התקנה
את קובץ ה-manifest מורידים מהאתר שנוצר, והוא כבר כולל את הכתובות הנכונות:
`https://elibz-elcom.github.io/mail-to-planner/manifest.xml`

**לבדיקה מהירה, רק אצלך:**
Outlook on the web → פתח מייל → **Apps** → **Add apps** → **My add-ins** → **Custom add-ins** → **Add from file** → בחר את `manifest.xml`.

**לכל הארגון, כולל טלפונים:**
M365 Admin Center → **Settings → Integrated apps → Upload custom apps** → **Office Add-in** → העלה את `manifest.xml` → בחר משתמשים.
ההפצה לתיבות לוקחת בדרך כלל כמה שעות, ולכל היותר כ-24 שעות.

## שימוש בטלפון
פתח מייל באפליקציית Outlook ← **`...`** ← **שלח ל-Planner** ← בחר תוכנית (או לחץ על מועדף) ← **צור משימה**.

## התקנה אצל לקוח
1. האדמין של הלקוח, או אתה עם הרשאות GDAP, פותח את הכתובת הזו ומאשר:
   ```
   https://login.microsoftonline.com/<tenant-של-הלקוח>/adminconsent?client_id=<clientId>
   ```
2. מעלים את אותו `manifest.xml` ב-Integrated Apps של הלקוח.

אין צורך באירוח נפרד ולא ברישום נפרד.

## Publisher Verification
App registrations → האפליקציה → **Branding & properties**:
1. Publisher domain = `elcom.co.il`.
2. **Add Partner One ID to verify publisher** → `3873011` → **Verify and save**.

## עדכונים
- **שינוי בקוד:** `git push` מספיק. אאוטלוק טוען את הגרסה החדשה אוטומטית, בלי התקנה מחדש.
- **שינוי ב-manifest** (שם, אייקון, הרשאות): העלה את `version` ב-`config.json` (למשל `1.0.1.0`), בנה, והעלה את ה-manifest מחדש ב-Integrated Apps.

## בנייה מקומית (לא חובה)
```bash
npm install
npm run build
```
ה-build יוצר את `dist/` ואת `manifest.xml` עם הכתובות הנכונות, ומדפיס את ה-Redirect URIs שצריך להגדיר ב-Entra.

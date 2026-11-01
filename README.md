# 🍍 PineApple Attendance Bot (ប្រព័ន្ធគ្រប់គ្រងវត្តមានសិស្សតាម Telegram)

ប្រព័ន្ធកត់ត្រាវត្តមានសិស្ស និងទាញយករបាយការណ៍ Excel ស្វ័យប្រវត្តិតាមកាលវិភាគសិក្សា សម្រាប់លោកគ្រូ-អ្នកគ្រូ និងប្រធានថ្នាក់ នៃ **វិទ្យាល័យ នេត យ៉ង់ - ថ្នាក់ទី 10D**។

---

## 🌟 លក្ខណៈពិសេសចម្បងៗ (Key Features)

1. **⚡ កត់ត្រាវត្តមានរហ័ស & ងាយស្រួល:**
   - ប្រធានថ្នាក់វាយតែអ្នក **អវត្តមាន (`A`)**, **ច្បាប់ (`P`)**, ឬ **យឺត (`L`)** ប៉ុណ្ណោះ។
   - សិស្សដទៃទៀតទាំងអស់នឹងទទួលបាន **វត្តមាន (`✓`) ដោយស្វ័យប្រវត្តិ**។
   - មានប៊ូតុងចុចជ្រើសរើសថ្ងៃ $\rightarrow$ ម៉ោង $\rightarrow$ មុខវិជ្ជា (`/record`)។

2. **📊 របាយការណ៍ Excel ស្វ័យប្រវត្តិ (Excel Reports):**
   - **របាយការណ៍ពេញមួយថ្ងៃ (Full Day):** បង្ហាញគ្រប់ម៉ោងសិក្សាទាំងអស់ក្នុងថ្ងៃនោះ និងសម្គាល់ម៉ោងដែលប្រធានថ្នាក់ភ្លេចកត់ត្រា (`⚠️ ភ្លេចកត់`)។
   - **របាយការណ៍ប្រចាំខែ និងចន្លោះថ្ងៃ (Monthly & Date Range):**
     - សន្លឹកទី១៖ សង្ខេបប្រចាំខែ ចំណាត់ថ្នាក់ និងការវាយតម្លៃវិន័យសិស្ស។
     - សន្លឹកទី២៖ តារាងវត្តមានតាមថ្ងៃពេញមួយខែ (Daily Attendance Grid)។
     - សន្លឹកទី៣៖ ស្ថិតិវត្តមានតាមមុខវិជ្ជានីមួយៗ។
   - **ប្តូរឈ្មោះឯកសារដោយសេរី:** អាចកំណត់ឈ្មោះឯកសារ និងចំណងជើងតាមចិត្ត (`/range 15-10-2026 to 14-11-2026 ឈ្មោះឯកសារ`)។

3. **🔒 សុវត្ថិភាពទិន្នន័យ (Teacher Only Privacy):**
   - មានតែលោកគ្រូទេដែលអាចទាញយករបាយការណ៍ និងគ្រប់គ្រងសិស្ស/មុខវិជ្ជាបាន។
   - នៅពេលហៅរបាយការណ៍ក្នុងគ្រុប Bot នឹងផ្ញើឯកសារ Excel ទៅកាន់ **ឆាតផ្ទាល់ខ្លួន (Private Chat)** របស់លោកគ្រូដោយសម្ងាត់។

4. **☁️ Google Sheet Database + Vercel 24/7:**
   - ប្រើ **Google Sheet** ក្នុង Google Drive ជា Database រក្សាទុកទិន្នន័យអចិន្ត្រៃយ៍។
   - ដំណើរការលើ **Vercel** ដោយឥតគិតថ្លៃ ២៤ម៉ោង/៧ថ្ងៃ។

---

## 🚀 របៀបតម្លើង និង Host លើ Vercel + Google Sheet

### ជំហានទី ១៖ បង្កើត Google Sheet Database លើ Google Drive
1. ចូលទៅកាន់ [Google Drive](https://drive.google.com)
2. បង្កើត Folder មួយឈ្មោះថា **`PineApple_Attendance`**
3. ចូលក្នុង Folder នោះ រួចបង្កើត Google Sheet មួយឈ្មោះថា **`PineApple_Attendance_DB`**
4. នៅលើ menu ខាងលើ ចុច៖ **Extensions (ផ្នែកបន្ថែម) -> Apps Script**
5. លុបកូដចាស់ៗចោល រួច Copy កូដទាំងអស់ពីឯកសារ [`google_apps_script.js`](file:///d:/PineApple_Attendance/google_apps_script.js) យកទៅ Paste ចូល
6. ចុច **Save (រូបថាស)**
7. ចុចប៊ូតុង **Deploy (ពណ៌ខៀវខាងលើស្តាំ)** -> ជ្រើសរើស **New deployment**
   - ចុចរូបកង់ធ្មេញ (Select type) -> ជ្រើសរើស **Web app**
   - Description: `PineApple Attendance API`
   - Execute as: **Me**
   - Who has access: **Anyone**
8. ចុច **Deploy** -> ចុច **Authorize access** (ជ្រើសរើស Google Account របស់លោកគ្រូ)
9. Copy យក **Web app URL** (ឧ. `https://script.google.com/macros/s/AKfycb.../exec`)

---

### ជំហានទី ២៖ Push Code ទៅកាន់ GitHub
1. ចូលទៅកាន់ [GitHub Create New Repo](https://github.com/new)
2. ដាក់ឈ្មោះ Repository: **`PineApple_Attendance_Bot`** (ជ្រើសរើស Public ឬ Private)
3. ចុច **Create repository**
4. នៅលើកុំព្យូទ័រលោកគ្រូ បើក Terminal ក្នុង folder នេះរួចវាយ៖
   ```bash
   git add .
   git commit -m "Initial commit for PineApple Attendance Bot"
   git remote add origin https://github.com/<YOUR_GITHUB_USERNAME>/PineApple_Attendance_Bot.git
   git push -u origin main
   ```

---

### ជំហានទី ៣៖ Deploy ទៅកាន់ Vercel
1. ចូលទៅកាន់ [Vercel Dashboard](https://vercel.com)
2. ចុច **Add New...** -> **Project** -> ជ្រើសរើស Repository **`PineApple_Attendance_Bot`**
3. ត្រង់កន្លែង **Environment Variables** បន្ថែមអថេរខាងក្រោម៖
   - `BOT_TOKEN` = `<BOT_TOKEN_ដែលបានមកពី_BotFather>`
   - `TEACHER_NAME` = `សក់ វាសនា`
   - `CLASS_NAME` = `វិទ្យាល័យ នេត យ៉ង់ - ថ្នាក់ទី 10D`
   - `TIMEZONE` = `Asia/Phnom_Penh`
   - `GOOGLE_SHEET_API_URL` = `<Web_app_URL_ដែលបានមកពីជំហានទី១>`
4. ចុច **Deploy**!
5. នៅពេល Deploy ចប់ Vercel នឹងផ្តល់ Domain មួយ (ឧ. `https://pineapple-attendance-bot.vercel.app`)

---

### ជំហានទី ៤៖ ភ្ជាប់ Telegram Webhook ទៅ Vercel
បើក Terminal លើកុំព្យូទ័រលោកគ្រូរួចវាយ៖
```bash
node scripts/setWebhook.js https://<YOUR_VERCEL_DOMAIN>.vercel.app
```
*(ឬបើក Browser វាយ: `https://api.telegram.org/bot<TOKEN>/setWebhook?url=https://<YOUR_VERCEL_DOMAIN>.vercel.app/api/webhook`)*

🎉 **រួចរាល់! Bot នឹងដំណើរការ ២៤ម៉ោង/៧ថ្ងៃ ដោយស្វ័យប្រវត្តិ!**

# Gemini TTS Pro

تطبيق سطح مكتب (PyQt6) لتحويل النصوص العربية إلى كلام باستخدام Google Gemini TTS API.

## المتطلبات
- Python 3.10+
- (اختياري) ffmpeg مثبت على النظام لدمج الصوت بصيغة MP3

## التثبيت محليًا
```bash
pip install -r requirements.txt
python Gemini_TTS_Pro_v3_8_13.py
```

## البناء كتطبيق one-dir (مجلد وليس ملف واحد)
الطريقة السريعة — شغّل ملف البناء الجاهز:
- على ويندوز: `build.bat` (دبل كليك عليه أو `build.bat` من الـ cmd)
- على ماك/لينكس: `./build.sh`

أو يدويًا:
```bash
pip install pyinstaller
pyinstaller gemini_tts_pro.spec
```
الناتج هيبقى في `dist/GeminiTTSPro/` — مجلد فيه الملف التنفيذي (`GeminiTTSPro.exe` على ويندوز) وكل الملفات المصاحبة.
دي طريقة `--onedir` (عكس `--onefile`)، بتشتغل أسرع في الفتح وأسهل في التحديث لأن كل الملفات ظاهرة في المجلد.

> ملاحظة: نفّذ الأمر على نفس نظام التشغيل اللي هتوزّع عليه التطبيق (ابني على ويندوز عشان تطلع .exe لويندوز).

## رفع المشروع على GitHub
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<username>/<repo-name>.git
git push -u origin main
```

الملفات دي (`requirements.txt`, `.gitignore`, `gemini_tts_pro.spec`) لازم تترفع مع الكود عشان أي حد يقدر يبني التطبيق من جديد.

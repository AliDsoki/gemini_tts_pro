@echo off
chcp 65001 >nul
echo ============================================
echo   بناء Gemini TTS Pro (one-dir)
echo ============================================

echo [1/3] تثبيت المتطلبات...
pip install -r requirements.txt pyinstaller
if errorlevel 1 (
    echo فشل تثبيت المتطلبات.
    pause
    exit /b 1
)

echo [2/3] تنظيف مجلدات البناء القديمة...
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul

echo [3/3] بناء التطبيق...
pyinstaller gemini_tts_pro.spec
if errorlevel 1 (
    echo فشلت عملية البناء.
    pause
    exit /b 1
)

echo ============================================
echo   تم البناء بنجاح! الملفات في dist\GeminiTTSPro
echo ============================================
pause

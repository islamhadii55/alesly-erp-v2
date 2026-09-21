# تطبيق سطح المكتب — الأصلي لقطع الغيار

تطبيق سطح مكتب يشغّل سيرفر Flask محلياً على الجهاز (SQLite)، ثم يفتحه في نافذة تطبيق مستقلة.
يعمل بدون إنترنت بالكامل، ويزامن البيانات تلقائياً عند توفر الاتصال.

## التشغيل السريع
```bash
python3 desktop/alasly_desktop.py
```
- إذا كانت مكتبة `pywebview` مثبتة: تُفتح نافذة تطبيق مستقلة.
- إذا لم تكن مثبتة: يُفتح النظام في المتصفح الافتراضي مع إبقاء الطرفية مفتوحة.

## الاعتماديات
```bash
pip install --break-system-packages flask pywebview
```

## بناء ملف تنفيذي واحد (Windows / Linux / macOS)
```bash
pip install --break-system-packages pyinstaller
cd desktop
pyinstaller alasly.spec
```
الملف الناتج: `desktop/dist/AlaslyERP` (أو `AlaslyERP.exe` على ويندوز).

## المزامنة
- المزامنة المحلية: كل التعديلات تُخزَّن في `original_auto.db` بجانب التطبيق.
- عند عودة الإنترنت: يرسل العميل الطابور المؤجل إلى `/api/sync/push` تلقائياً كل 20 ثانية.
- لإدارة المزامنة: صفحة "المزامنة بدون إنترنت" داخل النظام.

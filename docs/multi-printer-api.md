# API الطابعات المتعددة

تتم تهيئة جداول SQLite تلقائيًا عند تشغيل `app.py` عبر `printer_support.py`.
جميع المسارات تتطلب جلسة دخول، ومسارات الإدارة تتطلب دور `مدير`.

## الملفات

- `printer_support.py`: المخطط، الترحيل التلقائي، ومسارات API.
- `app.py`: تسجيل الوحدة وتهيئتها عند بدء التطبيق.

## أنواع الطابعات

- `thermal`: إيصالات، التنسيق المعتاد `esc_pos`.
- `barcode`: ملصقات باركود، التنسيقات `zpl` أو `tspl`.
- `invoice`: فواتير وتقارير، التنسيق المعتاد `pdf`.

## المسارات

| الطريقة | المسار | الصلاحية | الوصف |
|---|---|---|---|
| `GET` | `/api/printer-profiles` | دخول | عرض ملفات الطابعات |
| `POST` | `/api/printer-profiles` | مدير | إنشاء ملف طابعة |
| `GET` | `/api/printers` | دخول | عرض الطابعات |
| `POST` | `/api/printers` | مدير | إنشاء طابعة |
| `GET/PATCH/DELETE` | `/api/printers/:id` | دخول/مدير | عرض أو تعديل أو تعطيل طابعة |
| `POST` | `/api/printers/:id/default` | مدير | تعيين الافتراضية |
| `POST` | `/api/printer-assignments` | مدير | ربط نوع مهمة بطابعة |
| `GET` | `/api/printer-assignments` | دخول | عرض الربط |
| `POST` | `/api/print-jobs` | دخول | إضافة مهمة إلى الطابور |
| `GET` | `/api/print-jobs/:id` | دخول | متابعة مهمة |
| `POST` | `/api/print-jobs/:id/cancel` | دخول | إلغاء مهمة queued |
| `POST` | `/api/printers/:id/health` | دخول | تسجيل فحص اتصال |

## مثال إنشاء مهمة

```json
{
  "job_type": "invoice",
  "payload_format": "pdf",
  "payload": {"invoice_id": 123},
  "copies": 1,
  "branch_id": 1,
  "workstation_id": "POS-01"
}
```

يمكن إرسال `Idempotency-Key` في رأس الطلب لمنع إنشاء المهمة نفسها مرتين.
يتم اختيار الطابعة الصريحة أولًا، ثم الأكثر تخصصًا حسب محطة العمل والفرع، ثم الافتراضية.

> فحص الاتصال الفعلي عبر TCP أو USB أو Windows Spooler يحتاج Adapter خاص ببيئة التشغيل. المسار الحالي يسجل عدم توفر الـ Adapter بدل الإدعاء بنجاح اتصال غير منفذ.

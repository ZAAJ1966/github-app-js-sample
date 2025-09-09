# نظام إدارة فواتير عوائد المناولة والتخزين

هذا المستودع يحوي نسخة **Offline Desktop** مبنية على Excel VBA لإدارة فواتير عوائد المناولة والتخزين.

## محتويات المشروع

```
InvoiceManagement/
│   README_AR.md
│
├── إدارة_الفواتير.xlsm   ← سيتم توليده من خلال استيراد الوحدات والنماذج
├── Documents/             ← يخزن المستندات المرفقة آليًا
├── PDFReports/            ← يحفظ تقارير الفواتير بصيغة PDF
├── import_data/           ← ملفات Excel/CSV جاهزة للاستيراد
└── src/
    ├── modules/           ← ملفات .bas (كود VBA مقسم لوحدات)
    └── forms/             ← ملفات .frm و .frx لنماذج UserForm
```

## المتطلبات
1. نظام تشغيل Windows مع Microsoft Office (Excel) إصدار 2016 أو أحدث.
2. تمكين الماكرو في Excel.
3. صلاحية إنشاء ملفات في نفس مجلد المشروع (لإنشاء المجلدات الفرعية وحفظ الوثائق).

## خطوات الإعداد السريع
1. افتح ملف **إدارة_الفواتير.xlsm** (غير موجود بعد – ستنشئه من خلال استيراد الكود).
2. فعّل الماكرو عند الطلب.
3. سيظهر نموذج تسجيل الدخول تلقائيًا.
4. بيانات دخول افتراضية:
   - المستخدم: `admin`، كلمة المرور: `1234` (دور: مدير).

## استيراد الكود إلى مصنف Excel جديد

> ملاحظة: يمكنك أيضاً استخدام أدوات الذكاء الاصطناعي لتوليد مصنف مكتمل تلقائياً وفق البرومبت المرفق، ولكن في حال الرغبة بالاستيراد يدويًا:

1. أنشئ مصنف Excel جديد واحفظه باسم `إدارة_الفواتير.xlsm` داخل هذا المجلد.
2. من **VBA Editor** (Alt+F11):
   * اضغط يمين على المشروع → `Import File…` → استورد كل الملفات داخل `src/modules`.
   * اضغط يمين على قسم النماذج → `Import File…` → استورد كل النماذج داخل `src/forms`.
3. افتح `ThisWorkbook` وأضف الكود التالي لتشغيل نموذج الدخول:
   ```vba
   Private Sub Workbook_Open()
       UserForm_Login.Show
   End Sub
   ```
4. أضف أوراق العمل التالية في المصنف (إن لم تكن موجودة):
   - `Users`
   - `Invoices`
   - `Payments`
   - `Attachments`
   - `InvoiceReport`
5. انسخ رؤوس الأعمدة من ملفات CSV داخل `import_data` لكل ورقة.
6. احفظ وأعد فتح المصنف للتجربة.

## بنية الجداول

### ورقة Users
| Username | Password | Role |
|----------|----------|------|
| admin | 1234 | Manager |
| clearance1 | abc1 | ClearanceOfficer |
| audit1 | auditpass | Auditor |
| accountant1 | passacc | Accountant |

### ورقة Invoices (العناوين الإنجليزية لتسهيل الكود)
| SerialNo | ClaimNo | ClaimDate | InvoiceNo | InvoiceDate | AmountWithoutStamp | StampDuty | TotalAmount | PolicyNo | LCNo | ContractNo | CustomsDecNo | PortOfEntry |

### ورقة Payments
| ClaimNo | ClaimDate | InvoiceNo | InvoiceDate | PaidStatus | PaymentDate |

### ورقة Attachments
| InvoiceNo | DocType | FileName | FullPath |

## الدعم
لأي استفسارات راسل مسؤول النظام أو راجع الكود داخل الوحدات لفهم سير العمل.
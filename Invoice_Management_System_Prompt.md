# نظام إدارة فواتير عوائد المناولة والتخزين - Excel VBA
## Prompt احترافي لتطوير نظام متكامل

---

## 🎯 **وصف المشروع**

أنشئ نظاماً متكاملاً لإدارة فواتير عوائد المناولة والتخزين باستخدام Excel VBA، يعمل كتطبيق سطح مكتب مستقل (Offline) على أنظمة Windows. النظام يجب أن يكون احترافياً وسهل الاستخدام مع واجهات مستخدم جذابة ووظائف شاملة.

---

## 🏗️ **المتطلبات الأساسية**

### 1. **البيئة التقنية**
- **اللغة**: Excel VBA (Visual Basic for Applications)
- **النسخة**: متوافق مع Excel 2016 أو أحدث
- **نظام التشغيل**: Windows 10/11
- **نوع التطبيق**: Offline Desktop Application
- **تنسيق الملف**: .xlsm (Excel Macro-Enabled Workbook)

### 2. **الخصائص العامة**
- واجهة مستخدم باللغة العربية
- تصميم احترافي وسهل الاستخدام
- نظام صلاحيات متعدد المستويات
- إمكانية العمل بدون اتصال بالإنترنت
- نظام أمان وحماية للبيانات

---

## 📊 **هيكل قاعدة البيانات**

### **Sheet 1: Users** (جدول المستخدمين)
```
| العمود | الحقل | نوع البيانات | الوصف |
|--------|-------|-------------|--------|
| A | UserID | Text | معرف المستخدم الفريد |
| B | Username | Text | اسم المستخدم |
| C | Password | Text | كلمة المرور (مشفرة) |
| D | UserType | Text | نوع المستخدم (مدير/موظف تخليص/مدقق/محاسب) |
| E | FullName | Text | الاسم الكامل |
| F | IsActive | Boolean | حالة النشاط |
| G | LastLogin | Date | آخر تسجيل دخول |
```

### **Sheet 2: Invoices** (جدول الفواتير الرئيسي)
```
| العمود | الحقل | نوع البيانات | الوصف |
|--------|-------|-------------|--------|
| A | SerialNumber | Number | الرقم التسلسلي (Auto-increment) |
| B | ClaimNumber | Text | رقم المطالبة |
| C | ClaimDate | Date | تاريخ المطالبة |
| D | InvoiceNumber | Text | رقم الفاتورة |
| E | InvoiceDate | Date | تاريخ الفاتورة |
| F | AmountWithoutStamp | Currency | قيمة الفاتورة بدون رسوم الدمغة |
| G | StampDuty | Currency | رسوم الدمغة |
| H | TotalAmount | Currency | قيمة الفاتورة + رسوم الدمغة |
| I | PolicyNumber | Text | رقم البوليصة |
| J | LCNumber | Text | رقم الاعتماد |
| K | ContractNumber | Text | رقم العقد |
| L | CustomsDeclaration | Text | رقم الإقرار الجمركي |
| M | EntryPort | Text | منفذ الدخول (ميناء/مطار/بر) |
| N | PaymentStatus | Text | حالة السداد (مسدد/غير مسدد) |
| O | PaymentDate | Date | تاريخ السداد |
| P | CreatedBy | Text | أنشئ بواسطة |
| Q | CreatedDate | Date | تاريخ الإنشاء |
| R | ModifiedBy | Text | عُدل بواسطة |
| S | ModifiedDate | Date | تاريخ التعديل |
```

### **Sheet 3: PaymentTracking** (متابعة السداد)
```
| العمود | الحقل | نوع البيانات | الوصف |
|--------|-------|-------------|--------|
| A | TrackingID | Number | معرف المتابعة |
| B | ClaimNumber | Text | رقم المطالبة |
| C | InvoiceNumber | Text | رقم الفاتورة |
| D | HandlingStorageFees | Currency | رسوم عوائد المناولة والتخزين |
| E | StampDuty | Currency | رسوم الدمغة |
| F | TotalFees | Currency | إجمالي الرسوم |
| G | PaymentStatus | Text | حالة السداد |
| H | TransferredTo | Text | محول إلى (محاسبة المشروعات/الحسابات التشغيلية) |
| I | TransferDate | Date | تاريخ التحويل |
```

### **Sheet 4: Attachments** (المستندات المرفقة)
```
| العمود | الحقل | نوع البيانات | الوصف |
|--------|-------|-------------|--------|
| A | AttachmentID | Number | معرف المرفق |
| B | InvoiceNumber | Text | رقم الفاتورة |
| C | DocumentType | Text | نوع المستند |
| D | FileName | Text | اسم الملف |
| E | FilePath | Text | مسار الملف الكامل |
| F | FileSize | Number | حجم الملف |
| G | UploadDate | Date | تاريخ الرفع |
| H | UploadedBy | Text | رفع بواسطة |
```

### **Sheet 5: InvoiceReport** (قالب التقارير)
قالب لتوليد تقارير PDF للفواتير

---

## 🎨 **واجهات المستخدم (UserForms)**

### **1. UserForm_Login** (نموذج تسجيل الدخول)
```vba
' عناصر النموذج:
- Label: lblTitle = "نظام إدارة فواتير عوائد المناولة والتخزين"
- Label: lblUsername = "اسم المستخدم:"
- TextBox: txtUsername
- Label: lblPassword = "كلمة المرور:"
- TextBox: txtPassword (PasswordChar = "*")
- CommandButton: cmdLogin = "دخول"
- CommandButton: cmdExit = "خروج"
- Image: imgLogo (شعار الشركة - اختياري)

' خصائص النموذج:
- StartUpPosition = 2 (CenterScreen)
- BorderStyle = 1 (Fixed Single)
- ControlBox = False
- MinButton = False
- MaxButton = False
```

### **2. UserForm_MainMenu** (القائمة الرئيسية)
```vba
' عناصر النموذج:
- Label: lblWelcome = "مرحباً، [اسم المستخدم] - [نوع المستخدم]"
- CommandButton: btnInvoiceEntry = "🧾 إدارة الفواتير"
- CommandButton: btnPaymentTracking = "💰 متابعة السداد"
- CommandButton: btnSearchFilter = "🔍 البحث والتصفية"
- CommandButton: btnAttachments = "📎 إدارة المستندات"
- CommandButton: btnReports = "📊 التقارير"
- CommandButton: btnImportData = "📥 استيراد البيانات"
- CommandButton: btnUserManagement = "👥 إدارة المستخدمين" (للمدير فقط)
- CommandButton: btnLogout = "🚪 تسجيل الخروج"
- CommandButton: btnExit = "❌ إغلاق النظام"

' إحصائيات سريعة:
- Label: lblTotalInvoices = "إجمالي الفواتير: [عدد]"
- Label: lblPaidInvoices = "الفواتير المسددة: [عدد]"
- Label: lblUnpaidInvoices = "الفواتير غير المسددة: [عدد]"
- Label: lblTotalAmount = "إجمالي المبالغ: [مبلغ]"
```

### **3. UserForm_InvoiceEntry** (إدخال وتعديل الفواتير)
```vba
' عناصر النموذج:
- ComboBox: cmbWorksheet = "اختيار ورقة العمل"
- ListBox: lstInvoices = "قائمة الفواتير" (لعرض البيانات الموجودة)

' حقول إدخال البيانات:
- TextBox: txtClaimNumber = "رقم المطالبة"
- TextBox: txtClaimDate = "تاريخ المطالبة" (DatePicker)
- ComboBox: cmbInvoiceNumber = "رقم الفاتورة"
- TextBox: txtInvoiceDate = "تاريخ الفاتورة" (DatePicker)
- TextBox: txtAmountWithoutStamp = "القيمة بدون رسوم الدمغة"
- TextBox: txtStampDuty = "رسوم الدمغة"
- TextBox: txtTotalAmount = "إجمالي القيمة" (محسوب تلقائياً)
- ComboBox: cmbPolicyNumber = "رقم البوليصة"
- ComboBox: cmbLCNumber = "رقم الاعتماد"
- ComboBox: cmbContractNumber = "رقم العقد"
- ComboBox: cmbCustomsDeclaration = "رقم الإقرار الجمركي"
- ComboBox: cmbEntryPort = "منفذ الدخول" (ميناء/مطار/بر)

' أزرار التحكم:
- CommandButton: cmdAdd = "➕ إضافة جديد"
- CommandButton: cmdUpdate = "✏️ تعديل"
- CommandButton: cmdDelete = "🗑️ حذف"
- CommandButton: cmdClear = "🔄 مسح الحقول"
- CommandButton: cmdSave = "💾 حفظ"
- CommandButton: cmdCancel = "❌ إلغاء"
- CommandButton: cmdClose = "🚪 إغلاق"

' حقل البحث السريع:
- TextBox: txtQuickSearch = "البحث السريع..."
- CommandButton: cmdQuickSearch = "🔍"
```

### **4. UserForm_PaymentTracking** (متابعة السداد)
```vba
' عناصر النموذج:
- ListBox: lstPaymentTracking = "قائمة متابعة السداد"
- ComboBox: cmbClaimNumber = "رقم المطالبة"
- ComboBox: cmbInvoiceNumber = "رقم الفاتورة"
- TextBox: txtHandlingFees = "رسوم عوائد المناولة والتخزين"
- TextBox: txtStampDuty = "رسوم الدمغة"
- TextBox: txtTotalFees = "إجمالي الرسوم" (محسوب تلقائياً)
- ComboBox: cmbPaymentStatus = "حالة السداد" (مسدد/غير مسدد)
- ComboBox: cmbTransferredTo = "محول إلى" (محاسبة المشروعات/الحسابات التشغيلية)
- TextBox: txtTransferDate = "تاريخ التحويل" (DatePicker)

' أزرار التحكم:
- CommandButton: cmdUpdateStatus = "تحديث الحالة"
- CommandButton: cmdSendNotification = "إرسال إشعار"
- CommandButton: cmdGenerateReport = "تقرير السداد"
```

### **5. UserForm_SearchFilter** (البحث والتصفية المتقدمة)
```vba
' معايير البحث:
- TextBox: txtSearchInvoice = "رقم الفاتورة"
- TextBox: txtSearchClaim = "رقم المطالبة"
- TextBox: txtSearchPolicy = "رقم البوليصة"
- TextBox: txtSearchContract = "رقم العقد"
- TextBox: txtSearchLC = "رقم الاعتماد"
- TextBox: txtSearchCustoms = "رقم الإقرار الجمركي"
- ComboBox: cmbSearchPort = "منفذ الدخول"
- ComboBox: cmbPaymentStatus = "حالة السداد"

' تصفية التواريخ:
- TextBox: txtDateFrom = "من تاريخ" (DatePicker)
- TextBox: txtDateTo = "إلى تاريخ" (DatePicker)
- ComboBox: cmbDateType = "نوع التاريخ" (تاريخ الفاتورة/تاريخ المطالبة)

' نتائج البحث:
- ListBox: lstSearchResults = "نتائج البحث"
- Label: lblResultCount = "عدد النتائج: [عدد]"

' أزرار التحكم:
- CommandButton: cmdSearch = "🔍 بحث"
- CommandButton: cmdClearSearch = "مسح البحث"
- CommandButton: cmdOpenSelected = "فتح المحدد"
- CommandButton: cmdPrintSelected = "طباعة المحدد"
- CommandButton: cmdExportResults = "تصدير النتائج"
```

### **6. UserForm_Attachments** (إدارة المستندات)
```vba
' عناصر النموذج:
- ComboBox: cmbInvoiceNumber = "رقم الفاتورة"
- ComboBox: cmbDocumentType = "نوع المستند"
  (بوليصة الشحن/أمر التسليم/الفاتورة/قائمة التعبئة/شهادة المنشأ/أخرى)
- Label: lblSelectedFile = "الملف المحدد: [لا يوجد]"
- ListBox: lstAttachments = "المستندات المرفقة"

' أزرار التحكم:
- CommandButton: cmdSelectFile = "📁 اختيار ملف"
- CommandButton: cmdUpload = "⬆️ رفع المستند"
- CommandButton: cmdViewDocument = "👁️ عرض المستند"
- CommandButton: cmdDeleteDocument = "🗑️ حذف المستند"
- CommandButton: cmdOrganizeFiles = "📂 تنظيم الملفات"
```

### **7. UserForm_Reports** (التقارير)
```vba
' أنواع التقارير:
- OptionButton: optSingleInvoice = "تقرير فاتورة واحدة"
- OptionButton: optMultipleInvoices = "تقرير متعدد الفواتير"
- OptionButton: optPaymentStatus = "تقرير حالة السداد"
- OptionButton: optFinancialSummary = "التقرير المالي الشامل"

' معايير التقرير:
- ComboBox: cmbInvoiceNumber = "رقم الفاتورة" (للتقرير المفرد)
- TextBox: txtDateFrom = "من تاريخ" (DatePicker)
- TextBox: txtDateTo = "إلى تاريخ" (DatePicker)
- ComboBox: cmbPort = "منفذ الدخول"
- ComboBox: cmbPaymentStatus = "حالة السداد"

' خيارات التصدير:
- CheckBox: chkIncludeAttachments = "تضمين المستندات المرفقة"
- CheckBox: chkIncludeSignature = "تضمين التوقيع"
- ComboBox: cmbReportFormat = "تنسيق التقرير" (PDF/Excel)

' أزرار التحكم:
- CommandButton: cmdGenerateReport = "📊 إنشاء التقرير"
- CommandButton: cmdPreview = "👁️ معاينة"
- CommandButton: cmdPrint = "🖨️ طباعة"
- CommandButton: cmdSave = "💾 حفظ"
```

### **8. UserForm_ImportData** (استيراد البيانات)
```vba
' مصدر البيانات:
- OptionButton: optFromExcel = "من ملف Excel"
- OptionButton: optFromPDF = "من ملف PDF"
- OptionButton: optFromWord = "من ملف Word"
- OptionButton: optFromCSV = "من ملف CSV"

' إعدادات الاستيراد:
- Label: lblSelectedFile = "الملف المحدد: [لا يوجد]"
- CheckBox: chkValidateData = "التحقق من صحة البيانات"
- CheckBox: chkSkipDuplicates = "تخطي المكرر"
- ComboBox: cmbTargetSheet = "ورقة الهدف"

' معاينة البيانات:
- ListBox: lstPreviewData = "معاينة البيانات المستوردة"
- Label: lblRecordCount = "عدد السجلات: [عدد]"

' أزرار التحكم:
- CommandButton: cmdSelectFile = "📁 اختيار ملف"
- CommandButton: cmdPreviewData = "👁️ معاينة البيانات"
- CommandButton: cmdImport = "📥 استيراد"
- CommandButton: cmdCancel = "❌ إلغاء"
```

### **9. UserForm_UserManagement** (إدارة المستخدمين - للمدير فقط)
```vba
' قائمة المستخدمين:
- ListBox: lstUsers = "قائمة المستخدمين"

' بيانات المستخدم:
- TextBox: txtUsername = "اسم المستخدم"
- TextBox: txtPassword = "كلمة المرور"
- TextBox: txtFullName = "الاسم الكامل"
- ComboBox: cmbUserType = "نوع المستخدم" (مدير/موظف تخليص/مدقق/محاسب)
- CheckBox: chkIsActive = "نشط"

' أزرار التحكم:
- CommandButton: cmdAddUser = "➕ إضافة مستخدم"
- CommandButton: cmdUpdateUser = "✏️ تعديل مستخدم"
- CommandButton: cmdDeleteUser = "🗑️ حذف مستخدم"
- CommandButton: cmdResetPassword = "🔄 إعادة تعيين كلمة المرور"
```

---

## 🔧 **وحدات VBA (Modules)**

### **Module 1: GlobalVariables**
```vba
' متغيرات عامة للنظام
Public Const APP_NAME As String = "نظام إدارة فواتير عوائد المناولة والتخزين"
Public Const APP_VERSION As String = "1.0"

Public CurrentUser As String
Public CurrentUserType As String
Public CurrentUserFullName As String
Public IsLoggedIn As Boolean

' مسارات المجلدات
Public Const DOCUMENTS_FOLDER As String = "Documents"
Public Const REPORTS_FOLDER As String = "PDFReports"
Public Const IMPORTS_FOLDER As String = "ImportedData"

' أنواع المستخدمين
Public Enum UserTypes
    Manager = 1
    ClearanceOfficer = 2
    Auditor = 3
    Accountant = 4
End Enum

' حالات السداد
Public Enum PaymentStatus
    Paid = 1
    Unpaid = 2
    Partial = 3
End Enum
```

### **Module 2: AuthenticationModule**
```vba
' وظائف تسجيل الدخول والأمان
Public Function ValidateLogin(username As String, password As String) As Boolean
Public Function EncryptPassword(password As String) As String
Public Function DecryptPassword(encryptedPassword As String) As String
Public Sub LogUserActivity(activity As String)
Public Function CheckUserPermissions(userType As String, action As String) As Boolean
Public Sub ShowAccessDeniedMessage()
```

### **Module 3: DatabaseOperations**
```vba
' عمليات قاعدة البيانات
Public Function AddNewInvoice(invoiceData As Variant) As Boolean
Public Function UpdateInvoice(invoiceNumber As String, invoiceData As Variant) As Boolean
Public Function DeleteInvoice(invoiceNumber As String) As Boolean
Public Function SearchInvoices(searchCriteria As Variant) As Variant
Public Function GetInvoiceByNumber(invoiceNumber As String) As Variant
Public Function GetAllInvoices() As Variant
Public Function ValidateInvoiceData(invoiceData As Variant) As String
Public Function CheckDuplicateInvoice(invoiceNumber As String) As Boolean
```

### **Module 4: FileOperations**
```vba
' عمليات الملفات والمستندات
Public Function CreateDirectoryStructure() As Boolean
Public Function UploadDocument(invoiceNumber As String, documentType As String, filePath As String) As Boolean
Public Function OrganizeDocuments() As Boolean
Public Function GetDocumentsByInvoice(invoiceNumber As String) As Variant
Public Function DeleteDocument(attachmentID As Long) As Boolean
Public Function OpenDocument(filePath As String) As Boolean
```

### **Module 5: ImportExportModule**
```vba
' استيراد وتصدير البيانات
Public Function ImportFromExcel(filePath As String, targetSheet As String) As Boolean
Public Function ImportFromPDF(filePath As String) As Variant
Public Function ImportFromWord(filePath As String) As Variant
Public Function ImportFromCSV(filePath As String, delimiter As String) As Boolean
Public Function ExportToExcel(data As Variant, filePath As String) As Boolean
Public Function ExportToPDF(data As Variant, filePath As String) As Boolean
```

### **Module 6: ReportGenerator**
```vba
' توليد التقارير
Public Function GenerateInvoiceReport(invoiceNumber As String) As Boolean
Public Function GenerateMultipleInvoicesReport(criteria As Variant) As Boolean
Public Function GeneratePaymentStatusReport(dateFrom As Date, dateTo As Date) As Boolean
Public Function GenerateFinancialSummaryReport(period As String) As Boolean
Public Function FormatReportTemplate(reportType As String, data As Variant) As Boolean
Public Function PrintReport(reportPath As String) As Boolean
```

### **Module 7: CalculationModule**
```vba
' العمليات الحسابية
Public Function CalculateTotalAmount(amountWithoutStamp As Double, stampDuty As Double) As Double
Public Function CalculateStampDuty(amount As Double, stampRate As Double) As Double
Public Function GetFinancialSummary(dateFrom As Date, dateTo As Date) As Variant
Public Function CalculatePaymentStatistics() As Variant
Public Function ValidateNumericInput(value As String) As Boolean
```

### **Module 8: UIHelpers**
```vba
' مساعدات واجهة المستخدم
Public Sub PopulateComboBox(cmbBox As ComboBox, dataSource As Variant)
Public Sub PopulateListBox(lstBox As ListBox, dataSource As Variant)
Public Sub ClearFormControls(formName As UserForm)
Public Sub SetFormPermissions(formName As UserForm, userType As String)
Public Sub ShowProgressBar(message As String, percentage As Integer)
Public Sub HideProgressBar()
Public Sub ShowNotification(message As String, messageType As String)
```

### **Module 9: ValidationModule**
```vba
' التحقق من صحة البيانات
Public Function ValidateInvoiceNumber(invoiceNumber As String) As Boolean
Public Function ValidateDate(dateValue As String) As Boolean
Public Function ValidateAmount(amount As String) As Boolean
Public Function ValidateRequiredFields(formControls As Variant) As String
Public Function ValidateFileFormat(filePath As String, allowedFormats As String) As Boolean
Public Function SanitizeInput(inputValue As String) As String
```

### **Module 10: NotificationModule**
```vba
' نظام الإشعارات
Public Sub SendPaymentNotification(invoiceNumber As String, recipientType As String)
Public Sub SendOverdueNotification(overdueInvoices As Variant)
Public Sub CreateSystemAlert(alertType As String, message As String)
Public Function GetPendingNotifications() As Variant
Public Sub MarkNotificationAsRead(notificationID As Long)
```

---

## ⚙️ **الوظائف المتقدمة**

### **1. نظام الصلاحيات**
```vba
' تحديد الصلاحيات لكل نوع مستخدم:

' المدير (Manager):
- إضافة/تعديل/حذف جميع البيانات
- إدارة المستخدمين
- الوصول لجميع التقارير
- تصدير البيانات
- إعدادات النظام

' موظف التخليص (Clearance Officer):
- إضافة/تعديل الفواتير
- رفع المستندات
- البحث والتصفية
- طباعة التقارير الأساسية

' المدقق (Auditor):
- عرض البيانات فقط (قراءة)
- البحث والتصفية
- طباعة التقارير
- تصدير البيانات للمراجعة

' المحاسب (Accountant):
- عرض البيانات المالية
- تحديث حالة السداد
- إنشاء التقارير المالية
- متابعة السداد
```

### **2. نظام النسخ الاحتياطي التلقائي**
```vba
Public Sub CreateAutoBackup()
    ' إنشاء نسخة احتياطية تلقائية كل يوم
    ' حفظ النسخ في مجلد منفصل مع التاريخ والوقت
End Sub
```

### **3. نظام التدقيق (Audit Trail)**
```vba
Public Sub LogActivity(userID As String, action As String, details As String)
    ' تسجيل جميع العمليات في جدول منفصل
    ' تتبع من قام بماذا ومتى
End Sub
```

### **4. إشعارات الفواتير المتأخرة**
```vba
Public Sub CheckOverdueInvoices()
    ' فحص الفواتير المتأخرة في السداد
    ' إرسال إشعارات تلقائية
End Sub
```

---

## 📁 **هيكل المجلدات**

```
📁 نظام إدارة الفواتير/
│
├── 📄 InvoiceManagementSystem.xlsm (الملف الرئيسي)
│
├── 📁 Documents/ (المستندات المرفقة)
│   ├── 📁 فاتورة_1001/
│   │   ├── بوليصة_شحن.pdf
│   │   ├── شهادة_منشأ.pdf
│   │   └── فاتورة.pdf
│   └── 📁 فاتورة_1002/
│       └── أمر_تسليم.pdf
│
├── 📁 PDFReports/ (تقارير PDF)
│   ├── تقرير_فاتورة_1001_2025-01-15.pdf
│   └── تقرير_مالي_شامل_2025-01-15.pdf
│
├── 📁 ImportedData/ (البيانات المستوردة)
│   ├── استيراد_2025-01-15.xlsx
│   └── backup_data.csv
│
├── 📁 Backups/ (النسخ الاحتياطية)
│   ├── backup_2025-01-15_09-30.xlsm
│   └── backup_2025-01-14_17-00.xlsm
│
├── 📁 Templates/ (قوالب التقارير)
│   ├── قالب_تقرير_فاتورة.xlsx
│   └── قالب_تقرير_مالي.xlsx
│
├── 📁 Logs/ (سجلات النظام)
│   ├── system_log_2025-01-15.txt
│   └── user_activity_2025-01-15.txt
│
└── 📄 دليل_الاستخدام.pdf
```

---

## 🚀 **خطوات التنفيذ المرحلية**

### **المرحلة الأولى: الإعداد الأساسي**
1. إنشاء ملف Excel جديد (.xlsm)
2. تصميم جداول قاعدة البيانات
3. إنشاء النماذج الأساسية (Login, MainMenu)
4. تطوير نظام تسجيل الدخول

### **المرحلة الثانية: إدارة البيانات**
1. تطوير نموذج إدخال الفواتير
2. إنشاء وظائف CRUD للبيانات
3. تطوير نظام التحقق من صحة البيانات
4. إضافة خاصية البحث والتصفية

### **المرحلة الثالثة: المستندات والمرفقات**
1. تطوير نظام رفع المستندات
2. إنشاء هيكل المجلدات التلقائي
3. ربط المستندات بالفواتير
4. تطوير نظام الأرشفة

### **المرحلة الرابعة: التقارير والطباعة**
1. تصميم قوالب التقارير
2. تطوير مولد التقارير
3. إضافة خاصية التصدير للـ PDF
4. تطوير التقارير المالية

### **المرحلة الخامسة: الميزات المتقدمة**
1. تطوير نظام الاستيراد من ملفات خارجية
2. إضافة نظام الإشعارات
3. تطوير نظام النسخ الاحتياطي
4. إضافة نظام التدقيق

### **المرحلة السادسة: الاختبار والتحسين**
1. اختبار جميع الوظائف
2. تحسين الأداء
3. إضافة معالجة الأخطاء
4. إنشاء دليل المستخدم

---

## 📋 **متطلبات التشغيل**

### **متطلبات النظام:**
- Windows 10 أو أحدث
- Microsoft Excel 2016 أو أحدث
- 4 GB RAM كحد أدنى
- 500 MB مساحة حرة على القرص الصلب
- دقة شاشة 1024x768 كحد أدنى

### **الإعدادات المطلوبة:**
- تفعيل الماكرو في Excel
- السماح بتشغيل VBA
- إعدادات الأمان: متوسط أو منخفض للماكرو

---

## 🛡️ **الأمان وحماية البيانات**

### **1. حماية الملف:**
```vba
' حماية الملف بكلمة مرور
ActiveWorkbook.Password = "SecurePassword123"

' حماية هيكل المصنف
ActiveWorkbook.Protect Password:="StructurePassword", Structure:=True
```

### **2. حماية الأوراق:**
```vba
' حماية أوراق البيانات من التعديل المباشر
Worksheets("Users").Protect Password:="UsersPassword"
Worksheets("Invoices").Protect Password:="InvoicesPassword"
```

### **3. تشفير كلمات المرور:**
```vba
Public Function EncryptPassword(password As String) As String
    ' استخدام خوارزمية تشفير بسيطة
    ' يمكن تحسينها باستخدام خوارزميات أقوى
End Function
```

---

## 📖 **دليل الاستخدام السريع**

### **التشغيل لأول مرة:**
1. فتح ملف `InvoiceManagementSystem.xlsm`
2. السماح بتشغيل الماكرو عند السؤال
3. تسجيل الدخول باستخدام:
   - اسم المستخدم: `admin`
   - كلمة المرور: `1234`

### **إضافة فاتورة جديدة:**
1. من القائمة الرئيسية اختر "إدارة الفواتير"
2. اختر ورقة العمل المناسبة
3. اضغط "إضافة جديد"
4. املأ جميع الحقول المطلوبة
5. اضغط "حفظ"

### **رفع مستند:**
1. من القائمة الرئيسية اختر "إدارة المستندات"
2. اختر رقم الفاتورة
3. حدد نوع المستند
4. اختر الملف من الكمبيوتر
5. اضغط "رفع المستند"

### **إنشاء تقرير:**
1. من القائمة الرئيسية اختر "التقارير"
2. اختر نوع التقرير المطلوب
3. حدد المعايير (التواريخ، الفواتير، إلخ)
4. اضغط "إنشاء التقرير"
5. احفظ أو اطبع التقرير

---

## 🔧 **كود التشغيل التلقائي**

### **في وحدة ThisWorkbook:**
```vba
Private Sub Workbook_Open()
    ' إخفاء جميع الأوراق عدا ورقة البيانات
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Dashboard" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    
    ' إنشاء هيكل المجلدات إذا لم يكن موجوداً
    Call CreateDirectoryStructure
    
    ' تشغيل نموذج تسجيل الدخول
    UserForm_Login.Show
    
    Application.ScreenUpdating = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' إنشاء نسخة احتياطية تلقائية قبل الإغلاق
    Call CreateAutoBackup
    
    ' تسجيل تسجيل الخروج
    Call LogUserActivity("تسجيل خروج من النظام")
End Sub
```

---

## 🎨 **تخصيص الواجهات**

### **الألوان والتصميم:**
```vba
' ألوان النظام
Public Const PRIMARY_COLOR As Long = RGB(41, 128, 185)    ' أزرق
Public Const SECONDARY_COLOR As Long = RGB(52, 152, 219)  ' أزرق فاتح
Public Const SUCCESS_COLOR As Long = RGB(39, 174, 96)     ' أخضر
Public Const WARNING_COLOR As Long = RGB(241, 196, 15)    ' أصفر
Public Const DANGER_COLOR As Long = RGB(231, 76, 60)      ' أحمر
Public Const BACKGROUND_COLOR As Long = RGB(236, 240, 241) ' رمادي فاتح

' تطبيق التصميم على النماذج
Public Sub ApplyFormStyling(frm As UserForm)
    frm.BackColor = BACKGROUND_COLOR
    ' تطبيق الألوان على الأزرار والعناصر
End Sub
```

### **خطوط النصوص:**
```vba
' إعدادات الخطوط
Public Const MAIN_FONT As String = "Tahoma"
Public Const TITLE_FONT_SIZE As Integer = 14
Public Const NORMAL_FONT_SIZE As Integer = 10
Public Const SMALL_FONT_SIZE As Integer = 8
```

---

## 📊 **مؤشرات الأداء**

### **إحصائيات النظام:**
- إجمالي عدد الفواتير
- عدد الفواتير المسددة/غير المسددة
- إجمالي المبالغ المالية
- عدد المستندات المرفقة
- متوسط وقت معالجة الفاتورة

### **تقارير الاستخدام:**
- عدد مرات تسجيل الدخول لكل مستخدم
- أكثر الوظائف استخداماً
- أوقات الذروة في الاستخدام
- معدل الأخطاء والمشاكل

---

## 🔄 **التحديثات والصيانة**

### **نظام التحديثات:**
```vba
Public Function CheckForUpdates() As Boolean
    ' فحص وجود إصدار جديد
    ' تنزيل وتطبيق التحديثات تلقائياً
End Function
```

### **صيانة قاعدة البيانات:**
```vba
Public Sub OptimizeDatabase()
    ' ضغط وتحسين قاعدة البيانات
    ' حذف السجلات المكررة
    ' إعادة تنظيم الفهارس
End Sub
```

---

## 🎯 **الخلاصة**

هذا النظام المتكامل يوفر حلاً شاملاً لإدارة فواتير عوائد المناولة والتخزين باستخدام Excel VBA. النظام مصمم ليكون:

- **سهل الاستخدام**: واجهات بسيطة ومفهومة
- **آمن**: نظام صلاحيات متعدد المستويات
- **مرن**: قابل للتخصيص والتوسع
- **موثوق**: نسخ احتياطية ونظام تدقيق
- **عملي**: يعمل بدون إنترنت على أي جهاز Windows

النظام جاهز للتطوير والتنفيذ وفقاً للمواصفات المذكورة أعلاه.

---

**تاريخ الإنشاء:** 2025-01-15  
**الإصدار:** 1.0  
**المطور:** نظام إدارة الفواتير - Excel VBA  

---

*ملاحظة: هذا المستند يحتوي على المواصفات الكاملة والتفصيلية لتطوير النظام. يُنصح بمراجعة كل قسم بعناية قبل البدء في التطوير.*
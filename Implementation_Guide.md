# دليل التنفيذ العملي
## تحويل المشروع إلى ملف Excel يعمل على حاسوبك

---

## 🎯 **نظرة عامة**

هذا الدليل سيأخذك خطوة بخطوة لتحويل مواصفات المشروع إلى ملف Excel فعلي يعمل على حاسوبك الشخصي.

---

## 📋 **المتطلبات الأساسية**

### **البرامج المطلوبة:**
- ✅ Microsoft Excel 2016 أو أحدث
- ✅ Windows 10/11
- ✅ صلاحيات المدير على الحاسوب

### **الإعدادات المطلوبة:**
1. **تفعيل Developer Tab في Excel:**
   - File → Options → Customize Ribbon
   - ✅ تأشير على "Developer"

2. **إعدادات الماكرو:**
   - File → Options → Trust Center → Trust Center Settings
   - Macro Settings → "Enable all macros"
   - ✅ Trust access to the VBA project object model

---

## 🚀 **الخطوة الأولى: إنشاء الملف الأساسي**

### **1. إنشاء ملف Excel جديد:**
```
1. افتح Excel
2. اختر "Blank workbook"
3. احفظ الملف باسم: "InvoiceManagementSystem.xlsm"
4. تأكد من اختيار نوع الملف: "Excel Macro-Enabled Workbook (*.xlsm)"
```

### **2. إنشاء أوراق العمل الأساسية:**
```
قم بإنشاء الأوراق التالية (كليك يمين على تبويب الورقة → Insert):

📊 Sheet1 → أعد تسميتها إلى: "Users"
📊 Sheet2 → أعد تسميتها إلى: "Invoices"  
📊 Sheet3 → أعد تسميتها إلى: "PaymentTracking"
📊 Sheet4 → أعد تسميتها إلى: "Attachments"
📊 Sheet5 → أعد تسميتها إلى: "SystemLog"
📊 Sheet6 → أعد تسميتها إلى: "SystemSettings"
📊 Sheet7 → أعد تسميتها إلى: "InvoiceReport"
📊 Sheet8 → أعد تسميتها إلى: "Dashboard"
```

---

## 📊 **الخطوة الثانية: إعداد قاعدة البيانات**

### **ورقة "Users" - جدول المستخدمين:**

```
في ورقة "Users"، في الصف الأول اكتب العناوين التالية:

A1: UserID
B1: Username  
C1: PasswordHash
D1: UserType
E1: FullName
F1: IsActive
G1: CreatedDate
H1: LastLoginDate
I1: LoginAttempts
J1: IsLocked

ثم في الصف الثاني أضف المستخدم الافتراضي:
A2: USR001
B2: admin
C2: 5e884898da28047151d0e56f8dc6292773603d0d6aabbdd62a11ef721d1542d8
D2: مدير
E2: مدير النظام
F2: TRUE
G2: =TODAY()
H2: (اتركها فارغة)
I2: 0
J2: FALSE
```

### **ورقة "Invoices" - جدول الفواتير:**

```
في ورقة "Invoices"، في الصف الأول:

A1: InvoiceID
B1: SerialNumber
C1: ClaimNumber
D1: ClaimDate
E1: InvoiceNumber
F1: InvoiceDate
G1: AmountWithoutStamp
H1: StampDuty
I1: TotalAmount
J1: PolicyNumber
K1: LCNumber
L1: ContractNumber
M1: CustomsDeclaration
N1: EntryPort
O1: PaymentStatus
P1: PaymentDate
Q1: CreatedBy
R1: CreatedDate
S1: ModifiedBy
T1: ModifiedDate
U1: IsDeleted
V1: Notes
```

### **ورقة "PaymentTracking" - متابعة السداد:**

```
A1: TrackingID
B1: InvoiceID
C1: ClaimNumber
D1: InvoiceNumber
E1: HandlingStorageFees
F1: StampDuty
G1: TotalFees
H1: PaymentStatus
I1: PaymentMethod
J1: PaymentReference
K1: TransferredTo
L1: TransferDate
M1: TransferReference
N1: ReceivedBy
O1: ReceivedDate
P1: CreatedBy
Q1: CreatedDate
R1: UpdatedBy
S1: UpdatedDate
```

### **ورقة "Attachments" - المرفقات:**

```
A1: AttachmentID
B1: InvoiceID
C1: InvoiceNumber
D1: DocumentType
E1: OriginalFileName
F1: StoredFileName
G1: FilePath
H1: FileSize
I1: FileExtension
J1: MimeType
K1: UploadedBy
L1: UploadDate
M1: IsActive
N1: Description
O1: CheckSum
```

### **ورقة "SystemLog" - سجل النظام:**

```
A1: LogID
B1: UserID
C1: Action
D1: TableName
E1: RecordID
F1: OldValues
G1: NewValues
H1: IPAddress
I1: UserAgent
J1: LogDate
K1: LogLevel
L1: Description
```

### **ورقة "SystemSettings" - إعدادات النظام:**

```
A1: SettingID
B1: SettingKey
C1: SettingValue
D1: SettingType
E1: Description
F1: IsEditable
G1: Category
H1: UpdatedBy
I1: UpdatedDate

وأضف بعض الإعدادات الافتراضية:
A2: 1    B2: STAMP_DUTY_RATE    C2: 0.05    D2: NUMBER
A3: 2    B3: AUTO_BACKUP_ENABLED    C3: TRUE    D3: BOOLEAN
A4: 3    B4: BACKUP_RETENTION_DAYS    C4: 30    D4: NUMBER
```

---

## 🎨 **الخطوة الثالثة: إنشاء النماذج (UserForms)**

### **فتح محرر VBA:**
```
1. اضغط Alt + F11 لفتح محرر VBA
2. في النافيسة اليسرى ستجد "VBAProject (InvoiceManagementSystem.xlsm)"
```

### **إنشاء النموذج الأول - تسجيل الدخول:**

```
1. كليك يمين على VBAProject → Insert → UserForm
2. غير اسم النموذج إلى: "UserForm_Login"
3. في نافذة Properties:
   - Caption: "تسجيل الدخول - نظام إدارة الفواتير"
   - Width: 400
   - Height: 300
   - StartUpPosition: 2-CenterScreen

4. أضف العناصر التالية من Toolbox:

📋 Label (lblTitle):
   - Caption: "نظام إدارة فواتير عوائد المناولة والتخزين"
   - Font: Tahoma, 12pt, Bold
   - Left: 50, Top: 20, Width: 300

📋 Label (lblUsername):
   - Caption: "اسم المستخدم:"
   - Left: 50, Top: 80

📋 TextBox (txtUsername):
   - Left: 150, Top: 77, Width: 200

📋 Label (lblPassword):
   - Caption: "كلمة المرور:"
   - Left: 50, Top: 110

📋 TextBox (txtPassword):
   - PasswordChar: "*"
   - Left: 150, Top: 107, Width: 200

📋 CommandButton (cmdLogin):
   - Caption: "دخول"
   - Left: 270, Top: 150, Width: 80, Height: 30

📋 CommandButton (cmdExit):
   - Caption: "خروج"
   - Left: 180, Top: 150, Width: 80, Height: 30
```

---

## 💻 **الخطوة الرابعة: إضافة الأكواد الأساسية**

### **إنشاء Module للمتغيرات العامة:**

```
1. كليك يمين على VBAProject → Insert → Module
2. غير اسم الـ Module إلى: "GlobalVariables"
3. انسخ الكود التالي:
```

```vba
Option Explicit

' ثوابت النظام
Public Const APP_NAME As String = "نظام إدارة فواتير عوائد المناولة والتخزين"
Public Const APP_VERSION As String = "1.0.0"

' متغيرات المستخدم الحالي
Public Type UserInfo
    UserID As String
    Username As String
    FullName As String
    UserType As String
    IsActive As Boolean
    LastLogin As Date
    Permissions As String
End Type

Public CurrentUser As UserInfo
Public IsLoggedIn As Boolean

' مسارات المجلدات
Public Const DOCUMENTS_FOLDER As String = "Documents"
Public Const REPORTS_FOLDER As String = "PDFReports"
Public Const BACKUPS_FOLDER As String = "Backups"
```

### **إنشاء Module للوظائف الأساسية:**

```
1. أنشئ Module جديد باسم: "DatabaseOperations"
2. انسخ الكود التالي:
```

```vba
Option Explicit

Public Function InitializeSystem() As Boolean
    On Error GoTo ErrorHandler
    
    ' إنشاء هيكل المجلدات
    Call CreateDirectoryStructure
    
    ' تنسيق الجداول
    Call FormatDatabaseTables
    
    InitializeSystem = True
    Exit Function
    
ErrorHandler:
    MsgBox "خطأ في تهيئة النظام: " & Err.Description, vbCritical
    InitializeSystem = False
End Function

Public Sub CreateDirectoryStructure()
    Dim basePath As String
    Dim fso As Object
    
    basePath = ThisWorkbook.Path & "\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' إنشاء المجلدات الأساسية
    If Not fso.FolderExists(basePath & DOCUMENTS_FOLDER) Then
        fso.CreateFolder basePath & DOCUMENTS_FOLDER
    End If
    
    If Not fso.FolderExists(basePath & REPORTS_FOLDER) Then
        fso.CreateFolder basePath & REPORTS_FOLDER
    End If
    
    If Not fso.FolderExists(basePath & BACKUPS_FOLDER) Then
        fso.CreateFolder basePath & BACKUPS_FOLDER
    End If
End Sub

Public Sub FormatDatabaseTables()
    Dim ws As Worksheet
    
    ' تنسيق جدول المستخدمين
    Set ws = ThisWorkbook.Worksheets("Users")
    With ws.Range("A1:J1")
        .Font.Bold = True
        .Interior.Color = RGB(52, 152, 219)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    ws.Columns.AutoFit
    
    ' تنسيق جدول الفواتير
    Set ws = ThisWorkbook.Worksheets("Invoices")
    With ws.Range("A1:V1")
        .Font.Bold = True
        .Interior.Color = RGB(39, 174, 96)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    ws.Columns.AutoFit
    
    ' تنسيق باقي الجداول بنفس الطريقة
    Call FormatSheet("PaymentTracking", "A1:S1", RGB(241, 196, 15))
    Call FormatSheet("Attachments", "A1:O1", RGB(155, 89, 182))
    Call FormatSheet("SystemLog", "A1:L1", RGB(231, 76, 60))
    Call FormatSheet("SystemSettings", "A1:I1", RGB(52, 73, 94))
End Sub

Private Sub FormatSheet(sheetName As String, headerRange As String, color As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    With ws.Range(headerRange)
        .Font.Bold = True
        .Interior.Color = color
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    ws.Columns.AutoFit
End Sub
```

### **إضافة كود تسجيل الدخول:**

```
1. ارجع إلى نموذج UserForm_Login
2. اضغط مرتين على زر "دخول" (cmdLogin)
3. انسخ الكود التالي:
```

```vba
Private Sub cmdLogin_Click()
    Dim username As String
    Dim password As String
    
    username = Trim(txtUsername.Text)
    password = Trim(txtPassword.Text)
    
    ' التحقق من الحقول الفارغة
    If username = "" Or password = "" Then
        MsgBox "يرجى إدخال اسم المستخدم وكلمة المرور", vbExclamation
        Exit Sub
    End If
    
    ' التحقق من بيانات الدخول
    If ValidateLogin(username, password) Then
        MsgBox "مرحباً " & CurrentUser.FullName, vbInformation
        Me.Hide
        ' هنا ستفتح الواجهة الرئيسية لاحقاً
        ' UserForm_MainMenu.Show
    Else
        MsgBox "اسم المستخدم أو كلمة المرور غير صحيحة", vbCritical
        txtPassword.Text = ""
        txtUsername.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    ThisWorkbook.Close False
End Sub
```

### **إضافة دالة التحقق من تسجيل الدخول:**

```
في Module "DatabaseOperations" أضف:
```

```vba
Public Function ValidateLogin(username As String, password As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim storedHash As String
    Dim inputHash As String
    
    Set ws = ThisWorkbook.Worksheets("Users")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    inputHash = SimpleHash(password)
    
    For i = 2 To lastRow
        If ws.Cells(i, 2).Value = username Then
            storedHash = ws.Cells(i, 3).Value
            
            If storedHash = inputHash And ws.Cells(i, 6).Value = True Then
                ' تحميل بيانات المستخدم
                With CurrentUser
                    .UserID = ws.Cells(i, 1).Value
                    .Username = ws.Cells(i, 2).Value
                    .FullName = ws.Cells(i, 5).Value
                    .UserType = ws.Cells(i, 4).Value
                    .IsActive = ws.Cells(i, 6).Value
                End With
                
                IsLoggedIn = True
                ws.Cells(i, 8).Value = Now() ' تحديث آخر دخول
                
                ValidateLogin = True
                Exit Function
            End If
        End If
    Next i
    
    ValidateLogin = False
    Exit Function
    
ErrorHandler:
    ValidateLogin = False
End Function

Public Function SimpleHash(password As String) As String
    ' دالة تشفير بسيطة (للاختبار فقط)
    Dim i As Integer
    Dim hashValue As Long
    
    hashValue = 5381
    
    For i = 1 To Len(password)
        hashValue = ((hashValue * 33) Xor Asc(Mid(password, i, 1))) And &H7FFFFFFF
    Next i
    
    SimpleHash = Hex(hashValue)
End Function
```

---

## 🔧 **الخطوة الخامسة: إعداد التشغيل التلقائي**

### **إضافة كود التشغيل التلقائي:**

```
1. في محرر VBA، اضغط مرتين على "ThisWorkbook"
2. انسخ الكود التالي:
```

```vba
Private Sub Workbook_Open()
    ' تهيئة النظام
    Application.ScreenUpdating = False
    
    ' إخفاء جميع الأوراق عدا Dashboard
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Dashboard" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    
    ' تهيئة النظام
    Call InitializeSystem
    
    ' عرض نموذج تسجيل الدخول
    UserForm_Login.Show
    
    Application.ScreenUpdating = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' حفظ تلقائي
    ThisWorkbook.Save
End Sub
```

---

## 🧪 **الخطوة السادسة: الاختبار الأولي**

### **اختبار تسجيل الدخول:**

```
1. احفظ الملف (Ctrl + S)
2. أغلق الملف وأعد فتحه
3. يجب أن يظهر نموذج تسجيل الدخول تلقائياً
4. جرب تسجيل الدخول بـ:
   - اسم المستخدم: admin
   - كلمة المرور: 1234

إذا نجح تسجيل الدخول، ستظهر رسالة ترحيب
```

### **التحقق من إنشاء المجلدات:**

```
1. اذهب إلى مجلد الملف على الكمبيوتر
2. يجب أن تجد المجلدات التالية قد تم إنشاؤها:
   📁 Documents
   📁 PDFReports  
   📁 Backups
```

### **التحقق من تنسيق الجداول:**

```
1. في Excel، اضغط Alt + F11 للخروج من محرر VBA
2. تحقق من أن عناوين الجداول منسقة بألوان مختلفة
3. تأكد من أن الأعمدة متناسقة الحجم
```

---

## 🚨 **حل المشاكل الشائعة**

### **مشكلة: لا يظهر نموذج تسجيل الدخول**
```
الحل:
1. تأكد من حفظ الملف بصيغة .xlsm
2. تأكد من تفعيل الماكرو عند فتح الملف
3. في محرر VBA تحقق من وجود الكود في ThisWorkbook
```

### **مشكلة: خطأ في تشغيل الكود**
```
الحل:
1. تحقق من أسماء الأوراق (Users, Invoices, etc.)
2. تأكد من وضع العناوين في الصف الأول
3. تحقق من عدم وجود أخطاء إملائية في الكود
```

### **مشكلة: لا يتم إنشاء المجلدات**
```
الحل:
1. تحقق من صلاحيات الكتابة في مجلد الملف
2. شغل Excel كمدير (Run as Administrator)
3. تأكد من عدم وجود برامج حماية تمنع إنشاء المجلدات
```

---

## ✅ **نقاط التحقق النهائية**

قبل الانتقال للمرحلة التالية، تأكد من:

- ✅ تم إنشاء الملف بصيغة .xlsm
- ✅ تم إنشاء جميع الأوراق (8 أوراق)
- ✅ تم إضافة العناوين لكل جدول
- ✅ تم إنشاء نموذج تسجيل الدخول
- ✅ تم إضافة الأكواد الأساسية
- ✅ تم اختبار تسجيل الدخول بنجاح
- ✅ تم إنشاء المجلدات تلقائياً
- ✅ الجداول منسقة بالألوان

---

## 🎯 **الخطوة التالية**

بعد إنجاز هذه المرحلة بنجاح، ستكون جاهزاً لـ:

1. **إنشاء الواجهة الرئيسية** (Main Menu)
2. **إضافة نموذج إدخال الفواتير**
3. **تطوير وظائف البحث والتصفية**
4. **إضافة نظام التقارير**

**هل تريد المتابعة لإنشاء الواجهة الرئيسية؟** 🚀
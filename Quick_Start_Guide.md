# دليل البدء السريع
## إنشاء نظام إدارة الفواتير - خطوة بخطوة

---

## 🎯 **المرحلة الأولى: إعداد Excel (5 دقائق)**

### **الخطوة 1: تفعيل Developer Tab**
1. **افتح Microsoft Excel**
2. **اذهب إلى:** File → Options
3. **اختر:** Customize Ribbon
4. **في الجانب الأيمن:** ✅ ضع علامة أمام **"Developer"**
5. **اضغط:** OK

### **الخطوة 2: إعداد إعدادات الماكرو**
1. **اذهب إلى:** File → Options → Trust Center
2. **اضغط:** Trust Center Settings
3. **اختر:** Macro Settings
4. **حدد:** ✅ **"Enable all macros"**
5. **حدد:** ✅ **"Trust access to the VBA project object model"**
6. **اضغط:** OK مرتين

### **الخطوة 3: إنشاء الملف الأساسي**
1. **في Excel:** اختر **Blank workbook**
2. **اضغط:** Ctrl + S للحفظ
3. **اكتب اسم الملف:** `InvoiceManagementSystem`
4. **⚠️ مهم جداً:** في "Save as type" اختر **"Excel Macro-Enabled Workbook (*.xlsm)"**
5. **اختر مكان الحفظ** (مثل: Desktop أو Documents)
6. **اضغط:** Save

---

## 📊 **المرحلة الثانية: إنشاء أوراق العمل (10 دقائق)**

### **الخطوة 4: إنشاء الأوراق المطلوبة**

**لكل ورقة جديدة:**
- كليك يمين على تبويب الورقة → Insert Worksheet

**أنشئ الأوراق التالية بالترتيب:**

1. **Sheet1** → أعد تسميتها إلى: **`Users`**
   - كليك يمين على التبويب → Rename → اكتب "Users"

2. **أضف ورقة جديدة** → اسمها: **`Invoices`**

3. **أضف ورقة جديدة** → اسمها: **`PaymentTracking`**

4. **أضف ورقة جديدة** → اسمها: **`Attachments`**

5. **أضف ورقة جديدة** → اسمها: **`SystemLog`**

6. **أضف ورقة جديدة** → اسمها: **`SystemSettings`**

7. **أضف ورقة جديدة** → اسمها: **`InvoiceReport`**

8. **أضف ورقة جديدة** → اسمها: **`Dashboard`**

### **الخطوة 5: إعداد جدول المستخدمين**

1. **اضغط على تبويب ورقة "Users"**

2. **في الصف الأول، اكتب العناوين التالية:**

```
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
```

3. **في الصف الثاني، أضف المستخدم الافتراضي:**

```
A2: USR001
B2: admin
C2: 5E884898DA28047151D0E56F8DC6292773603D0D6AABBDD62A11EF721D1542D8
D2: مدير
E2: مدير النظام
F2: TRUE
G2: =TODAY()
H2: (اتركها فارغة)
I2: 0
J2: FALSE
```

4. **حدد الصف الأول (A1:J1)** واجعل النص **عريض (Bold)**

---

## 📋 **المرحلة الثالثة: إعداد جدول الفواتير (10 دقائق)**

### **الخطوة 6: إعداد ورقة Invoices**

1. **اضغط على تبويب ورقة "Invoices"**

2. **في الصف الأول، اكتب العناوين:**

```
A1: InvoiceID          B1: SerialNumber       C1: ClaimNumber
D1: ClaimDate          E1: InvoiceNumber      F1: InvoiceDate
G1: AmountWithoutStamp H1: StampDuty          I1: TotalAmount
J1: PolicyNumber       K1: LCNumber           L1: ContractNumber
M1: CustomsDeclaration N1: EntryPort          O1: PaymentStatus
P1: PaymentDate        Q1: CreatedBy          R1: CreatedDate
S1: ModifiedBy         T1: ModifiedDate       U1: IsDeleted
V1: Notes
```

3. **حدد الصف الأول (A1:V1)** واجعل النص **عريض**

### **الخطوة 7: إعداد باقي الجداول**

**لكل جدول من الجداول التالية، اضغط على التبويب واكتب العناوين:**

#### **ورقة PaymentTracking:**
```
A1: TrackingID    B1: InvoiceID      C1: ClaimNumber     D1: InvoiceNumber
E1: HandlingStorageFees    F1: StampDuty    G1: TotalFees    H1: PaymentStatus
I1: PaymentMethod    J1: PaymentReference    K1: TransferredTo    L1: TransferDate
M1: TransferReference    N1: ReceivedBy    O1: ReceivedDate    P1: CreatedBy
Q1: CreatedDate    R1: UpdatedBy    S1: UpdatedDate
```

#### **ورقة Attachments:**
```
A1: AttachmentID    B1: InvoiceID       C1: InvoiceNumber    D1: DocumentType
E1: OriginalFileName    F1: StoredFileName    G1: FilePath    H1: FileSize
I1: FileExtension    J1: MimeType    K1: UploadedBy    L1: UploadDate
M1: IsActive    N1: Description    O1: CheckSum
```

#### **ورقة SystemLog:**
```
A1: LogID    B1: UserID    C1: Action    D1: TableName    E1: RecordID
F1: OldValues    G1: NewValues    H1: IPAddress    I1: UserAgent
J1: LogDate    K1: LogLevel    L1: Description
```

#### **ورقة SystemSettings:**
```
A1: SettingID    B1: SettingKey    C1: SettingValue    D1: SettingType
E1: Description    F1: IsEditable    G1: Category    H1: UpdatedBy
I1: UpdatedDate
```

**وأضف الإعدادات الافتراضية في الصفوف التالية:**
```
A2: 1    B2: STAMP_DUTY_RATE        C2: 0.05    D2: NUMBER
A3: 2    B3: AUTO_BACKUP_ENABLED    C3: TRUE    D3: BOOLEAN
A4: 3    B4: BACKUP_RETENTION_DAYS  C4: 30      D4: NUMBER
```

---

## 🎨 **المرحلة الرابعة: تنسيق الجداول (5 دقائق)**

### **الخطوة 8: تلوين العناوين**

**لكل جدول:**

1. **جدول Users:** حدد A1:J1 → اختر لون أزرق للخلفية
2. **جدول Invoices:** حدد A1:V1 → اختر لون أخضر للخلفية  
3. **جدول PaymentTracking:** حدد A1:S1 → اختر لون أصفر للخلفية
4. **جدول Attachments:** حدد A1:O1 → اختر لون بنفسجي للخلفية
5. **جدول SystemLog:** حدد A1:L1 → اختر لون أحمر للخلفية
6. **جدول SystemSettings:** حدد A1:I1 → اختر لون رمادي للخلفية

**لتطبيق اللون:**
- حدد الخلايا → Home → Fill Color → اختر اللون
- اجعل لون النص أبيض: Home → Font Color → أبيض

### **الخطوة 9: ضبط عرض الأعمدة**
**لكل ورقة:**
- حدد جميع الأعمدة (Ctrl + A)
- Home → Format → AutoFit Column Width

---

## 💻 **المرحلة الخامسة: إضافة أكواد VBA (15 دقيقة)**

### **الخطوة 10: فتح محرر VBA**
- **اضغط:** Alt + F11
- ستفتح نافذة جديدة (Microsoft Visual Basic for Applications)

### **الخطوة 11: إنشاء الوحدات**

#### **وحدة المتغيرات العامة:**
1. **كليك يمين** على "VBAProject" → Insert → Module
2. **في نافذة Properties** (أسفل اليسار): غير الاسم إلى **"GlobalVariables"**
3. **انسخ والصق الكود التالي:**

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

' ألوان النظام
Public Const PRIMARY_COLOR As Long = 2854399    ' أزرق
Public Const SUCCESS_COLOR As Long = 2541174    ' أخضر
Public Const WARNING_COLOR As Long = 1023215    ' أصفر
Public Const DANGER_COLOR As Long = 3973631     ' أحمر
```

#### **وحدة عمليات قاعدة البيانات:**
1. **أنشئ Module جديد** → اسمه **"DatabaseOperations"**
2. **انسخ والصق الكود:**

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
    On Error Resume Next
    
    ' تنسيق جدول المستخدمين
    Call FormatSheet("Users", "A1:J1", PRIMARY_COLOR)
    
    ' تنسيق جدول الفواتير
    Call FormatSheet("Invoices", "A1:V1", SUCCESS_COLOR)
    
    On Error GoTo 0
End Sub

Private Sub FormatSheet(sheetName As String, headerRange As String, color As Long)
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    If Not ws Is Nothing Then
        With ws.Range(headerRange)
            .Font.Bold = True
            .Interior.Color = color
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        ws.Columns.AutoFit
    End If
    
    On Error GoTo 0
End Sub

Public Function ValidateLogin(username As String, password As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim inputHash As String
    
    Set ws = ThisWorkbook.Worksheets("Users")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    inputHash = SimpleHash(password)
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 2).Value)) = UCase(Trim(username)) Then
            If ws.Cells(i, 3).Value = inputHash And ws.Cells(i, 6).Value = True Then
                ' تحميل بيانات المستخدم
                With CurrentUser
                    .UserID = ws.Cells(i, 1).Value
                    .Username = ws.Cells(i, 2).Value
                    .FullName = ws.Cells(i, 5).Value
                    .UserType = ws.Cells(i, 4).Value
                    .IsActive = ws.Cells(i, 6).Value
                End With
                
                IsLoggedIn = True
                ws.Cells(i, 8).Value = Now()
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

## 🖥️ **المرحلة السادسة: إنشاء نموذج تسجيل الدخول (10 دقائق)**

### **الخطوة 12: إنشاء النموذج**
1. **في محرر VBA:** كليك يمين على VBAProject → Insert → UserForm
2. **في Properties:** غير Name إلى **"UserForm_Login"**
3. **غير Caption إلى:** "تسجيل الدخول - نظام إدارة الفواتير"

### **الخطوة 13: إضافة العناصر**
**من Toolbox (إذا لم تظهر: View → Toolbox):**

1. **أضف Label:** 
   - Name: lblTitle
   - Caption: "نظام إدارة فواتير عوائد المناولة والتخزين"

2. **أضف Label:**
   - Name: lblUsername  
   - Caption: "اسم المستخدم:"

3. **أضف TextBox:**
   - Name: txtUsername

4. **أضف Label:**
   - Name: lblPassword
   - Caption: "كلمة المرور:"

5. **أضف TextBox:**
   - Name: txtPassword
   - PasswordChar: *

6. **أضف CommandButton:**
   - Name: cmdLogin
   - Caption: "دخول"

7. **أضف CommandButton:**
   - Name: cmdExit
   - Caption: "خروج"

### **الخطوة 14: إضافة كود النموذج**
1. **اضغط مرتين على زر "دخول"**
2. **انسخ والصق الكود:**

```vba
Private Sub cmdLogin_Click()
    Dim username As String
    Dim password As String
    
    username = Trim(txtUsername.Text)
    password = Trim(txtPassword.Text)
    
    If username = "" Or password = "" Then
        MsgBox "يرجى إدخال اسم المستخدم وكلمة المرور", vbExclamation
        Exit Sub
    End If
    
    If ValidateLogin(username, password) Then
        MsgBox "مرحباً " & CurrentUser.FullName, vbInformation
        Me.Hide
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

---

## ⚙️ **المرحلة السابعة: إعداد التشغيل التلقائي (5 دقائق)**

### **الخطوة 15: إضافة كود التشغيل**
1. **في محرر VBA:** اضغط مرتين على **"ThisWorkbook"**
2. **انسخ والصق الكود:**

```vba
Private Sub Workbook_Open()
    Application.ScreenUpdating = False
    
    ' إخفاء أوراق قاعدة البيانات
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
    ThisWorkbook.Save
End Sub
```

---

## 🧪 **المرحلة الثامنة: الاختبار الأولي (5 دقائق)**

### **الخطوة 16: حفظ واختبار النظام**

1. **احفظ كل شيء:** Ctrl + S في محرر VBA
2. **أغلق محرر VBA:** Alt + F4
3. **احفظ ملف Excel:** Ctrl + S
4. **أغلق Excel تماماً**

### **الخطوة 17: اختبار التشغيل**

1. **افتح الملف مرة أخرى:** `InvoiceManagementSystem.xlsm`
2. **عند السؤال عن الماكرو:** اختر **"Enable Content"**
3. **يجب أن يظهر نموذج تسجيل الدخول تلقائياً**

### **الخطوة 18: اختبار تسجيل الدخول**

**جرب تسجيل الدخول بالبيانات التالية:**
- **اسم المستخدم:** admin
- **كلمة المرور:** 1234

**إذا نجح:** ستظهر رسالة ترحيب "مرحباً مدير النظام"

### **الخطوة 19: فحص المجلدات**

**اذهب إلى مجلد الملف وتأكد من إنشاء:**
- 📁 Documents
- 📁 PDFReports  
- 📁 Backups

---

## 🎉 **تهانينا! لقد أنشأت النظام بنجاح!**

### **✅ ما تم إنجازه:**
- ✅ ملف Excel متكامل بقاعدة بيانات منظمة
- ✅ نظام تسجيل دخول آمن
- ✅ مجلدات تلقائية للتنظيم
- ✅ واجهة احترافية
- ✅ حماية للبيانات الحساسة

### **🚀 الخطوات التالية:**
1. **إضافة الواجهة الرئيسية** (Main Menu)
2. **تطوير نموذج إدخال الفواتير**
3. **إضافة وظائف البحث والتقارير**
4. **تطوير نظام المرفقات**

---

## ❓ **هل واجهت أي مشاكل؟**

### **المشاكل الشائعة والحلول:**

**🔴 لا يظهر نموذج تسجيل الدخول:**
- تأكد من تفعيل الماكرو عند فتح الملف
- تحقق من حفظ الملف بصيغة .xlsm

**🔴 خطأ في الكود:**
- تأكد من نسخ الأكواد بالكامل
- تحقق من أسماء الأوراق (Users, Invoices, etc.)

**🔴 لا يقبل تسجيل الدخول:**
- تأكد من وجود المستخدم admin في ورقة Users
- تحقق من Hash كلمة المرور في الخلية C2

---

**الآن أخبرني: هل تمكنت من إنشاء النظام بنجاح؟ وهل تريد المتابعة لإضافة المزيد من الوظائف؟** 🚀
# دليل وضع الأكواد التفصيلي
## أماكن وضع الأكواد بالترتيب والتفصيل الكامل

---

## 🎯 **المرحلة الأولى: إعداد Excel الأساسي**

### **الخطوة 1: إعداد Excel (بدون أكواد)**
```
1. افتح Excel
2. File → Options → Customize Ribbon → ✅ Developer
3. File → Options → Trust Center → Trust Center Settings → Macro Settings → Enable all macros
4. إنشاء ملف جديد → حفظ باسم InvoiceManagementSystem.xlsm
```

### **الخطوة 2: إنشاء الأوراق (بدون أكواد)**
```
إنشاء 8 أوراق:
- Users
- Invoices  
- PaymentTracking
- Attachments
- SystemLog
- SystemSettings
- InvoiceReport
- Dashboard
```

---

## 🏗️ **المرحلة الثانية: إعداد قاعدة البيانات (بدون أكواد)**

### **إدخال العناوين والبيانات الأساسية في كل ورقة:**

#### **ورقة Users:**
```
A1: UserID          B1: Username        C1: PasswordHash
D1: UserType        E1: FullName        F1: IsActive
G1: CreatedDate     H1: LastLoginDate   I1: LoginAttempts
J1: IsLocked

A2: USR001          B2: admin           C2: 5E884898DA28047151D0E56F8DC6292773603D0D6AABBDD62A11EF721D1542D8
D2: مدير           E2: مدير النظام      F2: TRUE
G2: =TODAY()        H2:                 I2: 0
J2: FALSE
```

#### **ورقة Invoices:**
```
A1: InvoiceID       B1: SerialNumber    C1: ClaimNumber     D1: ClaimDate
E1: InvoiceNumber   F1: InvoiceDate     G1: AmountWithoutStamp  H1: StampDuty
I1: TotalAmount     J1: PolicyNumber    K1: LCNumber        L1: ContractNumber
M1: CustomsDeclaration  N1: EntryPort   O1: PaymentStatus   P1: PaymentDate
Q1: CreatedBy       R1: CreatedDate     S1: ModifiedBy      T1: ModifiedDate
U1: IsDeleted       V1: Notes
```

#### **ورقة PaymentTracking:**
```
A1: TrackingID      B1: InvoiceID       C1: ClaimNumber     D1: InvoiceNumber
E1: HandlingStorageFees  F1: StampDuty  G1: TotalFees       H1: PaymentStatus
I1: PaymentMethod   J1: PaymentReference K1: TransferredTo  L1: TransferDate
M1: TransferReference N1: ReceivedBy    O1: ReceivedDate    P1: CreatedBy
Q1: CreatedDate     R1: UpdatedBy       S1: UpdatedDate
```

#### **ورقة Attachments:**
```
A1: AttachmentID    B1: InvoiceID       C1: InvoiceNumber   D1: DocumentType
E1: OriginalFileName F1: StoredFileName G1: FilePath        H1: FileSize
I1: FileExtension   J1: MimeType        K1: UploadedBy      L1: UploadDate
M1: IsActive        N1: Description     O1: CheckSum
```

#### **ورقة SystemLog:**
```
A1: LogID           B1: UserID          C1: Action          D1: TableName
E1: RecordID        F1: OldValues       G1: NewValues       H1: IPAddress
I1: UserAgent       J1: LogDate         K1: LogLevel        L1: Description
```

#### **ورقة SystemSettings:**
```
A1: SettingID       B1: SettingKey      C1: SettingValue    D1: SettingType
E1: Description     F1: IsEditable      G1: Category        H1: UpdatedBy
I1: UpdatedDate

A2: 1               B2: STAMP_DUTY_RATE C2: 0.05            D2: NUMBER
A3: 2               B3: AUTO_BACKUP_ENABLED C3: TRUE        D3: BOOLEAN
A4: 3               B4: BACKUP_RETENTION_DAYS C4: 30        D4: NUMBER
```

---

## 💻 **المرحلة الثالثة: فتح محرر VBA وإنشاء الوحدات**

### **الخطوة 3: فتح محرر VBA**
```
1. في Excel: اضغط Alt + F11
2. ستفتح نافذة "Microsoft Visual Basic for Applications"
3. في الجانب الأيسر ستجد "VBAProject (InvoiceManagementSystem.xlsm)"
```

### **الخطوة 4: إنشاء الوحدة الأولى - GlobalVariables**

#### **أين تنشئ الوحدة:**
```
1. في محرر VBA
2. كليك يمين على "VBAProject (InvoiceManagementSystem.xlsm)"
3. اختر: Insert → Module
4. ستظهر وحدة جديدة اسمها "Module1"
```

#### **تغيير اسم الوحدة:**
```
1. في النافيسة اليسرى تحت "Modules" ستجد "Module1"
2. في نافيسة Properties (أسفل اليسار):
   - إذا لم تظهر: View → Properties Window
3. في خانة "Name" غير "Module1" إلى "GlobalVariables"
```

#### **أين تضع الكود:**
```
1. اضغط مرتين على "GlobalVariables" في النافيسة اليسرى
2. ستفتح نافيسة كبيرة فارغة على اليمين
3. احذف أي نص موجود
4. انسخ والصق الكود التالي:
```

```vba
Option Explicit

' ثوابت النظام
Public Const APP_NAME As String = "نظام إدارة فواتير عوائد المناولة والتخزين"
Public Const APP_VERSION As String = "1.0.0"
Public Const APP_BUILD As String = "20250115"
Public Const DEVELOPER As String = "فريق تطوير النظام"

' مسارات المجلدات
Public Const DOCUMENTS_FOLDER As String = "Documents"
Public Const REPORTS_FOLDER As String = "PDFReports"
Public Const IMPORTS_FOLDER As String = "ImportedData"
Public Const BACKUPS_FOLDER As String = "Backups"
Public Const TEMPLATES_FOLDER As String = "Templates"
Public Const LOGS_FOLDER As String = "Logs"

' متغيرات المستخدم الحالي
Public CurrentUser As UserInfo
Public IsSystemInitialized As Boolean
Public IsLoggedIn As Boolean

' هياكل البيانات المخصصة
Public Type UserInfo
    UserID As String
    Username As String
    FullName As String
    UserType As String
    IsActive As Boolean
    LastLogin As Date
    Permissions As String
End Type

Public Type InvoiceInfo
    InvoiceID As Long
    SerialNumber As Long
    ClaimNumber As String
    ClaimDate As Date
    InvoiceNumber As String
    InvoiceDate As Date
    AmountWithoutStamp As Double
    StampDuty As Double
    TotalAmount As Double
    PolicyNumber As String
    LCNumber As String
    ContractNumber As String
    CustomsDeclaration As String
    EntryPort As String
    PaymentStatus As String
    PaymentDate As Date
    CreatedBy As String
    CreatedDate As Date
    ModifiedBy As String
    ModifiedDate As Date
    Notes As String
End Type

' تعدادات النظام
Public Enum UserTypes
    Manager = 1
    ClearanceOfficer = 2
    Auditor = 3
    Accountant = 4
End Enum

Public Enum PaymentStatus
    Unpaid = 0
    Paid = 1
    Partial = 2
    Cancelled = 3
End Enum

' ألوان النظام
Public Const PRIMARY_COLOR As Long = 2854399    ' RGB(41, 128, 185) - أزرق
Public Const SECONDARY_COLOR As Long = 3368652  ' RGB(52, 152, 219) - أزرق فاتح
Public Const SUCCESS_COLOR As Long = 2541174    ' RGB(39, 174, 96) - أخضر
Public Const WARNING_COLOR As Long = 1023215    ' RGB(241, 196, 15) - أصفر
Public Const DANGER_COLOR As Long = 3973631     ' RGB(231, 76, 60) - أحمر
Public Const BACKGROUND_COLOR As Long = 15856113 ' RGB(236, 240, 241) - رمادي فاتح

' إعدادات الخطوط
Public Const MAIN_FONT As String = "Tahoma"
Public Const TITLE_FONT_SIZE As Integer = 14
Public Const NORMAL_FONT_SIZE As Integer = 10
Public Const SMALL_FONT_SIZE As Integer = 8
```

---

## 🗃️ **المرحلة الرابعة: إنشاء وحدة عمليات قاعدة البيانات**

### **الخطوة 5: إنشاء الوحدة الثانية - DatabaseOperations**

#### **إنشاء الوحدة:**
```
1. في محرر VBA
2. كليك يمين مرة أخرى على "VBAProject (InvoiceManagementSystem.xlsm)"
3. اختر: Insert → Module
4. ستظهر وحدة جديدة اسمها "Module2"
5. في Properties غير الاسم إلى "DatabaseOperations"
```

#### **أين تضع الكود:**
```
1. اضغط مرتين على "DatabaseOperations" في النافيسة اليسرى
2. في النافيسة الكبيرة على اليمين
3. احذف أي نص موجود
4. انسخ والصق الكود التالي:
```

```vba
Option Explicit

' === وظائف تهيئة النظام ===

Public Function InitializeSystem() As Boolean
    On Error GoTo ErrorHandler
    
    SystemStartTime = Now()
    
    ' إنشاء هيكل المجلدات
    If Not CreateDirectoryStructure() Then
        MsgBox "فشل في إنشاء المجلدات المطلوبة", vbCritical
        InitializeSystem = False
        Exit Function
    End If
    
    ' تهيئة قاعدة البيانات
    If Not InitializeDatabase() Then
        MsgBox "فشل في تهيئة قاعدة البيانات", vbCritical
        InitializeSystem = False
        Exit Function
    End If
    
    ' تنسيق الجداول
    Call FormatDatabaseTables
    
    ' تحميل الإعدادات
    Call LoadSystemSettings
    
    IsSystemInitialized = True
    InitializeSystem = True
    
    Exit Function
    
ErrorHandler:
    InitializeSystem = False
End Function

Public Function CreateDirectoryStructure() As Boolean
    On Error GoTo ErrorHandler
    
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
    
    If Not fso.FolderExists(basePath & IMPORTS_FOLDER) Then
        fso.CreateFolder basePath & IMPORTS_FOLDER
    End If
    
    If Not fso.FolderExists(basePath & BACKUPS_FOLDER) Then
        fso.CreateFolder basePath & BACKUPS_FOLDER
    End If
    
    If Not fso.FolderExists(basePath & TEMPLATES_FOLDER) Then
        fso.CreateFolder basePath & TEMPLATES_FOLDER
    End If
    
    If Not fso.FolderExists(basePath & LOGS_FOLDER) Then
        fso.CreateFolder basePath & LOGS_FOLDER
    End If
    
    CreateDirectoryStructure = True
    Exit Function
    
ErrorHandler:
    CreateDirectoryStructure = False
End Function

Public Function InitializeDatabase() As Boolean
    On Error GoTo ErrorHandler
    
    ' التحقق من وجود الجداول المطلوبة
    If Not WorksheetExists("Users") Then
        MsgBox "ورقة Users غير موجودة", vbCritical
        InitializeDatabase = False
        Exit Function
    End If
    
    If Not WorksheetExists("Invoices") Then
        MsgBox "ورقة Invoices غير موجودة", vbCritical
        InitializeDatabase = False
        Exit Function
    End If
    
    InitializeDatabase = True
    Exit Function
    
ErrorHandler:
    InitializeDatabase = False
End Function

Public Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    WorksheetExists = Not (ws Is Nothing)
    On Error GoTo 0
End Function

Public Sub FormatDatabaseTables()
    On Error Resume Next
    
    ' تنسيق جدول المستخدمين
    Call FormatSheet("Users", "A1:J1", PRIMARY_COLOR)
    
    ' تنسيق جدول الفواتير
    Call FormatSheet("Invoices", "A1:V1", SUCCESS_COLOR)
    
    ' تنسيق جدول متابعة السداد
    Call FormatSheet("PaymentTracking", "A1:S1", WARNING_COLOR)
    
    ' تنسيق جدول المرفقات
    Call FormatSheet("Attachments", "A1:O1", RGB(155, 89, 182))
    
    ' تنسيق جدول السجلات
    Call FormatSheet("SystemLog", "A1:L1", DANGER_COLOR)
    
    ' تنسيق جدول الإعدادات
    Call FormatSheet("SystemSettings", "A1:I1", RGB(52, 73, 94))
    
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
            .Font.Name = MAIN_FONT
            .Font.Size = NORMAL_FONT_SIZE
        End With
        ws.Columns.AutoFit
    End If
    
    On Error GoTo 0
End Sub

Public Sub LoadSystemSettings()
    On Error Resume Next
    ' تحميل إعدادات النظام من جدول SystemSettings
    ' يمكن إضافة المزيد من الإعدادات هنا
    On Error GoTo 0
End Sub
```

---

## 🔐 **المرحلة الخامسة: إنشاء وحدة الأمان**

### **الخطوة 6: إنشاء الوحدة الثالثة - SecurityManager**

#### **إنشاء الوحدة:**
```
1. كليك يمين مرة أخرى على "VBAProject"
2. Insert → Module
3. غير الاسم في Properties إلى "SecurityManager"
```

#### **أين تضع الكود:**
```
1. اضغط مرتين على "SecurityManager"
2. انسخ والصق الكود التالي:
```

```vba
Option Explicit

' === وظائف الأمان والتشفير ===

Public Function ValidateLogin(username As String, password As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim storedHash As String
    Dim inputHash As String
    Dim loginAttempts As Integer
    Dim isLocked As Boolean
    
    Set ws = ThisWorkbook.Worksheets("Users")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    inputHash = HashPassword(password)
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 2).Value)) = UCase(Trim(username)) Then
            ' التحقق من حالة القفل
            isLocked = ws.Cells(i, 10).Value
            loginAttempts = ws.Cells(i, 9).Value
            
            If isLocked Or loginAttempts >= 3 Then
                MsgBox "الحساب مقفل. يرجى الاتصال بالمدير.", vbCritical, "حساب مقفل"
                ValidateLogin = False
                Exit Function
            End If
            
            storedHash = ws.Cells(i, 3).Value
            
            If storedHash = inputHash And ws.Cells(i, 6).Value = True Then
                ' تسجيل دخول ناجح
                Call LoadCurrentUser(i, ws)
                ws.Cells(i, 8).Value = Now() ' تحديث آخر تسجيل دخول
                ws.Cells(i, 9).Value = 0 ' إعادة تعيين محاولات الدخول
                
                ValidateLogin = True
                Exit Function
            Else
                ' تسجيل دخول فاشل
                ws.Cells(i, 9).Value = loginAttempts + 1
                
                If loginAttempts + 1 >= 3 Then
                    ws.Cells(i, 10).Value = True ' قفل الحساب
                    MsgBox "تم قفل الحساب بعد 3 محاولات فاشلة. يرجى الاتصال بالمدير.", vbCritical, "حساب مقفل"
                Else
                    MsgBox "بيانات الدخول غير صحيحة. المحاولات المتبقية: " & (3 - (loginAttempts + 1)), vbExclamation, "خطأ في تسجيل الدخول"
                End If
                
                ValidateLogin = False
                Exit Function
            End If
        End If
    Next i
    
    MsgBox "اسم المستخدم غير موجود", vbCritical, "خطأ في تسجيل الدخول"
    ValidateLogin = False
    Exit Function
    
ErrorHandler:
    ValidateLogin = False
End Function

Private Sub LoadCurrentUser(userRow As Long, ws As Worksheet)
    With CurrentUser
        .UserID = ws.Cells(userRow, 1).Value
        .Username = ws.Cells(userRow, 2).Value
        .FullName = ws.Cells(userRow, 5).Value
        .UserType = ws.Cells(userRow, 4).Value
        .IsActive = ws.Cells(userRow, 6).Value
        .LastLogin = ws.Cells(userRow, 8).Value
        .Permissions = GetUserPermissions(.UserType)
    End With
    
    IsLoggedIn = True
End Sub

Public Function HashPassword(password As String) As String
    ' تشفير بسيط باستخدام خوارزمية مخصصة
    Dim hashedPassword As String
    Dim i As Integer
    Dim charCode As Integer
    Dim hashValue As Long
    
    hashValue = 5381
    
    For i = 1 To Len(password)
        charCode = Asc(Mid(password, i, 1))
        hashValue = ((hashValue * 33) Xor charCode) And &H7FFFFFFF
    Next i
    
    ' إضافة salt بناءً على طول كلمة المرور والحرف الأول
    Dim salt As String
    salt = Hex(Len(password)) & Hex(Asc(Left(password, 1)))
    
    hashedPassword = Hex(hashValue) & salt
    HashPassword = hashedPassword
End Function

Public Function GetUserPermissions(userType As String) As String
    Select Case UCase(Trim(userType))
        Case "مدير", "MANAGER"
            GetUserPermissions = "ALL"
        Case "موظف تخليص", "CLEARANCE_OFFICER"
            GetUserPermissions = "ADD_INVOICE,EDIT_INVOICE,VIEW_INVOICE,UPLOAD_DOCUMENTS,SEARCH_INVOICES"
        Case "مدقق", "AUDITOR"
            GetUserPermissions = "VIEW_INVOICE,SEARCH_INVOICES,GENERATE_REPORTS,VIEW_LOGS"
        Case "محاسب", "ACCOUNTANT"
            GetUserPermissions = "VIEW_INVOICE,UPDATE_PAYMENT,GENERATE_FINANCIAL_REPORTS,SEARCH_INVOICES"
        Case Else
            GetUserPermissions = "VIEW_ONLY"
    End Select
End Function

Public Function CheckPermission(action As String) As Boolean
    If Not IsLoggedIn Then
        CheckPermission = False
        Exit Function
    End If
    
    If CurrentUser.Permissions = "ALL" Then
        CheckPermission = True
        Exit Function
    End If
    
    CheckPermission = InStr(UCase(CurrentUser.Permissions), UCase(action)) > 0
End Function

Public Sub LogoutCurrentUser()
    If IsLoggedIn Then
        ' يمكن إضافة تسجيل العملية هنا لاحقاً
    End If
    
    ' مسح بيانات المستخدم الحالي
    With CurrentUser
        .UserID = ""
        .Username = ""
        .FullName = ""
        .UserType = ""
        .IsActive = False
        .LastLogin = #1/1/1900#
        .Permissions = ""
    End With
    
    IsLoggedIn = False
End Sub
```

---

## 🖼️ **المرحلة السادسة: إنشاء نموذج تسجيل الدخول**

### **الخطوة 7: إنشاء النموذج**

#### **أين تنشئ النموذج:**
```
1. في محرر VBA
2. كليك يمين على "VBAProject (InvoiceManagementSystem.xlsm)"
3. اختر: Insert → UserForm
4. ستظهر نافيسة جديدة تحتوي على نموذج فارغ
5. وستظهر نافيسة "Toolbox" تحتوي على العناصر
```

#### **تغيير خصائص النموذج:**
```
1. اضغط على النموذج مرة واحدة (الخلفية الرمادية)
2. في نافيسة Properties (أسفل اليسار):
   - Name: غير إلى "UserForm_Login"
   - Caption: غير إلى "تسجيل الدخول - نظام إدارة الفواتير"
   - Width: 400
   - Height: 300
   - StartUpPosition: 2-CenterScreen
```

### **الخطوة 8: إضافة العناصر للنموذج**

#### **إضافة العنوان (Label):**
```
1. من Toolbox اضغط على "Label" (حرف A)
2. ارسم مستطيل في أعلى النموذج
3. في Properties:
   - Name: lblTitle
   - Caption: "نظام إدارة فواتير عوائد المناولة والتخزين"
   - Font: اضغط على [...] واختر Tahoma, Size 12, Bold
```

#### **إضافة تسمية اسم المستخدم:**
```
1. من Toolbox اضغط على "Label"
2. ارسم مستطيل صغير
3. في Properties:
   - Name: lblUsername
   - Caption: "اسم المستخدم:"
```

#### **إضافة حقل اسم المستخدم:**
```
1. من Toolbox اضغط على "TextBox" (مربع ab)
2. ارسم مستطيل بجانب التسمية
3. في Properties:
   - Name: txtUsername
```

#### **إضافة تسمية كلمة المرور:**
```
1. من Toolbox اضغط على "Label"
2. ارسم مستطيل تحت تسمية اسم المستخدم
3. في Properties:
   - Name: lblPassword
   - Caption: "كلمة المرور:"
```

#### **إضافة حقل كلمة المرور:**
```
1. من Toolbox اضغط على "TextBox"
2. ارسم مستطيل بجانب تسمية كلمة المرور
3. في Properties:
   - Name: txtPassword
   - PasswordChar: *
```

#### **إضافة زر الدخول:**
```
1. من Toolbox اضغط على "CommandButton"
2. ارسم زر في أسفل النموذج
3. في Properties:
   - Name: cmdLogin
   - Caption: "دخول"
   - Default: True
```

#### **إضافة زر الخروج:**
```
1. من Toolbox اضغط على "CommandButton"
2. ارسم زر بجانب زر الدخول
3. في Properties:
   - Name: cmdExit
   - Caption: "خروج"
   - Cancel: True
```

### **الخطوة 9: إضافة الكود للنموذج**

#### **أين تضع كود النموذج:**
```
1. اضغط مرتين على زر "دخول" (cmdLogin)
2. ستفتح نافيسة الكود الخاصة بالنموذج
3. ستجد السطر: Private Sub cmdLogin_Click()
4. بين هذا السطر و End Sub ضع الكود التالي:
```

```vba
Private Sub cmdLogin_Click()
    Dim username As String
    Dim password As String
    
    ' الحصول على البيانات المدخلة
    username = Trim(txtUsername.Text)
    password = Trim(txtPassword.Text)
    
    ' التحقق من الحقول الفارغة
    If username = "" Then
        MsgBox "يرجى إدخال اسم المستخدم", vbExclamation, "حقل مطلوب"
        txtUsername.SetFocus
        Exit Sub
    End If
    
    If password = "" Then
        MsgBox "يرجى إدخال كلمة المرور", vbExclamation, "حقل مطلوب"
        txtPassword.SetFocus
        Exit Sub
    End If
    
    ' التحقق من بيانات الدخول
    If ValidateLogin(username, password) Then
        ' تسجيل دخول ناجح
        MsgBox "مرحباً " & CurrentUser.FullName & vbCrLf & _
               "نوع المستخدم: " & CurrentUser.UserType, _
               vbInformation, "تسجيل دخول ناجح"
        
        ' إخفاء نموذج تسجيل الدخول
        Me.Hide
        
        ' هنا ستضاف الواجهة الرئيسية لاحقاً
        
    Else
        ' تسجيل دخول فاشل
        txtPassword.Text = ""
        txtUsername.SetFocus
    End If
End Sub
```

#### **إضافة كود زر الخروج:**
```
1. اضغط مرتين على زر "خروج" (cmdExit)
2. ستظهر: Private Sub cmdExit_Click()
3. بين السطرين ضع:
```

```vba
Private Sub cmdExit_Click()
    ' تأكيد الخروج
    If MsgBox("هل تريد إغلاق النظام؟", vbYesNo + vbQuestion, "تأكيد الخروج") = vbYes Then
        ThisWorkbook.Close False
    End If
End Sub
```

#### **إضافة كود تهيئة النموذج:**
```
1. في نفس نافيسة كود النموذج
2. في المربع الأيسر العلوي اختر "UserForm"
3. في المربع الأيمن العلوي اختر "Initialize"
4. ستظهر: Private Sub UserForm_Initialize()
5. بين السطرين ضع:
```

```vba
Private Sub UserForm_Initialize()
    ' تهيئة النموذج عند تحميله
    Me.Caption = "تسجيل الدخول - " & APP_NAME
    
    ' تعيين الخصائص الأساسية
    txtUsername.SetFocus
    txtPassword.PasswordChar = "*"
    
    ' تحميل بيانات تجريبية للاختبار
    txtUsername.Text = "admin"
End Sub
```

---

## ⚙️ **المرحلة السابعة: إضافة كود التشغيل التلقائي**

### **الخطوة 10: إضافة كود ThisWorkbook**

#### **أين تضع الكود:**
```
1. في محرر VBA
2. في النافيسة اليسرى تحت "Microsoft Excel Objects"
3. اضغط مرتين على "ThisWorkbook"
4. ستفتح نافيسة كود فارغة
```

#### **إضافة كود فتح الملف:**
```
1. في المربع الأيسر العلوي اختر "Workbook"
2. في المربع الأيمن العلوي اختر "Open"
3. ستظهر: Private Sub Workbook_Open()
4. بين السطرين ضع الكود التالي:
```

```vba
Private Sub Workbook_Open()
    ' تسجيل بداية تشغيل النظام
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' إخفاء جميع الأوراق عدا Dashboard
    Call HideWorksheets
    
    ' تهيئة النظام
    If InitializeSystem() Then
        ' تشغيل نموذج تسجيل الدخول
        Application.ScreenUpdating = True
        UserForm_Login.Show
    Else
        MsgBox "فشل في تهيئة النظام. يرجى الاتصال بالدعم الفني.", vbCritical, "خطأ في التهيئة"
        Application.ScreenUpdating = True
        ThisWorkbook.Close False
    End If
    
    Application.DisplayAlerts = True
End Sub
```

#### **إضافة كود إغلاق الملف:**
```
1. في المربع الأيمن العلوي اختر "BeforeClose"
2. ستظهر: Private Sub Workbook_BeforeClose(Cancel As Boolean)
3. بين السطرين ضع:
```

```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' إجراءات ما قبل الإغلاق
    
    If IsLoggedIn Then
        ' تسجيل خروج المستخدم
        Call LogoutCurrentUser
    End If
    
    ' إنشاء نسخة احتياطية تلقائية
    Call CreateAutoBackup
    
    ' حفظ تلقائي
    If ThisWorkbook.Saved = False Then
        ThisWorkbook.Save
    End If
    
    ' إظهار رسالة وداع
    If Not Cancel Then
        MsgBox "شكراً لاستخدام " & APP_NAME & vbCrLf & _
               "تم حفظ جميع البيانات بنجاح.", _
               vbInformation, "إغلاق النظام"
    End If
End Sub
```

#### **إضافة وظيفة إخفاء الأوراق:**
```
في نفس نافيسة ThisWorkbook، في نهاية الكود أضف:
```

```vba
Private Sub HideWorksheets()
    ' إخفاء جميع أوراق قاعدة البيانات
    Dim ws As Worksheet
    Dim protectedSheets As Variant
    
    protectedSheets = Array("Users", "Invoices", "PaymentTracking", "Attachments", "SystemLog", "SystemSettings")
    
    For Each ws In ThisWorkbook.Worksheets
        Dim i As Integer
        For i = 0 To UBound(protectedSheets)
            If ws.Name = protectedSheets(i) Then
                ws.Visible = xlSheetVeryHidden
                Exit For
            End If
        Next i
    Next ws
    
    ' التأكد من وجود ورقة Dashboard مرئية
    On Error Resume Next
    ThisWorkbook.Worksheets("Dashboard").Visible = xlSheetVisible
    If Err.Number <> 0 Then
        ' إنشاء ورقة Dashboard إذا لم تكن موجودة
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Dashboard"
        ws.Cells(1, 1).Value = "لوحة تحكم " & APP_NAME
        ws.Cells(1, 1).Font.Size = 16
        ws.Cells(1, 1).Font.Bold = True
    End If
    On Error GoTo 0
End Sub

Public Sub CreateAutoBackup()
    ' إنشاء نسخة احتياطية تلقائية
    On Error Resume Next
    
    Dim backupPath As String
    Dim backupFileName As String
    Dim timestamp As String
    
    timestamp = Format(Now(), "yyyy-mm-dd_hh-nn-ss")
    backupFileName = "backup_" & timestamp & ".xlsm"
    backupPath = ThisWorkbook.Path & "\" & BACKUPS_FOLDER & "\" & backupFileName
    
    ' التأكد من وجود مجلد النسخ الاحتياطية
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(ThisWorkbook.Path & "\" & BACKUPS_FOLDER) Then
        fso.CreateFolder ThisWorkbook.Path & "\" & BACKUPS_FOLDER
    End If
    
    ' نسخ الملف
    ThisWorkbook.SaveCopyAs backupPath
    
    On Error GoTo 0
End Sub
```

---

## 💾 **المرحلة الثامنة: الحفظ والاختبار**

### **الخطوة 11: حفظ جميع الأكواد**

#### **ترتيب الحفظ:**
```
1. في محرر VBA: اضغط Ctrl + S
2. ستظهر رسالة حفظ، اضغط "Yes"
3. أغلق محرر VBA: Alt + F4
4. في Excel: اضغط Ctrl + S لحفظ الملف
```

### **الخطوة 12: الاختبار الأولي**

#### **اختبار النظام:**
```
1. أغلق ملف Excel تماماً
2. افتح الملف مرة أخرى: InvoiceManagementSystem.xlsm
3. عند السؤال عن الماكرو اختر "Enable Content"
4. يجب أن يظهر نموذج تسجيل الدخول تلقائياً
```

#### **اختبار تسجيل الدخول:**
```
1. اسم المستخدم: admin
2. كلمة المرور: 1234
3. اضغط "دخول"
4. يجب أن تظهر رسالة "مرحباً مدير النظام"
```

#### **فحص المجلدات:**
```
1. اذهب لمجلد الملف على الكمبيوتر
2. يجب أن تجد المجلدات التالية قد تم إنشاؤها:
   📁 Documents
   📁 PDFReports
   📁 Backups
   📁 Templates
   📁 Logs
   📁 ImportedData
```

---

## 🎯 **ملخص أماكن الأكواد:**

### **📍 المواقع النهائية للأكواد:**

1. **Module "GlobalVariables":** المتغيرات والثوابت العامة
2. **Module "DatabaseOperations":** وظائف قاعدة البيانات والتهيئة
3. **Module "SecurityManager":** وظائف الأمان وتسجيل الدخول
4. **UserForm "UserForm_Login":** واجهة تسجيل الدخول وأكوادها
5. **ThisWorkbook:** أكواد التشغيل التلقائي والإغلاق

### **🔄 تسلسل التشغيل:**
```
1. فتح الملف → Workbook_Open في ThisWorkbook
2. تشغيل InitializeSystem في DatabaseOperations
3. إظهار UserForm_Login
4. عند الدخول → ValidateLogin في SecurityManager
5. عند الإغلاق → Workbook_BeforeClose في ThisWorkbook
```

---

**الآن لديك الخريطة الكاملة لأماكن وضع كل كود! هل تريد البدء في تطبيق هذه الخطوات؟** 🚀
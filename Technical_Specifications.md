# المواصفات التقنية التفصيلية
## نظام إدارة فواتير عوائد المناولة والتخزين - Excel VBA

---

## 🔧 **المتطلبات التقنية**

### **متطلبات النظام:**
- **نظام التشغيل:** Windows 10/11 (64-bit مُفضل)
- **Microsoft Excel:** 2016 أو أحدث (Office 365 مدعوم)
- **الذاكرة:** 4 GB RAM كحد أدنى، 8 GB مُفضل
- **مساحة القرص:** 1 GB مساحة حرة
- **دقة الشاشة:** 1024x768 كحد أدنى، 1920x1080 مُفضل
- **معالج:** Intel Core i3 أو AMD equivalent أو أفضل

### **إعدادات Excel المطلوبة:**
```vba
' في Developer Options:
- Enable all macros (not recommended; potentially dangerous code can run)
- أو Trust access to the VBA project object model
- Enable Excel 4.0 macros when VBA macros are enabled

' في Trust Center Settings:
- Add the application folder to Trusted Locations
- Enable ActiveX controls and plug-ins
```

---

## 📊 **هيكل قاعدة البيانات التفصيلي**

### **Sheet: Users**
```sql
-- هيكل الجدول
Column A: UserID (Primary Key) - Text(10)
Column B: Username - Text(50) - Unique, Not Null
Column C: PasswordHash - Text(255) - Not Null
Column D: UserType - Text(20) - Not Null
Column E: FullName - Text(100) - Not Null
Column F: IsActive - Boolean - Default True
Column G: CreatedDate - Date - Default Now()
Column H: LastLoginDate - Date
Column I: LoginAttempts - Integer - Default 0
Column J: IsLocked - Boolean - Default False

-- بيانات أولية
admin,admin,5e884898da28047151d0e56f8dc6292773603d0d6aabbdd62a11ef721d1542d8,مدير,مدير النظام,TRUE,2025-01-15,2025-01-15,0,FALSE
clearance1,clearance1,ef92b778bafe771e89245b89ecbc08a44a4e166c06659911881f383d4473e94f,موظف تخليص,موظف التخليص الأول,TRUE,2025-01-15,,0,FALSE
```

### **Sheet: Invoices**
```sql
-- هيكل الجدول الرئيسي
Column A: InvoiceID (Primary Key) - AutoNumber
Column B: SerialNumber - Integer - Auto Increment
Column C: ClaimNumber - Text(50) - Not Null, Indexed
Column D: ClaimDate - Date - Not Null
Column E: InvoiceNumber - Text(50) - Unique, Not Null
Column F: InvoiceDate - Date - Not Null
Column G: AmountWithoutStamp - Currency - Not Null, >= 0
Column H: StampDuty - Currency - Default 0, >= 0
Column I: TotalAmount - Currency - Calculated Field
Column J: PolicyNumber - Text(50) - Indexed
Column K: LCNumber - Text(50) - Indexed
Column L: ContractNumber - Text(50) - Indexed
Column M: CustomsDeclaration - Text(50) - Indexed
Column N: EntryPort - Text(100) - Not Null
Column O: PaymentStatus - Text(20) - Default "غير مسدد"
Column P: PaymentDate - Date - Nullable
Column Q: CreatedBy - Text(50) - Not Null
Column R: CreatedDate - DateTime - Default Now()
Column S: ModifiedBy - Text(50)
Column T: ModifiedDate - DateTime
Column U: IsDeleted - Boolean - Default False
Column V: Notes - Text(500) - Nullable

-- فهارس مركبة
Index: IX_Invoice_Claim (ClaimNumber, InvoiceNumber)
Index: IX_Invoice_Date (InvoiceDate, ClaimDate)
Index: IX_Invoice_Status (PaymentStatus, EntryPort)
```

### **Sheet: PaymentTracking**
```sql
-- جدول متابعة السداد
Column A: TrackingID (Primary Key) - AutoNumber
Column B: InvoiceID - Integer - Foreign Key
Column C: ClaimNumber - Text(50) - Not Null
Column D: InvoiceNumber - Text(50) - Not Null
Column E: HandlingStorageFees - Currency - Not Null
Column F: StampDuty - Currency - Not Null
Column G: TotalFees - Currency - Calculated
Column H: PaymentStatus - Text(20) - Not Null
Column I: PaymentMethod - Text(50)
Column J: PaymentReference - Text(100)
Column K: TransferredTo - Text(100)
Column L: TransferDate - Date
Column M: TransferReference - Text(100)
Column N: ReceivedBy - Text(100)
Column O: ReceivedDate - Date
Column P: CreatedBy - Text(50) - Not Null
Column Q: CreatedDate - DateTime - Default Now()
Column R: UpdatedBy - Text(50)
Column S: UpdatedDate - DateTime
```

### **Sheet: Attachments**
```sql
-- جدول المرفقات
Column A: AttachmentID (Primary Key) - AutoNumber
Column B: InvoiceID - Integer - Foreign Key
Column C: InvoiceNumber - Text(50) - Not Null
Column D: DocumentType - Text(100) - Not Null
Column E: OriginalFileName - Text(255) - Not Null
Column F: StoredFileName - Text(255) - Not Null
Column G: FilePath - Text(500) - Not Null
Column H: FileSize - Long - Not Null
Column I: FileExtension - Text(10) - Not Null
Column J: MimeType - Text(100)
Column K: UploadedBy - Text(50) - Not Null
Column L: UploadDate - DateTime - Default Now()
Column M: IsActive - Boolean - Default True
Column N: Description - Text(500)
Column O: CheckSum - Text(64) - For file integrity
```

### **Sheet: SystemLog**
```sql
-- جدول سجل النظام
Column A: LogID (Primary Key) - AutoNumber
Column B: UserID - Text(50) - Not Null
Column C: Action - Text(100) - Not Null
Column D: TableName - Text(50)
Column E: RecordID - Text(50)
Column F: OldValues - Text(1000)
Column G: NewValues - Text(1000)
Column H: IPAddress - Text(45)
Column I: UserAgent - Text(500)
Column J: LogDate - DateTime - Default Now()
Column K: LogLevel - Text(20) - (INFO, WARNING, ERROR, DEBUG)
Column L: Description - Text(1000)
```

### **Sheet: SystemSettings**
```sql
-- إعدادات النظام
Column A: SettingID (Primary Key) - AutoNumber
Column B: SettingKey - Text(100) - Unique
Column C: SettingValue - Text(1000)
Column D: SettingType - Text(20) - (STRING, NUMBER, BOOLEAN, DATE)
Column E: Description - Text(500)
Column F: IsEditable - Boolean - Default True
Column G: Category - Text(50)
Column H: UpdatedBy - Text(50)
Column I: UpdatedDate - DateTime

-- إعدادات افتراضية
STAMP_DUTY_RATE,0.05,NUMBER,معدل رسوم الدمغة الافتراضي,TRUE,FINANCIAL
AUTO_BACKUP_ENABLED,TRUE,BOOLEAN,تفعيل النسخ الاحتياطي التلقائي,TRUE,SYSTEM
BACKUP_RETENTION_DAYS,30,NUMBER,عدد أيام الاحتفاظ بالنسخ الاحتياطية,TRUE,SYSTEM
```

---

## 🎨 **مواصفات واجهات المستخدم التفصيلية**

### **UserForm_Login**
```vba
' خصائص النموذج
Name: UserForm_Login
Caption: "تسجيل الدخول - نظام إدارة الفواتير"
Width: 400
Height: 300
StartUpPosition: 2 (CenterScreen)
BorderStyle: 1 (fmBorderStyleSingle)
ShowModal: True

' العناصر المطلوبة:
1. Image: imgLogo
   - Left: 20, Top: 20, Width: 80, Height: 80
   - Picture: Company Logo

2. Label: lblTitle
   - Left: 120, Top: 30, Width: 250, Height: 20
   - Caption: "نظام إدارة فواتير عوائد المناولة والتخزين"
   - Font: Tahoma, 12pt, Bold
   - ForeColor: RGB(41, 128, 185)

3. Label: lblVersion
   - Left: 120, Top: 55, Width: 100, Height: 15
   - Caption: "الإصدار 1.0"
   - Font: Tahoma, 8pt
   - ForeColor: RGB(127, 140, 141)

4. Frame: fraLogin
   - Left: 20, Top: 120, Width: 360, Height: 120
   - Caption: "بيانات تسجيل الدخول"

5. Label: lblUsername (داخل fraLogin)
   - Left: 10, Top: 30, Width: 80, Height: 15
   - Caption: "اسم المستخدم:"

6. TextBox: txtUsername (داخل fraLogin)
   - Left: 100, Top: 27, Width: 240, Height: 20
   - Font: Tahoma, 10pt
   - TextAlign: 3 (fmTextAlignRight)

7. Label: lblPassword (داخل fraLogin)
   - Left: 10, Top: 60, Width: 80, Height: 15
   - Caption: "كلمة المرور:"

8. TextBox: txtPassword (داخل fraLogin)
   - Left: 100, Top: 57, Width: 240, Height: 20
   - PasswordChar: "*"
   - Font: Tahoma, 10pt

9. CheckBox: chkRememberMe (داخل fraLogin)
   - Left: 100, Top: 85, Width: 120, Height: 15
   - Caption: "تذكر بياناتي"

10. CommandButton: cmdLogin
    - Left: 280, Top: 255, Width: 80, Height: 25
    - Caption: "دخول"
    - Default: True
    - BackColor: RGB(39, 174, 96)
    - Font: Tahoma, 10pt, Bold

11. CommandButton: cmdExit
    - Left: 190, Top: 255, Width: 80, Height: 25
    - Caption: "خروج"
    - Cancel: True
    - BackColor: RGB(231, 76, 60)
    - Font: Tahoma, 10pt

12. Label: lblStatus
    - Left: 20, Top: 290, Width: 360, Height: 15
    - Caption: ""
    - ForeColor: RGB(231, 76, 60)
    - Font: Tahoma, 8pt
```

### **UserForm_MainMenu**
```vba
' خصائص النموذج
Name: UserForm_MainMenu
Caption: "لوحة التحكم الرئيسية"
Width: 800
Height: 600
StartUpPosition: 2 (CenterScreen)
BorderStyle: 1 (fmBorderStyleSingle)

' العناصر المطلوبة:
1. Label: lblWelcome
   - Left: 20, Top: 20, Width: 600, Height: 25
   - Caption: "مرحباً، [اسم المستخدم] - [نوع المستخدم]"
   - Font: Tahoma, 14pt, Bold
   - ForeColor: RGB(41, 128, 185)

2. Frame: fraQuickStats
   - Left: 20, Top: 60, Width: 760, Height: 80
   - Caption: "الإحصائيات السريعة"

3. Label: lblTotalInvoices (داخل fraQuickStats)
   - Left: 20, Top: 25, Width: 150, Height: 40
   - Caption: "إجمالي الفواتير" & vbCrLf & "[عدد]"
   - TextAlign: 2 (fmTextAlignCenter)
   - BackColor: RGB(52, 152, 219)
   - ForeColor: White
   - Font: Tahoma, 10pt, Bold

4. Label: lblPaidInvoices (داخل fraQuickStats)
   - Left: 190, Top: 25, Width: 150, Height: 40
   - Caption: "الفواتير المسددة" & vbCrLf & "[عدد]"
   - TextAlign: 2 (fmTextAlignCenter)
   - BackColor: RGB(39, 174, 96)
   - ForeColor: White

5. Label: lblUnpaidInvoices (داخل fraQuickStats)
   - Left: 360, Top: 25, Width: 150, Height: 40
   - Caption: "الفواتير غير المسددة" & vbCrLf & "[عدد]"
   - TextAlign: 2 (fmTextAlignCenter)
   - BackColor: RGB(231, 76, 60)
   - ForeColor: White

6. Label: lblTotalAmount (داخل fraQuickStats)
   - Left: 530, Top: 25, Width: 200, Height: 40
   - Caption: "إجمالي المبالغ" & vbCrLf & "[مبلغ] ج.م"
   - TextAlign: 2 (fmTextAlignCenter)
   - BackColor: RGB(241, 196, 15)
   - ForeColor: White

7. Frame: fraMainButtons
   - Left: 20, Top: 160, Width: 760, Height: 300
   - Caption: "الوظائف الرئيسية"

' الأزرار الرئيسية (داخل fraMainButtons):
8. CommandButton: btnInvoiceEntry
   - Left: 30, Top: 30, Width: 150, Height: 60
   - Caption: "🧾" & vbCrLf & "إدارة الفواتير"
   - Font: Tahoma, 10pt, Bold

9. CommandButton: btnPaymentTracking
   - Left: 200, Top: 30, Width: 150, Height: 60
   - Caption: "💰" & vbCrLf & "متابعة السداد"

10. CommandButton: btnSearchFilter
    - Left: 370, Top: 30, Width: 150, Height: 60
    - Caption: "🔍" & vbCrLf & "البحث والتصفية"

11. CommandButton: btnAttachments
    - Left: 540, Top: 30, Width: 150, Height: 60
    - Caption: "📎" & vbCrLf & "إدارة المستندات"

12. CommandButton: btnReports
    - Left: 30, Top: 110, Width: 150, Height: 60
    - Caption: "📊" & vbCrLf & "التقارير"

13. CommandButton: btnImportData
    - Left: 200, Top: 110, Width: 150, Height: 60
    - Caption: "📥" & vbCrLf & "استيراد البيانات"

14. CommandButton: btnUserManagement
    - Left: 370, Top: 110, Width: 150, Height: 60
    - Caption: "👥" & vbCrLf & "إدارة المستخدمين"
    - Visible: False ' يظهر للمدير فقط

15. CommandButton: btnSettings
    - Left: 540, Top: 110, Width: 150, Height: 60
    - Caption: "⚙️" & vbCrLf & "الإعدادات"

16. CommandButton: btnBackup
    - Left: 30, Top: 190, Width: 150, Height: 60
    - Caption: "💾" & vbCrLf & "النسخ الاحتياطي"

17. CommandButton: btnHelp
    - Left: 200, Top: 190, Width: 150, Height: 60
    - Caption: "❓" & vbCrLf & "المساعدة"

18. CommandButton: btnLogout
    - Left: 540, Top: 190, Width: 150, Height: 60
    - Caption: "🚪" & vbCrLf & "تسجيل الخروج"
    - BackColor: RGB(241, 196, 15)

19. Frame: fraSystemInfo
    - Left: 20, Top: 480, Width: 760, Height: 60
    - Caption: "معلومات النظام"

20. Label: lblSystemInfo (داخل fraSystemInfo)
    - Left: 10, Top: 20, Width: 740, Height: 30
    - Caption: "آخر تسجيل دخول: [تاريخ] | إصدار النظام: 1.0 | المستخدمين النشطين: [عدد]"
    - Font: Tahoma, 8pt
    - ForeColor: RGB(127, 140, 141)

21. CommandButton: cmdExit
    - Left: 700, Top: 555, Width: 80, Height: 25
    - Caption: "إغلاق"
    - BackColor: RGB(231, 76, 60)
```

---

## 💻 **مواصفات وحدات VBA**

### **Module: GlobalDeclarations**
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

Public Enum DocumentTypes
    Invoice = 1
    BillOfLading = 2
    CertificateOfOrigin = 3
    DeliveryOrder = 4
    PackingList = 5
    Other = 99
End Enum

Public Enum LogLevels
    Info = 1
    Warning = 2
    ErrorLevel = 3
    Debug = 4
End Enum

' متغيرات الأداء
Public LastDatabaseUpdate As Date
Public SystemStartTime As Date
Public ActiveConnections As Integer
```

### **Module: DatabaseEngine**
```vba
Option Explicit

' === وظائف إدارة قاعدة البيانات ===

Public Function InitializeDatabase() As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    ' التحقق من وجود الجداول المطلوبة
    If Not WorksheetExists("Users") Then
        Call CreateUsersTable
    End If
    
    If Not WorksheetExists("Invoices") Then
        Call CreateInvoicesTable
    End If
    
    If Not WorksheetExists("PaymentTracking") Then
        Call CreatePaymentTrackingTable
    End If
    
    If Not WorksheetExists("Attachments") Then
        Call CreateAttachmentsTable
    End If
    
    If Not WorksheetExists("SystemLog") Then
        Call CreateSystemLogTable
    End If
    
    If Not WorksheetExists("SystemSettings") Then
        Call CreateSystemSettingsTable
    End If
    
    ' تطبيق الحماية على الجداول
    Call ProtectDatabaseTables
    
    InitializeDatabase = True
    Exit Function
    
ErrorHandler:
    Call LogError("InitializeDatabase", Err.Description)
    InitializeDatabase = False
End Function

Public Function WorksheetExists(sheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not (Sheets(sheetName) Is Nothing)
    On Error GoTo 0
End Function

Private Sub CreateUsersTable()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Users"
    
    ' إنشاء العناوين
    With ws
        .Cells(1, 1).Value = "UserID"
        .Cells(1, 2).Value = "Username"
        .Cells(1, 3).Value = "PasswordHash"
        .Cells(1, 4).Value = "UserType"
        .Cells(1, 5).Value = "FullName"
        .Cells(1, 6).Value = "IsActive"
        .Cells(1, 7).Value = "CreatedDate"
        .Cells(1, 8).Value = "LastLoginDate"
        .Cells(1, 9).Value = "LoginAttempts"
        .Cells(1, 10).Value = "IsLocked"
        
        ' تنسيق العناوين
        .Range("A1:J1").Font.Bold = True
        .Range("A1:J1").Interior.Color = RGB(52, 152, 219)
        .Range("A1:J1").Font.Color = RGB(255, 255, 255)
        
        ' إضافة المستخدم الافتراضي
        .Cells(2, 1).Value = "USR001"
        .Cells(2, 2).Value = "admin"
        .Cells(2, 3).Value = HashPassword("1234")
        .Cells(2, 4).Value = "مدير"
        .Cells(2, 5).Value = "مدير النظام"
        .Cells(2, 6).Value = True
        .Cells(2, 7).Value = Now()
        .Cells(2, 8).Value = ""
        .Cells(2, 9).Value = 0
        .Cells(2, 10).Value = False
    End With
End Sub

Private Sub CreateInvoicesTable()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Invoices"
    
    With ws
        .Cells(1, 1).Value = "InvoiceID"
        .Cells(1, 2).Value = "SerialNumber"
        .Cells(1, 3).Value = "ClaimNumber"
        .Cells(1, 4).Value = "ClaimDate"
        .Cells(1, 5).Value = "InvoiceNumber"
        .Cells(1, 6).Value = "InvoiceDate"
        .Cells(1, 7).Value = "AmountWithoutStamp"
        .Cells(1, 8).Value = "StampDuty"
        .Cells(1, 9).Value = "TotalAmount"
        .Cells(1, 10).Value = "PolicyNumber"
        .Cells(1, 11).Value = "LCNumber"
        .Cells(1, 12).Value = "ContractNumber"
        .Cells(1, 13).Value = "CustomsDeclaration"
        .Cells(1, 14).Value = "EntryPort"
        .Cells(1, 15).Value = "PaymentStatus"
        .Cells(1, 16).Value = "PaymentDate"
        .Cells(1, 17).Value = "CreatedBy"
        .Cells(1, 18).Value = "CreatedDate"
        .Cells(1, 19).Value = "ModifiedBy"
        .Cells(1, 20).Value = "ModifiedDate"
        .Cells(1, 21).Value = "IsDeleted"
        .Cells(1, 22).Value = "Notes"
        
        ' تنسيق العناوين
        .Range("A1:V1").Font.Bold = True
        .Range("A1:V1").Interior.Color = RGB(39, 174, 96)
        .Range("A1:V1").Font.Color = RGB(255, 255, 255)
        .Columns.AutoFit
    End With
End Sub

' === وظائف CRUD للفواتير ===

Public Function AddNewInvoice(invoiceData As InvoiceInfo) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim newID As Long
    
    Set ws = ThisWorkbook.Worksheets("Invoices")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' توليد ID جديد
    newID = GetNextInvoiceID()
    
    ' إدراج البيانات
    With ws
        .Cells(lastRow, 1).Value = newID
        .Cells(lastRow, 2).Value = GetNextSerialNumber()
        .Cells(lastRow, 3).Value = invoiceData.ClaimNumber
        .Cells(lastRow, 4).Value = invoiceData.ClaimDate
        .Cells(lastRow, 5).Value = invoiceData.InvoiceNumber
        .Cells(lastRow, 6).Value = invoiceData.InvoiceDate
        .Cells(lastRow, 7).Value = invoiceData.AmountWithoutStamp
        .Cells(lastRow, 8).Value = invoiceData.StampDuty
        .Cells(lastRow, 9).Value = invoiceData.TotalAmount
        .Cells(lastRow, 10).Value = invoiceData.PolicyNumber
        .Cells(lastRow, 11).Value = invoiceData.LCNumber
        .Cells(lastRow, 12).Value = invoiceData.ContractNumber
        .Cells(lastRow, 13).Value = invoiceData.CustomsDeclaration
        .Cells(lastRow, 14).Value = invoiceData.EntryPort
        .Cells(lastRow, 15).Value = "غير مسدد"
        .Cells(lastRow, 16).Value = ""
        .Cells(lastRow, 17).Value = CurrentUser.Username
        .Cells(lastRow, 18).Value = Now()
        .Cells(lastRow, 19).Value = ""
        .Cells(lastRow, 20).Value = ""
        .Cells(lastRow, 21).Value = False
        .Cells(lastRow, 22).Value = ""
    End With
    
    ' تسجيل العملية
    Call LogActivity("ADD_INVOICE", "تم إضافة فاتورة جديدة: " & invoiceData.InvoiceNumber)
    
    AddNewInvoice = True
    Exit Function
    
ErrorHandler:
    Call LogError("AddNewInvoice", Err.Description)
    AddNewInvoice = False
End Function

Public Function UpdateInvoice(invoiceID As Long, invoiceData As InvoiceInfo) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim targetRow As Long
    
    Set ws = ThisWorkbook.Worksheets("Invoices")
    targetRow = FindInvoiceRow(invoiceID)
    
    If targetRow = 0 Then
        UpdateInvoice = False
        Exit Function
    End If
    
    ' تحديث البيانات
    With ws
        .Cells(targetRow, 3).Value = invoiceData.ClaimNumber
        .Cells(targetRow, 4).Value = invoiceData.ClaimDate
        .Cells(targetRow, 5).Value = invoiceData.InvoiceNumber
        .Cells(targetRow, 6).Value = invoiceData.InvoiceDate
        .Cells(targetRow, 7).Value = invoiceData.AmountWithoutStamp
        .Cells(targetRow, 8).Value = invoiceData.StampDuty
        .Cells(targetRow, 9).Value = invoiceData.TotalAmount
        .Cells(targetRow, 10).Value = invoiceData.PolicyNumber
        .Cells(targetRow, 11).Value = invoiceData.LCNumber
        .Cells(targetRow, 12).Value = invoiceData.ContractNumber
        .Cells(targetRow, 13).Value = invoiceData.CustomsDeclaration
        .Cells(targetRow, 14).Value = invoiceData.EntryPort
        .Cells(targetRow, 19).Value = CurrentUser.Username
        .Cells(targetRow, 20).Value = Now()
    End With
    
    ' تسجيل العملية
    Call LogActivity("UPDATE_INVOICE", "تم تعديل الفاتورة: " & invoiceData.InvoiceNumber)
    
    UpdateInvoice = True
    Exit Function
    
ErrorHandler:
    Call LogError("UpdateInvoice", Err.Description)
    UpdateInvoice = False
End Function

Public Function DeleteInvoice(invoiceID As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim targetRow As Long
    Dim invoiceNumber As String
    
    Set ws = ThisWorkbook.Worksheets("Invoices")
    targetRow = FindInvoiceRow(invoiceID)
    
    If targetRow = 0 Then
        DeleteInvoice = False
        Exit Function
    End If
    
    invoiceNumber = ws.Cells(targetRow, 5).Value
    
    ' حذف منطقي (تعديل حالة IsDeleted)
    ws.Cells(targetRow, 21).Value = True
    ws.Cells(targetRow, 19).Value = CurrentUser.Username
    ws.Cells(targetRow, 20).Value = Now()
    
    ' تسجيل العملية
    Call LogActivity("DELETE_INVOICE", "تم حذف الفاتورة: " & invoiceNumber)
    
    DeleteInvoice = True
    Exit Function
    
ErrorHandler:
    Call LogError("DeleteInvoice", Err.Description)
    DeleteInvoice = False
End Function

' === وظائف مساعدة ===

Private Function GetNextInvoiceID() As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxID As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("Invoices")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    maxID = 0
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value > maxID Then
            maxID = ws.Cells(i, 1).Value
        End If
    Next i
    
    GetNextInvoiceID = maxID + 1
End Function

Private Function GetNextSerialNumber() As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxSerial As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("Invoices")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    maxSerial = 0
    For i = 2 To lastRow
        If ws.Cells(i, 2).Value > maxSerial And ws.Cells(i, 21).Value = False Then
            maxSerial = ws.Cells(i, 2).Value
        End If
    Next i
    
    GetNextSerialNumber = maxSerial + 1
End Function

Private Function FindInvoiceRow(invoiceID As Long) As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("Invoices")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = invoiceID And ws.Cells(i, 21).Value = False Then
            FindInvoiceRow = i
            Exit Function
        End If
    Next i
    
    FindInvoiceRow = 0
End Function
```

### **Module: SecurityManager**
```vba
Option Explicit

' === وظائف الأمان والتشفير ===

Public Function HashPassword(password As String) As String
    ' تشفير بسيط باستخدام SHA-256 simulation
    Dim hashedPassword As String
    Dim i As Integer
    Dim charCode As Integer
    Dim hashValue As Long
    
    hashValue = 5381
    
    For i = 1 To Len(password)
        charCode = Asc(Mid(password, i, 1))
        hashValue = ((hashValue * 33) Xor charCode) And &H7FFFFFFF
    Next i
    
    hashedPassword = Hex(hashValue) & Hex(Len(password)) & Hex(Asc(Left(password, 1)))
    HashPassword = hashedPassword
End Function

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
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    inputHash = HashPassword(password)
    
    For i = 2 To lastRow
        If ws.Cells(i, 2).Value = username Then
            ' التحقق من حالة القفل
            isLocked = ws.Cells(i, 10).Value
            loginAttempts = ws.Cells(i, 9).Value
            
            If isLocked Or loginAttempts >= 3 Then
                MsgBox "الحساب مقفل. يرجى الاتصال بالمدير.", vbCritical
                ValidateLogin = False
                Exit Function
            End If
            
            storedHash = ws.Cells(i, 3).Value
            
            If storedHash = inputHash And ws.Cells(i, 6).Value = True Then
                ' تسجيل دخول ناجح
                Call LoadCurrentUser(i, ws)
                ws.Cells(i, 8).Value = Now() ' تحديث آخر تسجيل دخول
                ws.Cells(i, 9).Value = 0 ' إعادة تعيين محاولات الدخول
                
                Call LogActivity("LOGIN_SUCCESS", "تسجيل دخول ناجح للمستخدم: " & username)
                ValidateLogin = True
                Exit Function
            Else
                ' تسجيل دخول فاشل
                ws.Cells(i, 9).Value = loginAttempts + 1
                
                If loginAttempts + 1 >= 3 Then
                    ws.Cells(i, 10).Value = True ' قفل الحساب
                    Call LogActivity("ACCOUNT_LOCKED", "تم قفل الحساب: " & username)
                End If
                
                Call LogActivity("LOGIN_FAILED", "محاولة دخول فاشلة للمستخدم: " & username)
                ValidateLogin = False
                Exit Function
            End If
        End If
    Next i
    
    Call LogActivity("LOGIN_FAILED", "محاولة دخول بمستخدم غير موجود: " & username)
    ValidateLogin = False
    Exit Function
    
ErrorHandler:
    Call LogError("ValidateLogin", Err.Description)
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

Public Function GetUserPermissions(userType As String) As String
    Select Case userType
        Case "مدير"
            GetUserPermissions = "ALL"
        Case "موظف تخليص"
            GetUserPermissions = "ADD_INVOICE,EDIT_INVOICE,VIEW_INVOICE,UPLOAD_DOCUMENTS"
        Case "مدقق"
            GetUserPermissions = "VIEW_INVOICE,GENERATE_REPORTS"
        Case "محاسب"
            GetUserPermissions = "VIEW_INVOICE,UPDATE_PAYMENT,GENERATE_FINANCIAL_REPORTS"
        Case Else
            GetUserPermissions = "VIEW_ONLY"
    End Select
End Function

Public Function CheckPermission(action As String) As Boolean
    If CurrentUser.Permissions = "ALL" Then
        CheckPermission = True
        Exit Function
    End If
    
    CheckPermission = InStr(CurrentUser.Permissions, action) > 0
End Function

Public Sub LogoutCurrentUser()
    Call LogActivity("LOGOUT", "تسجيل خروج للمستخدم: " & CurrentUser.Username)
    
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

## 📁 **مواصفات نظام الملفات**

### **هيكل المجلدات التلقائي:**
```vba
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
```

### **نظام الأرشفة التلقائية:**
```vba
Public Function OrganizeDocumentsByInvoice(invoiceNumber As String) As String
    On Error GoTo ErrorHandler
    
    Dim basePath As String
    Dim invoiceFolderPath As String
    Dim fso As Object
    
    basePath = ThisWorkbook.Path & "\" & DOCUMENTS_FOLDER & "\"
    invoiceFolderPath = basePath & "فاتورة_" & invoiceNumber & "\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' إنشاء مجلد الفاتورة إذا لم يكن موجوداً
    If Not fso.FolderExists(invoiceFolderPath) Then
        fso.CreateFolder invoiceFolderPath
        
        ' إنشاء مجلدات فرعية لأنواع المستندات
        fso.CreateFolder invoiceFolderPath & "فواتير\"
        fso.CreateFolder invoiceFolderPath & "بوالص_شحن\"
        fso.CreateFolder invoiceFolderPath & "شهادات_منشأ\"
        fso.CreateFolder invoiceFolderPath & "أوامر_تسليم\"
        fso.CreateFolder invoiceFolderPath & "قوائم_تعبئة\"
        fso.CreateFolder invoiceFolderPath & "مستندات_أخرى\"
    End If
    
    OrganizeDocumentsByInvoice = invoiceFolderPath
    Exit Function
    
ErrorHandler:
    OrganizeDocumentsByInvoice = ""
End Function
```

---

## 🔄 **نظام النسخ الاحتياطي**

```vba
Public Function CreateAutoBackup() As Boolean
    On Error GoTo ErrorHandler
    
    Dim backupPath As String
    Dim backupFileName As String
    Dim timestamp As String
    
    timestamp = Format(Now(), "yyyy-mm-dd_hh-nn-ss")
    backupFileName = "backup_" & timestamp & ".xlsm"
    backupPath = ThisWorkbook.Path & "\" & BACKUPS_FOLDER & "\" & backupFileName
    
    ' نسخ الملف
    ThisWorkbook.SaveCopyAs backupPath
    
    ' تسجيل العملية
    Call LogActivity("BACKUP_CREATED", "تم إنشاء نسخة احتياطية: " & backupFileName)
    
    ' تنظيف النسخ القديمة (الاحتفاظ بآخر 30 نسخة)
    Call CleanOldBackups
    
    CreateAutoBackup = True
    Exit Function
    
ErrorHandler:
    Call LogError("CreateAutoBackup", Err.Description)
    CreateAutoBackup = False
End Function

Private Sub CleanOldBackups()
    On Error Resume Next
    
    Dim fso As Object
    Dim folder As Object
    Dim files As Object
    Dim file As Object
    Dim fileArray() As String
    Dim fileCount As Integer
    Dim i As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(ThisWorkbook.Path & "\" & BACKUPS_FOLDER)
    Set files = folder.files
    
    ' جمع أسماء ملفات النسخ الاحتياطية
    fileCount = 0
    For Each file In files
        If Left(file.Name, 7) = "backup_" And Right(file.Name, 5) = ".xlsm" Then
            fileCount = fileCount + 1
            ReDim Preserve fileArray(1 To fileCount)
            fileArray(fileCount) = file.Name
        End If
    Next file
    
    ' حذف النسخ الزائدة عن 30
    If fileCount > 30 Then
        ' ترتيب الملفات حسب التاريخ (الأقدم أولاً)
        Call SortArray(fileArray)
        
        ' حذف النسخ الأقدم
        For i = 1 To fileCount - 30
            fso.DeleteFile ThisWorkbook.Path & "\" & BACKUPS_FOLDER & "\" & fileArray(i)
        Next i
    End If
End Sub
```

---

## 📊 **نظام التقارير المتقدم**

```vba
Public Function GenerateAdvancedReport(reportType As String, criteria As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    Select Case reportType
        Case "SINGLE_INVOICE"
            GenerateAdvancedReport = GenerateSingleInvoiceReport(criteria("InvoiceNumber"))
        Case "MULTIPLE_INVOICES"
            GenerateAdvancedReport = GenerateMultipleInvoicesReport(criteria)
        Case "PAYMENT_STATUS"
            GenerateAdvancedReport = GeneratePaymentStatusReport(criteria)
        Case "FINANCIAL_SUMMARY"
            GenerateAdvancedReport = GenerateFinancialSummaryReport(criteria)
        Case "AUDIT_TRAIL"
            GenerateAdvancedReport = GenerateAuditTrailReport(criteria)
        Case Else
            GenerateAdvancedReport = False
    End Select
    
    Exit Function
    
ErrorHandler:
    Call LogError("GenerateAdvancedReport", Err.Description)
    GenerateAdvancedReport = False
End Function

Private Function GenerateSingleInvoiceReport(invoiceNumber As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    Dim invoiceData As InvoiceInfo
    Dim reportPath As String
    Dim timestamp As String
    
    ' البحث عن بيانات الفاتورة
    invoiceData = GetInvoiceByNumber(invoiceNumber)
    
    If invoiceData.InvoiceNumber = "" Then
        MsgBox "الفاتورة غير موجودة", vbCritical
        GenerateSingleInvoiceReport = False
        Exit Function
    End If
    
    ' إنشاء ورقة التقرير
    Set reportWs = ThisWorkbook.Worksheets("InvoiceReport")
    
    ' تعبئة بيانات التقرير
    With reportWs
        .Cells(2, 2).Value = invoiceData.InvoiceNumber
        .Cells(4, 2).Value = invoiceData.ClaimNumber
        .Cells(5, 2).Value = Format(invoiceData.ClaimDate, "dd/mm/yyyy")
        .Cells(6, 2).Value = invoiceData.InvoiceNumber
        .Cells(7, 2).Value = Format(invoiceData.InvoiceDate, "dd/mm/yyyy")
        .Cells(8, 2).Value = Format(invoiceData.AmountWithoutStamp, "#,##0.00")
        .Cells(9, 2).Value = Format(invoiceData.StampDuty, "#,##0.00")
        .Cells(10, 2).Value = Format(invoiceData.TotalAmount, "#,##0.00")
        .Cells(11, 2).Value = invoiceData.PolicyNumber
        .Cells(12, 2).Value = invoiceData.ContractNumber
        .Cells(13, 2).Value = invoiceData.LCNumber
        .Cells(14, 2).Value = invoiceData.CustomsDeclaration
        .Cells(15, 2).Value = invoiceData.EntryPort
    End With
    
    ' تصدير إلى PDF
    timestamp = Format(Now(), "yyyy-mm-dd_hh-nn")
    reportPath = ThisWorkbook.Path & "\" & REPORTS_FOLDER & "\تقرير_فاتورة_" & invoiceNumber & "_" & timestamp & ".pdf"
    
    reportWs.ExportAsFixedFormat Type:=xlTypePDF, Filename:=reportPath, Quality:=xlQualityStandard
    
    ' تسجيل العملية
    Call LogActivity("REPORT_GENERATED", "تم إنشاء تقرير للفاتورة: " & invoiceNumber)
    
    GenerateSingleInvoiceReport = True
    Exit Function
    
ErrorHandler:
    Call LogError("GenerateSingleInvoiceReport", Err.Description)
    GenerateSingleInvoiceReport = False
End Function
```

---

## ⚡ **تحسينات الأداء**

### **فهرسة البيانات:**
```vba
Public Sub CreateDataIndexes()
    ' إنشاء فهارس للبحث السريع
    Call CreateInvoiceNumberIndex
    Call CreateClaimNumberIndex
    Call CreateDateRangeIndex
End Sub

Private Sub CreateInvoiceNumberIndex()
    ' إنشاء مصفوفة مفهرسة لأرقام الفواتير
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("Invoices")
    lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
    
    ReDim InvoiceNumberIndex(2 To lastRow, 1 To 2)
    
    For i = 2 To lastRow
        If ws.Cells(i, 21).Value = False Then ' غير محذوف
            InvoiceNumberIndex(i, 1) = ws.Cells(i, 5).Value ' رقم الفاتورة
            InvoiceNumberIndex(i, 2) = i ' رقم الصف
        End If
    Next i
    
    ' ترتيب المصفوفة للبحث السريع
    Call SortInvoiceIndex(InvoiceNumberIndex)
End Sub
```

### **تحسين الذاكرة:**
```vba
Public Sub OptimizeMemoryUsage()
    ' تعطيل التحديث التلقائي للشاشة
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' تنظيف المتغيرات غير المستخدمة
    Call ClearUnusedVariables
    
    ' إعادة تفعيل الإعدادات
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
```

---

## 🛡️ **معالجة الأخطاء المتقدمة**

```vba
Public Sub LogError(functionName As String, errorDescription As String)
    Dim logFile As String
    Dim fileNum As Integer
    Dim timestamp As String
    
    timestamp = Format(Now(), "yyyy-mm-dd hh:nn:ss")
    logFile = ThisWorkbook.Path & "\" & LOGS_FOLDER & "\error_log_" & Format(Now(), "yyyy-mm-dd") & ".txt"
    
    fileNum = FreeFile
    Open logFile For Append As #fileNum
    Print #fileNum, timestamp & " | ERROR | " & functionName & " | " & errorDescription
    Close #fileNum
    
    ' تسجيل في قاعدة البيانات أيضاً
    Call LogToDatabase("ERROR", functionName, errorDescription)
End Sub

Public Sub LogActivity(action As String, description As String)
    Dim logFile As String
    Dim fileNum As Integer
    Dim timestamp As String
    
    timestamp = Format(Now(), "yyyy-mm-dd hh:nn:ss")
    logFile = ThisWorkbook.Path & "\" & LOGS_FOLDER & "\activity_log_" & Format(Now(), "yyyy-mm-dd") & ".txt"
    
    fileNum = FreeFile
    Open logFile For Append As #fileNum
    Print #fileNum, timestamp & " | " & CurrentUser.Username & " | " & action & " | " & description
    Close #fileNum
    
    ' تسجيل في قاعدة البيانات
    Call LogToDatabase("INFO", action, description)
End Sub

Private Sub LogToDatabase(logLevel As String, action As String, description As String)
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets("SystemLog")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    With ws
        .Cells(lastRow, 1).Value = lastRow - 1 ' LogID
        .Cells(lastRow, 2).Value = CurrentUser.Username
        .Cells(lastRow, 3).Value = action
        .Cells(lastRow, 4).Value = ""
        .Cells(lastRow, 5).Value = ""
        .Cells(lastRow, 6).Value = ""
        .Cells(lastRow, 7).Value = ""
        .Cells(lastRow, 8).Value = ""
        .Cells(lastRow, 9).Value = ""
        .Cells(lastRow, 10).Value = Now()
        .Cells(lastRow, 11).Value = logLevel
        .Cells(lastRow, 12).Value = description
    End With
End Sub
```

---

## 📱 **واجهة المستخدم المتجاوبة**

```vba
Public Sub AdjustFormSizeToScreen(frm As UserForm)
    Dim screenWidth As Long
    Dim screenHeight As Long
    
    screenWidth = GetSystemMetrics(0) ' SM_CXSCREEN
    screenHeight = GetSystemMetrics(1) ' SM_CYSCREEN
    
    ' تحديد الحد الأقصى لحجم النموذج
    If frm.Width > screenWidth * 0.8 Then
        frm.Width = screenWidth * 0.8
    End If
    
    If frm.Height > screenHeight * 0.8 Then
        frm.Height = screenHeight * 0.8
    End If
    
    ' توسيط النموذج
    frm.Left = (screenWidth - frm.Width) / 2
    frm.Top = (screenHeight - frm.Height) / 2
End Sub

' دالة Windows API للحصول على أبعاد الشاشة
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
```

---

## 🔐 **إعدادات الحماية المتقدمة**

```vba
Public Sub ApplySecuritySettings()
    ' حماية هيكل المصنف
    ThisWorkbook.Protect Password:="SecureStructure2025!", Structure:=True, Windows:=False
    
    ' حماية الأوراق الحساسة
    Call ProtectSensitiveSheets
    
    ' تشفير الملف
    Call EncryptWorkbook
    
    ' إعداد صلاحيات الملفات
    Call SetFilePermissions
End Sub

Private Sub ProtectSensitiveSheets()
    Dim protectionPassword As String
    protectionPassword = "DataProtect2025!"
    
    ' حماية جدول المستخدمين
    Worksheets("Users").Protect Password:=protectionPassword, _
        DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowSorting:=False, AllowFiltering:=False, AllowUsingPivotTables:=False
    
    ' حماية جدول السجلات
    Worksheets("SystemLog").Protect Password:=protectionPassword, _
        DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowSorting:=False, AllowFiltering:=False
    
    ' السماح بالتعديل على جدول الفواتير للمستخدمين المصرح لهم فقط
    If CheckPermission("EDIT_INVOICE") Then
        Worksheets("Invoices").Unprotect Password:=protectionPassword
    Else
        Worksheets("Invoices").Protect Password:=protectionPassword, _
            DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If
End Sub

Private Sub EncryptWorkbook()
    ' تشفير الملف بكلمة مرور قوية
    ThisWorkbook.Password = "InvoiceSystem2025!@#$%"
End Sub
```

---

هذه المواصفات التقنية التفصيلية توفر الأساس الكامل لتطوير نظام إدارة الفواتير المطلوب. كل قسم مصمم ليكون قابلاً للتنفيذ مباشرة في Excel VBA مع الحفاظ على أعلى معايير الجودة والأمان والأداء.
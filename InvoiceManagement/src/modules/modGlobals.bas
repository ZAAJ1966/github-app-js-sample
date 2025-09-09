Attribute VB_Name = "modGlobals"
Option Explicit

'=====================
'  Global Variables
'=====================
Public gCurrentUser As String
Public gCurrentRole As String

'Folder paths – adjust if workbook moved.
Public Const DOCS_FOLDER As String = "Documents"
Public Const PDF_FOLDER As String = "PDFReports"

'=====================
' Ensure required folders exist at runtime
'=====================
Public Sub EnsureFolders()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim basePath As String
    basePath = ThisWorkbook.Path & Application.PathSeparator
    If Not fso.FolderExists(basePath & DOCS_FOLDER) Then fso.CreateFolder basePath & DOCS_FOLDER
    If Not fso.FolderExists(basePath & PDF_FOLDER) Then fso.CreateFolder basePath & PDF_FOLDER
End Sub
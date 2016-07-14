Attribute VB_Name = "Export_Routines"
Option Explicit
'Remember to add a reference to Microsoft Visual Basic for Applications Extensibility
'Exports all VBA project components containing code to a folder in the same directory as this spreadsheet.
Public Sub ExportAllComponents()
    Dim VBComp As VBIDE.VBComponent
    Dim destDir As String, fName As String, ext As String
    Dim destPath As String
    'Create the directory where code will be created.
    'Alternatively, you could change this so that the user is prompted
    Dim unhide As Boolean
    On Error Resume Next
    'workbooks.Open ("C:\Users\almendez\AppData\Roaming\Microsoft\Excel\XLSTART\Personal.xlsb")
    'unhide = Windows("Aldo Personal.XLSB").Visible
    
    Debug.Print ActiveWorkbook.Path
    If ActiveWorkbook.Path = "" Then
        Debug.Print ActiveWorkbook.Path
        'MsgBox "You must first save this workbook somewhere so that it has a path.", , "Error"
        'Exit Sub
    End If
    Debug.Print ActiveWorkbook.Path
    
    'Workbooks.Open("C:\Users\Ryan\AppData\Roaming\Microsoft\Excel\XLSTART\Personal.xlsb")
    
    destPath = "C:\apps\oee_master_report"
    
    destDir = destPath & "\" & ActiveWorkbook.Name & " Modules"
    If Dir(destDir, vbDirectory) = vbNullString Then MkDir destDir
    
    'Export all non-blank components to the directory
    For Each VBComp In ActiveWorkbook.VBProject.VBComponents
        If VBComp.CodeModule.CountOfLines > 0 Then
            'Determine the standard extention of the exported file.
            'These can be anything, but for re-importing, should be the following:
            Select Case VBComp.Type
                Case vbext_ct_ClassModule: ext = ".cls"
                Case vbext_ct_Document: ext = ".cls"
                Case vbext_ct_StdModule: ext = ".bas"
                Case vbext_ct_MSForm: ext = ".frm"
                Case Else: ext = vbNullString
            End Select
            If ext <> vbNullString Then
                fName = destDir & "\" & VBComp.Name & ext
                'Overwrite the existing file
                'Alternatively, you can prompt the user before killing the file.
                If Dir(fName, vbNormal) <> vbNullString Then Kill (fName)
                VBComp.export (fName)
            End If
        End If
    Next VBComp
End Sub

Sub Copy_PMWB_To_Mark()
Dim unhide As Boolean
On Error Resume Next
unhide = Windows("PERSONAL.XLSB").Visible
Windows("PERSONAL.XLSB").Close
On Error GoTo 0
On Error Resume Next
    Kill "Mark's Drive\Macro Backups\Mark's Personal Macro Workbook\*.*"
    On Error GoTo 0
 
 
    FileCopy "C:\Users\almendez\AppData\Roaming\Microsoft\Excel\XLSTART\Personal.xlsb", "C:\apps\oee_master_report\Aldo Personal.xlsb"
    workbooks.Open ("C:\apps\oee_master_report\Aldo Personal.xlsb")
 Windows("Aldo Personal.XLSB").Visible = True

 
End Sub

Sub workbooks()
    Dim wkb As Workbook
    For Each wkb In Application.workbooks
        Debug.Print wkb.Name
        wkb.Activate
        Call ExportAllComponents
    Next wkb
End Sub


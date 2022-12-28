Attribute VB_Name = "DEMO"
Option Explicit

Public Sub DEMO()
    Const actionButton As Long = -1
    Dim filePath As String
    '
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Please select an Excel file"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel files", "*.xlsx;*.xlsm"
        If .Show = actionButton Then filePath = .SelectedItems(1)
    End With
    If LenB(filePath) = 0 Then Exit Sub
    '
    Dim zip As New ExcelZIP: zip.Init filePath
    Dim b() As Byte
    Dim s As String
    '
    Debug.Print "Archive has " & zip.Count & " files"
    '
    zip.ReadData "xl/workbook.xml", b
    s = StrConv(b, vbUnicode)
    '
    Dim startIndex As Long
    Dim endIndex As Long
    '
    'The preffered way would be to parse the XML but the below should suffice as demo
    Do
        startIndex = InStr(endIndex + 1, s, "<sheet name=", vbTextCompare)
        endIndex = InStr(startIndex + 1, s, " sheetId=", vbTextCompare)
        '
        If startIndex > 0 Then
            Debug.Print "Sheet found: " & Mid$(s, startIndex + 13, endIndex - startIndex - 14)
        End If
    Loop Until startIndex = 0
End Sub

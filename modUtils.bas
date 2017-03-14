Option Explicit

' Contains general functions and subroutines


' Return the path to the selected folder.
Public Function GetFolderPath() As String
    Dim fldr As FileDialog
    Dim selectedFolder As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        selectedFolder = .SelectedItems(1)
    End With
NextCode:
    GetFolderPath = selectedFolder
    Set fldr = Nothing
End Function
' Loop through all the sheets in the given excel file and return true if the given sheet name
'  is present in this excel file. Otherwise return false.
Public Function SheetExists(sheetName As String, excelFile As Workbook) As Boolean
    Dim wkSh As Worksheet
    For Each wkSh In excelFile.Worksheets
        If wkSh.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next wkSh
    SheetExists = False
End Function
'Adds metadata to a sheet that it names "FileMetaData" and it then hides and protects the sheet using
'  the provided password. Can add more meta data here as needed.
Public Function AddMetaDataToNewExcelFile(excelFile As Workbook, metaDataText As String, pwdProtect As String, assayType As String, experimentName As String) As Boolean
    Dim wkSh As Worksheet
    If SheetExists("Sheet1", excelFile) Then
        Set wkSh = excelFile.Worksheets("Sheet1")
    Else
        Set wkSh = excelFile.Worksheets.Add
        wkSh.Name = "Sheet1"
    End If
    wkSh.Name = "FileMetaData"
    wkSh.Range("A1").Value = "Code Version"
    wkSh.Range("B1").Value = modCodeInfo.CODE_VERSION
    wkSh.Range("A2").Value = "Code Date"
    wkSh.Range("B2").Value = modCodeInfo.CODE_DATE
    wkSh.Range("A3").Value = "File Description"
    wkSh.Range("B3").Value = metaDataText
    wkSh.Range("A4").Value = "Selected Assay Type"
    wkSh.Range("B4").Value = assayType
    wkSh.Range("A5").Value = "Experiment Name"
    wkSh.Range("B5").Value = experimentName
    wkSh.Range("A1:B5").Name = "FileMetaData"
    wkSh.Protect pwdProtect
    wkSh.Visible = xlSheetHidden
End Function
' Used by the button on sheet "RunCode".
Public Sub ShowForm()
    uiAddWellIds.Show
End Sub

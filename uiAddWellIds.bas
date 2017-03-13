Option Explicit

Private plMap As PlateMap

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnRun_Click()
    RunAddWellIds
End Sub


Private Sub btnSaveToNewWorkbook_Click()
    Dim folderPath As String
    Dim excelSaveToFilename As String
    Dim fullOutputPath As String
    Dim newExcelFile As Workbook
    Dim sheetToCopy As Worksheet
    Set sheetToCopy = ThisWorkbook.Worksheets("DataSheet")
    folderPath = modUtils.GetFolderPath
    excelSaveToFilename = Me.txtOutputFilename
    fullOutputPath = folderPath & "\" & excelSaveToFilename
    Set newExcelFile = Application.Workbooks.Add
    sheetToCopy.Copy after:=newExcelFile.Sheets(newExcelFile.Sheets.Count)
    Set sheetToCopy = ThisWorkbook.Worksheets("PlateMaps")
    sheetToCopy.Copy after:=newExcelFile.Sheets(newExcelFile.Sheets.Count)
    modUtils.AddMetaDataToNewExcelFile newExcelFile, Me.txtExcelFileMetaData, Me.txtPasswordToProtect
    newExcelFile.SaveAs fullOutputPath, 51
    newExcelFile.Close True
End Sub

Private Sub UserForm_Click()
    
End Sub

Private Sub UserForm_Initialize()
    Set plMap = New PlateMap
End Sub

Private Sub RunAddWellIds()
    Dim plMapRng As Range
    Dim embeddedWellIdRng As Range
    Set plMapRng = Range(Me.refeditPlateMapRange.Value)
    MsgBox plMapRng.Address
    Set embeddedWellIdRng = Range(Me.refeditEmbeddedWellIdRange)
    MsgBox embeddedWellIdRng.Address
    plMap.SetPlateMapRng plMapRng
    plMap.SetWellIdRng embeddedWellIdRng
    plMap.AddWellIds
    MsgBox "Done!"
End Sub


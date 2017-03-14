Option Explicit

Private plMap As PlateMap
Private arrAssayTypes As Variant 'Used to populate the listbox
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnRun_Click()
    RunAddWellIds
End Sub

'Copies "DataSheet" and "PlateMaps" sheets to a new workbook, writes in metadata to a hidden sheet, protects the metadata sheet,
' saves the new Excel file and clears out the pasted data from this workbook.
Private Sub btnSaveToNewWorkbook_Click()
    Dim folderPath As String
    Dim excelSaveToFilename As String
    Dim fullOutputPath As String
    Dim newExcelFile As Workbook
    Dim sheetToCopy As Worksheet
    Dim selectedAssayType As String
    Set sheetToCopy = ThisWorkbook.Worksheets("DataSheet")
    folderPath = modUtils.GetFolderPath
    excelSaveToFilename = Me.txtOutputFilename
    fullOutputPath = folderPath & "\" & excelSaveToFilename
    Set newExcelFile = Application.Workbooks.Add
    sheetToCopy.Copy after:=newExcelFile.Sheets(newExcelFile.Sheets.Count)
    Set sheetToCopy = ThisWorkbook.Worksheets("PlateMaps")
    sheetToCopy.Copy after:=newExcelFile.Sheets(newExcelFile.Sheets.Count)
    selectedAssayType = GetSelectedAssayType()
    modUtils.AddMetaDataToNewExcelFile newExcelFile, Me.txtExcelFileMetaData, Me.txtPasswordToProtect, selectedAssayType, Me.txtExperimentName
    newExcelFile.SaveAs fullOutputPath, 51
    newExcelFile.Close True
    ThisWorkbook.Worksheets("DataSheet").Cells.Clear
    ThisWorkbook.Worksheets("PlateMaps").Cells.Clear
    MsgBox "Generated Excel file saved as: " & fullOutputPath
End Sub

Private Sub lblExperimentName_Click()

End Sub

Private Sub UserForm_Click()

End Sub
Private Sub LoadAssayTypesLbox()
    Dim i As Integer
    For i = 0 To UBound(arrAssayTypes)
        Me.lboxAssayTypes.AddItem arrAssayTypes(i)
    Next
End Sub
Private Function GetSelectedAssayType() As String
    Dim i As Integer
    For i = 0 To Me.lboxAssayTypes.ListCount - 1
        If Me.lboxAssayTypes.Selected(i) Then
            GetSelectedAssayType = Me.lboxAssayTypes.List(i)
            Exit Function
        End If
    Next i
    GetSelectedAssayType = ""
End Function
Private Sub UserForm_Initialize()
    Set plMap = New PlateMap
    Dim lastRowNumber As Long
    arrAssayTypes = Array("% CD4+CD25+", _
                           "MFI CD4+CD25+", _
                           "% CD8+CD25+", _
                           "MFI CD8+CD25+", _
                           "% CD4+CD137+", _
                           "MFI CD4+CD137+", _
                           "% CD8+CD137+", _
                           "MFI CD8+CD137+", _
                           "% CD4+proliferation+", _
                           "MFI CD4+proliferation+", _
                           "% CD8+proliferation+", _
                           "MFI CD8+proliferation+", _
                           "% CD4+proliferation+", _
                           "MFI CD4+proliferation+", _
                           "% CD8+proliferation+", _
                           "MFI CD8+proliferation+")
    LoadAssayTypesLbox
    lastRowNumber = Worksheets("DataSheet").UsedRange.Rows.Count
    'Set default for the RefEdit
    Me.refeditEmbeddedWellIdRange.Value = "DataSheet!$A$1:$A$" & CStr(lastRowNumber)
End Sub
' Uses the PlateMap instance to match the sample names in the wells to extracted well IDs.
Private Sub RunAddWellIds()
    Dim plMapRng As Range
    Dim embeddedWellIdRng As Range
    Set plMapRng = Range(Me.refeditPlateMapRange.Value)
    'MsgBox plMapRng.Address
    Set embeddedWellIdRng = Range(Me.refeditEmbeddedWellIdRange)
    'MsgBox embeddedWellIdRng.Address
    plMap.SetPlateMapRng plMapRng
    plMap.SetWellIdRng embeddedWellIdRng
    plMap.AddWellIds
    MsgBox "Sample IDs added to DataSheet!"
End Sub

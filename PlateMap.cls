Option Explicit
Private Const PLATE_RNG_NAME As String = "CurrentPlateMap" 'A name for the 96 well plate map containing the well names
Private Const WELL_ID_RINDEX As Integer = 2 'Index from right of Well ID (e.g. "A1") in array profuced by splitting on "_".
Private Const WELL_ID_RNG_NAME As String = "WellIdRng" ' Section of column containing embedded well IDs, e.g. "TSK01_vitro_024_day3 donor2_A1_A01_001.fcs"
'' Class to take a plate map range, create a dictionary from it mapping well IDs to its value, extract embedded well IDs from a column and
'' use these as keys to extract and add the plate map values to sheet.

'' TEST CODE
'' Michael Maguire - 2017-03-10
'' Test version!
' Added Microsoft Scripting Runtime for dictionaries and Microsoft VBScript Regular Expressions 5.5

'''''Public Sub test_Setup()
'''''    Dim rng As Range
'''''    Dim plMap As PlateMap: Set plMap = New PlateMap
'''''    Set rng = Selection
'''''    Debug.Assert plMap.Setup(rng)
'''''    Set plMap = Nothing
'''''End Sub
'''''
'''''Public Sub test_GetWellIdValueMap()
'''''    Dim rng As Range:    Set rng = Selection
'''''    Dim plMap As PlateMap: Set plMap = New PlateMap
'''''    Dim wellIdValueMap As Dictionary
'''''    Dim key As Variant
'''''    plMap.Setup rng
'''''    Set wellIdValueMap = plMap.GetWellIdValueMap()
'''''    For Each key In wellIdValueMap.Keys
'''''        Debug.Print key, wellIdValueMap(key)
'''''    Next key
'''''    Set plMap = Nothing
'''''End Sub
'''''
'''''Public Function test_AddWellIds()
'''''    Dim rng As Range: Set rng = Range("WellIdRng")
'''''    Dim plMap As PlateMap: Set plMap = New PlateMap
'''''    plMap.AddWellIds
'''''    Set plMap = Nothing
'''''End Function


Sub class_initialize()
End Sub
' Ensure that range containing the plate map is the correct dimension.
Private Function CheckPlateMapRng(rng As Range) As Boolean
    If rng.Columns.Count = 12 And rng.Rows.Count = 8 Then
        CheckPlateMapRng = True
        Exit Function
    End If
    CheckPlateMapRng = False
End Function
'Define the plate map range and name it.
Public Function SetPlateMapRng(rng As Range) As Boolean
    Dim rowCount As Integer
    Dim colCount As Integer
    If Not CheckPlateMapRng(rng) Then
        rowCount = rng.Rows.Count
        colCount = rng.Columns.Count
        Err.Raise 1000, "Setup", "Expected dimensions for plate map range are 12 columns and 8 rows. Given row count: " & rowCount & vbCr & _
        ", Given column count: " & colCount
        Exit Function
    End If
    rng.Name = PLATE_RNG_NAME
    Setup = True
End Function
'Define and name the column section containing the embedded well IDs.
Public Function SetWellIdRng(rng As Range) As Boolean
    If rng.Columns.Count > 1 Then
        Err.Raise 1000, "SetWellIdRng", "This range can only contain a single column!"
    End If
    rng.Name = WELL_ID_RNG_NAME
    SetWellIdRng = True
End Function
'Taking the range named given by constant PLATE_RNG_NAME, return a dictionary mapping
' the well IDs to the actual values.
Public Function GetWellIdValueMap() As Dictionary
    Dim wellIdValueMap As Dictionary: Set wellIdValueMap = New Dictionary
    Dim arrWellRowLetters As Variant: arrWellRowLetters = Array("A", "B", "C", "D", "E", "F", "G", "H")
    Dim i As Integer
    Dim j As Integer
    Dim wellId As String
    Dim cellValueForWellId As String
    For i = 0 To UBound(arrWellRowLetters)
        For j = 1 To 12
            wellId = arrWellRowLetters(i) & CStr(j)
            cellValueForWellId = Range(PLATE_RNG_NAME).Cells(i + 1, j).Value
            wellIdValueMap.Add wellId, cellValueForWellId
        Next j
    Next i
    Set GetWellIdValueMap = wellIdValueMap
End Function
' Return the constant used to name the input plate map range.
Public Function GetRngNameForPlateMap() As String
    GetRngNameForPlateMap = PLATE_RNG_NAME
End Function
'Extract the embedded well IDs and use them as keys to get the plate map value from
' the dictionary returned by "GetWellIdValueMap()" and write them to the first empty rightmost column.
' Need to make this more robust! Case where input column is last in used range will result in an error due to .End(xlToRight).Offset call!
Public Function AddWellIds() As Boolean
    Dim wellIdValueMap As Dictionary: Set wellIdValueMap = GetWellIdValueMap()
    Dim cell As Range
    Dim cellValsSplit() As String
    Dim wellId As String
    Dim wellIdIdx As Integer
    Dim wellVal As String

    For Each cell In Range(WELL_ID_RNG_NAME).Cells
        cellValsSplit = Split(cell.Value, "_")
        wellIdIdx = UBound(cellValsSplit) - WELL_ID_RINDEX
        wellId = cellValsSplit(wellIdIdx)
        wellVal = wellIdValueMap(wellId)
        cell.End(xlToRight).Offset(0, 1).Value = wellVal
    Next cell
    AddWellIds = True
End Function

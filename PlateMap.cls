Option Explicit
Private Const PLATE_RNG_NAME As String = "CurrentPlateMap" 'A name for the 96 well plate map containing the well names
Private Const WELL_ID_RNG_NAME As String = "WellIdRng" ' Section of column containing embedded well IDs, e.g. "TSK01_vitro_024_day3 donor2_A1_A01_001.fcs"
Private Const WELL_ID_OUTPUT_RNG_NAME = "WellIdOutputRng"
Private arrWellIds As Variant ' Array values set in constructor.

'' Class to take a plate map range, create a dictionary from it mapping well IDs to its value, extract embedded well IDs from a column and
'' use these as keys to extract and add the plate map values to sheet.

'' TEST CODE
'' Michael Maguire - 2017-03-10
'' Test version!
' Added Microsoft Scripting Runtime for dictionaries and Microsoft VBScript Regular Expressions 5.5


Sub class_initialize()
    arrWellIds = Array("A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12", _
                       "B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12", _
                       "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", _
                       "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", _
                       "E1", "E2", "E3", "E4", "E5", "E6", "E7", "E8", "E9", "E10", "E11", "E12", _
                       "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12", _
                       "G1", "G2", "G3", "G4", "G5", "G6", "G7", "G8", "G9", "G10", "G11", "G12", _
                       "H1", "H2", "H3", "H4", "H5", "H6", "H7", "H8", "H9", "H10", "H11", "H12")
End Sub
' Ensure that range containing the plate map is the correct dimension.
Private Function CheckPlateMapRng(rng As Range) As Boolean
    If rng.Columns.Count = 12 And rng.Rows.Count = 8 Then
        CheckPlateMapRng = True
        Exit Function
    End If
    CheckPlateMapRng = False
End Function
' Given a worksheet object, assign a range name to the first blank column on the RHS.
' Used to set a range target for the well values taken from the plate map.
Private Function GiveNameToFirstBlankColumn(dataSheet As Worksheet) As String
    Dim usedRng As Range: Set usedRng = dataSheet.UsedRange
    Dim usedRngColCount As Integer: usedRngColCount = usedRng.Columns.Count
    Dim firstBlankColumnNumber As Integer: firstBlankColumnNumber = usedRngColCount + 1
    Dim usedRngRowCount As Long: usedRngRowCount = usedRng.Rows.Count
    Range(usedRng.Cells(1, usedRngColCount).Offset(0, 1), usedRng.Cells(1, usedRngColCount).Offset(usedRngRowCount - 1, 1)).Name = WELL_ID_OUTPUT_RNG_NAME
    GiveNameToFirstBlankColumn = WELL_ID_OUTPUT_RNG_NAME
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
    SetPlateMapRng = True
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
' Well IDs are embedded in larger cell values. This function splits the input string on a given delimiter (default underscore)
'  and the uses a nested "for" loop to determine if any of the substring elements are present in the array of well IDs.
'If a well ID is found, then return it and exit the function. Note that the use of "Ucase" mmakes the code case-insensitive so
' that, for example, "A1" and "a1" are treated the same but the uppercase version is returned.
' If no well ID candidate substring is found, an empty string is returned.
Private Function GetWellIdFromString(stringWithEmbeddedWellId As String, Optional delimiter As String = "_") As String
    Dim substringElements() As String
    Dim i As Integer
    Dim j As Integer
    substringElements = Split(stringWithEmbeddedWellId, delimiter)
    For i = 0 To UBound(substringElements)
        For j = 0 To UBound(arrWellIds)
            If UCase(substringElements(i)) = arrWellIds(j) Then
                GetWellIdFromString = UCase(substringElements(i))
                Exit Function
            End If
        Next j
    Next i
    GetWellIdFromString = ""
End Function
'Extract the embedded well IDs and use them as keys to get the plate map value from
' the dictionary returned by "GetWellIdValueMap()" and write them to the first empty rightmost column.
Public Function AddWellIds() As Boolean
    Dim wellIdValueMap As Dictionary: Set wellIdValueMap = GetWellIdValueMap()
    Dim cell As Range
    Dim cellValsSplit() As String
    Dim wellId As String
    Dim wellIdIdx As Integer
    Dim wellVal As String
    Dim thisSheet As Worksheet: Set thisSheet = Range(WELL_ID_RNG_NAME).Worksheet
    GiveNameToFirstBlankColumn thisSheet ' Set a range name for the output to follow.
    For Each cell In Range(WELL_ID_RNG_NAME).Cells
        wellId = GetWellIdFromString(cell.Value)
        wellVal = wellIdValueMap(wellId)
        Range(WELL_ID_OUTPUT_RNG_NAME).Cells(cell.Row, 1).Value = wellVal
    Next cell
    AddWellIds = True
End Function

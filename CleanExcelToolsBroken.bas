Sub ConvertText2Num()
  For Each xCell In Selection
   
    If IsNumeric(xCell.Value) Then
      xCell.Value = WorksheetFunction.Clean(xCell.Value)
      xCell.Value = WorksheetFunction.Trim(xCell.Value)
      xCell.Value = xCell.Value
    End If
  
  Next xCell
End Sub

Sub FillBlankCellWithValue()
  'Fill an empty or blank cell in selection with value specified in InputBox
  Dim cell As Range
  Dim InputValue As String
  On Error Resume Next

  'Prompt for value to fill empty cells in selection
  InputValue = InputBox("Enter value that will fill empty cells in selection", _
  "Fill Empty Cells")

  'Test for empty cell. If empty, fill cell with value given
  For Each cell In Selection
    If IsEmpty(cell) Then
      cell.Value = InputValue
    End If
  Next
End Sub

Sub RemoveCarriageReturns()
    Dim MyRange As Range
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For Each MyRange In ActiveSheet.UsedRange
        If 0 < InStr(MyRange, Chr(10)) Then
            MyRange = Replace(MyRange, Chr(10), "")
        End If
    Next

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub StripCharacters()
  'Fill an empty or blank cell in selection with value specified in InputBox
  Dim cell As Range
  Dim InputValue As String
  On Error Resume Next

  'Prompt for value to fill empty cells in selection
  InputValue = InputBox("Enter the value to strip out")

  'Test for empty cell. If empty, fill cell with value given
  For Each cell In Selection
    cell.Value = Replace(cell.Value, InputValue, "")
  Next

End Sub

Sub CleanSelection()
For Each xCell In Selection
  xCell = WorksheetFunction.Clean(xCell)
Next
End Sub

Function ReplaceClean1(sText As String, Optional sSubText As String = " ")
    Dim J As Integer
    Dim vAddText

    vAddText = Array(Chr(129), Chr(141), Chr(143), Chr(144), Chr(157))
    
    For J = 1 To 31
        sText = Replace(sText, Chr(J), sSubText)
    Next
    
    For J = 0 To UBound(vAddText)
        sText = Replace(sText, vAddText(J), sSubText)
    Next
    
    ReplaceClean = sText
End Function

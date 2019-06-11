
Sub ConvertText2Num()
  For Each xCell In Selection
  
  If IsNumeric(xCell.Value) Then
    xCell.Value = Val(xCell.Value)
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
    If Len(cell) = 1 Then
      If Asc(cell.Value) > 0 And Asc(cell.Value) < 33 Then
        cell.Value = ""
      End If
    End If
    
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

  'Prompt for value remove
  InputValue = InputBox("Enter the value to strip out (it will be contigious)")

  'Replace the given value with nulls
  For Each cell In Selection
    cell.Value = Replace(cell.Value, InputValue, "")
  Next

End Sub

Sub CleanCrap()
  Dim i As Integer
  
  ' Go through each cell in the selection
  For Each xCell In Selection
  
    ' Strip out all bogus characters 1 through 31
    For i = 1 To 31
      '' Make the bogus values null
      xCell.Value = Replace(xCell.Value, Chr(i), "")
    Next
    
    ' Take out one other stupid value
    xCell.Value = Replace(xCell.Value, Chr(127), "")
  Next
End Sub

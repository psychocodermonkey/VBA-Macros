Attribute VB_Name = "FillBlankCellWithValue"

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


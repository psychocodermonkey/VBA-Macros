Attribute VB_Name = "ConvertText2Num"

' Convert fields with numbers in them that are formatted as text to numeric
Sub ConvertText2Num()
  For Each xCell In Selection

  If IsNumeric(xCell.Value) Then
    xCell.Value = xCell.Value
  End If

  Next xCell
End Sub


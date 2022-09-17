VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Data")

If Target.Row > 7 And Target.Row <= 12 And Target.column = 2 Then
Cancel = True
sh.Range("Q1") = Target.Value
End If
Call column_
End Sub


Sub column_()

Columns("D").Select
Selection.EntireColumn.AutoFit
Range("D1").Select

End Sub

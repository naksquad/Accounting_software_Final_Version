VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
If Sheets("feuil28").Range("a20") = 1 Then
Cancel = False
Else
Cancel = True
End If
End Sub
Private Sub Workbook_Open()
Sheets("feuil28").Range("a20").ClearContents
MsgBox Now
End Sub


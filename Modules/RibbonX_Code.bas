Attribute VB_Name = "RibbonX_Code"
Option Explicit
Const sResourcePrefix As String = "RES_"

'Get Culture
Private Function GetATPUICultureTag() As String
    Dim shTemp As Worksheet
    Dim sCulture As String
    Dim sSheetName As String
    
    sCulture = Application.International(xlUICultureTag)
    sSheetName = sResourcePrefix + sCulture
    
    On Error Resume Next
    Set shTemp = ThisWorkbook.Worksheets(sSheetName)
    On Error GoTo 0
    If shTemp Is Nothing Then sCulture = GetFallbackTag(sCulture)
    
    GetATPUICultureTag = sCulture
End Function

'Entry point for RibbonX button click
Sub ShowATPDialog(control As IRibbonControl)
    Application.Run ("fDialog")
End Sub

'Callback for RibbonX button label
Sub GetATPLabel(control As IRibbonControl, ByRef label)
    label = ThisWorkbook.Sheets(sResourcePrefix + GetATPUICultureTag()).Range("RibbonCommand").Value
End Sub

'Callback for screentip
Public Sub GetATPScreenTip(control As IRibbonControl, ByRef label)
    label = ThisWorkbook.Sheets(sResourcePrefix + GetATPUICultureTag()).Range("ScreenTip").Value
End Sub

'Callback for Super Tip
Public Sub GetATPSuperTip(control As IRibbonControl, ByRef label)
    label = ThisWorkbook.Sheets(sResourcePrefix + GetATPUICultureTag()).Range("SuperTip").Value
End Sub

Public Sub GetGroupName(control As IRibbonControl, ByRef label)
    label = ThisWorkbook.Sheets(sResourcePrefix + GetATPUICultureTag()).Range("GroupName").Value
End Sub

'Check for Fallback Languages
Private Function GetFallbackTag(szCulture As String) As String
    'Sorted alphabetically by returned culture tag, then input culture tag
    Select Case (szCulture)
        Case "rm-CH"
            GetFallbackTag = "de-DE"
        Case "ca-ES", "ca-ES-valencia", "eu-ES", "gl-ES"
            GetFallbackTag = "es-ES"
        Case "lb-LU"
            GetFallbackTag = "fr-FR"
        Case "nn-NO"
            GetFallbackTag = "nb-NO"
        Case "be-BY", "ky-KG", "tg-Cyrl-TJ", "tt-RU", "uz-Latn-UZ"
            GetFallbackTag = "ru-RU"
        Case Else
            GetFallbackTag = "en-US"
    End Select
End Function

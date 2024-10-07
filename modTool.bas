Attribute VB_Name = "modTool"

Option Explicit

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWMAXIMIZED As Integer = 3

Public Sub SelectExpression(ByRef Box As TextBox)

    Box.SelStart = 0
    Box.SelLength = Len(Box.Text)

End Sub

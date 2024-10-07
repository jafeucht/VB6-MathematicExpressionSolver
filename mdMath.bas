Attribute VB_Name = "mdMathTool"
' Jonathan A. Feucht
' Mathematics simulator
'-------------------------
' This module contains stuff useful to both the Exp and MathTool class.

Option Explicit

Enum Errors
    None
    Infinity
    Syntax
    FuncInvalid
    Assignment
End Enum

' Math tool supports variables, although it does not allow assignment within the
' input expression. Those must be modified through code using Sub MathTool.SetVar.
Public Vars() As Variable

' Returns the number of variables.
Public Function VarCnt() As Integer

    On Error Resume Next
      VarCnt = UBound(Vars)

End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MathTool"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Jonathan A. Feucht
' Mathematics simulator
'-------------------------
' This is the main class. It can be used by remote classes. It works to solve
' expressions and to set or modify variables.

Option Explicit

Type Variable   ' A variable has a reference name and a value
    Name As String
    Constant As Boolean
    Value As Double
End Type

' Initialize useful constants
Private Sub Class_Initialize()

    SetVar "pi", Atn(1) * 4, True
    SetVar "e", Exp(1), True

End Sub

' Removes all unwanted characters from an expression.
Public Function CleanExpression(Expression As String) As String

  Dim i As Integer, CurChar As String, RetStr As String

    Expression = LCase$(Expression)
    For i = 1 To Len(Expression)
        CurChar = Mid$(Expression, i, 1)
        If Asc(CurChar) > 32 Then   ' If it's a good character, add it to the return string.
            ' We don't allow parenteses in an expression, since IsNumeric ("(7)")
            ' is true and IsNumeric ("[7]") is false. This is important.
            If CurChar = "(" Or CurChar = "{" Then
              CurChar = "["
            ElseIf CurChar = ")" Or CurChar = "}" Then
              CurChar = "]"
            End If
            RetStr = RetStr & CurChar
        End If
    Next i
    CleanExpression = RetStr

End Function

' Returns the number of entries in the Vars() array
Public Function GetVarCount() As Integer

    GetVarCount = VarCnt

End Function

' Returns the Name property for the Vars(Index) array entry
Public Function GetVarName(Index As Integer) As String

    GetVarName = Vars(Index).Name

End Function

' Returns the Value property for the Vars(Index) array entry
Public Function GetVarValue(Index As Integer) As Double

    GetVarValue = Vars(Index).Value

End Function

' Returns the Constant property for the Vars(Index) array entry
Public Function IsVarConstant(Index As Integer) As Boolean

    IsVarConstant = Vars(Index).Constant

End Function

' Create or modify a variable
Public Sub SetVar(VarName As String, Value As Double, Optional Constant As Boolean)

  Dim i As Integer, VarCount As Integer, CurChar As String

    On Error Resume Next
      ' Minimum 1 character
      If VarName = "" Then
          Err.Raise Assignment, , "Variable names can have at least 1 character"
          Exit Sub '>---> Bottom
      End If
      ' Maximum 15 characters
      If Len(VarName) > 15 Then
          Err.Raise Assignment, , "Variable names can have at most 15 characters"
          Exit Sub '>---> Bottom
      End If
      ' First character must be a letter
      If IsNumeric(Left$(VarName, 1)) Then
          Err.Raise Assignment, , "First character of variable name must be a letter"
          Exit Sub '>---> Bottom
      End If
      ' Capitalize the name
      VarName = UCase$(VarName)
      ' Check for invalid characters
      For i = 1 To Len(VarName)
          CurChar = Mid$(VarName, i, 1)
          ' Only allow letters, numbers, and an underscore
          If Not ((CurChar >= "A" And CurChar <= "Z") Or (CurChar >= "0" And CurChar <= "9") Or (CurChar = "_")) Then
              Err.Raise Assignment, , "Invalid character in variable name"
          End If
      Next i
      VarCount = VarCnt
      ' Check to see if variable already exists
      For i = 1 To VarCount
          ' If so, change the value
          If Vars(i).Name = VarName Then
              If Vars(i).Constant Then
                  Err.Raise Assignment, , VarName & " is constant"
                  Exit Sub '>---> Bottom
              End If
              Vars(i).Value = Value
              Exit Sub '>---> Bottom
          End If
      Next i
      ' Otherwise, add the variable
      ReDim Preserve Vars(1 To VarCount + 1)
      Vars(VarCount + 1).Name = VarName
      Vars(VarCount + 1).Value = Value
      Vars(VarCount + 1).Constant = Constant

End Sub

' Solves an expression, returns the answer
Public Function Solve(Expression As String) As Double

  Dim Equat As New Exp

    If Err.Number > 0 Then Exit Function
    ' Remove unwanted characters
    Equat.Expression = CleanExpression(Expression)
    ' Solve the equation
    Equat.SolveExpression
    ' Return the value of the equation
    Solve = Equat.GetValue
    
End Function

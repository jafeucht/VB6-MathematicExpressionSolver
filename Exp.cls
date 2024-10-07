VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Exp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Jonathan A. Feucht
' Mathematics simulator
'-------------------------
' This private class describes a part of an expression, determined by the expression's
' parenthesis. For instance, 3*(5+4^2) has two parts, part1=3*(part2) and part2=5+4^2.
' The math tool recursively splits an expression into Exp classes containing data
' found in the parenthesis. Afterwards, each part is solved, and replaced into the
' prior expression. All this is found in Sub SolveExpression.

' Part2=5+4^2 becomes part2=21, then that is substituted into part1.
' Part1 becomes 3*9, which is 27. The process of carrying out the mathematical
' operations is described in Function Solve.

Option Explicit
Option Compare Text

Const OPS As String = "^$*/%\+-"  ' Operator characters organized by priority in sets of two

Private Type SepData
    ParseInt As Variant
    Operator As String
    NextOp As Integer
End Type

Private Enum OpModes
    OpExp
    OpMod
    OpMult
    OpAdd
End Enum

Public Expression As String
Dim Expressions() As New Exp
Dim eqVal As Double

'This function returns the number of elements in the array Expressions
Private Function GetExpressionCnt() As Integer

    On Error Resume Next
      GetExpressionCnt = UBound(Expressions)

End Function

'This property is used to return the value for the solved equation.
Function GetValue() As Double ':( Missing Scope

    GetValue = eqVal

End Function

' GoFunct executes functions, parenthesized expressions, and variables
Private Function GoFunct(ExpresVal As String) As Double

  Dim i As Integer
  Dim OpenPar As Integer, ClosePar As Integer
  Dim FunctName As String, RefVal As Double
  Dim PI As Double
  Dim Dec As Double

    On Error Resume Next

      PI = Atn(1) * 4

      ' Find next closing and ending parentheses
      OpenPar = InStr(1, ExpresVal, "[")
      ClosePar = InStr(1, ExpresVal, "]")
    
      ' If there are no parentheses, it must be a variable
      If OpenPar = 0 And ClosePar = 0 Then
          ' If it is blank, then be totally confused
          If ExpresVal = "" Or ExpresVal = "+" Or ExpresVal = "-" Then Exit Function ':( Expand Structure or consider reversing Condition
          ' Otherwise, look up the variable
          For i = 1 To VarCnt
              ' If we find the variable name, return the variable value
              If ExpresVal = Vars(i).Name Then
                  GoFunct = Vars(i).Value
                  Exit Function '>---> Bottom
              End If
          Next i
          ' If the variable doesn't exist, be totally confused
          Err.Raise FuncInvalid, , "Not recognized: """ & UCase$(ExpresVal) & """"
        Else 'NOT OPENPAR...
          ' Get the data preceeding the opening parenthesis
          FunctName = Trim$(Left$(ExpresVal, OpenPar - 1))
          ' Get the part referenced by the function
          RefVal = Expressions(Val(Mid$(ExpresVal, OpenPar + 1, ClosePar - OpenPar - 1))).GetValue
        
          ' Execute the function
          Select Case FunctName
            Case "SIN"  ' Sine
              GoFunct = Sin(RefVal)
            Case "CSC"  ' Cosecant
              GoFunct = 1 / Sin(RefVal)
            Case "COS"  ' Cosine
              GoFunct = Cos(RefVal)
            Case "ARCSIN"   ' Arc sine
              ' The arc sine of 1 or -1 normally returns an error in Visual Basic
              If Abs(RefVal) = 1 Then
                  GoFunct = Atn(1) * 4 / 2 * Sgn(RefVal)
                Else 'NOT ABS(REFVAL)...
                  GoFunct = Atn(RefVal / Sqr(-RefVal * RefVal + 1))
              End If
            Case "ARCCOS"   ' Arc cosine
              ' The arc cosine of 1 or -1 normally returns an error in Visual Basic
              If RefVal = 1 Then
                  GoFunct = 0
                ElseIf RefVal = -1 Then 'NOT REFVAL...
                  GoFunct = PI
                Else 'NOT REFVAL...
                  GoFunct = Atn(-RefVal / Sqr(-RefVal * RefVal + 1)) + 2 * Atn(1)
              End If
            Case "ARCTAN"   ' Arc tangent
              GoFunct = Atn(RefVal)
            Case "SEC"  ' Secant
              GoFunct = 1 / Cos(RefVal)
            Case "TAN"  ' Tangent
              GoFunct = Tan(RefVal)
            Case "COT"  ' Cotangent
              GoFunct = 1 / Tan(RefVal)
            Case "ABS"  ' Absolute value
              GoFunct = Abs(RefVal)
            Case "EXP"  ' e raised to a power
              GoFunct = Exp(RefVal)
            Case "INT"  ' Integer value
              GoFunct = Int(RefVal)
            Case "FIX"  ' RefVal without the decimal part
              GoFunct = Fix(RefVal)
            Case "DECPART"  ' Decimal part
              GoFunct = RefVal - Fix(RefVal)
            Case "ROUND"    ' Round to nearest integer
              Dec = Abs(RefVal - Fix(RefVal))
              GoFunct = Fix(RefVal)
              If Dec >= 0.5 Then GoFunct = GoFunct + (1 * Sgn(RefVal)) ':( Expand Structure
            Case "LN"   ' Natural log
              GoFunct = Log(RefVal)
            Case "LOG"  ' Log base 10
              GoFunct = Log(RefVal) / Log(10)
            Case "RND"  ' Random
              Randomize Timer
              GoFunct = Rnd * RefVal
            Case "SGN"  ' Positive/Negative sign
              GoFunct = Sgn(RefVal)
            Case "DTR"  ' Decimal to radians
              GoFunct = RefVal * (Atn(1) * 4) / 180
            Case "RTD"  ' Radians to decimal
              GoFunct = RefVal * 180 / (Atn(1) * 4)
            Case "" ' No function name, must be just an expression
              GoFunct = RefVal
            Case Else   ' Function name not found, so be totally confused
              Err.Raise FuncInvalid, , "No function """ & UCase$(FunctName) & """"
          End Select
      End If
      
      If Err.Number > 4 Then
          Err.Raise Infinity, , "Invalid argument"
      End If

End Function

' This function returns the solved part
Private Function Solve() As Double

  Dim OpMode As OpModes, OpVal As Integer
  Dim NewOp As Integer, OldOp As Integer
  Dim ParseOps() As SepData

    On Error Resume Next
    
      ' Split the expression into parts separated by operators
      Split ParseOps(), Expression
      ' If we hit an error in the process, be totally confused
      If Err.Number > 0 Then Exit Function
      ' Loop through operator priorities
      For OpMode = OpExp To OpAdd
          ' Reset the operator pointers
          NewOp = 1
          Do
              ' Point to next operator, save old pointer
              OldOp = NewOp
              NewOp = ParseOps(OldOp).NextOp
            
              ' If there are no more operators, move on to next operator priority
              If NewOp = 0 Then Exit Do
            
              ' Get current operator priority
              OpVal = InStr(1, OPS, ParseOps(OldOp).Operator)
            
              ' If the current operator has the priority we're looking for, then
              ' solve
              If OpMode = (OpVal - 1) \ 2 Then
            
                  Select Case ParseOps(OldOp).Operator
                    Case "^" ' Exponant Operator
                      ParseOps(OldOp).ParseInt = ParseOps(OldOp).ParseInt ^ ParseOps(NewOp).ParseInt
                      ' Root of a negative number deserves utter confusion
                      If Err.Number = 5 Then Err.Raise Infinity, , "Imaginary number" ':( Expand Structure
                    Case "$" ' Square root operator
                      ParseOps(OldOp).ParseInt = ParseOps(OldOp).ParseInt ^ (1 / ParseOps(NewOp).ParseInt)
                      If Err.Number = 5 Then Err.Raise Infinity, , "Imaginary number" ':( Expand Structure
                    Case "*" ' Multiplication Operator
                      ParseOps(OldOp).ParseInt = ParseOps(OldOp).ParseInt * ParseOps(NewOp).ParseInt
                    Case "/" ' Division Operator
                      ParseOps(OldOp).ParseInt = ParseOps(OldOp).ParseInt / ParseOps(NewOp).ParseInt
                    Case "%" ' Mod Operator
                      ParseOps(OldOp).ParseInt = ParseOps(OldOp).ParseInt Mod ParseOps(NewOp).ParseInt
                    Case "\" ' Div Operator
                      ParseOps(OldOp).ParseInt = ParseOps(OldOp).ParseInt \ ParseOps(NewOp).ParseInt
                    Case "+" ' Addition Operator
                      ParseOps(OldOp).ParseInt = ParseOps(OldOp).ParseInt + ParseOps(NewOp).ParseInt
                    Case "-" ' Subtraction Operator
                      ParseOps(OldOp).ParseInt = ParseOps(OldOp).ParseInt - ParseOps(NewOp).ParseInt
                  End Select
                
                  With ParseOps(NewOp)
                      ParseOps(OldOp).NextOp = .NextOp
                      ParseOps(OldOp).Operator = .Operator
                  End With 'PARSEOPS(NEWOP)
                
                  NewOp = OldOp
              End If
            
          Loop
      Next OpMode
    
      If Err.Number >= 5 Or Err.Number = 1 Then
          Select Case Err.Number
            Case 11
              Err.Raise Infinity, , "Division by 0"
            Case Else
              Err.Raise Infinity, , "Overflow"
          End Select
      End If
    
      Solve = ParseOps(1).ParseInt

End Function

'This procedure solves the mathematical expression
Public Sub SolveExpression()

  Dim BegPos As Integer, EndPos As Integer
  Dim BegPar As Integer, EndPar As Integer
  Dim CurPos As Integer, ParDepth As Integer
  Dim NewParse As Integer, OldLen As Integer

    ' Add a space at the end of the expression if there is not one already
    If Not Right$(Expression, 1) = " " Then Expression = Expression & " "
    
    CurPos = 1
    Do
        ' Find parentheses positions
        BegPos = InStr(CurPos, Expression, "[")
        EndPos = InStr(CurPos, Expression, "]")
        
        ' If the parentheses don't exist, point them to the last character
        ' of the expression
        If BegPos = 0 Then BegPos = Len(Expression)
        If EndPos = 0 Then EndPos = Len(Expression)
        If BegPos < EndPos Then ' If the next parenthesis is a "["...
            ' Point to the data within the parentheses
            CurPos = BegPos + 1
            ' We only read data within the level 1 parenthesis. For instance, in the
            ' expression "6*[9+4*{6-SIN(PI/2)}]+[4-3*{5}]", the brackets are level 1
            ' parenthesis, braces are level 2, and parentheses are level 3.
            If ParDepth = 0 Then BegPar = CurPos - 1
            ' We're getting deeper
            ParDepth = ParDepth + 1
          ElseIf BegPos > EndPos Then ' If the next parenthesis is a "]"...
            ' Point to data outside parenthesis
            CurPos = EndPos + 1
            ' We're getting shallower
            ParDepth = ParDepth - 1
            ' Check if we hit ending parenthesis for a level 1 statement
            If ParDepth = 0 Then
                ' If so, then point to last character within parentheses
                EndPar = CurPos - 1
                ' Create a new expression containing the data within the parentheses
                ' marked by the BegPar and EndPar pointers
                NewParse = GetExpressionCnt + 1
                ReDim Preserve Expressions(1 To NewParse)
                Expressions(NewParse).Expression = Mid$(Expression, BegPar + 1, EndPar - BegPar - 1)
                ' Solve the new expression
                Expressions(NewParse).SolveExpression
                ' If there was an error in the solving process, be totally confused
                If Err.Number > 0 Then Exit Sub
                ' Modify the expression to contain a link to the new expression
                Expression = Left$(Expression, BegPar - 1) & "[" & NewParse & "]" & Right$(Expression, Len(Expression) - EndPar)
                ' Point to the new position for the closing parenthesis
                CurPos = InStr(BegPar, Expression, "]") + 1
            End If
          Else    ' There are no parentheses, so exit
            Exit Do '>---> Loop
        End If
    Loop
    
    If ParDepth <> 0 Then   ' Wrong number of parenthesis, be totally confused
        Err.Number = Syntax
        Err.Description = "Parenthesis expected"
      Else 'NOT PARDEPTH...
        eqVal = Solve() ' Solve the current part
    End If

End Sub

Private Function Split(ByRef PartArray() As SepData, StrVal As String)

  Dim i As Integer, Done As Boolean
  Dim StartPart As Integer, EndPart As Integer, CurPart As String
  Dim SepVal As Integer, ArrPos As Integer, CurOp As String
  Dim CurChar As String, DecHit As Boolean, SignDat As Integer
  Dim OpPos As Integer
    
    On Error Resume Next
    
      SignDat = 1
    
      ' Point to first character in StrVal
      StartPart = 1
      '
      If Left$(StrVal, 1) = "-" Then StrVal = "0" & StrVal ':( Expand Structure
      ' Loop through characters in StrVal
      Do
          ' Add a character to the string selection
          EndPart = EndPart + 1
          ' Get the last character in the selection
          CurChar = Mid$(StrVal, EndPart, 1)
          ' Check if the character is in the list of separating characters
          SepVal = InStr(1, OPS, CurChar)
          ' Check for decimal place
          If CurChar = "." Then
              ' If there is a double decimal place, be totally confused
              If DecHit Then
                  Err.Raise Syntax, , "Double decimal"
                  Exit Function '>---> Bottom
              End If
              DecHit = True
          End If
          Done = EndPart >= Len(StrVal)
          OpPos = InStr(1, OPS, CurChar)
          ' If the current character is an operator, add it to the parts list
          If OpPos > 0 Or Done Then
              CurPart = Mid$(StrVal, StartPart, EndPart - StartPart)
              If Not Done Then CurOp = Mid$(OPS, SepVal, 1)
            
              If Not (IsNumeric(CurPart) Or CurPart = "") Then
                
                  If Not Right$(Trim$(CurPart), 1) = "]" And InStr(1, CurPart, "[") > 0 Then Err.Raise Syntax, , "Operator expected"
                
                  CurPart = CStr(GoFunct(CurPart))
                ElseIf CurPart = "" And Done Then 'NOT NOT...
                  Err.Raise Syntax, , "Expected expression"
                  Exit Function '>---> Bottom
              End If
              If Err.Number > 1 Then Exit Function
              If Not IsNumeric(CurPart) And Not Err.Number = Infinity Then
                  Select Case CurOp
                    Case "-"
                      SignDat = -SignDat
                    Case "+"
                    Case Else
                      Err.Raise Syntax, , "Unexpected operator"
                  End Select
                Else
                  ArrPos = ArrPos + 1
                  ReDim Preserve PartArray(1 To ArrPos)
                  ' If the current part isn't isnumeric, then run it as a function /
                  ' expression / variable
                  ' Store the data into the parts list
                  With PartArray(ArrPos)
                      .ParseInt = Val(CurPart) * SignDat
                      If Not Done Then
                          .Operator = Mid$(OPS, SepVal, 1)
                          .NextOp = ArrPos + 1
                      End If
                  End With 'PARTARRAY(ARRPOS)
                  SignDat = 1
              End If
              ' Reset the pointer
              StartPart = EndPart + 1
          End If
      Loop Until Done

End Function

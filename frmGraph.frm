VERSION 5.00
Begin VB.Form frmGraph 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Graph"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pctcontainer 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   4800
      ScaleHeight     =   4455
      ScaleWidth      =   4815
      TabIndex        =   14
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtClarity 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Text            =   "1"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "&Help..."
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox Equation 
         Height          =   285
         Left            =   480
         TabIndex        =   0
         Text            =   "X"
         Top             =   360
         Width           =   4215
      End
      Begin VB.CommandButton cmdSolve 
         Caption         =   "Solve!"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox txtFromY 
         Height          =   285
         Left            =   480
         TabIndex        =   3
         Text            =   "-10"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtToY 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Text            =   "10"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtFromX 
         Height          =   285
         Left            =   480
         TabIndex        =   1
         Text            =   "-10"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtToX 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Text            =   "10"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblClarity 
         AutoSize        =   -1  'True
         Caption         =   "Clarity:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label lblEquation 
         AutoSize        =   -1  'True
         Caption         =   "Equation:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   675
      End
      Begin VB.Label lblY 
         Caption         =   "y ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         Caption         =   "From:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   390
      End
      Begin VB.Label lblFromY 
         Caption         =   "y ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblToY 
         Caption         =   "y ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         Caption         =   "To:"
         Height          =   195
         Left            =   1440
         TabIndex        =   9
         Top             =   840
         Width           =   240
      End
      Begin VB.Label lblFromX 
         Caption         =   "x ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblToX 
         Caption         =   "x ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   1080
         Width           =   375
      End
   End
   Begin VB.PictureBox pctGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   120
      ScaleHeight     =   298
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   298
      TabIndex        =   15
      Top             =   120
      Width           =   4500
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   4680
      Width           =   45
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Jonathan A. Feucht
' MathTool Graph Sample
' ----------------------------------------------------------------------

Option Explicit

' API call used to load the help document
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWMAXIMIZED = 3

Private Type PointAPI
    X As Double
    Y As Double
End Type

' Graph data
Dim Clarity As Double
Dim Range As Double, Domain As Double
Dim FromPos As PointAPI, ToPos As PointAPI
Dim Step As PointAPI

Sub SetStatus(Description As String, Status As String)
    If Len(Description) > 0 Then Description = Description & ":"
    lblStatus = Description & " " & Status
End Sub

Private Sub cmdHelp_Click()
Dim HelpPath As String
    HelpPath = App.Path
    If Not Right$(HelpPath, 1) = "\" Then HelpPath = HelpPath & "\"
    ShellExecute hwnd, "Open", "MathTool.doc", "", HelpPath, SW_SHOWMAXIMIZED
End Sub

Private Sub cmdSolve_Click()
    On Error GoTo FoundErr
    
    EnableContainerObjects Me, pctcontainer, False
    
    Clarity = Val(txtClarity)
    
    FromPos.X = SolveEq(txtFromX)
    FromPos.Y = SolveEq(txtFromY)
    ToPos.X = SolveEq(txtToX)
    ToPos.Y = SolveEq(txtToY)
    
    Range = ToPos.X - FromPos.X
    Domain = ToPos.Y - FromPos.Y
       
    If Range <= 0 Then Err.Raise 1, , "Empty range"
    If Domain <= 0 Then Err.Raise 1, , "Empty domain"
            
    Step.X = (pctGraph.ScaleWidth - 1) / Range
    Step.Y = (pctGraph.ScaleHeight - 1) / Domain
    
    ' Test the equation for syntax errors
    SetVar "X", 0
    
    If Err.Number > 0 Then Err.Raise Err.Number, , Err.Description
    
    On Error Resume Next
    
    SolveEq Equation
    
    If Err.Number > 1 And Err.Number < 5 Then GoTo FoundErr
    
    pctGraph.Cls
    DrawGrid
    GraphEq
    
    EnableContainerObjects Me, pctcontainer, True
    
    SetStatus "Equation", "y=" & CleanExpression(Equation)
    
    Exit Sub
    
FoundErr:
    MsgBox Err.Description, vbCritical, "Error"
    
    Err.Clear
    
    EnableContainerObjects Me, pctcontainer, True
End Sub

Sub GraphEq()
Dim i As Long
Dim CurPt As PointAPI, OldPt As PointAPI

    On Error Resume Next
    
    pctGraph.DrawWidth = 1
    For i = 0 To pctGraph.ScaleWidth * Range * Clarity
            
        OldPt = CurPt
        CurPt.X = i / (Range * Clarity)
        SetVar "X", CurPt.X / Step.X + FromPos.X
        CurPt.Y = pctGraph.ScaleHeight - (Solve(Equation) - FromPos.Y) * Step.Y
  
        If CurPt.Y <= pctGraph.ScaleHeight And CurPt.Y >= 0 And Err.Number = 0 Then
            pctGraph.PSet (CurPt.X, CurPt.Y), 0
        End If
        
    Next i
End Sub

Sub DrawGrid()
Dim i As Double, Pos As Double
Dim Location As PointAPI
Dim StepVal As PointAPI, Start As PointAPI, DecStr As String
Dim CurPos As Double

    DecStr = CStr(Range)
    StepVal.X = 1 * Mid$(DecStr, InStr(1, DecStr, ".") + 1) / 10
    DecStr = CStr(Domain)
    StepVal.Y = 1 * Mid$(DecStr, InStr(1, DecStr, ".") + 1) / 10

    DrawWidth = 1
    
    pctGraph.DrawWidth = 1
    pctGraph.ForeColor = RGB(225, 225, 225)
    
    For i = FromPos.X To ToPos.X + 1 Step 0.1
        Pos = (Fix(i) - FromPos.X) * Step.X
        pctGraph.Line (Pos, 0)-(Pos, pctGraph.ScaleHeight)
    Next i
    For i = FromPos.Y To ToPos.Y + 1 Step Sgn(Domain)
        Pos = pctGraph.ScaleHeight - (Fix(i) - FromPos.Y) * Step.Y - 1
        pctGraph.Line (0, Pos)-(pctGraph.ScaleWidth, Pos)
    Next i
    
    pctGraph.DrawWidth = 2
    pctGraph.ForeColor = RGB(225, 125, 255)
    
    Pos = -FromPos.X * Step.X
    pctGraph.Line (Pos, 0)-(Pos, pctGraph.ScaleHeight)
    
    pctGraph.ForeColor = RGB(125, 225, 225)
    
    Pos = pctGraph.ScaleHeight + FromPos.Y * Step.Y - 1
    pctGraph.Line (0, Pos)-(pctGraph.ScaleWidth, Pos)

End Sub

Function SolveEq(ByRef Equation As TextBox) As Double
    If Err.Number > 0 Then Exit Function
    SolveEq = Solve(Equation.Text)
    If Err.Number > 1 And Err.Number < 5 Then
        Equation.ForeColor = RGB(150, 0, 0)
    Else
        Equation.ForeColor = 0
    End If
End Function

Function GetFraction(Number As Double) As Double
    GetFraction = Number - Int(Number)
End Function

Public Sub EnableContainerObjects(ContainerForm As Form, ContainerName As Control, Enabled As Boolean)
Dim i As Integer
    For i = 0 To ContainerForm.Controls.Count - 1
        If ContainerForm.Controls(i).Container.Name = ContainerName.Name Then
            ContainerForm.Controls(i).Enabled = Enabled
        End If
    Next i
End Sub

Sub UpperCase(ByRef KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

Private Sub Equation_KeyPress(KeyAscii As Integer)
    UpperCase KeyAscii
End Sub

Private Sub txtFromX_KeyPress(KeyAscii As Integer)
    UpperCase KeyAscii
End Sub

Private Sub txtFromY_KeyPress(KeyAscii As Integer)
    UpperCase KeyAscii
End Sub

Private Sub txtToX_KeyPress(KeyAscii As Integer)
    UpperCase KeyAscii
End Sub

Private Sub txtToY_KeyPress(KeyAscii As Integer)
    UpperCase KeyAscii
End Sub

VERSION 5.00
Begin VB.Form frmTool 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MathTool, Mathematics Engine"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help..."
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Answer 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox Expression 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Variables"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   7815
      Begin VB.CommandButton cmdChange 
         Caption         =   "&Change..."
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtVarValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
      Begin VB.ComboBox cboVars 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Label lblAnswer 
      Caption         =   "Answer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   825
   End
End
Attribute VB_Name = "frmTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Answer_GotFocus()

    SelectExpression Answer

End Sub

Private Sub cboVars_Click()

    txtVarValue = GetVarValue(cboVars.ListIndex + 1)

End Sub

Private Sub cmdChange_Click()

    frmVar.ChangeVar cboVars.ListIndex + 1
    RefreshVars

End Sub

Private Sub cmdExit_Click()

    End

End Sub

Private Sub cmdHelp_Click()

  Dim HelpPath As String

    HelpPath = App.Path
    If Not Right$(HelpPath, 1) = "\" Then HelpPath = HelpPath & "\"
    ShellExecute hwnd, "Open", "MathTool.doc", "", HelpPath, SW_SHOWMAXIMIZED

End Sub

Private Sub cmdSolve_Click()

    Answer = Solve(Expression)
    If Err.Number > 0 Then
        MsgBox Err.Description, vbCritical, "Error"
        Expression.SetFocus
    End If

End Sub

Private Sub Expression_Change()

    Answer = Solve(Expression)
    If Err.Number > 0 Then Answer = Err.Description

End Sub

Private Sub Expression_GotFocus()

    SelectExpression Expression

End Sub

Private Sub Expression_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub Form_Load()

    RefreshVars
    cboVars.ListIndex = 0
    Expression = 0

End Sub

Sub RefreshVars()

  Dim i As Integer, OldIdx As Integer

    OldIdx = cboVars.ListIndex
    cboVars.Clear
    For i = 1 To getVarCount
        cboVars.AddItem GetVarName(i)
    Next i
    If cboVars.ListCount > 0 Then cboVars.ListIndex = OldIdx

End Sub

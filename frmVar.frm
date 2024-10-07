VERSION 5.00
Begin VB.Form frmVar 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtVarValue 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox txtVarName 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblVarValue 
      AutoSize        =   -1  'True
      Caption         =   "Variable Value:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1065
   End
   Begin VB.Label lblVarVal 
      AutoSize        =   -1  'True
      Caption         =   "Variable Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frmVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ChangeVar(VarIndex As Integer)
    Err.Clear
    txtVarName = GetVarName(VarIndex)
    txtVarValue = GetVarValue(VarIndex)
    Show vbModal
    If Err.Number > 0 Then MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSet_Click()
    SetVar txtVarName, Solve(txtVarValue)
    Unload Me
End Sub

Private Sub txtVarName_GotFocus()
    SelectExpression txtVarName
End Sub

Private Sub txtVarName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtVarValue_GotFocus()
    SelectExpression txtVarValue
End Sub

Private Sub txtVarValue_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

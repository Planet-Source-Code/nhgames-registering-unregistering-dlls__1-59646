VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Registering Dlls "
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Silent UnRegister"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Silent Register"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Unregister"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "C:\"
      ToolTipText     =   "Enter Server Path"
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Registering Dlls For Dummies
'NHGames
'http://nhgames.tk

Private Sub Command1_Click()
'Register
 Shell "Regsvr32 " & Text1.Text, vbMaximizedFocus
End Sub

Private Sub Command2_Click()
'Unregister
 Shell "Regsvr32 /u " & Text1.Text, vbMaximizedFocus
End Sub

Private Sub Command4_Click()
' Silent Unregister
 Shell "Regsvr32 /s /u " & Text1.Text, vbMaximizedFocus
End Sub

Private Sub Command3_Click()
' Silent register
 Shell "Regsvr32 /s " & Text1.Text, vbMaximizedFocus
End Sub

Private Sub Form_Load()
'//
End Sub

Private Sub Text1_Change()
'//
End Sub

VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Enter Handle"
   ClientHeight    =   1455
   ClientLeft      =   3510
   ClientTop       =   3540
   ClientWidth     =   4455
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4455
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter your handle:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
End
End Sub
Private Sub cmdOk_Click()
If txtName.Text = "" Then
MsgBox "Please enter your name"
txtName.SetFocus
Else
frmMain.Show
End If
End Sub


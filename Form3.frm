VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Edit IP List"
   ClientHeight    =   2040
   ClientLeft      =   7455
   ClientTop       =   1995
   ClientWidth     =   2490
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   2490
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save to List"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Enter computer IP address:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Enter computer name:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
On Error GoTo error
Dim Filepath As String
Filepath = "C:\program files\LanChat32\IPList.txt"
On Error GoTo error
Open Filepath For Append As #1
Print #1, txtName.Text & ":   " & txtIP.Text
Close #1
Form4.Show
Form3.Hide
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub
Private Sub Form_Load()
txtName.Text = ""
txtIP.Text = ""
End Sub


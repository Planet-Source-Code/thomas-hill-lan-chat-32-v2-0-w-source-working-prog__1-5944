VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Local Computer List"
   ClientHeight    =   3705
   ClientLeft      =   6390
   ClientTop       =   1995
   ClientWidth     =   4575
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4575
   Begin VB.TextBox txtList 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Computer to List"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   4335
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtConnect 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Enter IP to connect to and click connect"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Computer Name: IP Number"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Local computer list:"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Unload Form4
Form3.Show
End Sub
Private Sub cmdConnect_Click()
Form4.Hide
frmMain.txtRemote.Text = txtConnect.Text
On Error Resume Next
frmMain.wData.Close
frmMain.wData.Connect frmMain.txtRemote.Text, frmMain.txtPort.Text
frmMain.lblStatus.Caption = "Connecting..."
frmMain.cmdDisconnect.Visible = True
frmMain.cmdListen.Visible = False
frmMain.cmdCancel.Visible = True
If Err Then frmMain.lblStatus.Caption = Err.Description
End Sub
Private Sub Form_Load()
Open ("C:\program files\lanchat32\iplist.txt") For Input As #1
txtList.Text = Input(LOF(1), #1)
Close #1
End Sub

VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lan Chat 32"
   ClientHeight    =   2190
   ClientLeft      =   5520
   ClientTop       =   720
   ClientWidth     =   4380
   FillColor       =   &H00808080&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4380
   Begin VB.CommandButton cmdMin2 
      Caption         =   "min"
      Height          =   375
      Left            =   3960
      TabIndex        =   32
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmdSend2 
      Caption         =   "Send"
      Height          =   288
      Left            =   3480
      TabIndex        =   31
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   ">>"
      Height          =   375
      Left            =   3960
      TabIndex        =   30
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "<<"
      Height          =   375
      Left            =   3960
      TabIndex        =   29
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H80000008&
      Caption         =   "&Send"
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "About Lan Chat 32"
      Height          =   255
      Left            =   4800
      TabIndex        =   27
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Minimize"
      Height          =   255
      Left            =   4800
      TabIndex        =   25
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Options"
      Height          =   255
      Left            =   4800
      TabIndex        =   21
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H80000004&
      Caption         =   "Check3"
      Height          =   255
      Left            =   4440
      TabIndex        =   24
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000008&
      Caption         =   "Keep On Top"
      Height          =   255
      Left            =   4800
      TabIndex        =   23
      Top             =   840
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   200
      Left            =   4440
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   22
      Top             =   120
      Width           =   200
   End
   Begin VB.Frame frmConnection 
      BackColor       =   &H80000004&
      Caption         =   "Connection"
      ForeColor       =   &H00000000&
      Height          =   1572
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   6255
      Begin VB.CommandButton cmdIP 
         Caption         =   "IP List"
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   5040
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtRemote 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtLocal 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0.0.0.0"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdListen 
         Caption         =   "&Listen"
         Height          =   255
         Left            =   5040
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "&Disconnect"
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   3120
         TabIndex        =   10
         Text            =   "1113"
         Top             =   240
         Width           =   492
      End
      Begin VB.TextBox txtNick 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   960
         TabIndex        =   9
         Text            =   "YourNickHere"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblRemote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remote Host:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   984
      End
      Begin VB.Label lblLocal 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Local IP:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3840
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2640
         TabIndex        =   18
         Top             =   240
         Width           =   330
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   6015
      End
      Begin VB.Label llnNick 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nickname:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   768
      End
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000004&
      Caption         =   "Check2"
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   600
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Enable Sound"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disable Sound"
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000004&
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      Caption         =   "Select Local Color"
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   840
      Top             =   240
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wData 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtText 
      Height          =   288
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   3255
   End
   Begin RichTextLib.RichTextBox txtData 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2778
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0442
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      BorderWidth     =   175
      X1              =   0
      X2              =   6240
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim frmHeight As Integer
Public Sub AlwaysOnTop(frmMain As Form, SetOnTop As Boolean)
If SetOnTop Then
lFlag = HWND_TOPMOST
Else
lFlag = HWND_NOTOPMOST
End If
SetWindowPos frmMain.hwnd, lFlag, frmMain.Left / Screen.TwipsPerPixelX, _
frmMain.Top / Screen.TwipsPerPixelY, frmMain.Width / Screen.TwipsPerPixelX, _
frmMain.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub
Private Sub cmdCancel_Click()
wData.Close
cmdCancel.Visible = False
cmdListen.Visible = True
lblStatus.Caption = "Operation Canceled"
End Sub

Private Sub cmdHide_Click()
If Me.Width > 4485 Then
Me.Width = 4485
End If
If Me.Height > 2565 Then
Me.Height = 2565
End If
txtText.Width = 3255
cmdSend2.Visible = True
Command4.Caption = "Options"
cmdHide.Visible = False
cmdShow.Visible = True
cmdMin2.Visible = True
End Sub

Private Sub cmdIP_Click()
Form4.Show
End Sub

Private Sub cmdMin2_Click()
frmMain.Hide
Form1.Show
End Sub

Private Sub cmdSend2_Click()
Dim SendStr As String
On Error Resume Next
SendStr = txtNick & ":" & vbTab & txtText.Text
wData.SendData SendStr
txtData.SelStart = Len(txtData.Text)
txtData.SelColor = vbBlue
txtData.SelText = txtNick & ":" & vbTab
txtData.SelStart = Len(txtData.Text)
txtData.SelColor = CommonDialog1.Color
txtData.SelText = txtText.Text & vbCrLf
txtText.Text = ""
If Err Then lblStatus.Caption = Err.Description
End Sub

Private Sub cmdShow_Click()
If Me.Width < 6630 Then
Me.Width = 6630
End If
cmdSend2.Visible = False
txtText.Width = 3735
cmdShow.Visible = False
cmdHide.Visible = True
cmdSend2.Visible = False
cmdMin2.Visible = False
End Sub

Private Sub Command1_Click()
CommonDialog1.ShowColor
Picture1.BackColor = CommonDialog1.Color
txtText.ForeColor = CommonDialog1.Color
End Sub
Private Sub Command2_Click()
If Check1.Enabled = False Then Check1.Enabled = True
If Check2.Enabled = True Then Check2.Enabled = False
Check1.Value = 1
Check2.Value = 0
End Sub
Private Sub Command3_Click()
If Check1.Enabled = True Then Check1.Enabled = False
If Check2.Enabled = False Then Check2.Enabled = True
Check2.Value = 1
Check1.Value = 0
End Sub
Private Sub Command4_Click()
If frmMain.Height < 4335 Then
frmMain.Height = 4335
Command4.Caption = "Hide Options"
Else
If frmMain.Height > 2565 Then
frmMain.Height = 2565
Command4.Caption = "Options"
End If
End If
End Sub
Private Sub Command5_Click()
frmAbout.Show
End Sub
Private Sub Command7_Click()
frmMain.Hide
Form1.Show
End Sub
Private Sub Form_Load()
txtLocal.Text = Winsock1.LocalIP
cmdDisconnect.Left = cmdConnect.Left
cmdDisconnect.Top = cmdConnect.Top
Timer1.Interval = 1
frmHeight = frmMain.Height
frmMain.Height = 100
wData.Listen
lblStatus.Caption = "Listening..."
Form2.Hide
txtNick.Text = Form2.txtName.Text
frmMain.Caption = "Hello, " & Form2.txtName.Text
cmdShow.Visible = True
cmdSend2.Visible = True
cmdHide.Visible = False
End Sub
Private Sub Timer1_Timer()
While frmMain.Height < frmHeight
frmMain.Height = frmMain.Height + 8
Wend
Timer1.Enabled = False
End Sub
Private Sub cmdConnect_Click()
On Error Resume Next
wData.Close
wData.Connect txtRemote.Text, txtPort.Text
lblStatus.Caption = "Connecting..."
cmdDisconnect.Visible = True
cmdListen.Visible = False
cmdCancel.Visible = True
If Err Then lblStatus.Caption = Err.Description
End Sub
Private Sub cmdDisconnect_Click()
wData.Close
wData.Listen
lblStatus.Caption = "Listening..."
cmdDisconnect.Visible = False
End Sub
Private Sub cmdListen_Click()
On Error Resume Next
wData.LocalPort = txtPort.Text
wData.Listen
lblStatus.Caption = "Listening..."
cmdListen.Visible = False
cmdCancel.Visible = True
If Err Then lblStatus.Caption = Err.Description
End Sub
Private Sub cmdSend_Click()
Dim SendStr As String
On Error Resume Next
SendStr = txtNick & ":" & vbTab & txtText.Text
wData.SendData SendStr
txtData.SelStart = Len(txtData.Text)
txtData.SelColor = vbBlue
txtData.SelText = txtNick & ":" & vbTab
txtData.SelStart = Len(txtData.Text)
txtData.SelColor = CommonDialog1.Color
txtData.SelText = txtText.Text & vbCrLf
txtText.Text = ""
If Err Then lblStatus.Caption = Err.Description
End Sub
Private Sub txtText_GotFocus()
cmdSend.Default = True
End Sub
Private Sub txtText_LostFocus()
cmdSend.Default = False
End Sub
Private Sub wData_Close()
wData.Close
wData.Listen
lblStatus.Caption = "Listening..."
cmdDisconnect.Visible = False
End Sub
Private Sub wData_Connect()
lblStatus.Caption = "Connected!"
End Sub
Private Sub wData_ConnectionRequest(ByVal requestID As Long)
wData.Close
wData.Accept requestID
lblStatus.Caption = "Connection Accepted!"
End Sub
Private Sub wData_DataArrival(ByVal bytesTotal As Long)
Dim nData As String
On Error Resume Next
wData.GetData nData
txtData.SelStart = Len(txtData.Text)
txtData.SelColor = vbRed
txtData.SelText = Left(nData, InStr(1, nData, ":"))
txtData.SelStart = Len(txtData.Text)
txtData.SelColor = vbBalck
txtData.SelText = Mid(nData, InStr(1, nData, ":") + 1) & vbCrLf
If Err Then lblStatus.Caption = Err.Description
End Sub
Private Sub wData_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
lblStatus.Caption = Description
End Sub
Private Sub lblStatus_Change()
If Check1.Enabled = True Then GoTo SoundError Else
Select Case lblStatus.Caption
Case "Connection Accepted!"
Case "Listening..."
Case "Connected!"
Case "Connection Closed"
Case "Connecting..."
Case "Disconnected!"
End Select
If lblStatus.Caption = "Connection Accepted!" Then sndPlaySound "c:\Program Files\LanChat32\Tada.wav", 0
If lblStatus.Caption = "Disconnected!" Then sndPlaySound "c:\Program Files\LanChat32\Ding.wav", 0
If lblStatus.Caption = "Connection Closed" Then sndPlaySound "c:\Program Files\LanChat32\chord.wav", 0
If frmMain.WindowState = 1 Then
frmMain.WindowState = 0
Else
End If
SoundError:
If lblStatus.Caption = "Listening..." Then cmdDisconnect.Visible = True
If lblStatus.Caption = "Connection Closed" Then cmdConnect.Visible = True
End Sub
Private Sub txtData_Change()
If Check1.Enabled = True Then GoTo SoundError1
sndPlaySound "c:\Program Files\LanChat32\Chimes.wav", 1
If frmMain.WindowState = 1 Then
frmMain.WindowState = 0
Else
End If
SoundError1:
End Sub
Private Sub Command6_Click()
If Check3.Value = 0 Then
AlwaysOnTop frmMain, True
Check3.Value = 1
Else
If Check3.Value = 1 Then
AlwaysOnTop frmMain, False
Check3.Value = 0
End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Form1
Unload Form2
End
End Sub

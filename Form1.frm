VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Lan Chat32"
   ClientHeight    =   3195
   ClientLeft      =   3495
   ClientTop       =   3615
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4905
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   2880
      Width           =   975
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Hyperswede"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   0
      Width           =   2175
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Lan Chat 32"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   2175
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   855
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Starfield By:"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   2520
      Width           =   975
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Thomas Hill,"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Arthur Chaparyan"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   855
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Written By:"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   360
      Top             =   480
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Thomas Baker,"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim starX(0 To 100) As Double
Dim starY(0 To 100) As Double
Dim starDist(0 To 100) As Double
Dim starSpeed As Double
Dim formMidX As Double
Dim formMidY As Double
Private Sub Form_KeyPress(KeyAscii As Integer)
End
End Sub
Private Sub Form_Load()
Randomize
frmAbout.BackColor = &H0&
frmAbout.ForeColor = &HFFFFFF
frmAbout.FillColor = &HFFFFFF
frmAbout.FillStyle = 0
frmAbout.DrawWidth = 2
formMidX = (frmAbout.Width / 2)
formMidY = (frmAbout.Height / 2)
For X = 0 To 100
Do
starX(X) = Int(Rnd * frmAbout.Width)
starY(X) = Int(Rnd * frmAbout.Height)
starDist(X) = Int(Rnd * 5)
Loop While (starX(X) = formMidY And starY(Y) = formMidY)
starDist(X) = 0
Next X
starSpeed = 0.025
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
formMidX = X
formMidY = Y
End Sub
Private Sub Text2_Click()
End Sub
Private Sub Text3_Click()
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload frmAbout
MsgBox "Special Thanx To: ToidyMan, CorpWhore, Bebabo, AllSystemsGo and anyone else I have missed.  If you do not see your name here and you should, please e-mail me so I can add it.  Thanx for trying Lan Chat 32!", vbOKOnly + vbExclamation, "Thanx!"
wrap$ = Chr$(10) + Chr$(13)
End Sub
Private Sub Timer1_Timer()
For X = 0 To 100
frmAbout.FillColor = frmAbout.BackColor
Circle (starX(X), starY(X)), starDist(X), BackColor
starDist(X) = starDist(X) + 0.1
If starX(X) > (formMidX) Then
starX(X) = starX(X) + Int(Abs(formMidX - starX(X)) * starSpeed) * (starDist(X) * 0.2)
Else
starX(X) = starX(X) - Int(Abs(formMidX - starX(X)) * starSpeed) * (starDist(X) * 0.2)
End If
If starY(X) > (formMidY) Then
starY(X) = starY(X) + Int(Abs(formMidY - starY(X)) * starSpeed) * (starDist(X) * 0.2)
Else
starY(X) = starY(X) - Int(Abs(formMidY - starY(X)) * starSpeed) * (starDist(X) * 0.2)
End If
If starX(X) > frmAbout.Width Or starX(X) < 0 Or starY(X) > frmAbout.Height Or starY(X) < 0 Then
Do
starX(X) = Int(Rnd * frmAbout.Width)
starY(X) = Int(Rnd * frmAbout.Height)
Loop While (starX(X) = formMidX Or starY(Y) = formMidY)
starDist(X) = 1
End If
If starDist(X) > 30 Then
starDist(X) = 1
starX(X) = Int(Rnd * frmAbout.Width)
starY(X) = Int(Rnd * frmAbout.Height)
End If
frmAbout.FillColor = &HFFFFFF
Circle (starX(X), starY(X)), starDist(X)
Next X
End Sub

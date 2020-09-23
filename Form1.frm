VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0218EB5B-AB19-4D19-885A-EB44273A08F0}#1.0#0"; "iSoftFPL01md.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin Fade_Player.FadePlayer FadePlayer1 
      Left            =   3480
      Top             =   4320
      _ExtentX        =   873
      _ExtentY        =   873
      PosChangeInterval=   100
   End
   Begin VB.CommandButton Command3 
      Caption         =   "[ ]"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "| |"
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "|>"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   5280
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   327682
      Max             =   100
   End
   Begin VB.FileListBox File2 
      Height          =   3405
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Change This to the path where your music is located:
Const mPath = "c:\backups\backup2\music\"
'''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub Command1_Click()
FadePlayer1.fPlay
End Sub

Private Sub Command2_Click()
FadePlayer1.fPause
End Sub

Private Sub Command3_Click()
FadePlayer1.fStop
End Sub

Private Sub File1_Click()
FadePlayer1.FileName = File1.Path & "\" & File1.FileName
File2.ListIndex = File1.ListIndex
End Sub


Private Sub File2_Click()
FadePlayer1.QFileName = File2.Path & "\" & File2.FileName
End Sub

Private Sub Form_Load()
File1.Path = mPath
File2.Path = mPath
End Sub
Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1.Enabled = False
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
FadePlayer1.CurrentPosition = Slider1.Value
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Slider1 = FadePlayer1.CurrentPosition
Me.Caption = Int(FadePlayer1.CurrentTime) & "  -  " & Int(FadePlayer1.Duration)
End Sub





VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H80000012&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dragon Fire"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10230
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MMControl1 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   360
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   5640
      Top             =   4920
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Picture         =   "frmStart.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      Picture         =   "frmStart.frx":8404
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   1935
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   2010
      ItemData        =   "frmStart.frx":107FC
      Left            =   5280
      List            =   "frmStart.frx":107FE
      TabIndex        =   0
      Top             =   2160
      Width           =   4815
   End
   Begin VB.Image Image3 
      Height          =   1275
      Index           =   0
      Left            =   2280
      Picture         =   "frmStart.frx":10800
      Top             =   120
      Width           =   5865
   End
   Begin VB.Image Image2 
      Height          =   615
      Index           =   0
      Left            =   240
      Picture         =   "frmStart.frx":182D6
      Top             =   1800
      Width           =   4860
   End
   Begin VB.Image Image1 
      Height          =   2565
      Left            =   960
      Picture         =   "frmStart.frx":1E22B
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   2955
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FrameCount As Long


Private Sub Command1_Click()
    'determine what mode the user choose
    Mode = List1.ListIndex + 1
    'end determining what mode the user choose
    'MMControl1.Command = "Stop"
    LoadMode (Mode)
End Sub

Private Sub Command2_Click()
    Mbox.Label1.Caption = "Exit?"
    Mbox.Label2.Caption = "Are you sure you want to exit Dragon Fire?"
    Mbox.Command1.Caption = "Yup..."
    Mbox.Show vbModal, Me
    If MBoxReturn = False Then Exit Sub
    If MBoxReturn = True Then End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Command2_Click
    End If
End Sub

Public Sub LoadGifs()
    Timer1.Enabled = False
  If LoadGif(App.Path & "\Images\mode.gif", Image2) Then
     FrameCount = 0
     Timer1.Interval = CLng(Image2(0).Tag)
     Timer1.Enabled = True
  End If

End Sub

Private Sub Form_Load()
    'populate the listbox
    List1.AddItem "Offline (1 Player)"
    List1.AddItem "Online (Requires Internet Connection) (Many Players)"
    List1.ListIndex = 0
    List1.Refresh
    'end of listbox functions
    
    'load all needed forms
    Load Mbox
    titleclose = False
    LoadGifs
    'MMControl1.Command = "Open"
    'MMControl1.Command = "Play"
End Sub

Private Sub List1_DblClick()
    Command1_Click
End Sub

Private Sub Timer1_Timer()
    If FrameCount < TotalFrames Then
        Image2(FrameCount).Visible = False
        FrameCount = FrameCount + 1
        Image2(FrameCount).Visible = True
        Timer1.Interval = CLng(Image2(FrameCount).Tag)
    Else
        FrameCount = 0
        For i = 1 To Image2.Count - 1
            Image2(i).Visible = False
        Next i
        Image2(FrameCount).Visible = True
        Timer1.Interval = CLng(Image2(FrameCount).Tag)
    End If
End Sub

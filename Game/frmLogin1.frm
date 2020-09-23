VERSION 5.00
Begin VB.Form frmLogin1 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1500
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   7185
   Icon            =   "frmLogin1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   886.25
   ScaleMode       =   0  'User
   ScaleWidth      =   6746.326
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   405
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   -240
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   1920
      Top             =   840
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   615
      Index           =   0
      Left            =   120
      Picture         =   "frmLogin1.frx":000C
      Top             =   120
      Width           =   6795
   End
End
Attribute VB_Name = "frmLogin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Mode As String
Public Account As String
Public Password As String
Dim FrameCount As Long


Private Sub Form_Load()
    LoadGifs
    Mode = "AN"
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

Public Sub LoadGifs()
    Timer1.Enabled = False
  If LoadGif(App.Path & "\Images\accountname.gif", Image2) Then
     FrameCount = 0
     Timer1.Interval = CLng(Image2(0).Tag)
     Timer1.Enabled = True
  End If
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Mode = "AN" Then
        txt.Locked = True
        LoadGifs2
        Account = txt.Text
        txt.Text = ""
        Mode = "PW"
        txt.Locked = False
        txt.SetFocus
        Exit Sub
    End If
    If KeyCode = vbKeyReturn And Mode = "PW" Then
        Password = txt.Text
        LoadGifs3
        DoEvents
        Player.Online.Account = Account
        Player.Online.Password = Password
        CheckPW
        Exit Sub
    End If
    
End Sub

Public Sub LoadGifs2()

          Timer1.Enabled = False
        If LoadGif(App.Path & "\Images\password.gif", Image2) Then
           FrameCount = 0
           Timer1.Interval = CLng(Image2(0).Tag)
           Timer1.Enabled = True
        End If
End Sub

Public Sub LoadGifs3()

          Timer1.Enabled = False
        If LoadGif(App.Path & "\Images\connecting.gif", Image2) Then
           FrameCount = 0
           Timer1.Interval = CLng(Image2(0).Tag)
           Timer1.Enabled = True
        End If

End Sub

Public Sub LoadInvalid()

          Timer1.Enabled = False
        If LoadGif(App.Path & "\Images\invalid.gif", Image2) Then
           FrameCount = 0
           Timer1.Interval = CLng(Image2(0).Tag)
           Timer1.Enabled = True
        End If

End Sub


Public Sub CheckPW()
    Dim Data, Data2 As String
    Dim i As String
    DoEvents
    Load frmDB
    Load frmMainOnline
    Data = "clientlogin:" & Player.Online.Account & "," & Player.Online.Password
    Data2 = Data
    'frmMainOnline.Sock.Connect
    frmMainOnline.Sock.SendData Data
    i = 0
    Do Until Data <> "" And Data <> ("clientlogin:" & Player.Online.Account & "," & Player.Online.Password)
        frmMainOnline.Sock.GetData Data
        DoEvents
        i = i + 1
        If i = 1001 Then i = 0: GoTo nextt
    Loop
nextt:
    frmSetStats2.SetFocus
    If Data = True Then
            Timer1.Enabled = False
            Me.Hide
            frmPatch.Show
            frmPatch.Refresh
    Else
        LoadInvalid
        Mode = "AN"
        txt.Text = "Accout Name?"
    End If
End Sub

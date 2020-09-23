VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "FrmMp3"
   ClientHeight    =   5835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   LinkTopic       =   "Form2"
   ScaleHeight     =   5835
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   4995
      TabIndex        =   2
      Top             =   5520
      Width           =   5055
      Begin VB.Label Label4 
         Caption         =   "Select Link to Visit"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   4575
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   120
      ScaleHeight     =   4395
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   480
      Width           =   4695
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.simplemp3s.com/newmp3s.htm"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3720
         Width           =   3615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.mp3hits.net"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3360
         Width           =   4215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "http://21century-mp3.com"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   3255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.dafreak.com/indispensable/hits.html"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   4455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "http://home.wanadoo.nl/musicinside"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   4095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.eurotipten.com/tips.htm"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   3975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "http://home.wanadoo.nl/~jan.empels/"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   3975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.mp3delivery.com"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.geocities.com/daarzijnweweer"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.newmp3.com/tophits.shtml"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Mp3  2 k net... (Offers New hits.)"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   " EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Below Are Some Good Mp3 Sites Ihave found WhileSurfing the net."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Label13_Click()
Unload Me
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShellExecute hwnd, "open", "http://63.217.28.192/new.php3", vbNullString, vbNullString, conSwNormal
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShellExecute hwnd, "open", "http://www.newmp3.com/tophits.shtml", vbNullString, vbNullString, conSwNormal
End Sub
Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShellExecute hwnd, "open", "http://www.geocities.com/daarzijnweweer", vbNullString, vbNullString, conSwNormal
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShellExecute hwnd, "open", "http://www.mp3delivery.com", vbNullString, vbNullString, conSwNormal
End Sub
Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShellExecute hwnd, "open", "http://home.wanadoo.nl/~jan.empels/", vbNullString, vbNullString, conSwNormal
End Sub
Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShellExecute hwnd, "open", "http://www.eurotipten.com/tips.htm", vbNullString, vbNullString, conSwNormal
End Sub
Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShellExecute hwnd, "open", "http://home.wanadoo.nl/musicinside", vbNullString, vbNullString, conSwNormal
End Sub
Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShellExecute hwnd, "open", "http://www.dafreak.com/indispensable/hits.html", vbNullString, vbNullString, conSwNormal
End Sub
Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShellExecute hwnd, "open", "http://21century-mp3.com", vbNullString, vbNullString, conSwNormal
End Sub
Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShellExecute hwnd, "open", "http://www.mp3hits.net", vbNullString, vbNullString, conSwNormal
End Sub
Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShellExecute hwnd, "open", "http://www.simplemp3s.com/newmp3s.htm", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = "Mp3 2k. Offers New hits. Updated daily"
Label2.ForeColor = &HFF&

End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = " Offers Euro hits and Us top 40"
Label3.ForeColor = &HFF&


End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = "New Hits "
Label14.ForeColor = &HFF&


End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = "New Mp3s. Dutch top 40. Us top40"
Label5.ForeColor = &HFF&


End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = "Huge Mp3 Archive and Lots of Lyrics"
Label6.ForeColor = &HFF&


End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = "Euro Tip Ten"
Label7.ForeColor = &HFF&


End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = "Latest Hits and Album Search"
Label8.ForeColor = &HFF&


End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = "New Releases, Summer hits, trance and club, full albums"
Label9.ForeColor = &HFF&


End Sub
Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = "New Releases,Dance Classix, trance and club, full albums"
Label10.ForeColor = &HFF&
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = "New Releases and Mp3 Archives"
Label11.ForeColor = &HFF&
End Sub
Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = "New mp3s"
Label12.ForeColor = &HFF&
End Sub

Private Sub Restore_Colors()
Label2.ForeColor = &HFF0000
Label3.ForeColor = &HFF0000
Label14.ForeColor = &HFF0000
Label5.ForeColor = &HFF0000
Label6.ForeColor = &HFF0000
Label7.ForeColor = &HFF0000
Label8.ForeColor = &HFF0000
Label9.ForeColor = &HFF0000
Label10.ForeColor = &HFF0000
Label11.ForeColor = &HFF0000
Label12.ForeColor = &HFF0000

End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Restore_Colors

End Sub

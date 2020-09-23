VERSION 5.00
Begin VB.Form frmDied 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "You Died...."
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   509
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Restart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Exit Program"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Died..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   6975
   End
End
Attribute VB_Name = "frmDied"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
On Error Resume Next
    frmStart.Show
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Image1.Picture = frmGraphics.PlayerDead.Picture
    frmDB.txtHealth.Text = frmDB.txtTotalHealth.Text
    Unload frmDummy
    Unload frmGraphics
    Unload frmLoading
    Unload frmMain
    Unload frmMainOnline
    Unload frmMusic
    Unload frmOptions
    Unload frmStart
    Unload Mbox
    Unload frmDB

End Sub


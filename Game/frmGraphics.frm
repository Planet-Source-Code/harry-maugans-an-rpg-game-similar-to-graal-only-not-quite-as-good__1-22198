VERSION 5.00
Begin VB.Form frmGraphics 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   352
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBlack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   4800
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image PlayerDown 
      Height          =   480
      Index           =   1
      Left            =   1440
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image PlayerUp 
      Height          =   480
      Index           =   1
      Left            =   960
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image PlayerRight 
      Height          =   480
      Index           =   1
      Left            =   480
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image PlayerLeft 
      Height          =   480
      Index           =   1
      Left            =   0
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image imgBush 
      Height          =   465
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   465
   End
   Begin VB.Image imgBomb 
      Height          =   465
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   960
      Width           =   465
   End
   Begin VB.Image imgArrow 
      Height          =   465
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   465
   End
   Begin VB.Image imgGold3 
      Height          =   495
      Left            =   2880
      Top             =   960
      Width           =   495
   End
   Begin VB.Image imgGold2 
      Height          =   495
      Left            =   2400
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image imgGold1 
      Height          =   495
      Left            =   2400
      Top             =   960
      Width           =   495
   End
   Begin VB.Image imgSwordLeft 
      Height          =   480
      Left            =   1920
      Top             =   960
      Width           =   480
   End
   Begin VB.Image imgSwordRight 
      Height          =   480
      Left            =   1440
      Top             =   960
      Width           =   480
   End
   Begin VB.Image imgSwordDown 
      Height          =   480
      Left            =   960
      Top             =   960
      Width           =   480
   End
   Begin VB.Image imgSwordUp 
      Height          =   480
      Left            =   480
      Top             =   960
      Width           =   480
   End
   Begin VB.Image PicBlank 
      Height          =   480
      Left            =   4800
      Picture         =   "frmGraphics.frx":0000
      Top             =   0
      Width           =   480
   End
   Begin VB.Image PlayerLeft 
      Height          =   480
      Index           =   0
      Left            =   0
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image PlayerRight 
      Height          =   480
      Index           =   0
      Left            =   480
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image PlayerUp 
      Height          =   480
      Index           =   0
      Left            =   960
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image PlayerDown 
      Height          =   480
      Index           =   0
      Left            =   1440
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image PlayerDead 
      Height          =   480
      Left            =   1920
      Top             =   1440
      Width           =   480
   End
End
Attribute VB_Name = "frmGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Picture1_Click()

End Sub

Private Sub Form_Load()
    imgGold1.Picture = LoadPicture(App.Path & "\Images\Gold1.gif")
    imgGold2.Picture = LoadPicture(App.Path & "\Images\Gold2.gif")
    imgGold3.Picture = LoadPicture(App.Path & "\Images\Gold3.gif")
    imgArrow.Picture = LoadPicture(App.Path & "\Images\Arrow.gif")
    imgBomb.Picture = LoadPicture(App.Path & "\Images\Bomb.gif")
    imgBush.Picture = LoadPicture(App.Path & "\Images\Bush.gif")
End Sub

VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "MP3 Player"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   4455
   End
   Begin VB.ComboBox cboScreen 
      Height          =   315
      ItemData        =   "frmOptions.frx":0000
      Left            =   2640
      List            =   "frmOptions.frx":000A
      TabIndex        =   10
      Text            =   "Yes"
      Top             =   600
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmOptions.frx":0017
      Left            =   2640
      List            =   "frmOptions.frx":0021
      TabIndex        =   8
      Text            =   "On"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmOptions.frx":002E
      Left            =   2640
      List            =   "frmOptions.frx":0038
      TabIndex        =   6
      Text            =   "On"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox txtScrollSpeed 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Text            =   "2"
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Screen Transitions:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Background Music:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Sound FX:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Scroll Speed (# of tiles per frame):"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "OPTIONS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If IsNumeric(txtScrollSpeed.Text) Then
        Open App.Path & "\Options.txt" For Output As #1
            Print #1, txtScrollSpeed.Text
            Print #1, Combo1.Text
            Print #1, Combo2.Text
            Print #1, cboScreen.Text
        Close #1
        If Combo2.Text = "No" Then frmMusic.PlayMusic ""
        Unload Me
    Else
        MsgBox "Please enter a numeric value for the scroll speed.", vbCritical, "Error"
        txtScrollSpeed.SetFocus
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Shell App.Path & "\MP3 Player\MP3.exe", vbNormalFocus
End Sub

Private Sub Form_Load()
    Dim ScrollSpeed As Integer
    Dim SoundFX As String
    Dim Music As String
    Dim ST As String
    
    Open App.Path & "\Options.txt" For Input As #1
        Input #1, ScrollSpeed
        Input #1, SoundFX
        Input #1, Music
        Input #1, ST
    Close #1
    txtScrollSpeed.Text = ScrollSpeed
    Combo1.Text = SoundFX
    Combo2.Text = Music
    cboScreen.Text = ST
End Sub

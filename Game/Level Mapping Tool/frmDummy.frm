VERSION 5.00
Begin VB.Form frmDummy 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dummy Buffer Form"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   448
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   360
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Image img 
      Height          =   1935
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim i As Integer
    
    Image1.Height = (20 * 8) * 32
    Image1.Width = (14 * 8) * 32
    i = 1
    Do Until i >= 65
        Load img(i)
        i = i + 1
    Loop
End Sub

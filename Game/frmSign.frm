VERSION 5.00
Begin VB.Form frmSign 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "Sign"
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   Picture         =   "frmSign.frx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   94800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1560
      Width           =   150
   End
   Begin VB.Label txtSign 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSign.frx":12E62
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   1305
      Left            =   240
      TabIndex        =   1
      Tag             =   "200 chr"
      Top             =   240
      Width           =   9000
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SignData As String
Dim Pages As Integer
Dim CPage As Integer

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Xme, Yme As Integer
    Me.Width = frmMain.Width
    Me.Left = frmMain.Left
    Me.Top = (frmMain.Top + frmMain.Height) - Me.Height
    
    SignData = frmMain.SignData

    Pages = Len(SignData) / 100
    If Len(SignData) > 100 * Pages Then Pages = Pages + 1
    If Pages <= 1 Then
        txtSign.Caption = SignData
        CPage = 1
        Exit Sub
    Else
        CPage = 1
        txtSign.Caption = Left(SignData, 100)
    End If
End Sub

Private Sub txtSign_Click()
    Unload Me
End Sub

Private Sub txtSign_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        If KeyCode = vbKeyUp Then
            If CPage <= 1 Then Exit Sub
            CPage = CPage - 1
            txtSign.Caption = Mid(SignData, (CPage - 1) * 100, 100)
        End If
        If KeyCode = vbKeyDown Then
            If CPage >= Pages Then Exit Sub
            CPage = CPage + 1
            txtSign.Caption = Mid(SignData, (CPage - 1) * 100, 100)
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub txt_Click()
    Unload Me
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        If KeyCode = vbKeyUp Then
            If CPage <= 1 Then Exit Sub
            CPage = CPage - 1
            txtSign.Caption = Mid(SignData, ((CPage - 1) * 100) + 1, 100)
        End If
        If KeyCode = vbKeyDown Then
            If CPage >= Pages Then Exit Sub
            CPage = CPage + 1
            txtSign.Caption = Mid(SignData, ((CPage - 1) * 100) + 1, 100)
        End If
    Else
        Unload Me
    End If
End Sub


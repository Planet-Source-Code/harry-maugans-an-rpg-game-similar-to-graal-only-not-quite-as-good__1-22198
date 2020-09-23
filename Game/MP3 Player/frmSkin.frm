VERSION 5.00
Begin VB.Form frmSkin 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3315
   ClientLeft      =   5625
   ClientTop       =   4095
   ClientWidth     =   4920
   LinkTopic       =   "Form2"
   ScaleHeight     =   3315
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option5 
      BackColor       =   &H00000000&
      Caption         =   "Skin5 (Iron)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtLoad 
      Height          =   285
      Left            =   3360
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   4875
      TabIndex        =   10
      Top             =   3000
      Width           =   4935
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose a skin from above and click ok or Apply"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   3375
      End
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00000000&
      Caption         =   "Skin4 (RED)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00000000&
      Caption         =   "Skin3 (spherized)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Skin2(Blue)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Default"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Skin Browser"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   960
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "OK"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "APPLY"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "CANCEL"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   " Skins Browser"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeyKey As String      'for ini
Dim KeyValue As String     'for ini

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.Caption = "Choose a skin and then click Ok or Apply"

End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Label2_Click()
frmList.Show
Equalizer.Show
frmList.mnuEqualizer.Checked = True
frmList.mnuPlaylist.Checked = True
If Option1.Value = True Then
frmMain.Picture = LoadPicture(App.Path & "\" & "main.jpg")
Equalizer.Picture = LoadPicture(App.Path & "\" & "main2.jpg")
frmList.Picture = LoadPicture(App.Path & "\" & "main3.jpg")
'===Here it saves the skin in the ini file so the next time you run
'the program it will load it as it was when you ended it.==========
txtLoad.Text = "default"
KeySection = "Load Default"
KeyKey = "load"
KeyValue = txtLoad.Text
saveini
'=============end save ini

ElseIf Option2.Value = True Then
frmMain.Picture = LoadPicture(App.Path & "\" & "mainblue.jpg")
Equalizer.Picture = LoadPicture(App.Path & "\" & "main2blue.jpg")
frmList.Picture = LoadPicture(App.Path & "\" & "main3blue.jpg")
txtLoad.Text = "blue"
KeySection = "Load Default"
KeyKey = "load"
KeyValue = txtLoad.Text
saveini
ElseIf Option3.Value = True Then
frmMain.Picture = LoadPicture(App.Path & "\" & "mainL.jpg")
Equalizer.Picture = LoadPicture(App.Path & "\" & "main2L.jpg")
frmList.Picture = LoadPicture(App.Path & "\" & "main3L.jpg")
txtLoad.Text = "L"
KeySection = "Load Default"
KeyKey = "load"
KeyValue = txtLoad.Text
saveini
ElseIf Option4.Value = True Then
frmMain.Picture = LoadPicture(App.Path & "\" & "mainred.jpg")
Equalizer.Picture = LoadPicture(App.Path & "\" & "main2red.jpg")
frmList.Picture = LoadPicture(App.Path & "\" & "main3red.jpg")
txtLoad.Text = "red"
KeySection = "Load Default"
KeyKey = "load"
KeyValue = txtLoad.Text
saveini
ElseIf Option5.Value = True Then
frmMain.Picture = LoadPicture(App.Path & "\" & "mainIron.jpg")
Equalizer.Picture = LoadPicture(App.Path & "\" & "main2Iron.jpg")
frmList.Picture = LoadPicture(App.Path & "\" & "main3Iron.jpg")
txtLoad.Text = "Iron"
KeySection = "Load Default"
KeyKey = "load"
KeyValue = txtLoad.Text
saveini

End If
Unload Me
End Sub


Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &H808080
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackColor = &H404040
End Sub

Private Sub Label3_Click()
Unload Me
End Sub



Private Sub Label4_Click()
Unload Me
End Sub

Private Sub Label5_Click()
frmList.Show
Equalizer.Show
frmList.mnuPlaylist.Checked = True
frmList.mnuEqualizer.Checked = True
If Option1.Value = True Then
frmMain.Picture = LoadPicture(App.Path & "\" & "main.jpg")
Equalizer.Picture = LoadPicture(App.Path & "\" & "main2.jpg")
frmList.Picture = LoadPicture(App.Path & "\" & "main3.jpg")
'===Here it saves the skin in the ini file so the next time you run
'the program it will load it as it was when you ended it.==========
txtLoad.Text = "default"
KeySection = "Load Default"
KeyKey = "load"
KeyValue = txtLoad.Text
saveini
'=============end save ini

ElseIf Option2.Value = True Then
frmMain.Picture = LoadPicture(App.Path & "\" & "mainblue.jpg")
Equalizer.Picture = LoadPicture(App.Path & "\" & "main2blue.jpg")
frmList.Picture = LoadPicture(App.Path & "\" & "main3blue.jpg")
txtLoad.Text = "blue"
KeySection = "Load Default"
KeyKey = "load"
KeyValue = txtLoad.Text
saveini
ElseIf Option3.Value = True Then
frmMain.Picture = LoadPicture(App.Path & "\" & "mainL.jpg")
Equalizer.Picture = LoadPicture(App.Path & "\" & "main2L.jpg")
frmList.Picture = LoadPicture(App.Path & "\" & "main3L.jpg")
txtLoad.Text = "L"
KeySection = "Load Default"
KeyKey = "load"
KeyValue = txtLoad.Text
saveini
ElseIf Option4.Value = True Then
frmMain.Picture = LoadPicture(App.Path & "\" & "mainred.jpg")
Equalizer.Picture = LoadPicture(App.Path & "\" & "main2red.jpg")
frmList.Picture = LoadPicture(App.Path & "\" & "main3red.jpg")
txtLoad.Text = "red"
KeySection = "Load Default"
KeyKey = "load"
KeyValue = txtLoad.Text
saveini
ElseIf Option5.Value = True Then
frmMain.Picture = LoadPicture(App.Path & "\" & "mainIron.jpg")
Equalizer.Picture = LoadPicture(App.Path & "\" & "main2Iron.jpg")
frmList.Picture = LoadPicture(App.Path & "\" & "main3Iron.jpg")
txtLoad.Text = "Iron"
KeySection = "Load Default"
KeyKey = "load"
KeyValue = txtLoad.Text
saveini

End If

End Sub



Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackColor = &H808080
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackColor = &H404040
End Sub

Private Sub Option1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.Caption = "Default Skin"

End Sub

Private Sub Option2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.Caption = "BLUE SKIN"
End Sub

Private Sub Option3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.Caption = "Spherized"

End Sub
Private Sub loadini()

Dim lngResult As Long
Dim strFileName
Dim strResult As String * 50
strFileName = App.Path & "\Myini.ini" 'Declare your ini file !
lngResult = GetPrivateProfileString(KeySection, _
KeyKey, strFileName, strResult, Len(strResult), _
strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
Else
KeyValue = Trim(strResult)
End If

End Sub


Private Sub saveini()

Dim lngResult As Long
Dim strFileName
strFileName = App.Path & "\Myini.ini" 'Declare your ini file !
lngResult = WritePrivateProfileString(KeySection, _
KeyKey, KeyValue, strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
End If

End Sub



Private Sub Option4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.Caption = "RED"
End Sub


Private Sub Option5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.Caption = "Iron"
End Sub

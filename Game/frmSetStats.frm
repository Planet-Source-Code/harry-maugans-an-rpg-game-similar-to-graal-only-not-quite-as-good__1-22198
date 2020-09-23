VERSION 5.00
Begin VB.Form frmSetStats 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Load or Create a Character"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   388
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   187
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtName 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "frmSetStats.frx":0000
      Left            =   840
      List            =   "frmSetStats.frx":0002
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Continue"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   1620
      ItemData        =   "frmSetStats.frx":0004
      Left            =   120
      List            =   "frmSetStats.frx":0006
      TabIndex        =   2
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   2640
      Left            =   120
      Picture         =   "frmSetStats.frx":0008
      Top             =   0
      Width           =   2565
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Left            =   2160
      Top             =   4200
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Preview:"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Character Type:"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Name:"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   495
   End
End
Attribute VB_Name = "frmSetStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim Temp As String
    
    If txtName.Text = Trim("") Then
        MsgBox "Please enter a name.", , "Error"
        Exit Sub
    End If
    
    Player.PlayerType = List1.Text
    Player.Name = txtName.Text
    
    Open App.Path & "\Nicknames.txt" For Input As #1
        Line Input #1, Temp
        Line Input #1, Temp
        Do Until EOF(1)
            Line Input #1, Temp
            If Temp = txtName.Text Then
                GoTo skipit
            End If
        Loop
        Temp = "Nope1234x"
skipit:
    Close #1
    If Temp = "Nope1234x" Then
        Open App.Path & "\Nicknames.txt" For Append As #1
            Print #1, txtName.Text
        Close #1
    End If
    
    Open App.Path & "\Nicknames.txt" For Input As #1
    Open App.Path & "\Temp1.txt" For Output As #2
        Line Input #1, Temp
        Print #2, txtName.Text
        Do Until EOF(1)
            Line Input #1, Temp
            Print #2, Temp
        Loop
    Close #1
    Close #2
        Open App.Path & "\Nicknames.txt" For Output As #1
        Open App.Path & "\Temp1.txt" For Input As #2
            Do Until EOF(2)
                Line Input #2, Temp
                Print #1, Temp
            Loop
        Close #1
        Close #2
    Kill App.Path & "\Temp1.txt"
    
    'Load the data from database
    Dim i As Integer
    
    frmDB.RS.MoveLast
    i = frmDB.RS!ID
    
    If frmDB.FindUsername(txtName.Text) = False Then
        frmDB.RS.AddNew
        frmDB.RS!AccessPriviledges = "Normal"
        frmDB.RS!Banned = "False"
        frmDB.RS!Deaths = "0"
        frmDB.RS!Gold = "0"
        frmDB.RS!Health = "10"
        frmDB.RS!ID = i + 1
        frmDB.RS!IdleTime = "0:0:0"
        frmDB.RS!Kills = "0"
        frmDB.RS!LastSignon = Date
        frmDB.RS!Level = "DemoMap.map"
        frmDB.RS!Password = ""
        frmDB.RS!PlayerX = "2"
        frmDB.RS!PlayerY = "4"
        frmDB.RS!Username = txtName.Text
        frmDB.RS.Update
        frmDB.LoadData
    Else
        frmDB.LoadDataNow
    End If
    'Done with the database stuff
    
    Unload Me
    frmMain.Show
End Sub

Private Sub Form_Load()
    Dim Temp As String
    
    Load frmGraphics
    
    Open App.Path & "\PlayerSet.txt" For Input As #1
        Do Until EOF(1)
            Line Input #1, Temp
            Open App.Path & "\Temp4.txt" For Output As #2
                Print #2, Temp
            Close #2
            Open App.Path & "\Temp4.txt" For Input As #2
                Input #2, Temp
                List1.AddItem Temp
            Close #2
            Kill App.Path & "\Temp4.txt"
        Loop
    Close #1
    If List1.ListCount >= 1 Then
        List1.Selected(0) = True
        List1_Click
    End If
    
    Open App.Path & "\Nicknames.txt" For Input As #1
        Line Input #1, Temp
        txtName.Text = Temp
        Line Input #1, Temp
        Do Until EOF(1)
            Line Input #1, Temp
            txtName.AddItem Temp
        Loop
    Close #1
    
    Load frmDB
End Sub



Private Sub List1_Click()
Dim FoundIt As Boolean

        FoundIt = False
        Open App.Path & "\PlayerSet.txt" For Input As #1
        Do Until EOF(1)
            Line Input #1, Temp
            Open App.Path & "\Temp4.txt" For Output As #2
                Print #2, Temp
            Close #2
            Open App.Path & "\Temp4.txt" For Input As #2
                Input #2, Temp
                If Temp = List1.Text Then
                    Input #2, Temp
                    frmGraphics.PlayerUp(0).Picture = LoadPicture(App.Path & "\Images\Player Images\" & Temp)
                    Input #2, Temp
                    frmGraphics.PlayerUp(1).Picture = LoadPicture(App.Path & "\Images\Player Images\" & Temp)
                    Input #2, Temp
                    frmGraphics.PlayerDown(0).Picture = LoadPicture(App.Path & "\Images\Player Images\" & Temp)
                    Image1.Picture = frmGraphics.PlayerDown(0).Picture
                    Input #2, Temp
                    frmGraphics.PlayerDown(1).Picture = LoadPicture(App.Path & "\Images\Player Images\" & Temp)
                    Input #2, Temp
                    frmGraphics.PlayerRight(0).Picture = LoadPicture(App.Path & "\Images\Player Images\" & Temp)
                    Input #2, Temp
                    frmGraphics.PlayerRight(1).Picture = LoadPicture(App.Path & "\Images\Player Images\" & Temp)
                    Input #2, Temp
                    frmGraphics.PlayerLeft(0).Picture = LoadPicture(App.Path & "\Images\Player Images\" & Temp)
                    Input #2, Temp
                    frmGraphics.PlayerLeft(1).Picture = LoadPicture(App.Path & "\Images\Player Images\" & Temp)
                    Input #2, Temp
                    frmGraphics.PlayerDead.Picture = LoadPicture(App.Path & "\Images\Player Images\" & Temp)
                    FoundIt = True
                End If
            Close #2
            Kill App.Path & "\Temp4.txt"
            If FoundIt = True Then GoTo mkay
        Loop
mkay:
    Close #1
End Sub

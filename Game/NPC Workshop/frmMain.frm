VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NPC Workshop"
   ClientHeight    =   10440
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11550
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txtScript 
      Height          =   5655
      Left            =   3240
      TabIndex        =   0
      Top             =   2880
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9975
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0442
   End
   Begin TabDlg.SSTab tbTemplates 
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   8640
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   2990
      _Version        =   393216
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "Bad Guy Templates"
      TabPicture(0)   =   "frmMain.frx":04FC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label12"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lstBadGuys"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lstPower"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lstSpeed"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lstWeapon"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Bartender Templates"
      TabPicture(1)   =   "frmMain.frx":0518
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Custom Template"
      TabPicture(2)   =   "frmMain.frx":0534
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.ListBox lstWeapon 
         Height          =   1035
         ItemData        =   "frmMain.frx":0550
         Left            =   9000
         List            =   "frmMain.frx":0557
         TabIndex        =   21
         Top             =   480
         Width           =   1575
      End
      Begin VB.ListBox lstSpeed 
         Height          =   1035
         ItemData        =   "frmMain.frx":0562
         Left            =   6600
         List            =   "frmMain.frx":0572
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
      Begin VB.ListBox lstPower 
         Height          =   1035
         ItemData        =   "frmMain.frx":05AB
         Left            =   3600
         List            =   "frmMain.frx":05C7
         TabIndex        =   17
         Top             =   480
         Width           =   1815
      End
      Begin VB.ListBox lstBadGuys 
         Height          =   1035
         ItemData        =   "frmMain.frx":0624
         Left            =   840
         List            =   "frmMain.frx":0637
         TabIndex        =   15
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Attacking Weapon:"
         Height          =   495
         Left            =   8160
         TabIndex        =   20
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Movement Speed:"
         Height          =   495
         Left            =   5760
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Damage:"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Image:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmMain.frx":0672
      Left            =   840
      List            =   "frmMain.frx":067C
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "No"
      Top             =   2925
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "NPC Info"
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   6120
      Width           =   3015
      Begin VB.ComboBox cboVisible 
         Height          =   315
         ItemData        =   "frmMain.frx":0689
         Left            =   1320
         List            =   "frmMain.frx":0693
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Yes"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "NPC Visible:"
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
         TabIndex        =   9
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "The name that shows up under it."
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Like: ""Bartender.npc"""
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "NPC Name:"
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Filename:"
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
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Stretch:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Scripting:"
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Image Preview:"
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
      TabIndex        =   1
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Image imgPreview 
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Left            =   120
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   2520
      Left            =   0
      Picture         =   "frmMain.frx":06A0
      Top             =   0
      Width           =   11580
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
    If Combo1.Text = "Yes" Then
        imgPreview.Stretch = True
    End If
    If Combo1.Text = "No" Then
        imgPreview.Stretch = False
    End If
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileOpen_Click()
    Dim FileName As String
    
    FileName = InputBox("Please type a filename to open: (ie. 'MyGuy.npc')", "NPC to Open")
    
    If Trim(FileName) = "" Then Exit Sub
    
    txtScript.Text = ""
    txtScript.LoadFile App.Path & "\..\NPCs\" & FileName
    
    txtScript.Text = Decrypt(txtScript.Text)
End Sub

Private Sub mnuFileSave_Click()
    If Trim(txtFile.Text) <> "" Then
        txtScript.Text = Encrypt(txtScript.Text)
        txtScript.SaveFile App.Path & "\..\NPCs\" & txtFile.Text
        txtScript.Text = Decrypt(txtScript.Text)
    Else
        MsgBox "Please enter a filename.", , "Filename Missing"
        txtFile.SetFocus
    End If
End Sub

Private Sub txtScript_Change()
  Dim sstring As String
  Dim FoundPos As Integer

    FoundPos = "0"
    sstring = "Name"
    
    FoundPos = txtScript.Find(sstring, FoundPos)
    If FoundPos >= 0 Then
        txtScript.SelStart = FoundPos
        txtScript.SelLength = Len(sstring)
        txtScript.SelColor = vbBlue
    End If

    FoundPos = "0"
    sstring = "Move"
    
    FoundPos = txtScript.Find(sstring, FoundPos)
    If FoundPos >= 0 Then
        txtScript.SelStart = FoundPos
        txtScript.SelLength = Len(sstring)
        txtScript.SelColor = vbBlue
    End If
    
    FoundPos = "0"
    sstring = "Take"
    
    FoundPos = txtScript.Find(sstring, FoundPos)
    If FoundPos >= 0 Then
        txtScript.SelStart = FoundPos
        txtScript.SelLength = Len(sstring)
        txtScript.SelColor = vbBlue
    End If
    
    FoundPos = "0"
    sstring = "Give"
    
    FoundPos = txtScript.Find(sstring, FoundPos)
    If FoundPos >= 0 Then
        txtScript.SelStart = FoundPos
        txtScript.SelLength = Len(sstring)
        txtScript.SelColor = vbBlue
    End If

    FoundPos = "0"
    sstring = "Set"
    
    FoundPos = txtScript.Find(sstring, FoundPos)
    If FoundPos >= 0 Then
        txtScript.SelStart = FoundPos
        txtScript.SelLength = Len(sstring)
        txtScript.SelColor = vbBlue
    End If
    

    
    FoundPos = "0"
    sstring = "Me"
    
    FoundPos = txtScript.Find(sstring, FoundPos)
    If FoundPos >= 0 Then
        txtScript.SelStart = FoundPos
        txtScript.SelLength = Len(sstring)
        txtScript.SelColor = vbRed
    End If
    
    FoundPos = "0"
    sstring = "Visible"
    
    FoundPos = txtScript.Find(sstring, FoundPos)
    If FoundPos >= 0 Then
        txtScript.SelStart = FoundPos
        txtScript.SelLength = Len(sstring)
        txtScript.SelColor = vbBlue
    End If

    FoundPos = "0"
    sstring = "True"
    
    FoundPos = txtScript.Find(sstring, FoundPos)
    If FoundPos >= 0 Then
        txtScript.SelStart = FoundPos
        txtScript.SelLength = Len(sstring)
        txtScript.SelColor = vbRed
    End If
    
    FoundPos = "0"
    sstring = "False"
    
    FoundPos = txtScript.Find(sstring, FoundPos)
    If FoundPos >= 0 Then
        txtScript.SelStart = FoundPos
        txtScript.SelLength = Len(sstring)
        txtScript.SelColor = vbRed
    End If
    
    FoundPos = "0"
    sstring = "("
    
    FoundPos = txtScript.Find(sstring, FoundPos)
    If FoundPos >= 0 Then
        txtScript.SelStart = FoundPos
        txtScript.SelLength = Len(sstring)
        txtScript.SelColor = &H80&
    End If
     
    FoundPos = "0"
    sstring = ")"
    
    FoundPos = txtScript.Find(sstring, FoundPos)
    If FoundPos >= 0 Then
        txtScript.SelStart = FoundPos
        txtScript.SelLength = Len(sstring)
        txtScript.SelColor = &H80&
    End If
    
    
    txtScript.SelStart = Len(txtScript.Text)
    txtScript.SelColor = vbBlack
    
End Sub

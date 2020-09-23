VERSION 5.00
Begin VB.Form frmIS 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inventory Select"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8325
   Icon            =   "frmIS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Inventory"
      ForeColor       =   &H0000FF00&
      Height          =   3615
      Left            =   240
      TabIndex        =   4
      Top             =   3960
      Width           =   7815
      Begin VB.ListBox lstInventory 
         Height          =   2985
         ItemData        =   "frmIS.frx":000C
         Left            =   240
         List            =   "frmIS.frx":000E
         TabIndex        =   5
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lblType 
         BackStyle       =   0  'Transparent
         Caption         =   "Type (Armor, Helmet, ect.)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label lblPower 
         BackStyle       =   0  'Transparent
         Caption         =   "Power"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label lblRareity 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rarity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   3960
         TabIndex        =   7
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Image imgSword 
      Height          =   2175
      Left            =   720
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1575
   End
   Begin VB.Image imgShield 
      Height          =   2175
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shield"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5760
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Armor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3120
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sword"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.Image imgArmor 
      Height          =   1335
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Helmet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3720
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Image imgHelmet 
      Height          =   855
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   360
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   2640
      Left            =   480
      Picture         =   "frmIS.frx":0010
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2040
   End
   Begin VB.Image Image2 
      Height          =   2640
      Left            =   5760
      Picture         =   "frmIS.frx":0C52
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2040
   End
   Begin VB.Image Image3 
      Height          =   1560
      Left            =   3360
      Picture         =   "frmIS.frx":1894
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1560
   End
   Begin VB.Image Image4 
      Height          =   1080
      Left            =   3600
      Picture         =   "frmIS.frx":24D6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "frmIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim i As Integer
    Dim Temp As String
    
    i = "-1"
    Do Until i = frmDB.lstInventory.ListCount
        i = i + 1
        lstInventory.AddItem frmDB.lstInventory.List(i)
    Loop
    
    'done loading the inventory, now load the current equipted items
    If frmDB.txtArmor.Text <> "" And frmDB.txtArmor.Text <> Chr(34) & " " & Chr(34) Then
        Open App.Path & "\Temper.txt" For Output As #1
            Print #1, frmDB.txtArmor.Text
        Close #1
        Open App.Path & "\Temper.txt" For Input As #1
            Input #1, Temp
            imgArmor.ToolTipText = Temp
            Input #1, Temp
            imgArmor.Tag = Temp
        Close #1
        Kill App.Path & "\Temper.txt"
        Open App.Path & "\Inventory\Armor.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Temp
            If Temp = imgArmor.ToolTipText Then
                    Input #1, Temp
                    Input #1, Temp
                    Input #1, Temp
                    Input #1, Temp
                    imgArmor.Picture = LoadPicture(App.Path & "\Images\Armor\" & Temp)
                    GoTo FoundIt6
            End If
        Loop
FoundIt6:
        Close #1
    End If
    
    If frmDB.txtShield.Text <> "" And frmDB.txtShield.Text <> Chr(34) & " " & Chr(34) Then
        Open App.Path & "\Temper.txt" For Output As #1
            Print #1, frmDB.txtShield.Text
        Close #1
        Open App.Path & "\Temper.txt" For Input As #1
            Input #1, Temp
            imgShield.ToolTipText = Temp
            Input #1, Temp
            imgShield.Tag = Temp
        Close #1
        Kill App.Path & "\Temper.txt"
        Open App.Path & "\Inventory\Shields.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Temp
            If Temp = imgShield.ToolTipText Then
                    Input #1, Temp
                    Input #1, Temp
                    Input #1, Temp
                    Input #1, Temp
                    imgShield.Picture = LoadPicture(App.Path & "\Images\Shields\" & Temp)
                GoTo FoundIt7
            End If
        Loop
FoundIt7:
        Close #1
    End If
    
    If frmDB.txtSword.Text <> "" And frmDB.txtSword.Text <> Chr(34) & " " & Chr(34) Then
        Open App.Path & "\Temper.txt" For Output As #1
            Print #1, frmDB.txtSword.Text
        Close #1
        Open App.Path & "\Temper.txt" For Input As #1
            Input #1, Temp
            imgSword.ToolTipText = Temp
            Input #1, Temp
            imgSword.Tag = Temp
        Close #1
        Kill App.Path & "\Temper.txt"
        Open App.Path & "\Inventory\Swords.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Temp
            If Temp = imgSword.ToolTipText Then
                Input #1, Temp
                Input #1, Temp
                Input #1, Temp
                imgSword.Picture = LoadPicture(App.Path & "\Images\Swords\" & Temp)
                Input #1, Temp
                GoTo FoundIt8
            End If
        Loop
FoundIt8:
        Close #1
    End If
        
    If frmDB.txtHelmet.Text <> "" And frmDB.txtHelmet.Text <> Chr(34) & " " & Chr(34) Then
        Open App.Path & "\Temper.txt" For Output As #1
            Print #1, frmDB.txtHelmet.Text
        Close #1
        Open App.Path & "\Temper.txt" For Input As #1
            Input #1, Temp
            imgHelmet.ToolTipText = Temp
            Input #1, Temp
            imgHelmet.Tag = Temp
        Close #1
        Kill App.Path & "\Temper.txt"
        Open App.Path & "\Inventory\Helmet.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Temp
            If Temp = imgHelmet.ToolTipText Then
                Input #1, Temp
                Input #1, Temp
                Input #1, Temp
                Input #1, Temp
                imgHelmet.Picture = LoadPicture(App.Path & "\Images\Helmets\" & Temp)
                GoTo FoundIt9
            End If
        Loop
FoundIt9:
        Close #1
    End If
    lstInventory_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If imgShield.Tag <> "" And imgShield.Tag <> "0" Then
        frmDB.txtShield.Text = imgShield.ToolTipText & "," & imgShield.Tag
    End If
    If imgSword.Tag <> "" And imgSword.Tag <> "0" Then
        frmDB.txtSword.Text = imgSword.ToolTipText & "," & imgSword.Tag
    End If
    If imgArmor.Tag <> "" And imgArmor.Tag <> "0" Then
        frmDB.txtArmor.Text = imgArmor.ToolTipText & "," & imgArmor.Tag
    End If
    If imgHelmet.Tag <> "" And imgHelmet.Tag <> "0" Then
        frmDB.txtHelmet.Text = imgHelmet.ToolTipText & "," & imgHelmet.Tag
    End If
    frmMain.Load_Resistances
    frmMain.Load_Sword
End Sub

Private Sub lstInventory_Click()
    Dim i As Integer
    Dim Temp As String

    If Right(lstInventory.Text, 6) = "-Sword" Then
    Open App.Path & "\Inventory\Swords.txt" For Input As #1
    Do Until EOF(1)
        Input #1, Temp
        If Temp = lstInventory.Text Then
            Input #1, Temp
            lblRareity.Caption = "Rarity: " & Temp
            Input #1, Temp
            lblPower.Caption = "Attack Power: " & Temp
            lblType.Caption = "Type: Sword"
            lblName.Caption = "Name: " & lstInventory.Text
            GoTo FoundIt1
        End If
    Loop
FoundIt1:
    Close #1
    End If
    If Right(lstInventory.Text, 7) = "-Shield" Then
    Open App.Path & "\Inventory\Shields.txt" For Input As #1
    Do Until EOF(1)
        Input #1, Temp
        If Temp = lstInventory.Text Then
            Input #1, Temp
            lblRareity.Caption = "Rarity: " & Temp
            Input #1, Temp
            lblPower.Caption = "Evasion: " & Temp & "%"
            lblType.Caption = "Type: Shield"
            lblName.Caption = "Name: " & lstInventory.Text
            GoTo FoundIt2
        End If
    Loop
FoundIt2:
    Close #1
    End If
    If Right(lstInventory.Text, 7) = "-Helmet" Then
    Open App.Path & "\Inventory\Helmets.txt" For Input As #1
    Do Until EOF(1)
        Input #1, Temp
        If Temp = lstInventory.Text Then
            Input #1, Temp
            lblRareity.Caption = "Rarity: " & Temp
            Input #1, Temp
            lblPower.Caption = "Defense: " & Temp
            lblType.Caption = "Type: Helmet"
            lblName.Caption = "Name: " & lstInventory.Text
            GoTo FoundIt3
        End If
    Loop
FoundIt3:
    Close #1
    End If
    If Right(lstInventory.Text, 6) = "-Armor" Then
    Open App.Path & "\Inventory\Armor.txt" For Input As #1
    Do Until EOF(1)
        Input #1, Temp
        If Temp = lstInventory.Text Then
            Input #1, Temp
            lblRareity.Caption = "Rarity: " & Temp
            Input #1, Temp
            lblPower.Caption = "Defense: " & Temp
            lblType.Caption = "Type: Armor"
            lblName.Caption = "Name: " & lstInventory.Text
            GoTo FoundIt4
        End If
    Loop
FoundIt4:
    Close #1
    End If
End Sub

Private Sub lstInventory_DblClick()
    Dim Temp As String
    
    If Right(lstInventory.Text, 7) = "-Shield" Then
        Open App.Path & "\Inventory\Shields.txt" For Input As #1
            Do Until EOF(1)
                Input #1, Temp
                If Temp = lstInventory.Text Then
                    Input #1, Temp
                    lblRareity.Caption = "Rarity: " & Temp
                    Input #1, Temp
                    lblPower.Caption = "Evasion: " & Temp & "%"
                    imgShield.Tag = Temp
                    lblType.Caption = "Type: Shield"
                    lblName.Caption = "Name: " & lstInventory.Text
                    imgShield.ToolTipText = lstInventory.Text
                    frmDB.txtShield.Text = lstInventory.Text
                    Input #1, Temp
                    Input #1, Temp
                    imgShield.Picture = LoadPicture(App.Path & "\Images\Shields\" & Temp)
                    GoTo FoundIt5
                End If
            Loop
FoundIt5:
        Close #1
    End If
    If Right(lstInventory.Text, 6) = "-Sword" Then
        Open App.Path & "\Inventory\Swords.txt" For Input As #1
            Do Until EOF(1)
                Input #1, Temp
                If Temp = lstInventory.Text Then
                    Input #1, Temp
                    lblRareity.Caption = "Rarity: " & Temp
                    Input #1, Temp
                    lblPower.Caption = "Attack Power: " & Temp
                    frmDB.txtSword.Text = lstInventory.Text & "," & Temp
                    imgSword.Tag = Temp
                    lblType.Caption = "Type: Sword"
                    lblName.Caption = "Name: " & lstInventory.Text
                    imgSword.ToolTipText = lstInventory.Text
                    Input #1, Temp
                    imgSword.Picture = LoadPicture(App.Path & "\Images\Swords\" & Temp)
                    Input #1, Temp
                    GoTo FoundIt10
                End If
            Loop
FoundIt10:
        Close #1
    End If
    If Right(lstInventory.Text, 6) = "-Armor" Then
        Open App.Path & "\Inventory\Armor.txt" For Input As #1
            Do Until EOF(1)
                Input #1, Temp
                If Temp = lstInventory.Text Then
                    Input #1, Temp
                    lblRareity.Caption = "Rarity: " & Temp
                    Input #1, Temp
                    lblPower.Caption = "Defense: " & Temp
                    imgSword.Tag = Temp
                    lblType.Caption = "Type: Armor"
                    lblName.Caption = "Name: " & lstInventory.Text
                    imgArmor.ToolTipText = lstInventory.Text
                    frmDB.txtArmor.Text = lstInventory.Text
                    Input #1, Temp
                    Input #1, Temp
                    imgArmor.Picture = LoadPicture(App.Path & "\Images\Armor\" & Temp)
                    GoTo FoundIt11
                End If
            Loop
FoundIt11:
        Close #1
    End If
    If Right(lstInventory.Text, 7) = "-Helmet" Then
        Open App.Path & "\Inventory\Helmets.txt" For Input As #1
            Do Until EOF(1)
                Input #1, Temp
                If Temp = lstInventory.Text Then
                    Input #1, Temp
                    lblRareity.Caption = "Rarity: " & Temp
                    Input #1, Temp
                    lblPower.Caption = "Defense: " & Temp
                    imgSword.Tag = Temp
                    lblType.Caption = "Type: Helmet"
                    lblName.Caption = "Name: " & lstInventory.Text
                    imgHelmet.ToolTipText = lstInventory.Text
                    frmDB.txtHelmet.Text = lstInventory.Text
                    Input #1, Temp
                    Input #1, Temp
                    imgHelmet.Picture = LoadPicture(App.Path & "\Images\Helmets\" & Temp)
                    GoTo FoundIt12
                End If
            Loop
FoundIt12:
        Close #1
    End If
End Sub

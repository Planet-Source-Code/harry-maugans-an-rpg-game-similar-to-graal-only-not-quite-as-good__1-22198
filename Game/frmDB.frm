VERSION 5.00
Begin VB.Form frmDB 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Database Form"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBombs 
      Height          =   285
      Left            =   7560
      TabIndex        =   51
      Text            =   "txtBombs"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtArrows 
      Height          =   285
      Left            =   7560
      TabIndex        =   49
      Text            =   "txtArrows"
      Top             =   2040
      Width           =   1935
   End
   Begin VB.ListBox lstInventory 
      Height          =   645
      Left            =   1320
      TabIndex        =   47
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txtHelmet 
      Height          =   285
      Left            =   4440
      TabIndex        =   45
      Text            =   "txtHelmet"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtSword 
      Height          =   285
      Left            =   4440
      TabIndex        =   43
      Text            =   "txtSword"
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtShield 
      Height          =   285
      Left            =   4440
      TabIndex        =   41
      Text            =   "txtShield"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtArmor 
      Height          =   285
      Left            =   4440
      TabIndex        =   39
      Text            =   "txtArmor"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtActiveWeapon 
      Height          =   285
      Left            =   1320
      TabIndex        =   37
      Text            =   "txtActiveWeapon"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.ListBox lstWeapons 
      Height          =   645
      Left            =   1320
      TabIndex        =   35
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtWeaponCount 
      Height          =   285
      Left            =   1320
      TabIndex        =   33
      Text            =   "txtWeaponCount"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtTotalHealth 
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "txtTotalHealth"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "txtUsername"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "txtPassword"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "txtID"
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtOnlineTime 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "txtOnlineTime"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtLastSignon 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "txtLastSignon"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtIdleTime 
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "txtIdleTime"
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtKills 
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "txtKills"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtDeaths 
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "txtDeaths"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtGold 
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "txtGold"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtHealth 
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "txtHealth"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtLevel 
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "txtLevel"
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtPlayerX 
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "txtPlayerX"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtPlayerY 
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "txtPlayerY"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtBanned 
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "txtBanned"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtAccess 
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "txtAccess"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label26 
      Caption         =   "Bombs:"
      Height          =   255
      Left            =   6600
      TabIndex        =   50
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "Arrows:"
      Height          =   255
      Left            =   6600
      TabIndex        =   48
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "Inventory:"
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label23 
      Caption         =   "Helmet:"
      Height          =   255
      Left            =   3480
      TabIndex        =   44
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "Sword:"
      Height          =   255
      Left            =   3480
      TabIndex        =   42
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Shield:"
      Height          =   255
      Left            =   3480
      TabIndex        =   40
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Armor:"
      Height          =   255
      Left            =   3480
      TabIndex        =   38
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Active Weapon:"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Weapons:"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Weapon Count:"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Total Health:"
      Height          =   375
      Left            =   3480
      TabIndex        =   31
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Member ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Online Time:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Last Signon:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Idle Time:"
      Height          =   255
      Left            =   3480
      TabIndex        =   24
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Kills:"
      Height          =   255
      Left            =   3480
      TabIndex        =   23
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Deaths:"
      Height          =   255
      Left            =   3480
      TabIndex        =   22
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "Gold:"
      Height          =   255
      Left            =   3480
      TabIndex        =   21
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label16 
      Caption         =   "Health:"
      Height          =   255
      Left            =   3480
      TabIndex        =   20
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "Level:"
      Height          =   255
      Left            =   6600
      TabIndex        =   19
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label18 
      Caption         =   "Player X:"
      Height          =   255
      Left            =   6600
      TabIndex        =   18
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label19 
      Caption         =   "Player Y:"
      Height          =   255
      Left            =   6600
      TabIndex        =   17
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label20 
      Caption         =   "Banned:"
      Height          =   255
      Left            =   6600
      TabIndex        =   16
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label21 
      Caption         =   "Access Priviledges:"
      Height          =   495
      Left            =   6600
      TabIndex        =   15
      Top             =   1545
      Width           =   855
   End
End
Attribute VB_Name = "frmDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DB As Database
Public RS As Recordset

Private Sub Form_Load()
    Set DB = OpenDatabase(App.Path & "\Databases\membersDB.mdb")
    Set RS = DB.OpenRecordset("Members")
    RS.Edit
    RS.MoveFirst
    LoadData
End Sub

Public Sub LoadData()
    Dim Temp As String
    
    With Me
        .txtID.Text = RS!ID
        .txtUsername.Text = RS!Username
        .txtPassword.Text = RS!Password
        .txtOnlineTime.Text = RS!OnlineTime
        .txtLastSignon.Text = RS!LastSignon
        .txtIdleTime.Text = RS!IdleTime
        .txtKills.Text = RS!Kills
        .txtDeaths.Text = RS!Deaths
        .txtGold.Text = RS!Gold
        .txtArrows.Text = RS!Arrows
        .txtBombs.Text = RS!Bombs
        .txtHealth.Text = RS!Health
        .txtTotalHealth.Text = RS!TotalHealth
        .txtLevel.Text = RS!Level
        .txtPlayerX.Text = RS!PlayerX
        .txtPlayerY.Text = RS!PlayerY
        .txtBanned.Text = RS!Banned
        .txtAccess.Text = RS!AccessPriviledges
        .txtWeaponCount.Text = RS!WeaponCount
        .lstWeapons.Clear
        .lstWeapons.AddItem RS!Weapon1, 0
        .lstWeapons.AddItem RS!Weapon2, 1
        .lstWeapons.AddItem RS!Weapon3, 2
        .lstWeapons.AddItem RS!Weapon4, 3
        .lstWeapons.AddItem RS!Weapon5, 4
        .lstWeapons.AddItem RS!Weapon6, 5
        .lstWeapons.AddItem RS!Weapon7, 6
        .lstWeapons.AddItem RS!Weapon8, 7
        .lstWeapons.AddItem RS!Weapon9, 8
        .lstWeapons.AddItem RS!Weapon10, 9
        .lstWeapons.AddItem RS!Weapon11, 10
        .lstWeapons.AddItem RS!Weapon12, 11
        .lstWeapons.AddItem RS!Weapon13, 12
        .lstWeapons.AddItem RS!Weapon14, 13
        .lstWeapons.AddItem RS!Weapon15, 14
        .txtActiveWeapon.Text = RS!ActiveWeapon
        .txtArmor.Text = RS!Armor
        If .txtArmor.Text = Chr(34) & " " & Chr(34) Then .txtArmor.Text = ""
        .txtHelmet.Text = RS!Helmet
        If .txtHelmet.Text = Chr(34) & " " & Chr(34) Then .txtHelmet.Text = ""
        .txtSword.Text = RS!Sword
        If .txtSword.Text = Chr(34) & " " & Chr(34) Then .txtSword.Text = ""
        .txtShield.Text = RS!Shield
        If .txtShield.Text = Chr(34) & " " & Chr(34) Then .txtShield.Text = ""
        .lstInventory.Clear
        Open App.Path & "\TempX.txt" For Output As #6
            Print #6, RS!Inventory
        Close #6
        Open App.Path & "\TempX.txt" For Input As #6
            Do Until EOF(6)
                Input #6, Temp
                .lstInventory.AddItem Temp
            Loop
        Close #6
        Kill App.Path & "\TempX.txt"
        .lstInventory.Selected(0) = True
        If .lstInventory.Text = Chr(34) & " " & Chr(34) Then .lstInventory.RemoveItem (0)
    End With
End Sub

Public Sub LoadDataNow()
    On Error Resume Next
    Dim Temp As String
    
    With Me
        .txtID.Text = RS!ID
        .txtUsername.Text = RS!Username
        .txtPassword.Text = RS!Password
        .txtOnlineTime.Text = RS!OnlineTime
        .txtLastSignon.Text = RS!LastSignon
        .txtIdleTime.Text = RS!IdleTime
        .txtKills.Text = RS!Kills
        .txtDeaths.Text = RS!Deaths
        .txtGold.Text = RS!Gold
        .txtArrows.Text = RS!Arrows
        .txtBombs.Text = RS!Bombs
        .txtHealth.Text = RS!Health
        .txtTotalHealth.Text = RS!TotalHealth
        .txtLevel.Text = RS!Level
        .txtPlayerX.Text = RS!PlayerX
        .txtPlayerY.Text = RS!PlayerY
        .txtBanned.Text = RS!Banned
        .txtAccess.Text = RS!AccessPriviledges
        .txtWeaponCount.Text = RS!WeaponCount
        .lstWeapons.Clear
        .lstWeapons.AddItem RS!Weapon1, 0
        .lstWeapons.AddItem RS!Weapon2, 1
        .lstWeapons.AddItem RS!Weapon3, 2
        .lstWeapons.AddItem RS!Weapon4, 3
        .lstWeapons.AddItem RS!Weapon5, 4
        .lstWeapons.AddItem RS!Weapon6, 5
        .lstWeapons.AddItem RS!Weapon7, 6
        .lstWeapons.AddItem RS!Weapon8, 7
        .lstWeapons.AddItem RS!Weapon9, 8
        .lstWeapons.AddItem RS!Weapon10, 9
        .lstWeapons.AddItem RS!Weapon11, 10
        .lstWeapons.AddItem RS!Weapon12, 11
        .lstWeapons.AddItem RS!Weapon13, 12
        .lstWeapons.AddItem RS!Weapon14, 13
        .lstWeapons.AddItem RS!Weapon15, 14
        .txtActiveWeapon.Text = RS!ActiveWeapon
        .txtArmor.Text = RS!Armor
        If .txtArmor.Text = Chr(34) & " " & Chr(34) Then .txtArmor.Text = ""
        .txtHelmet.Text = RS!Helmet
        If .txtHelmet.Text = Chr(34) & " " & Chr(34) Then .txtHelmet.Text = ""
        .txtSword.Text = RS!Sword
        If .txtSword.Text = Chr(34) & " " & Chr(34) Then .txtSword.Text = ""
        .txtShield.Text = RS!Shield
        If .txtShield.Text = Chr(34) & " " & Chr(34) Then .txtShield.Text = ""
        .lstInventory.Clear
        Open App.Path & "\TempX.txt" For Output As #6
            Print #6, RS!Inventory
        Close #6
        Open App.Path & "\TempX.txt" For Input As #6
            Do Until EOF(6)
                Input #6, Temp
                .lstInventory.AddItem Temp
            Loop
        Close #6
        Kill App.Path & "\TempX.txt"
        .lstInventory.Selected(0) = True
        If .lstInventory.Text = Chr(34) & " " & Chr(34) Then .lstInventory.RemoveItem (0)
    End With
End Sub

Public Function FindUsername(Username As String) As Boolean
    On Error Resume Next
    Dim i As Integer
    Dim Temp As String
    
    i = 1
    
    With frmDB
            RS.MoveFirst
        Do Until i > RS.RecordCount
            .txtID.Text = RS!ID
            .txtUsername.Text = RS!Username
            .txtPassword.Text = RS!Password
            .txtOnlineTime.Text = RS!OnlineTime
            .txtLastSignon.Text = RS!LastSignon
            .txtIdleTime.Text = RS!IdleTime
            .txtKills.Text = RS!Kills
            .txtDeaths.Text = RS!Deaths
            .txtGold.Text = RS!Gold
            .txtArrows.Text = RS!Arrows
            .txtBombs.Text = RS!Bombs
            .txtHealth.Text = RS!Health
            .txtTotalHealth.Text = RS!TotalHealth
            .txtLevel.Text = RS!Level
            .txtPlayerX.Text = RS!PlayerX
            .txtPlayerY.Text = RS!PlayerY
            .txtBanned.Text = RS!Banned
            .txtAccess.Text = RS!AccessPriviledges
            .txtWeaponCount.Text = RS!WeaponCount
            .txtActiveWeapon.Text = RS!ActiveWeapon
            .txtArmor.Text = RS!Armor
            .txtHelmet.Text = RS!Helmet
            .txtSword.Text = RS!Sword
            .txtShield.Text = RS!Shield
            .lstInventory.Clear
            Open App.Path & "\TempX.txt" For Output As #6
                Print #6, RS!Inventory
            Close #6
            Open App.Path & "\TempX.txt" For Input As #6
                Do Until EOF(6)
                    Input #6, Temp
                    .lstInventory.AddItem Temp
                Loop
            Close #6
            Kill App.Path & "\TempX.txt"
            
            If (.txtUsername.Text = Username) Then
                FindUsername = True
                LoadData
                Exit Function
            End If
            
        i = i + 1
        If i < RS.RecordCount + 1 Then
            RS.MoveNext
        Else
        End If
        Loop
        
    End With
End Function


Public Function SaveData2()
Dim i As Integer
Dim Temp As String
'save data after frmMain is unloaded.
    With frmDB
            RS.Edit
            RS.MoveFirst
            Do Until RS!ID = .txtID.Text Or RS.EOF = True
                RS.MoveNext
            Loop
            RS.Edit
            RS!ID = .txtID.Text
            RS!Username = .txtUsername.Text
            RS!Password = .txtPassword.Text
            RS!OnlineTime = .txtOnlineTime.Text
            RS!LastSignon = .txtLastSignon.Text
            RS!IdleTime = .txtIdleTime.Text
            RS!Kills = .txtKills.Text
            RS!Deaths = .txtDeaths.Text
            RS!Gold = .txtGold.Text
            RS!Arrows = .txtArrows.Text
            RS!Bombs = .txtBombs.Text
            RS!Health = .txtHealth.Text
            RS!TotalHealth = .txtTotalHealth.Text
            RS!Level = .txtLevel.Text
            RS!PlayerX = .txtPlayerX.Text
            RS!PlayerY = .txtPlayerY.Text
            RS!Banned = .txtBanned.Text
            RS!AccessPriviledges = .txtAccess.Text
            RS!WeaponCount = .txtWeaponCount.Text
            RS!Weapon1 = .lstWeapons.List(0)
            RS!Weapon2 = .lstWeapons.List(1)
            RS!Weapon3 = .lstWeapons.List(2)
            RS!Weapon4 = .lstWeapons.List(3)
            RS!Weapon5 = .lstWeapons.List(4)
            RS!Weapon6 = .lstWeapons.List(5)
            RS!Weapon7 = .lstWeapons.List(6)
            RS!Weapon8 = .lstWeapons.List(7)
            RS!Weapon9 = .lstWeapons.List(8)
            RS!Weapon10 = .lstWeapons.List(9)
            RS!Weapon11 = .lstWeapons.List(10)
            RS!Weapon12 = .lstWeapons.List(11)
            RS!Weapon13 = .lstWeapons.List(12)
            RS!Weapon14 = .lstWeapons.List(13)
            RS!Weapon15 = .lstWeapons.List(14)
            RS!ActiveWeapon = .txtActiveWeapon.Text
            If .txtArmor.Text = "" Then .txtArmor.Text = Chr(34) & " " & Chr(34)
            RS!Armor = .txtArmor.Text
            If .txtHelmet.Text = "" Then .txtHelmet.Text = Chr(34) & " " & Chr(34)
            RS!Helmet = .txtHelmet.Text
            If .txtSword.Text = "" Then .txtSword.Text = Chr(34) & " " & Chr(34)
            RS!Sword = .txtSword.Text
            If .txtShield.Text = "" Then .txtShield.Text = Chr(34) & " " & Chr(34)
            RS!Shield = .txtShield.Text

            .lstInventory.Clear
            Open App.Path & "\TempX.txt" For Output As #6
                Do Until i = .lstInventory.ListCount + 1
                    Write #6, .lstInventory.List(i)
                    i = i + 1
                Loop
            Close #6
            Open App.Path & "\TempX.txt" For Input As #6
                Do Until EOF(6)
                    Input #6, Temp
                    .lstInventory.AddItem Temp
                Loop
            Close #6
            Kill App.Path & "\TempX.txt"
            
            RS.Update
    End With
    Exit Function
    
End Function

Public Function SaveData3()
Dim i As Integer
Dim Temp As String
'save data after frmMain is unloaded.
    With frmDB
            RS.Edit
            RS.MoveFirst
            Do Until RS!ID = .txtID.Text Or RS.EOF = True
                RS.MoveNext
            Loop
            RS.Edit
            RS!ID = .txtID.Text
            RS!Username = .txtUsername.Text
            RS!Password = .txtPassword.Text
            RS!OnlineTime = .txtOnlineTime.Text
            RS!LastSignon = .txtLastSignon.Text
            RS!IdleTime = .txtIdleTime.Text
            RS!Kills = .txtKills.Text
            RS!Deaths = .txtDeaths.Text
            RS!Gold = .txtGold.Text
            RS!Arrows = .txtArrows.Text
            RS!Bombs = .txtBombs.Text
            RS!Health = .txtHealth.Text
            RS!TotalHealth = .txtTotalHealth.Text
            RS!Level = .txtLevel.Text
            RS!PlayerX = .txtPlayerX.Text
            RS!PlayerY = .txtPlayerY.Text
            RS!Banned = .txtBanned.Text
            RS!AccessPriviledges = .txtAccess.Text
            RS!WeaponCount = .txtWeaponCount.Text
            RS!Weapon1 = .lstWeapons.List(0)
            RS!Weapon2 = .lstWeapons.List(1)
            RS!Weapon3 = .lstWeapons.List(2)
            RS!Weapon4 = .lstWeapons.List(3)
            RS!Weapon5 = .lstWeapons.List(4)
            RS!Weapon6 = .lstWeapons.List(5)
            RS!Weapon7 = .lstWeapons.List(6)
            RS!Weapon8 = .lstWeapons.List(7)
            RS!Weapon9 = .lstWeapons.List(8)
            RS!Weapon10 = .lstWeapons.List(9)
            RS!Weapon11 = .lstWeapons.List(10)
            RS!Weapon12 = .lstWeapons.List(11)
            RS!Weapon13 = .lstWeapons.List(12)
            RS!Weapon14 = .lstWeapons.List(13)
            RS!Weapon15 = .lstWeapons.List(14)
            RS!ActiveWeapon = .txtActiveWeapon.Text
            If .txtArmor.Text = "" Then .txtArmor.Text = Chr(34) & " " & Chr(34)
            RS!Armor = .txtArmor.Text
            If .txtHelmet.Text = "" Then .txtHelmet.Text = Chr(34) & " " & Chr(34)
            RS!Helmet = .txtHelmet.Text
            If .txtSword.Text = "" Then .txtSword.Text = Chr(34) & " " & Chr(34)
            RS!Sword = .txtSword.Text
            If .txtShield.Text = "" Then .txtShield.Text = Chr(34) & " " & Chr(34)
            RS!Shield = .txtShield.Text

            .lstInventory.Clear
            Open App.Path & "\TempX.txt" For Output As #6
                Do Until i = .lstInventory.ListCount + 1
                    Write #6, .lstInventory.List(i)
                    i = i + 1
                Loop
            Close #6
            Open App.Path & "\TempX.txt" For Input As #6
                Do Until EOF(6)
                    Input #6, Temp
                    .lstInventory.AddItem Temp
                Loop
            Close #6
            Kill App.Path & "\TempX.txt"
            
            RS.Update
    End With
    Exit Function
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    RS.Close
    DB.Close
    FileCopy App.Path & "\Databases\membersDB.mdb", App.Path & "\Databases\server.mdb"
End Sub

Attribute VB_Name = "Module1"
Option Explicit

Public DB As Database
Public RS As Recordset

Public Function DatabaseStartup()
    Set DB = OpenDatabase(App.Path & "\..\Databases\server.mdb")
    Set RS = DB.OpenRecordset("Members")
    RS.Edit
End Function

Public Function LoadData()
    Dim Temp As String
    
    With Main
        .txtID.Text = RS!ID
        .txtUsername.Text = RS!Username
        .txtPassword.Text = RS!Password
        .txtOnlineTime.Text = RS!OnlineTime
        .txtLastSignon.Text = RS!LastSignon
        .txtIdleTime.Text = RS!IdleTime
        .txtKills.Text = RS!Kills
        .txtDeaths.Text = RS!Deaths
        .txtGold.Text = RS!Gold
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
End Function

Public Function CheckPassword(Username As String, Password As String) As Boolean
    Dim i As Integer
    i = 1
    
    With Main
        Do Until i > RS.RecordCount
            RS.MoveFirst
            .txtID.Text = RS!ID
            .txtUsername.Text = RS!Username
            .txtPassword.Text = RS!Password
            .txtOnlineTime.Text = RS!OnlineTime
            .txtLastSignon.Text = RS!LastSignon
            .txtIdleTime.Text = RS!IdleTime
            .txtKills.Text = RS!Kills
            .txtDeaths.Text = RS!Deaths
            .txtGold.Text = RS!Gold
            .txtHealth.Text = RS!Health
            .txtLevel.Text = RS!Level
            .txtPlayerX.Text = RS!PlayerX
            .txtPlayerY.Text = RS!PlayerY
            .txtBanned.Text = RS!Banned
            .txtAccess.Text = RS!AccessPriviledges
            
            If (.txtUsername.Text = Username) And (.txtPassword.Text = Password) Then
                CheckPassword = True
                Exit Function
            End If
            
        i = i + 1
        If i < RS.RecordCount + 1 Then
        Else
            RS.MoveNext
        End If
        Loop
        
    End With
End Function

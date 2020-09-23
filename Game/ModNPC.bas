Attribute VB_Name = "ModNPC"
Public JustWarped As Boolean

Public Function LoadScript(i As Integer)
    On Error GoTo FixIt
    Dim Temp2 As String
    Dim Temp As String
    Dim DwarpX As Integer, DwarpY As Integer, MapName As String
    Dim Num As Integer
    Dim a As Integer
    
    a = 0
    Num = FreeFile
    DoEvents
    Open App.Path & "\NPCs\" & Player.Map.NPCs.ScriptSource(i) For Input As #Num
        Do Until EOF(Num)
            If JustWarped = True Then JustWarped = False: GoTo closeit
            Line Input #Num, Temp
            
            DoEvents
            'find out what the heck it is!!!
            If Mid(UCase(Trim(Temp)), 1, 14) = "PLAYERTOUCHME(" And frmMain.imgNPC(i).Visible = True Then
                If frmMain.imgNPC(i).Left <= frmMain.PlayerImage(0).Left And frmMain.PlayerImage(0).Left <= (frmMain.imgNPC(i).Left + frmMain.imgNPC(i).Width) Then
                If frmMain.imgNPC(i).Top <= frmMain.PlayerImage(0).Top And frmMain.PlayerImage(0).Top <= (frmMain.imgNPC(i).Top + frmMain.imgNPC(i).Height) Then
                    Do Until EOF(Num) Or Temp = ")"
                    Line Input #Num, Temp
                    
                    If Mid(UCase(Trim(Temp)), 1, 7) = "PLAYER." And Right(Trim(Temp), 1) = "(" Then 'it's a Player.whatever(
                        Temp = Mid(UCase(Trim(Temp)), 8, Len(UCase(Trim(Temp))) - 8) 'player.whatever is now just whatever(
                        If Mid(UCase(Trim(Temp)), 1, 13) = "PRIVILEDGES =" Then
                            Temp = Mid(UCase(Trim(Temp)), 14, Len(UCase(Trim(Temp))) - 13)
                            If UCase(frmDB.txtAccess.Text) = Trim(UCase(Temp)) Then 'if player.priviledeges = whatever( passed and now we are going into the loop for it
                                Do Until Temp = ")" Or EOF(Num)
                                    Line Input #1, Temp
                                    If Trim(Temp) = ")" Then GoTo SkipLoop1
                                    If Temp <> "" Then
                                        DoSimpleAction Temp, i
                                    End If
                                Loop
SkipLoop1:
                            End If
                        End If
                    Else
                        DoSimpleAction Temp, i
                    End If
                    Loop
                End If 'Y coords are right for a playertouchme
                End If 'X coords are right for a playertouchme
            End If 'end if it is PlayerTouchMe
        
            'it is playerenters
            If Mid(UCase(Trim(Temp)), 1, 13) = "PLAYERENTERS(" And frmMain.imgNPC(i).Visible = True Then
                    Line Input #Num, Temp
                    
                    If Mid(UCase(Trim(Temp)), 1, 7) = "PLAYER." And Right(Trim(Temp), 1) = "(" Then 'it's a Player.whatever(
                        Temp = Mid(UCase(Trim(Temp)), 8, Len(UCase(Trim(Temp))) - 8) 'player.whatever is now just whatever(
                        If Mid(UCase(Trim(Temp)), 1, 13) = "PRIVILEDGES =" Then
                            Temp = Mid(UCase(Trim(Temp)), 14, Len(UCase(Trim(Temp))) - 13)
                            If UCase(frmDB.txtAccess.Text) = Trim(UCase(Temp)) Then 'if player.priviledeges = whatever( passed and now we are going into the loop for it
                                Do Until Temp = ")" Or EOF(Num)
                                    Line Input #1, Temp
                                    If Trim(Temp) = ")" Then GoTo SkipLoop2
                                    If Temp <> "" Then
                                        DoSimpleAction Temp, i
                                    End If
                                Loop
SkipLoop2:
                            End If
                        End If
                    End If
            End If
            'end of PlayerEnters(
loopp:
        Loop
closeit:
    Close #1
    
    Exit Function
FixIt:
    frmMain.LoadNPCs
    If a = 5 Then Exit Function
    a = a + 1
    Resume Next
End Function

Public Function LoadScript2(i As Integer)
    On Error GoTo FixIt
    Dim Temp2 As String
    Dim Temp As String
    Dim DwarpX As Integer, DwarpY As Integer, MapName As String
    Dim Num As Integer
    Dim a As Integer
    
    a = 0
    Num = FreeFile
    DoEvents
    Open App.Path & "\NPCs\" & Player.Map.NPCs.ScriptSource(i) For Input As #Num
        Do Until EOF(Num)
            If JustWarped = True Then JustWarped = False: GoTo closeit
            Line Input #Num, Temp
            
            DoEvents
            'find out what the heck it is!!!
            If Mid(UCase(Trim(Temp)), 1, 14) = "PLAYERTOUCHME(" And frmMainOnline.imgNPC(i).Visible = True Then
                If frmMainOnline.imgNPC(i).Left <= frmMainOnline.PlayerImage(0).Left And frmMainOnline.PlayerImage(0).Left <= (frmMainOnline.imgNPC(i).Left + frmMainOnline.imgNPC(i).Width) Then
                If frmMainOnline.imgNPC(i).Top <= frmMainOnline.PlayerImage(0).Top And frmMainOnline.PlayerImage(0).Top <= (frmMainOnline.imgNPC(i).Top + frmMainOnline.imgNPC(i).Height) Then
                    Line Input #Num, Temp
                    
                    If Mid(UCase(Trim(Temp)), 1, 7) = "PLAYER." And Right(Trim(Temp), 1) = "(" Then 'it's a Player.whatever(
                        Temp = Mid(UCase(Trim(Temp)), 8, Len(UCase(Trim(Temp))) - 8) 'player.whatever is now just whatever(
                        If Mid(UCase(Trim(Temp)), 1, 13) = "PRIVILEDGES =" Then
                            Temp = Mid(UCase(Trim(Temp)), 14, Len(UCase(Trim(Temp))) - 13)
                            If UCase(frmDB.txtAccess.Text) = Trim(UCase(Temp)) Then 'if player.priviledeges = whatever( passed and now we are going into the loop for it
                                Do Until Temp = ")" Or EOF(Num)
                                    Line Input #1, Temp
                                    If Trim(Temp) = ")" Then GoTo SkipLoop1
                                    If Temp <> "" Then
                                        DoSimpleAction Temp, i
                                    End If
                                Loop
SkipLoop1:
                            End If
                        End If
                    Else
                        DoSimpleAction Temp, i
                        GoTo loopp
                    End If
                End If 'Y coords are right for a playertouchme
                End If 'X coords are right for a playertouchme
            End If 'end if it is PlayerTouchMe
        
            'it is playerenters
            If Mid(UCase(Trim(Temp)), 1, 13) = "PLAYERENTERS(" And frmMainOnline.imgNPC(i).Visible = True Then
                    Line Input #Num, Temp
                    
                    If Mid(UCase(Trim(Temp)), 1, 7) = "PLAYER." And Right(Trim(Temp), 1) = "(" Then 'it's a Player.whatever(
                        Temp = Mid(UCase(Trim(Temp)), 8, Len(UCase(Trim(Temp))) - 8) 'player.whatever is now just whatever(
                        If Mid(UCase(Trim(Temp)), 1, 13) = "PRIVILEDGES =" Then
                            Temp = Mid(UCase(Trim(Temp)), 14, Len(UCase(Trim(Temp))) - 13)
                            If UCase(frmDB.txtAccess.Text) = Trim(UCase(Temp)) Then 'if player.priviledeges = whatever( passed and now we are going into the loop for it
                                Do Until Temp = ")" Or EOF(Num)
                                    Line Input #1, Temp
                                    If Trim(Temp) = ")" Then GoTo SkipLoop2
                                    If Temp <> "" Then
                                        DoSimpleAction Temp, i
                                    End If
                                Loop
SkipLoop2:
                            End If
                        End If
                    End If
            End If
            'end of PlayerEnters(
loopp:
        Loop
closeit:
    Close #1
    
    Exit Function
FixIt:
    frmMainOnline.LoadNPCs
    If a = 5 Then Exit Function
    a = a + 1
    Resume Next
End Function

Public Function DoSimpleAction(Temp As String, b As Integer)
    Dim it As String
    Dim Data As String
    Dim i As Integer
    Dim DwarpX As Integer, DwarpY As Integer, MapName As String
    Dim GifName, GifPath, WeaponSource As String
    Dim Num As Integer
    Dim Num2 As Integer
    
    Num = FreeFile
    Num2 = FreeFile
    
    'Beginning of Say Function
    If Left(Trim(UCase(Temp)), 5) = ("SAY " & Chr(34)) Then ' it is Say"
        Temp = Mid(Temp, 6, Len(Temp) - 5)
        it = ""
        i = 1
        Data = ""
        Do Until it = Chr(34)
            it = Mid(Temp, i, 1)
            If it = Chr(34) Then GoTo skipit1
            Data = Data & it
            i = i + 1
        Loop
skipit1:
        If i <= Len(Temp) Then
            frmIM.Show
            frmIM.txtMessage.Text = Data
            frmIM.Left = frmMain.Left + ((frmMain.imgNPC(b).Left + frmMain.imgNPC(b).Width) * Screen.TwipsPerPixelX)
            frmIM.Top = (((frmMain.imgNPC(b).Top * Screen.TwipsPerPixelY) + frmMain.Top) + (frmMain.imgNPC(b).Height * Screen.TwipsPerPixelY)) - frmIM.Height
            Exit Function
        End If
    End If
    'done with Say Function
    
    'Beginning of Warp Function
    If Left(Trim(UCase(Temp)), 5) = "WARP " Then ' it is Warp
        Temp = Mid(Trim(UCase(Temp)), 6, Len(Trim(Temp)) - 5)
        Open App.Path & "\Temp100.txt" For Output As #7
            Print #7, Temp
        Close #7
        Open App.Path & "\Temp100.txt" For Input As #7
            Input #7, DwarpX
            Input #7, DwarpY
            Input #7, MapName
        Close #7
        Kill App.Path & "\Temp100.txt"
        frmMain.Warp DwarpX, DwarpY, MapName
        JustWarped = True
        Exit Function
    End If
    'done with Warp Function
    
    'Beginning of AddWeapon Function
    If Left(Trim(UCase(Temp)), 10) = "ADDWEAPON " Then
        'format: AddWeapon WeaponName, WeaponGIF, SourceFile(WPN)
        Temp = Mid(Trim(Temp), 11, Len(Trim(Temp)) - 10)
        Open App.Path & "\Temp5.txt" For Output As #15
            Print #15, Temp
        Close #15
        Open App.Path & "\Temp5.txt" For Input As #15
            Input #15, GifName
            Input #15, GifPath
            Input #15, WeaponSource
        Close #15
        If frmDB.txtWeaponCount.Text >= 15 Then MsgBox "You cannot have more than 15 weapons, please drop one.": Exit Function
        frmDB.txtWeaponCount.Text = frmDB.txtWeaponCount.Text + 1
        frmMain.imgWeapon(frmDB.txtWeaponCount.Text - 1).ToolTipText = GifName
        frmMain.imgWeapon(frmDB.txtWeaponCount.Text - 1).Picture = LoadPicture(App.Path & "\Images\" & GifPath)
        frmMain.imgWeapon(frmDB.txtWeaponCount.Text - 1).Tag = WeaponSource
        frmMain.imgWeapon(frmDB.txtWeaponCount.Text - 1).Visible = True
        frmDB.lstWeapons.RemoveItem frmDB.txtWeaponCount.Text - 1
        frmDB.lstWeapons.AddItem GifName & "," & GifPath & "," & WeaponSource, frmDB.txtWeaponCount.Text - 1
        Kill App.Path & "\Temp5.txt"
        Exit Function
    End If
    'done with AddWeapon Function
    
    'Beginning of AddTextToSign Function
    If Left(Trim(UCase(Temp)), 14) = "ADDTEXTTOSIGN " Then
        Temp = Mid(Trim(Temp), 15, Len(Trim(Temp)) - 14)
        Open App.Path & "\Temp123.txt" For Output As Num
            Print #Num, Temp
        Close #Num
        Open App.Path & "\Temp123.txt" For Input As Num
            Input #Num, MapName
            Input #Num, DwarpX
            Input #Num, DwarpY
            Input #Num, Data
        Close #Num

        If Left(UCase(Data), 7) = "PLAYER." Then
            Data = Mid(UCase(Data), 8, Len(Data) - 7)
            If Data = "NAME" Then
                'data to add to the sign was "Player.Name"
                Data = Player.Name
            End If
        End If
        
        Open App.Path & "\Temp4.txt" For Output As #Num2
        Num = FreeFile
        Open App.Path & "\Maps\" & MapName For Input As #Num
            Temp = ""
            Line Input #Num, Temp
            Do Until Temp = "*SIGNS*"
                Print #Num2, Temp
                Line Input #Num, Temp
            Loop
            Do Until Left(Temp, Len(DwarpX & DwarpY) + 1) = DwarpX & "," & DwarpY
                Print #Num2, Temp
                Line Input #Num, Temp
            Loop
            
            Print #Num2, (Temp & Data)
            Do Until EOF(Num)
                Line Input #Num, Temp
                Print #Num2, Temp
            Loop
        Close #Num
        Close #Num2
        
        Open App.Path & "\Maps\" & MapName For Output As #Num
        Open App.Path & "\Temp4.txt" For Input As #Num2
            Do Until EOF(Num2)
                Line Input #Num2, Temp
                Print #Num, Temp
            Loop
        Close #Num
        Close #Num2
        Kill App.Path & "\Temp4.txt"
        Kill App.Path & "\Temp123.txt"
        Exit Function
    End If
    'done with AddTextToSign Function
    
    'beginning of me. function
    If Left(Trim(UCase(Temp)), 3) = "ME." Then ' it is Me.
        Temp = Mid(UCase(Trim(Temp)), 4, Len(Trim(Temp)) - 3)
        
        'beginning of Me.Direction function
        If Left(Trim(UCase(Temp)), 11) = "DIRECTION =" Then
            Temp = Mid(Trim(UCase(Temp)), 13, Len(Trim(Temp)) - 12)
            Select Case Temp
            Case "DOWN"
                GoingDown
            Case "UP"
                GoingUp
            Case "Left"
                GoingLeft
            Case "Right"
                GoingRight
            End Select
            Exit Function
        End If
        'end of Me.Direction function
        
        'beginning of Me.Visible function
        If Left(Trim(UCase(Temp)), 9) = "VISIBLE =" Then
            Temp = Mid(Trim(UCase(Temp)), 11, Len(Trim(Temp)) - 10)
            frmMain.imgNPC(b).Visible = Trim(Temp)
            Exit Function
        End If
        'end of Me.Visible function
    End If
    'end of me. function
End Function

Public Function Drop(Item As String, x As Integer, y As Integer)
    Dim i As Integer
    
    Randomize
    i = Int((5 * Rnd) + 1)
    
    Select Case Item:
        Case "Arrows":
            Player.Map.Arrows.Add i, x / 32 & "," & y / 32
            Load frmMain.imgArrow(Player.Map.Arrows.Count)
            frmMain.imgArrow(Player.Map.Arrows.Count).Picture = frmGraphics.imgArrow.Picture
            frmMain.imgArrow(Player.Map.Arrows.Count).Tag = i
            frmMain.imgArrow(Player.Map.Arrows.Count).Left = x
            frmMain.imgArrow(Player.Map.Arrows.Count).Top = y
            frmMain.imgArrow(Player.Map.Arrows.Count).Visible = True
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMain.imgArrow(Player.Map.Arrows.Count).Visible = False
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMain.imgArrow(Player.Map.Arrows.Count).Visible = True
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMain.imgArrow(Player.Map.Arrows.Count).Visible = False
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMain.imgArrow(Player.Map.Arrows.Count).Visible = True
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMain.imgArrow(Player.Map.Arrows.Count).Visible = False
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMain.imgArrow(Player.Map.Arrows.Count).Visible = True
        Case "Bombs":
            Player.Map.Bombs.Add i, x / 32 & "," & y / 32
            Load frmMain.imgBomb(Player.Map.Bombs.Count)
            frmMain.imgBomb(Player.Map.Bombs.Count).Picture = frmGraphics.imgBomb.Picture
            frmMain.imgBomb(Player.Map.Bombs.Count).Tag = i
            frmMain.imgBomb(Player.Map.Bombs.Count).Left = x
            frmMain.imgBomb(Player.Map.Bombs.Count).Top = y
            frmMain.imgBomb(Player.Map.Bombs.Count).Visible = True
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMain.imgBomb(Player.Map.Bombs.Count).Visible = False
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMain.imgBomb(Player.Map.Bombs.Count).Visible = True
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMain.imgBomb(Player.Map.Bombs.Count).Visible = False
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMain.imgBomb(Player.Map.Bombs.Count).Visible = True
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
        End Select
        
End Function




















































Public Function LoadScriptOnline(i As Integer)
    On Error GoTo FixIt
    Dim Temp2 As String
    Dim Temp As String
    Dim DwarpX As Integer, DwarpY As Integer, MapName As String
    Dim Num As Integer
    Dim a As Integer
    
    a = 0
    Num = FreeFile
    DoEvents
    Open App.Path & "\NPCs\" & Player.Map.NPCs.ScriptSource(i) For Input As #Num
        Do Until EOF(Num)
            If JustWarped = True Then JustWarped = False: GoTo closeit
            Line Input #Num, Temp
            
            DoEvents
            'find out what the heck it is!!!
            If Mid(UCase(Trim(Temp)), 1, 14) = "PLAYERTOUCHME(" And frmMainOnline.imgNPC(i).Visible = True Then
                If frmMainOnline.imgNPC(i).Left <= frmMainOnline.PlayerImage(0).Left And frmMainOnline.PlayerImage(0).Left <= (frmMainOnline.imgNPC(i).Left + frmMainOnline.imgNPC(i).Width) Then
                If frmMainOnline.imgNPC(i).Top <= frmMainOnline.PlayerImage(0).Top And frmMainOnline.PlayerImage(0).Top <= (frmMainOnline.imgNPC(i).Top + frmMainOnline.imgNPC(i).Height) Then
                    Do Until EOF(Num) Or Temp = ")"
                    Line Input #Num, Temp
                    
                    If Mid(UCase(Trim(Temp)), 1, 7) = "PLAYER." And Right(Trim(Temp), 1) = "(" Then 'it's a Player.whatever(
                        Temp = Mid(UCase(Trim(Temp)), 8, Len(UCase(Trim(Temp))) - 8) 'player.whatever is now just whatever(
                        If Mid(UCase(Trim(Temp)), 1, 13) = "PRIVILEDGES =" Then
                            Temp = Mid(UCase(Trim(Temp)), 14, Len(UCase(Trim(Temp))) - 13)
                            If UCase(frmDB.txtAccess.Text) = Trim(UCase(Temp)) Then 'if player.priviledeges = whatever( passed and now we are going into the loop for it
                                Do Until Temp = ")" Or EOF(Num)
                                    Line Input #1, Temp
                                    If Trim(Temp) = ")" Then GoTo SkipLoop1
                                    If Temp <> "" Then
                                        DoSimpleActionOnline Temp, i
                                    End If
                                Loop
SkipLoop1:
                            End If
                        End If
                    Else
                        DoSimpleActionOnline Temp, i
                    End If
                    Loop
                End If 'Y coords are right for a playertouchme
                End If 'X coords are right for a playertouchme
            End If 'end if it is PlayerTouchMe
        
            'it is playerenters
            If Mid(UCase(Trim(Temp)), 1, 13) = "PLAYERENTERS(" And frmMainOnline.imgNPC(i).Visible = True Then
                    Line Input #Num, Temp
                    
                    If Mid(UCase(Trim(Temp)), 1, 7) = "PLAYER." And Right(Trim(Temp), 1) = "(" Then 'it's a Player.whatever(
                        Temp = Mid(UCase(Trim(Temp)), 8, Len(UCase(Trim(Temp))) - 8) 'player.whatever is now just whatever(
                        If Mid(UCase(Trim(Temp)), 1, 13) = "PRIVILEDGES =" Then
                            Temp = Mid(UCase(Trim(Temp)), 14, Len(UCase(Trim(Temp))) - 13)
                            If UCase(frmDB.txtAccess.Text) = Trim(UCase(Temp)) Then 'if player.priviledeges = whatever( passed and now we are going into the loop for it
                                Do Until Temp = ")" Or EOF(Num)
                                    Line Input #1, Temp
                                    If Trim(Temp) = ")" Then GoTo SkipLoop2
                                    If Temp <> "" Then
                                        DoSimpleActionOnline Temp, i
                                    End If
                                Loop
SkipLoop2:
                            End If
                        End If
                    End If
            End If
            'end of PlayerEnters(
loopp:
        Loop
closeit:
    Close #1
    
    Exit Function
FixIt:
    frmMainOnline.LoadNPCs
    If a = 5 Then Exit Function
    a = a + 1
    Resume Next
End Function

Public Function LoadScript2Online(i As Integer)
    On Error GoTo FixIt
    Dim Temp2 As String
    Dim Temp As String
    Dim DwarpX As Integer, DwarpY As Integer, MapName As String
    Dim Num As Integer
    Dim a As Integer
    
    a = 0
    Num = FreeFile
    DoEvents
    Open App.Path & "\NPCs\" & Player.Map.NPCs.ScriptSource(i) For Input As #Num
        Do Until EOF(Num)
            If JustWarped = True Then JustWarped = False: GoTo closeit
            Line Input #Num, Temp
            
            DoEvents
            'find out what the heck it is!!!
            If Mid(UCase(Trim(Temp)), 1, 14) = "PLAYERTOUCHME(" And frmMainOnlineOnline.imgNPC(i).Visible = True Then
                If frmMainOnlineOnline.imgNPC(i).Left <= frmMainOnlineOnline.PlayerImage(0).Left And frmMainOnlineOnline.PlayerImage(0).Left <= (frmMainOnlineOnline.imgNPC(i).Left + frmMainOnlineOnline.imgNPC(i).Width) Then
                If frmMainOnlineOnline.imgNPC(i).Top <= frmMainOnlineOnline.PlayerImage(0).Top And frmMainOnlineOnline.PlayerImage(0).Top <= (frmMainOnlineOnline.imgNPC(i).Top + frmMainOnlineOnline.imgNPC(i).Height) Then
                    Line Input #Num, Temp
                    
                    If Mid(UCase(Trim(Temp)), 1, 7) = "PLAYER." And Right(Trim(Temp), 1) = "(" Then 'it's a Player.whatever(
                        Temp = Mid(UCase(Trim(Temp)), 8, Len(UCase(Trim(Temp))) - 8) 'player.whatever is now just whatever(
                        If Mid(UCase(Trim(Temp)), 1, 13) = "PRIVILEDGES =" Then
                            Temp = Mid(UCase(Trim(Temp)), 14, Len(UCase(Trim(Temp))) - 13)
                            If UCase(frmDB.txtAccess.Text) = Trim(UCase(Temp)) Then 'if player.priviledeges = whatever( passed and now we are going into the loop for it
                                Do Until Temp = ")" Or EOF(Num)
                                    Line Input #1, Temp
                                    If Trim(Temp) = ")" Then GoTo SkipLoop1
                                    If Temp <> "" Then
                                        DoSimpleActionOnline Temp, i
                                    End If
                                Loop
SkipLoop1:
                            End If
                        End If
                    Else
                        DoSimpleAction Temp, i
                        GoTo loopp
                    End If
                End If 'Y coords are right for a playertouchme
                End If 'X coords are right for a playertouchme
            End If 'end if it is PlayerTouchMe
        
            'it is playerenters
            If Mid(UCase(Trim(Temp)), 1, 13) = "PLAYERENTERS(" And frmMainOnlineOnline.imgNPC(i).Visible = True Then
                    Line Input #Num, Temp
                    
                    If Mid(UCase(Trim(Temp)), 1, 7) = "PLAYER." And Right(Trim(Temp), 1) = "(" Then 'it's a Player.whatever(
                        Temp = Mid(UCase(Trim(Temp)), 8, Len(UCase(Trim(Temp))) - 8) 'player.whatever is now just whatever(
                        If Mid(UCase(Trim(Temp)), 1, 13) = "PRIVILEDGES =" Then
                            Temp = Mid(UCase(Trim(Temp)), 14, Len(UCase(Trim(Temp))) - 13)
                            If UCase(frmDB.txtAccess.Text) = Trim(UCase(Temp)) Then 'if player.priviledeges = whatever( passed and now we are going into the loop for it
                                Do Until Temp = ")" Or EOF(Num)
                                    Line Input #1, Temp
                                    If Trim(Temp) = ")" Then GoTo SkipLoop2
                                    If Temp <> "" Then
                                        DoSimpleActionOnline Temp, i
                                    End If
                                Loop
SkipLoop2:
                            End If
                        End If
                    End If
            End If
            'end of PlayerEnters(
loopp:
        Loop
closeit:
    Close #1
    
    Exit Function
FixIt:
    frmMainOnlineOnline.LoadNPCs
    If a = 5 Then Exit Function
    a = a + 1
    Resume Next
End Function

Public Function DoSimpleActionOnline(Temp As String, b As Integer)
    Dim it As String
    Dim Data As String
    Dim i As Integer
    Dim DwarpX As Integer, DwarpY As Integer, MapName As String
    Dim GifName, GifPath, WeaponSource As String
    Dim Num As Integer
    Dim Num2 As Integer
    
    Num = FreeFile
    Num2 = FreeFile
    
    'Beginning of Say Function
    If Left(Trim(UCase(Temp)), 5) = ("SAY " & Chr(34)) Then ' it is Say"
        Temp = Mid(Temp, 6, Len(Temp) - 5)
        it = ""
        i = 1
        Data = ""
        Do Until it = Chr(34)
            it = Mid(Temp, i, 1)
            If it = Chr(34) Then GoTo skipit1
            Data = Data & it
            i = i + 1
        Loop
skipit1:
        If i <= Len(Temp) Then
            frmIM.Show
            frmIM.txtMessage.Text = Data
            frmIM.Left = frmMainOnline.Left + ((frmMainOnline.imgNPC(b).Left + frmMainOnline.imgNPC(b).Width) * Screen.TwipsPerPixelX)
            frmIM.Top = (((frmMainOnline.imgNPC(b).Top * Screen.TwipsPerPixelY) + frmMainOnline.Top) + (frmMainOnline.imgNPC(b).Height * Screen.TwipsPerPixelY)) - frmIM.Height
            Exit Function
        End If
    End If
    'done with Say Function
    
    'Beginning of Warp Function
    If Left(Trim(UCase(Temp)), 5) = "WARP " Then ' it is Warp
        Temp = Mid(Trim(UCase(Temp)), 6, Len(Trim(Temp)) - 5)
        Open App.Path & "\Temp100.txt" For Output As #7
            Print #7, Temp
        Close #7
        Open App.Path & "\Temp100.txt" For Input As #7
            Input #7, DwarpX
            Input #7, DwarpY
            Input #7, MapName
        Close #7
        Kill App.Path & "\Temp100.txt"
        frmMainOnline.Warp DwarpX, DwarpY, MapName
        JustWarped = True
        Exit Function
    End If
    'done with Warp Function
    
    'Beginning of AddWeapon Function
    If Left(Trim(UCase(Temp)), 10) = "ADDWEAPON " Then
        'format: AddWeapon WeaponName, WeaponGIF, SourceFile(WPN)
        Temp = Mid(Trim(Temp), 11, Len(Trim(Temp)) - 10)
        Open App.Path & "\Temp5.txt" For Output As #15
            Print #15, Temp
        Close #15
        Open App.Path & "\Temp5.txt" For Input As #15
            Input #15, GifName
            Input #15, GifPath
            Input #15, WeaponSource
        Close #15
        If frmDB.txtWeaponCount.Text >= 15 Then MsgBox "You cannot have more than 15 weapons, please drop one.": Exit Function
        frmDB.txtWeaponCount.Text = frmDB.txtWeaponCount.Text + 1
        frmMainOnline.imgWeapon(frmDB.txtWeaponCount.Text - 1).ToolTipText = GifName
        frmMainOnline.imgWeapon(frmDB.txtWeaponCount.Text - 1).Picture = LoadPicture(App.Path & "\Images\" & GifPath)
        frmMainOnline.imgWeapon(frmDB.txtWeaponCount.Text - 1).Tag = WeaponSource
        frmMainOnline.imgWeapon(frmDB.txtWeaponCount.Text - 1).Visible = True
        frmDB.lstWeapons.RemoveItem frmDB.txtWeaponCount.Text - 1
        frmDB.lstWeapons.AddItem GifName & "," & GifPath & "," & WeaponSource, frmDB.txtWeaponCount.Text - 1
        Kill App.Path & "\Temp5.txt"
        Exit Function
    End If
    'done with AddWeapon Function
    
    'Beginning of AddTextToSign Function
    If Left(Trim(UCase(Temp)), 14) = "ADDTEXTTOSIGN " Then
        Temp = Mid(Trim(Temp), 15, Len(Trim(Temp)) - 14)
        Open App.Path & "\Temp123.txt" For Output As Num
            Print #Num, Temp
        Close #Num
        Open App.Path & "\Temp123.txt" For Input As Num
            Input #Num, MapName
            Input #Num, DwarpX
            Input #Num, DwarpY
            Input #Num, Data
        Close #Num

        If Left(UCase(Data), 7) = "PLAYER." Then
            Data = Mid(UCase(Data), 8, Len(Data) - 7)
            If Data = "NAME" Then
                'data to add to the sign was "Player.Name"
                Data = Player.Name
            End If
        End If
        
        Open App.Path & "\Temp4.txt" For Output As #Num2
        Num = FreeFile
        Open App.Path & "\Maps\" & MapName For Input As #Num
            Temp = ""
            Line Input #Num, Temp
            Do Until Temp = "*SIGNS*"
                Print #Num2, Temp
                Line Input #Num, Temp
            Loop
            Do Until Left(Temp, Len(DwarpX & DwarpY) + 1) = DwarpX & "," & DwarpY
                Print #Num2, Temp
                Line Input #Num, Temp
            Loop
            
            Print #Num2, (Temp & Data)
            Do Until EOF(Num)
                Line Input #Num, Temp
                Print #Num2, Temp
            Loop
        Close #Num
        Close #Num2
        
        Open App.Path & "\Maps\" & MapName For Output As #Num
        Open App.Path & "\Temp4.txt" For Input As #Num2
            Do Until EOF(Num2)
                Line Input #Num2, Temp
                Print #Num, Temp
            Loop
        Close #Num
        Close #Num2
        Kill App.Path & "\Temp4.txt"
        Kill App.Path & "\Temp123.txt"
        Exit Function
    End If
    'done with AddTextToSign Function
    
    'beginning of me. function
    If Left(Trim(UCase(Temp)), 3) = "ME." Then ' it is Me.
        Temp = Mid(UCase(Trim(Temp)), 4, Len(Trim(Temp)) - 3)
        
        'beginning of Me.Direction function
        If Left(Trim(UCase(Temp)), 11) = "DIRECTION =" Then
            Temp = Mid(Trim(UCase(Temp)), 13, Len(Trim(Temp)) - 12)
            Select Case Temp
            Case "DOWN"
                GoingDown
            Case "UP"
                GoingUp
            Case "Left"
                GoingLeft
            Case "Right"
                GoingRight
            End Select
            Exit Function
        End If
        'end of Me.Direction function
        
        'beginning of Me.Visible function
        If Left(Trim(UCase(Temp)), 9) = "VISIBLE =" Then
            Temp = Mid(Trim(UCase(Temp)), 11, Len(Trim(Temp)) - 10)
            frmMainOnline.imgNPC(b).Visible = Trim(Temp)
            Exit Function
        End If
        'end of Me.Visible function
    End If
    'end of me. function
End Function

Public Function DropOnline(Item As String, x As Integer, y As Integer)
    Dim i As Integer
    
    Randomize
    i = Int((5 * Rnd) + 1)
    
    Select Case Item:
        Case "Arrows":
            Player.Map.Arrows.Add i, x / 32 & "," & y / 32
            Load frmMainOnline.imgArrow(Player.Map.Arrows.Count)
            frmMainOnline.imgArrow(Player.Map.Arrows.Count).Picture = frmGraphics.imgArrow.Picture
            frmMainOnline.imgArrow(Player.Map.Arrows.Count).Tag = i
            frmMainOnline.imgArrow(Player.Map.Arrows.Count).Left = x
            frmMainOnline.imgArrow(Player.Map.Arrows.Count).Top = y
            frmMainOnline.imgArrow(Player.Map.Arrows.Count).Visible = True
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMainOnline.imgArrow(Player.Map.Arrows.Count).Visible = False
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMainOnline.imgArrow(Player.Map.Arrows.Count).Visible = True
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMainOnline.imgArrow(Player.Map.Arrows.Count).Visible = False
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMainOnline.imgArrow(Player.Map.Arrows.Count).Visible = True
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMainOnline.imgArrow(Player.Map.Arrows.Count).Visible = False
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMainOnline.imgArrow(Player.Map.Arrows.Count).Visible = True
        Case "Bombs":
            Player.Map.Bombs.Add i, x / 32 & "," & y / 32
            Load frmMainOnline.imgBomb(Player.Map.Bombs.Count)
            frmMainOnline.imgBomb(Player.Map.Bombs.Count).Picture = frmGraphics.imgBomb.Picture
            frmMainOnline.imgBomb(Player.Map.Bombs.Count).Tag = i
            frmMainOnline.imgBomb(Player.Map.Bombs.Count).Left = x
            frmMainOnline.imgBomb(Player.Map.Bombs.Count).Top = y
            frmMainOnline.imgBomb(Player.Map.Bombs.Count).Visible = True
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMainOnline.imgBomb(Player.Map.Bombs.Count).Visible = False
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMainOnline.imgBomb(Player.Map.Bombs.Count).Visible = True
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMainOnline.imgBomb(Player.Map.Bombs.Count).Visible = False
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
            frmMainOnline.imgBomb(Player.Map.Bombs.Count).Visible = True
            i = 0
            Do Until i >= 15000
                i = i + 1
                DoEvents
            Loop
        End Select
        
End Function


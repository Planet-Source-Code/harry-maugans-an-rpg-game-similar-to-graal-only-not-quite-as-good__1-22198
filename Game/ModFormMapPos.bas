Attribute VB_Name = "ModFormMapPos"


Public Function Init()
    Dim ScaleH As Integer
    Dim ScaleW As Integer
    Dim i, j As Integer
    Dim a, b As String
    Dim Temp As Integer
    Dim Other As String
    Dim q As Integer
    Dim Num As Integer
    
    ScaleH = Player.Map.MapHeight
    ScaleW = Player.Map.MapWidth
    
    Player.Map.TotalMapData = Mid(Player.Map.TotalMapData, 1, Len(Player.Map.TotalMapData) - 1)
    
    j = 0
    i = 0
    a = 0
    b = 0
    Temp = 0
    O = 0
    
    ClearMapBoxes

    Open App.Path & "\Temp.dat" For Output As #9
        Print #9, Player.Map.TotalMapData
    Close #9
    
    Open App.Path & "\Temp.dat" For Input As #9
    Do Until i = ScaleH
        j = 0
        Do Until j = ScaleW
            
            Input #9, Num
            
            If Num = 1 Or Num = 4 Or Num = 5 Or Num = 7 Or Num = 15 Or Num = 16 Or Num = 17 Or Num = 18 Or Num = 19 Or Num = 20 Or Num = 21 Or Num = 22 Or Num = 23 Or Num = 24 Or Num = 25 Or Num = 26 Or Num = 27 Or Num = 28 Or Num = 31 Then
                Player.Map.MapBoxes.Add "walkable", j & "," & i
            Else
                Player.Map.MapBoxes.Add "unwalkable", j & "," & i
            End If
            j = j + 1
        Loop
        Temp = (Temp + ScaleW)
        i = i + 1
    Loop
    Close #9
    Kill App.Path & "\Temp.dat"
End Function


Public Function ClearMapBoxes()
    On Error Resume Next

    Dim i As Integer
    
    For i = 1 To Player.Map.MapBoxes.Count
        Player.Map.MapBoxes.Remove 1
    Next i
    
End Function


Public Function ClearWarps()
    On Error Resume Next

    Dim i As Integer
    
    For i = 1 To Player.Map.Warps.Count
        Player.Map.Warps.Remove 1
    Next i
    
End Function

Public Function ClearSigns()
    On Error Resume Next

    Dim i As Integer
    
    For i = 1 To Player.Map.Signs.Count
        Player.Map.Signs.Remove 1
    Next i
    
End Function

Public Function ClearGold()
    On Error Resume Next

    Dim i As Integer
    
    For i = 1 To Player.Map.Gold.Count
        Unload frmMain.imgGold(i)
        Player.Map.Gold.Remove 1
    Next i
    
End Function

Public Function ClearArrows()
    On Error Resume Next

    Dim i As Integer
    
    For i = 1 To Player.Map.Arrows.Count
        Unload frmMain.imgArrow(i)
        Player.Map.Arrows.Remove 1
    Next i
    
End Function

Public Function ClearBombs()
    On Error Resume Next

    Dim i As Integer
    
    For i = 1 To Player.Map.Bombs.Count
        Unload frmMain.imgBomb(i)
        Player.Map.Bombs.Remove 1
    Next i
    
End Function

Public Function ClearBushes()
    On Error Resume Next

    Dim i As Integer
    
    For i = 1 To Player.Map.Bushes.Count
        Unload frmMain.imgBush(i)
        Player.Map.Bushes.Remove 1
    Next i
    
End Function

Public Function ClearNPCs()
    On Error Resume Next

    Dim i As Integer
    
    For i = 1 To Player.Map.NPCs.GifPath.Count
        Unload frmMain.imgNPC(i)
    Next i
    
    For i = 1 To Player.Map.NPCs.GifPath.Count
        Player.Map.NPCs.GifPath.Remove 1
    Next i
    For i = 1 To Player.Map.NPCs.Height.Count
        Player.Map.NPCs.Height.Remove 1
    Next i
    For i = 1 To Player.Map.NPCs.Name.Count
        Player.Map.NPCs.Name.Remove 1
    Next i
    For i = 1 To Player.Map.NPCs.ScriptSource.Count
        Player.Map.NPCs.ScriptSource.Remove 1
    Next i
    For i = 1 To Player.Map.NPCs.Width.Count
        Player.Map.NPCs.Width.Remove 1
    Next i
    For i = 1 To Player.Map.NPCs.XPos.Count
        Player.Map.NPCs.XPos.Remove 1
    Next i
    For i = 1 To Player.Map.NPCs.YPos.Count
        Player.Map.NPCs.YPos.Remove 1
    Next i
    
    
End Function



































Public Function ClearGoldOnline()
    On Error Resume Next

    Dim i As Integer
    
    For i = 1 To Player.Map.Gold.Count
        Unload frmMainOnline.imgGold(i)
        Player.Map.Gold.Remove 1
    Next i
    
End Function

Public Function ClearArrowsOnline()
    On Error Resume Next

    Dim i As Integer
    
    For i = 1 To Player.Map.Arrows.Count
        Unload frmMainOnline.imgArrow(i)
        Player.Map.Arrows.Remove 1
    Next i
    
End Function

Public Function ClearBombsOnline()
    On Error Resume Next

    Dim i As Integer
    
    For i = 1 To Player.Map.Bombs.Count
        Unload frmMainOnline.imgBomb(i)
        Player.Map.Bombs.Remove 1
    Next i
    
End Function

Public Function ClearBushesOnline()
    On Error Resume Next

    Dim i As Integer
    
    For i = 1 To Player.Map.Bushes.Count
        Unload frmMainOnline.imgBush(i)
        Player.Map.Bushes.Remove 1
    Next i
    
End Function

Public Function ClearNPCsOnline()
    On Error Resume Next

    Dim i As Integer
    
    For i = 1 To Player.Map.NPCs.GifPath.Count
        Unload frmMainOnline.imgNPC(i)
    Next i
    
    For i = 1 To Player.Map.NPCs.GifPath.Count
        Player.Map.NPCs.GifPath.Remove 1
    Next i
    For i = 1 To Player.Map.NPCs.Height.Count
        Player.Map.NPCs.Height.Remove 1
    Next i
    For i = 1 To Player.Map.NPCs.Name.Count
        Player.Map.NPCs.Name.Remove 1
    Next i
    For i = 1 To Player.Map.NPCs.ScriptSource.Count
        Player.Map.NPCs.ScriptSource.Remove 1
    Next i
    For i = 1 To Player.Map.NPCs.Width.Count
        Player.Map.NPCs.Width.Remove 1
    Next i
    For i = 1 To Player.Map.NPCs.XPos.Count
        Player.Map.NPCs.XPos.Remove 1
    Next i
    For i = 1 To Player.Map.NPCs.YPos.Count
        Player.Map.NPCs.YPos.Remove 1
    Next i
    
    
End Function




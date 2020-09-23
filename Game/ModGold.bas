Attribute VB_Name = "ModGold"
Option Explicit

Public Function LoadGold(i As Integer)
    If frmMain.imgGold(i).Left <= frmMain.PlayerImage(0).Left And frmMain.PlayerImage(0).Left <= (frmMain.imgGold(i).Left + frmMain.imgGold(i).Width) Then
    If frmMain.imgGold(i).Top <= frmMain.PlayerImage(0).Top And frmMain.PlayerImage(0).Top <= (frmMain.imgGold(i).Top + frmMain.imgGold(i).Height) Then
        If frmMain.imgGold(i).Visible = False Then Exit Function
        PlayWav "Gold.wav"
        Player.Gold = Player.Gold + frmMain.imgGold(i).Tag
        frmMain.txtGold.Text = Player.Gold
        frmDB.txtGold.Text = Player.Gold
        frmMain.imgGold(i).Visible = False
    End If
    End If
End Function

Public Function LoadArrows(i As Integer)
    If frmMain.imgArrow(i).Left <= frmMain.PlayerImage(0).Left And frmMain.PlayerImage(0).Left <= (frmMain.imgArrow(i).Left + frmMain.imgArrow(i).Width) Then
    If frmMain.imgArrow(i).Top <= frmMain.PlayerImage(0).Top And frmMain.PlayerImage(0).Top <= (frmMain.imgArrow(i).Top + frmMain.imgArrow(i).Height) Then
        If frmMain.imgArrow(i).Visible = False Then Exit Function
        PlayWav "Gold.wav"
        Player.Arrows = Player.Arrows + frmMain.imgArrow(i).Tag
        frmMain.txtArrows.Text = Player.Arrows
        frmDB.txtArrows.Text = Player.Arrows
        frmMain.imgArrow(i).Visible = False
    End If
    End If
End Function

Public Function LoadBombs(i As Integer)
    If frmMain.imgBomb(i).Left <= frmMain.PlayerImage(0).Left And frmMain.PlayerImage(0).Left <= (frmMain.imgBomb(i).Left + frmMain.imgBomb(i).Width) Then
    If frmMain.imgBomb(i).Top <= frmMain.PlayerImage(0).Top And frmMain.PlayerImage(0).Top <= (frmMain.imgBomb(i).Top + frmMain.imgBomb(i).Height) Then
        If frmMain.imgBomb(i).Visible = False Then Exit Function
        PlayWav "Gold.wav"
        Player.Bombs = Player.Bombs + frmMain.imgBomb(i).Tag
        frmMain.txtBombs.Text = Player.Bombs
        frmDB.txtBombs.Text = Player.Bombs
        frmMain.imgBomb(i).Visible = False
    End If
    End If
End Function

Public Function LoadBushes(i As Integer)
    Dim b As Integer
    
    If frmMain.imgSword.Visible = False Then Exit Function
    If frmMain.imgBush(i).Visible = False Then Exit Function
    If (frmMain.imgSword.Left) = frmMain.imgBush(i).Left Then
    If (frmMain.imgSword.Top) = frmMain.imgBush(i).Top Then
        frmMain.imgBush(i).Visible = False
        If Player.Map.Bushes.Item(i) = "True" Then
            Player.Map.Bushes.Remove i
            Player.Map.Bushes.Add "False", , , (i - 1)
        End If
        Randomize
        b = Int((4 * Rnd) + 1)
        Select Case b
            Case 1:
                Drop "Bombs", frmMain.imgSword.Left, frmMain.imgSword.Top
            Case 2:
                Drop "Arrows", frmMain.imgSword.Left, frmMain.imgSword.Top
            Case Else:
        End Select
    End If
    End If
End Function





































Public Function LoadGoldOnline(i As Integer)
    If frmMainOnline.imgGold(i).Left <= frmMainOnline.PlayerImage(0).Left And frmMainOnline.PlayerImage(0).Left <= (frmMainOnline.imgGold(i).Left + frmMainOnline.imgGold(i).Width) Then
    If frmMainOnline.imgGold(i).Top <= frmMainOnline.PlayerImage(0).Top And frmMainOnline.PlayerImage(0).Top <= (frmMainOnline.imgGold(i).Top + frmMainOnline.imgGold(i).Height) Then
        If frmMainOnline.imgGold(i).Visible = False Then Exit Function
        PlayWav "Gold.wav"
        Player.Gold = Player.Gold + frmMainOnline.imgGold(i).Tag
        frmMainOnline.txtGold.Text = Player.Gold
        frmDB.txtGold.Text = Player.Gold
        frmMainOnline.imgGold(i).Visible = False
    End If
    End If
End Function

Public Function LoadArrowsOnline(i As Integer)
    If frmMainOnline.imgArrow(i).Left <= frmMainOnline.PlayerImage(0).Left And frmMainOnline.PlayerImage(0).Left <= (frmMainOnline.imgArrow(i).Left + frmMainOnline.imgArrow(i).Width) Then
    If frmMainOnline.imgArrow(i).Top <= frmMainOnline.PlayerImage(0).Top And frmMainOnline.PlayerImage(0).Top <= (frmMainOnline.imgArrow(i).Top + frmMainOnline.imgArrow(i).Height) Then
        If frmMainOnline.imgArrow(i).Visible = False Then Exit Function
        PlayWav "Gold.wav"
        Player.Arrows = Player.Arrows + frmMainOnline.imgArrow(i).Tag
        frmMainOnline.txtArrows.Text = Player.Arrows
        frmDB.txtArrows.Text = Player.Arrows
        frmMainOnline.imgArrow(i).Visible = False
    End If
    End If
End Function

Public Function LoadBombsOnline(i As Integer)
    If frmMainOnline.imgBomb(i).Left <= frmMainOnline.PlayerImage(0).Left And frmMainOnline.PlayerImage(0).Left <= (frmMainOnline.imgBomb(i).Left + frmMainOnline.imgBomb(i).Width) Then
    If frmMainOnline.imgBomb(i).Top <= frmMainOnline.PlayerImage(0).Top And frmMainOnline.PlayerImage(0).Top <= (frmMainOnline.imgBomb(i).Top + frmMainOnline.imgBomb(i).Height) Then
        If frmMainOnline.imgBomb(i).Visible = False Then Exit Function
        PlayWav "Gold.wav"
        Player.Bombs = Player.Bombs + frmMainOnline.imgBomb(i).Tag
        frmMainOnline.txtBombs.Text = Player.Bombs
        frmDB.txtBombs.Text = Player.Bombs
        frmMainOnline.imgBomb(i).Visible = False
    End If
    End If
End Function

Public Function LoadBushesOnline(i As Integer)
    Dim b As Integer
    
    If frmMainOnline.imgSword.Visible = False Then Exit Function
    If frmMainOnline.imgBush(i).Visible = False Then Exit Function
    If (frmMainOnline.imgSword.Left) = frmMainOnline.imgBush(i).Left Then
    If (frmMainOnline.imgSword.Top) = frmMainOnline.imgBush(i).Top Then
        frmMainOnline.imgBush(i).Visible = False
        If Player.Map.Bushes.Item(i) = "True" Then
            Player.Map.Bushes.Remove i
            Player.Map.Bushes.Add "False", , , (i - 1)
        End If
        Randomize
        b = Int((4 * Rnd) + 1)
        Select Case b
            Case 1:
                Drop "Bombs", frmMainOnline.imgSword.Left, frmMainOnline.imgSword.Top
            Case 2:
                Drop "Arrows", frmMainOnline.imgSword.Left, frmMainOnline.imgSword.Top
            Case Else:
        End Select
    End If
    End If
End Function


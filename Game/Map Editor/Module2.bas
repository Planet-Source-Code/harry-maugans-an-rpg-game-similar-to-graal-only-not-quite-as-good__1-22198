Attribute VB_Name = "Module2"
Public MBoxReturn As Boolean
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Type ItemList2
    Bow As Boolean
    Arrows As Integer
    LongBow As Boolean 'Farther Range than a bow
    LongBowArrows As Integer
    Bomb As Boolean
    Bombs As Integer
    PowerBomb As Boolean 'Twice as powerful as bomb
    PowerBombs As Integer
    RBomb As Boolean 'Twice the Range as bomb
    RBombs As Integer
    BlowGun As Boolean 'Arrow-like movement detonates like a normal bomb on impact
    
    PowerCrystal As Boolean 'Doubles attack power for 5 sec; can only be used once
    SpeedCrystal As Boolean 'Doubles speed for 5 sec; can only be used once
    DefenseCrystal As Boolean 'Doubles Defense for 5 sec; can only be used once
    PowerLocket As Boolean 'Doubles attack power for 10 sec; can only be used once
    SpeedLocket As Boolean 'Doubles speed for 10 sec; can only be used once
    DefenseLocket As Boolean 'Doubles Defense for 10 sec; can only be used once
    GoldenCrystal As Boolean 'Doubles attack, speed, and defense for 5 sec; can only be used once
    GoldenLocket As Boolean 'Doubles attack, speed, and defense for 10 sec; can only be used once
End Type

Public Type NPC3
    Name As New Collection
    Height As New Collection
    Width As New Collection
    XPos As New Collection
    YPos As New Collection
    GifPath As New Collection
    ScriptSource As New Collection
End Type

Public Type Map2
    Gold As New Collection
    Bushes As New Collection
    Arrows As New Collection
    Bombs As New Collection
    MapName As String
    MapFileName As String
    MapData As String
    Music As String
    MapHeight As Integer
    MapWidth As Integer
    MapX As Integer
    MapY As Integer
    MapTemp As String
    MapBoxes As New Collection
    NPCs As NPC3
    Warps As New Collection
    Signs As New Collection
    TotalMapData As String
End Type

Public Type Online2
    Account As String
    Password As String
End Type

Type Direction2
    Up As Boolean
    Down As Boolean
    Left As Boolean
    Right As Boolean
    Dead As Boolean
    Stopped As Boolean
End Type


Public Type Player2
    Gold As Integer
    Arrows As Integer
    Bombs As Integer
    PlayerType As String
    Direction As Direction2
    Name As String
    Health As Integer
    AttackPower As Integer
    Defense As Integer
    Evasion As Integer
    Speed As Integer
    ItemList As ItemList2
    PlayerX As Integer
    PlayerY As Integer
    MapPos As String
    Map As Map2
    MapTile(0 To 9, 0 To 9)
    SwingingSword As Boolean
    Online As Online2
End Type

Public Player As Player2
Const TileWidth = 32
Const TileHeight = 32
Dim titleclose As Boolean
Dim mode As Integer

Public Function PlayWav(Snd As String)
    Dim PlayIt
    Snd = App.Path & "\..\Sounds\" & Snd
    PlayIt = sndPlaySound(Snd, 1)
End Function

Public Function PlayMid(Snd As String)
    Snd = App.Path & "\..\Music\" & Snd
    sndPlaySound Snd, 1
End Function





Public Function GoingDown()
    Player.Direction.Down = True
    Player.Direction.Up = False
    Player.Direction.Left = False
    Player.Direction.Right = False
    Player.Direction.Dead = False
    Player.Direction.Stopped = False
    PlayWav "step.wav"
End Function

Public Function GoingLeft()
    Player.Direction.Down = False
    Player.Direction.Up = False
    Player.Direction.Left = True
    Player.Direction.Right = False
    Player.Direction.Dead = False
    Player.Direction.Stopped = False
    PlayWav "step.wav"
End Function

Public Function GoingRight()
    Player.Direction.Down = False
    Player.Direction.Up = False
    Player.Direction.Left = False
    Player.Direction.Right = True
    Player.Direction.Dead = False
    Player.Direction.Stopped = False
    PlayWav "step.wav"
End Function

Public Function GoingUp()
    Player.Direction.Down = False
    Player.Direction.Up = True
    Player.Direction.Left = False
    Player.Direction.Right = False
    Player.Direction.Dead = False
    Player.Direction.Stopped = False
    PlayWav "step.wav"
End Function

Public Function PlayerDead()
    Player.Direction.Down = False
    Player.Direction.Up = False
    Player.Direction.Left = False
    Player.Direction.Right = False
    Player.Direction.Dead = True
    Player.Direction.Stopped = False
    PlayWav "step.wav"
End Function

Public Function PlayerStopped()
    Player.Direction.Down = False
    Player.Direction.Up = False
    Player.Direction.Left = False
    Player.Direction.Right = False
    Player.Direction.Dead = False
    Player.Direction.Stopped = True
End Function


Public Function LoadGold(i As Integer)
    If frmMain.imgGold(i).Left <= frmMain.PlayerImage(0).Left And frmMain.PlayerImage(0).Left <= (frmMain.imgGold(i).Left + frmMain.imgGold(i).Width) Then
    If frmMain.imgGold(i).Top <= frmMain.PlayerImage(0).Top And frmMain.PlayerImage(0).Top <= (frmMain.imgGold(i).Top + frmMain.imgGold(i).Height) Then
        If frmMain.imgGold(i).Visible = False Then Exit Function
        PlayWav "Gold.wav"
        Player.Gold = Player.Gold + frmMain.imgGold(i).Tag
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
        frmMain.imgBomb(i).Visible = False
    End If
    End If
End Function

Public Function LoadBushes(i As Integer)
    
End Function



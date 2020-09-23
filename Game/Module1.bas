Attribute VB_Name = "Module1"
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
Public p2 As Player2
Const TileWidth = 32
Const TileHeight = 32
Dim titleclose As Boolean
Dim Mode As Integer

Public Function PlayWav(Snd As String)
    Dim PlayIt
    Dim SoundFX As String
    Dim Num As Integer
    
    Num = FreeFile
    Open App.Path & "\Options.txt" For Input As #Num
        Input #Num, SoundFX
        Input #Num, SoundFX
    Close #Num
    
    If SoundFX = "Off" Then Exit Function
    If frmMain.SoundCardInstalled = False Then Exit Function
    Snd = App.Path & "\Sounds\" & Snd
    PlayIt = sndPlaySound(Snd, 1)
End Function

Public Function PlayMid(Snd As String)
    Dim Music As String
    Dim Num As Integer
    
    Num = FreeFile
    Open App.Path & "\Options.txt" For Input As #Num
        Input #Num, Music
        Input #Num, Music
        Input #Num, Music
    Close #Num
    
    If Music = "Off" Then Exit Function
    If frmMain.SoundCardInstalled = False Then Exit Function

    Snd = App.Path & "\Music\" & Snd
    sndPlaySound Snd, 1
End Function


'mode of play
Function LoadMode(moded As Integer)
    Select Case moded
    Case 1
        Load frmSetStats
        'frmLoading.Show
        frmSetStats.Show
        Unload frmStart
        Exit Function
    Case 2
        Unload frmStart
        frmLogin1.Show
        Exit Function
    End Select
End Function
'end mode of play

Public Function Load_Map()
    Player.Map.MapWidth = 29
    Player.Map.MapHeight = 29
End Function








Public Function GoingDown()
    Player.Direction.Down = True
    Player.Direction.Up = False
    Player.Direction.Left = False
    Player.Direction.Right = False
    Player.Direction.Dead = False
    Player.Direction.Stopped = False
    PlayWav "step.wav"
    If frmMain.PlayerImage(0).Picture = frmGraphics.PlayerDown(0).Picture Then
        frmMain.PlayerImage(0).Picture = frmGraphics.PlayerDown(1).Picture
    Else
        frmMain.PlayerImage(0).Picture = frmGraphics.PlayerDown(0).Picture
    End If
    frmMain.Dir = 2
End Function

Public Function GoingLeft()
    Player.Direction.Down = False
    Player.Direction.Up = False
    Player.Direction.Left = True
    Player.Direction.Right = False
    Player.Direction.Dead = False
    Player.Direction.Stopped = False
    PlayWav "step.wav"
    If frmMain.PlayerImage(0).Picture = frmGraphics.PlayerLeft(0).Picture Then
        frmMain.PlayerImage(0).Picture = frmGraphics.PlayerLeft(1).Picture
    Else
        frmMain.PlayerImage(0).Picture = frmGraphics.PlayerLeft(0).Picture
    End If
    frmMain.Dir = 4
End Function

Public Function GoingRight()
    Player.Direction.Down = False
    Player.Direction.Up = False
    Player.Direction.Left = False
    Player.Direction.Right = True
    Player.Direction.Dead = False
    Player.Direction.Stopped = False
    PlayWav "step.wav"
    If frmMain.PlayerImage(0).Picture = frmGraphics.PlayerRight(0).Picture Then
        frmMain.PlayerImage(0).Picture = frmGraphics.PlayerRight(1).Picture
    Else
        frmMain.PlayerImage(0).Picture = frmGraphics.PlayerRight(0).Picture
    End If
    frmMain.Dir = 3
End Function

Public Function GoingUp()
    Player.Direction.Down = False
    Player.Direction.Up = True
    Player.Direction.Left = False
    Player.Direction.Right = False
    Player.Direction.Dead = False
    Player.Direction.Stopped = False
    PlayWav "step.wav"
    If frmMain.PlayerImage(0).Picture = frmGraphics.PlayerUp(0).Picture Then
        frmMain.PlayerImage(0).Picture = frmGraphics.PlayerUp(1).Picture
    Else
        frmMain.PlayerImage(0).Picture = frmGraphics.PlayerUp(0).Picture
    End If
    frmMain.Dir = 1
End Function

Public Function PlayerDead()
    Player.Direction.Down = False
    Player.Direction.Up = False
    Player.Direction.Left = False
    Player.Direction.Right = False
    Player.Direction.Dead = True
    Player.Direction.Stopped = False
    PlayWav "step.wav"
    frmMain.PlayerImage(0).Picture = frmGraphics.PlayerDead.Picture
End Function

Public Function PlayerStopped()
    Player.Direction.Down = False
    Player.Direction.Up = False
    Player.Direction.Left = False
    Player.Direction.Right = False
    Player.Direction.Dead = False
    Player.Direction.Stopped = True
End Function










Public Function GoingDownOnline()
    Player.Direction.Down = True
    Player.Direction.Up = False
    Player.Direction.Left = False
    Player.Direction.Right = False
    Player.Direction.Dead = False
    Player.Direction.Stopped = False
    PlayWav "step.wav"
    If frmMainOnline.PlayerImage(0).Picture = frmGraphics.PlayerDown(0).Picture Then
        frmMainOnline.PlayerImage(0).Picture = frmGraphics.PlayerDown(1).Picture
    Else
        frmMainOnline.PlayerImage(0).Picture = frmGraphics.PlayerDown(0).Picture
    End If
    frmMainOnline.Dir = 2
End Function

Public Function GoingLeftOnline()
    Player.Direction.Down = False
    Player.Direction.Up = False
    Player.Direction.Left = True
    Player.Direction.Right = False
    Player.Direction.Dead = False
    Player.Direction.Stopped = False
    PlayWav "step.wav"
    If frmMainOnline.PlayerImage(0).Picture = frmGraphics.PlayerLeft(0).Picture Then
        frmMainOnline.PlayerImage(0).Picture = frmGraphics.PlayerLeft(1).Picture
    Else
        frmMainOnline.PlayerImage(0).Picture = frmGraphics.PlayerLeft(0).Picture
    End If
    frmMainOnline.Dir = 4
End Function

Public Function GoingRightOnline()
    Player.Direction.Down = False
    Player.Direction.Up = False
    Player.Direction.Left = False
    Player.Direction.Right = True
    Player.Direction.Dead = False
    Player.Direction.Stopped = False
    PlayWav "step.wav"
    If frmMainOnline.PlayerImage(0).Picture = frmGraphics.PlayerRight(0).Picture Then
        frmMainOnline.PlayerImage(0).Picture = frmGraphics.PlayerRight(1).Picture
    Else
        frmMainOnline.PlayerImage(0).Picture = frmGraphics.PlayerRight(0).Picture
    End If
    frmMainOnline.Dir = 3
End Function

Public Function GoingUpOnline()
    Player.Direction.Down = False
    Player.Direction.Up = True
    Player.Direction.Left = False
    Player.Direction.Right = False
    Player.Direction.Dead = False
    Player.Direction.Stopped = False
    PlayWav "step.wav"
    If frmMainOnline.PlayerImage(0).Picture = frmGraphics.PlayerUp(0).Picture Then
        frmMainOnline.PlayerImage(0).Picture = frmGraphics.PlayerUp(1).Picture
    Else
        frmMainOnline.PlayerImage(0).Picture = frmGraphics.PlayerUp(0).Picture
    End If
    frmMainOnline.Dir = 1
End Function

Public Function PlayerDeadOnline()
    Player.Direction.Down = False
    Player.Direction.Up = False
    Player.Direction.Left = False
    Player.Direction.Right = False
    Player.Direction.Dead = True
    Player.Direction.Stopped = False
    PlayWav "step.wav"
    frmMainOnline.PlayerImage(0).Picture = frmGraphics.PlayerDead.Picture
End Function

Public Function PlayWavOnlineOnline(Snd As String)
    Dim PlayIt
    Dim SoundFX As String
    Dim Num As Integer
    
    Num = FreeFile
    Open App.Path & "\Options.txt" For Input As #Num
        Input #Num, SoundFX
        Input #Num, SoundFX
    Close #Num
    
    If SoundFX = "Off" Then Exit Function
    If frmMainOnline.SoundCardInstalled = False Then Exit Function
    Snd = App.Path & "\Sounds\" & Snd
    PlayIt = sndPlaySound(Snd, 1)
End Function

Public Function PlayMidOnline(Snd As String)
    Dim Music As String
    Dim Num As Integer
    
    Num = FreeFile
    Open App.Path & "\Options.txt" For Input As #Num
        Input #Num, Music
        Input #Num, Music
        Input #Num, Music
    Close #Num
    
    If Music = "Off" Then Exit Function
    If frmMainOnline.SoundCardInstalled = False Then Exit Function

    Snd = App.Path & "\Music\" & Snd
    sndPlaySound Snd, 1
End Function

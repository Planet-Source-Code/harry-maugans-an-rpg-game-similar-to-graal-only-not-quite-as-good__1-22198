VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Level Mapping Tool"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic 
      Height          =   375
      Left            =   2280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   66
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make The Map"
      Height          =   255
      Left            =   2400
      TabIndex        =   65
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   63
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   62
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   61
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   60
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   59
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   58
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   57
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   56
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   55
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   54
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   53
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   52
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   51
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   50
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   49
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   48
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   47
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   46
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   45
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   44
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   43
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   42
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   41
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   40
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   39
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   38
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   37
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   36
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   35
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   34
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   33
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   32
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   31
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   30
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   29
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   28
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   27
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   26
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   25
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   24
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   23
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   22
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   21
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   20
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   19
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   18
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   17
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   16
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   15
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   14
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   13
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   12
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   11
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   10
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   9
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   8
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   7
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   6
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   5
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   4
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   3
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   4965
      Left            =   0
      Pattern         =   "*.map"
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin ComctlLib.ImageList imlFloorTiles 
      Left            =   2760
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   36
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":24F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3148
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3D9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":49EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":563E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6290
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6EE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":7B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":8786
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":93D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":A02A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":AC7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":B8CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":C520
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":D172
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":DDC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":EA16
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":F668
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":102BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":10F0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":11B5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":127B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":13402
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":14054
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":14CA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":154F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1614A
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":16D9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":179EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":18640
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":19292
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":19EE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1AB36
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i As Integer
Dim a As Integer
Dim MapData As String
Dim TotalMapData As String
Dim CM As Integer

Private Sub Command1_Click()
    Dim z As Integer
    
    Me.Caption = "Loading..."
    
    Do Until z >= 64
        If txt(z).Text = "" Then txt(z).Text = "_(Default)-BlackMap.map"
        z = z + 1
    Loop
        
    
    a = 0
    Do Until a >= 8
    CM = 0
    Do Until CM >= 8
        Load_Map2
        CM = CM + 1
        DoEvents
    Loop
    a = a + 1
    Loop
    
    Me.Caption = "Level Mapping Tool"
    
    frmDummy.Cls
    frmDummy.Image1.Stretch = True
    z = 0
    Do Until z >= 65
        'frmDummy.img(z).Picture
        z = z + 1
    Loop
    frmDummy.Image1.Width = frmDummy.Width
    frmDummy.Image1.Height = frmDummy.Height
    frmDummy.Visible = True
    Unload Me
End Sub

Private Sub File1_Click()
    
    i = 0
    Do Until txt(i).Text = ""
        i = i + 1
    Loop
    
    txt(i).Text = File1.FileName
End Sub

Private Sub Form_Load()
    File1.Path = App.Path & "\..\Maps\"
End Sub




Public Sub Load_Map2()

    Dim TLine As Integer
    Dim Temp As String
    Dim Temp2 As String
    Dim Temp3 As String
    Dim MapName As String
    
    
    frmDummy.Cls
    Open App.Path & "\..\Maps\" & txt((a * 8) + CM).Text For Input As #1
    
    Input #1, Temp
    Input #1, Temp
    Input #1, Temp
    Input #1, Temp
        
    Do Until Temp = "*WARP*" Or EOF(1)
    
        Line Input #1, Temp
                
        MapData = ""
        If TotalMapData <> "" Then
            TotalMapData = TotalMapData & ","
        End If
        For i = 1 To Len(Temp)
        
            If IsNumeric(Mid(Temp, i, 1)) Or Mid(Temp, i, 1) = "," Then
                MapData = MapData + Mid(Temp, i, 1)
                TotalMapData = TotalMapData + Mid(Temp, i, 1)
            End If
            
        Next i
    
        DisplayMapLine2 TLine
    
        'Increment the Line Counter
        
        TLine = TLine + 1
        
    Loop
    
    Close #1
    
    
    'add to image
    frmDummy.img(CM + (a * 8)).Picture = frmDummy.Image
End Sub

Public Function DisplayMapLine(TLine As Integer)
    Dim i As Integer
    Dim p As Integer
    Dim Num As String
    Dim x As Integer

    i = 1
    x = 0
    Do Until i >= Len(MapData) + 1
        Num = ""
    Do
        If Mid(MapData, i, 1) = "," Or i >= Len(MapData) + 1 Then GoTo Next1
        Num = Num & Mid(MapData, i, 1)
        i = i + 1
    Loop
Next1:
    p = Num
    Set Pic.Picture = imlFloorTiles.ListImages(p).Picture
    frmDummy.PaintPicture Pic.Picture, x * 32, TLine * 32
    
    i = i + 1
    x = x + 1
    
    Loop
    
    
End Function

Public Function DisplayMapLine2(TLine As Integer)
    Dim i As Integer
    Dim p As Integer
    Dim Num As String
    Dim x As Integer

    i = 1
    x = 0
    Do Until i >= Len(MapData) + 1
        Num = ""
    Do
        If Mid(MapData, i, 1) = "," Or i >= Len(MapData) + 1 Then GoTo Next1
        Num = Num & Mid(MapData, i, 1)
        i = i + 1
    Loop
Next1:
    p = Num
    Set Pic.Picture = imlFloorTiles.ListImages(p).Picture
    frmDummy.PaintPicture Pic.Picture, x * 32, TLine * 32, 32, 32
    
    i = i + 1
    x = x + 1
    
    Loop
    
    
End Function


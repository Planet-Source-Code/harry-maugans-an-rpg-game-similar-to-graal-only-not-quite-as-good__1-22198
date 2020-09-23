VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Map Editor"
   ClientHeight    =   10650
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   710
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab2 
      Height          =   3375
      Left            =   3120
      TabIndex        =   27
      Top             =   7080
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   8
      Tab             =   7
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Warps"
      TabPicture(0)   =   "frmMain.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "List1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command13"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Signs"
      TabPicture(1)   =   "frmMain.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command14"
      Tab(1).Control(1)=   "lstSigns"
      Tab(1).Control(2)=   "Command11"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "NPCs"
      TabPicture(2)   =   "frmMain.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command10"
      Tab(2).Control(1)=   "lstNPCs"
      Tab(2).Control(2)=   "Command9"
      Tab(2).Control(3)=   "Image3"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Bushes"
      TabPicture(3)   =   "frmMain.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command19"
      Tab(3).Control(1)=   "Command18"
      Tab(3).Control(2)=   "lstBushes"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Gold"
      TabPicture(4)   =   "frmMain.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command20"
      Tab(4).Control(1)=   "lstGold"
      Tab(4).Control(2)=   "Command21"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Arrows"
      TabPicture(5)   =   "frmMain.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Command22"
      Tab(5).Control(1)=   "lstArrows"
      Tab(5).Control(2)=   "Command23"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Bombs"
      TabPicture(6)   =   "frmMain.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Command25"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "lstBombs"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Command24"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Baddies"
      TabPicture(7)   =   "frmMain.frx":0506
      Tab(7).ControlEnabled=   -1  'True
      Tab(7).Control(0)=   "lstBaddies"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Command27"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Command26"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "Option4"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "Option5"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).ControlCount=   5
      Begin VB.OptionButton Option5 
         Caption         =   "Baddie Bill"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   2040
         Width           =   2175
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Baddie Max"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Remove Baddie"
         Height          =   375
         Left            =   240
         TabIndex        =   51
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Add Baddie"
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   840
         Width           =   2175
      End
      Begin VB.ListBox lstBaddies 
         Height          =   2205
         Left            =   2520
         TabIndex        =   52
         Top             =   840
         Width           =   3495
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Remove Bombs"
         Height          =   375
         Left            =   -73080
         TabIndex        =   48
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ListBox lstBombs 
         Height          =   2205
         Left            =   -70680
         TabIndex        =   50
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Add Bombs"
         Height          =   375
         Left            =   -73080
         TabIndex        =   49
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Remove Arrows"
         Height          =   375
         Left            =   -73080
         TabIndex        =   45
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ListBox lstArrows 
         Height          =   2205
         Left            =   -70680
         TabIndex        =   47
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Add Arrows"
         Height          =   375
         Left            =   -73080
         TabIndex        =   46
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Remove Gold"
         Height          =   375
         Left            =   -73080
         TabIndex        =   42
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ListBox lstGold 
         Height          =   2205
         Left            =   -70680
         TabIndex        =   44
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Add Gold"
         Height          =   375
         Left            =   -73080
         TabIndex        =   43
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Delete Bush"
         Height          =   375
         Left            =   -73080
         TabIndex        =   41
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Add a Bush"
         Height          =   375
         Left            =   -73080
         TabIndex        =   40
         Top             =   960
         Width           =   2175
      End
      Begin VB.ListBox lstBushes 
         Height          =   2205
         Left            =   -70680
         TabIndex        =   39
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Delete NPC"
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ListBox lstNPCs 
         Height          =   2205
         ItemData        =   "frmMain.frx":0522
         Left            =   -72720
         List            =   "frmMain.frx":0524
         TabIndex        =   38
         Top             =   960
         Width           =   3735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Add NPC"
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Delete Selected Sign"
         Height          =   315
         Left            =   -71880
         TabIndex        =   33
         Top             =   2920
         Width           =   2895
      End
      Begin VB.ListBox lstSigns 
         Height          =   2010
         ItemData        =   "frmMain.frx":0526
         Left            =   -74760
         List            =   "frmMain.frx":0528
         TabIndex        =   35
         Top             =   885
         Width           =   5775
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Add A Sign"
         Height          =   315
         Left            =   -74760
         TabIndex        =   34
         Top             =   2920
         Width           =   2895
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Clear All Warps"
         Height          =   255
         Left            =   -73680
         TabIndex        =   28
         Top             =   1365
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Delete Selected Warp"
         Height          =   255
         Left            =   -73680
         TabIndex        =   30
         Top             =   1125
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   -71760
         TabIndex        =   32
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Make A Warp"
         Height          =   255
         Left            =   -73680
         TabIndex        =   31
         Top             =   885
         Width           =   1815
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Make a Screen Mirror Warp"
         Height          =   255
         Left            =   -74760
         TabIndex        =   29
         Top             =   2800
         Width           =   2775
      End
      Begin VB.Image Image3 
         Height          =   1200
         Left            =   -74640
         Picture         =   "frmMain.frx":052A
         Stretch         =   -1  'True
         Top             =   960
         Width           =   1440
      End
      Begin VB.Image Image4 
         Height          =   600
         Left            =   -74640
         Picture         =   "frmMain.frx":096C
         Stretch         =   -1  'True
         Top             =   960
         Width           =   600
      End
   End
   Begin VB.CommandButton Command17 
      Caption         =   "another button"
      Height          =   375
      Left            =   13320
      TabIndex        =   26
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Toggle Grid Lines"
      Height          =   375
      Left            =   13320
      TabIndex        =   25
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Map Converter"
      Height          =   375
      Left            =   13320
      TabIndex        =   24
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Load Map"
      Height          =   375
      Left            =   11520
      TabIndex        =   13
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Stop Test Run"
      Enabled         =   0   'False
      Height          =   375
      Left            =   11520
      TabIndex        =   23
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Map"
      Height          =   375
      Left            =   9720
      TabIndex        =   10
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Run a Map Test"
      Height          =   375
      Left            =   9720
      TabIndex        =   22
      Top             =   8280
      Width           =   1815
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Select Tile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   10200
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Vertical Line"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   9840
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Horizontal Line"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   9480
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   9960
      TabIndex        =   14
      Top             =   120
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   520
      TabCaption(0)   =   "1"
      TabPicture(0)   =   "frmMain.frx":0DAE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image1(23)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image1(22)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Image1(21)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Image1(20)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Image1(19)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Image1(18)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Image1(17)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Image1(16)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Image1(15)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Image1(14)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Image1(13)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Image1(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Image1(11)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Image1(10)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Image1(9)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Image1(8)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Image1(7)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Image1(6)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Image1(5)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Image1(4)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Image1(3)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Image1(2)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Image1(1)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Image1(0)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Image1(24)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Image1(25)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Image1(26)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Image1(27)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Image1(28)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Image1(29)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Image1(30)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Image1(31)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Image1(32)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Image1(33)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Image1(34)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Image1(35)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Image1(36)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Image1(37)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Image1(38)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Image1(39)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Image1(40)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Image1(41)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Image1(42)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Image1(43)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Image1(44)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Image1(45)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Image1(46)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Image1(47)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Image1(48)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Image1(49)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Image1(50)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Image1(51)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Image1(52)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Image1(53)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Image1(54)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Image1(55)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Image1(56)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "Image1(57)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "Image1(58)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "Image1(59)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "Image1(60)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "Image1(61)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "Image1(62)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "Image1(63)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "Image1(64)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "Image1(65)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "Image1(66)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "Image1(67)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "Image1(68)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "Image1(69)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "Image1(70)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "Image1(71)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "Image1(72)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "Image1(73)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "Image1(74)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "Image1(75)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "Image1(76)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "Image1(77)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "Image1(78)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "Image1(79)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "Image1(80)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "Image1(81)"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "Image1(82)"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "Image1(83)"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "Image1(84)"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "Image1(85)"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "Image1(86)"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "Image1(87)"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).ControlCount=   89
      TabCaption(1)   =   "2"
      TabPicture(1)   =   "frmMain.frx":0DCA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape1(1)"
      Tab(1).Control(1)=   "Image1(175)"
      Tab(1).Control(2)=   "Image1(174)"
      Tab(1).Control(3)=   "Image1(173)"
      Tab(1).Control(4)=   "Image1(172)"
      Tab(1).Control(5)=   "Image1(171)"
      Tab(1).Control(6)=   "Image1(170)"
      Tab(1).Control(7)=   "Image1(169)"
      Tab(1).Control(8)=   "Image1(168)"
      Tab(1).Control(9)=   "Image1(167)"
      Tab(1).Control(10)=   "Image1(166)"
      Tab(1).Control(11)=   "Image1(165)"
      Tab(1).Control(12)=   "Image1(164)"
      Tab(1).Control(13)=   "Image1(163)"
      Tab(1).Control(14)=   "Image1(162)"
      Tab(1).Control(15)=   "Image1(161)"
      Tab(1).Control(16)=   "Image1(160)"
      Tab(1).Control(17)=   "Image1(159)"
      Tab(1).Control(18)=   "Image1(158)"
      Tab(1).Control(19)=   "Image1(157)"
      Tab(1).Control(20)=   "Image1(156)"
      Tab(1).Control(21)=   "Image1(155)"
      Tab(1).Control(22)=   "Image1(154)"
      Tab(1).Control(23)=   "Image1(153)"
      Tab(1).Control(24)=   "Image1(152)"
      Tab(1).Control(25)=   "Image1(151)"
      Tab(1).Control(26)=   "Image1(150)"
      Tab(1).Control(27)=   "Image1(149)"
      Tab(1).Control(28)=   "Image1(148)"
      Tab(1).Control(29)=   "Image1(147)"
      Tab(1).Control(30)=   "Image1(146)"
      Tab(1).Control(31)=   "Image1(145)"
      Tab(1).Control(32)=   "Image1(144)"
      Tab(1).Control(33)=   "Image1(143)"
      Tab(1).Control(34)=   "Image1(142)"
      Tab(1).Control(35)=   "Image1(141)"
      Tab(1).Control(36)=   "Image1(140)"
      Tab(1).Control(37)=   "Image1(139)"
      Tab(1).Control(38)=   "Image1(138)"
      Tab(1).Control(39)=   "Image1(137)"
      Tab(1).Control(40)=   "Image1(136)"
      Tab(1).Control(41)=   "Image1(135)"
      Tab(1).Control(42)=   "Image1(134)"
      Tab(1).Control(43)=   "Image1(133)"
      Tab(1).Control(44)=   "Image1(132)"
      Tab(1).Control(45)=   "Image1(131)"
      Tab(1).Control(46)=   "Image1(130)"
      Tab(1).Control(47)=   "Image1(129)"
      Tab(1).Control(48)=   "Image1(128)"
      Tab(1).Control(49)=   "Image1(127)"
      Tab(1).Control(50)=   "Image1(126)"
      Tab(1).Control(51)=   "Image1(125)"
      Tab(1).Control(52)=   "Image1(124)"
      Tab(1).Control(53)=   "Image1(123)"
      Tab(1).Control(54)=   "Image1(122)"
      Tab(1).Control(55)=   "Image1(121)"
      Tab(1).Control(56)=   "Image1(120)"
      Tab(1).Control(57)=   "Image1(119)"
      Tab(1).Control(58)=   "Image1(118)"
      Tab(1).Control(59)=   "Image1(117)"
      Tab(1).Control(60)=   "Image1(116)"
      Tab(1).Control(61)=   "Image1(115)"
      Tab(1).Control(62)=   "Image1(114)"
      Tab(1).Control(63)=   "Image1(113)"
      Tab(1).Control(64)=   "Image1(112)"
      Tab(1).Control(65)=   "Image1(103)"
      Tab(1).Control(66)=   "Image1(102)"
      Tab(1).Control(67)=   "Image1(101)"
      Tab(1).Control(68)=   "Image1(100)"
      Tab(1).Control(69)=   "Image1(99)"
      Tab(1).Control(70)=   "Image1(98)"
      Tab(1).Control(71)=   "Image1(97)"
      Tab(1).Control(72)=   "Image1(96)"
      Tab(1).Control(73)=   "Image1(111)"
      Tab(1).Control(74)=   "Image1(110)"
      Tab(1).Control(75)=   "Image1(109)"
      Tab(1).Control(76)=   "Image1(108)"
      Tab(1).Control(77)=   "Image1(107)"
      Tab(1).Control(78)=   "Image1(106)"
      Tab(1).Control(79)=   "Image1(105)"
      Tab(1).Control(80)=   "Image1(104)"
      Tab(1).Control(81)=   "Image1(95)"
      Tab(1).Control(82)=   "Image1(94)"
      Tab(1).Control(83)=   "Image1(93)"
      Tab(1).Control(84)=   "Image1(92)"
      Tab(1).Control(85)=   "Image1(91)"
      Tab(1).Control(86)=   "Image1(90)"
      Tab(1).Control(87)=   "Image1(89)"
      Tab(1).Control(88)=   "Image1(88)"
      Tab(1).ControlCount=   89
      TabCaption(2)   =   "3"
      TabPicture(2)   =   "frmMain.frx":0DE6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape1(2)"
      Tab(2).Control(1)=   "Image1(263)"
      Tab(2).Control(2)=   "Image1(262)"
      Tab(2).Control(3)=   "Image1(261)"
      Tab(2).Control(4)=   "Image1(260)"
      Tab(2).Control(5)=   "Image1(259)"
      Tab(2).Control(6)=   "Image1(258)"
      Tab(2).Control(7)=   "Image1(257)"
      Tab(2).Control(8)=   "Image1(256)"
      Tab(2).Control(9)=   "Image1(255)"
      Tab(2).Control(10)=   "Image1(254)"
      Tab(2).Control(11)=   "Image1(253)"
      Tab(2).Control(12)=   "Image1(252)"
      Tab(2).Control(13)=   "Image1(251)"
      Tab(2).Control(14)=   "Image1(250)"
      Tab(2).Control(15)=   "Image1(249)"
      Tab(2).Control(16)=   "Image1(248)"
      Tab(2).Control(17)=   "Image1(247)"
      Tab(2).Control(18)=   "Image1(246)"
      Tab(2).Control(19)=   "Image1(245)"
      Tab(2).Control(20)=   "Image1(244)"
      Tab(2).Control(21)=   "Image1(243)"
      Tab(2).Control(22)=   "Image1(242)"
      Tab(2).Control(23)=   "Image1(241)"
      Tab(2).Control(24)=   "Image1(240)"
      Tab(2).Control(25)=   "Image1(239)"
      Tab(2).Control(26)=   "Image1(238)"
      Tab(2).Control(27)=   "Image1(237)"
      Tab(2).Control(28)=   "Image1(236)"
      Tab(2).Control(29)=   "Image1(235)"
      Tab(2).Control(30)=   "Image1(234)"
      Tab(2).Control(31)=   "Image1(233)"
      Tab(2).Control(32)=   "Image1(232)"
      Tab(2).Control(33)=   "Image1(231)"
      Tab(2).Control(34)=   "Image1(230)"
      Tab(2).Control(35)=   "Image1(229)"
      Tab(2).Control(36)=   "Image1(228)"
      Tab(2).Control(37)=   "Image1(227)"
      Tab(2).Control(38)=   "Image1(226)"
      Tab(2).Control(39)=   "Image1(225)"
      Tab(2).Control(40)=   "Image1(224)"
      Tab(2).Control(41)=   "Image1(223)"
      Tab(2).Control(42)=   "Image1(222)"
      Tab(2).Control(43)=   "Image1(221)"
      Tab(2).Control(44)=   "Image1(220)"
      Tab(2).Control(45)=   "Image1(219)"
      Tab(2).Control(46)=   "Image1(218)"
      Tab(2).Control(47)=   "Image1(217)"
      Tab(2).Control(48)=   "Image1(216)"
      Tab(2).Control(49)=   "Image1(215)"
      Tab(2).Control(50)=   "Image1(214)"
      Tab(2).Control(51)=   "Image1(213)"
      Tab(2).Control(52)=   "Image1(212)"
      Tab(2).Control(53)=   "Image1(211)"
      Tab(2).Control(54)=   "Image1(210)"
      Tab(2).Control(55)=   "Image1(209)"
      Tab(2).Control(56)=   "Image1(208)"
      Tab(2).Control(57)=   "Image1(207)"
      Tab(2).Control(58)=   "Image1(206)"
      Tab(2).Control(59)=   "Image1(205)"
      Tab(2).Control(60)=   "Image1(204)"
      Tab(2).Control(61)=   "Image1(203)"
      Tab(2).Control(62)=   "Image1(202)"
      Tab(2).Control(63)=   "Image1(201)"
      Tab(2).Control(64)=   "Image1(200)"
      Tab(2).Control(65)=   "Image1(199)"
      Tab(2).Control(66)=   "Image1(198)"
      Tab(2).Control(67)=   "Image1(197)"
      Tab(2).Control(68)=   "Image1(196)"
      Tab(2).Control(69)=   "Image1(195)"
      Tab(2).Control(70)=   "Image1(194)"
      Tab(2).Control(71)=   "Image1(193)"
      Tab(2).Control(72)=   "Image1(192)"
      Tab(2).Control(73)=   "Image1(191)"
      Tab(2).Control(74)=   "Image1(190)"
      Tab(2).Control(75)=   "Image1(189)"
      Tab(2).Control(76)=   "Image1(188)"
      Tab(2).Control(77)=   "Image1(187)"
      Tab(2).Control(78)=   "Image1(186)"
      Tab(2).Control(79)=   "Image1(185)"
      Tab(2).Control(80)=   "Image1(184)"
      Tab(2).Control(81)=   "Image1(183)"
      Tab(2).Control(82)=   "Image1(182)"
      Tab(2).Control(83)=   "Image1(181)"
      Tab(2).Control(84)=   "Image1(180)"
      Tab(2).Control(85)=   "Image1(179)"
      Tab(2).Control(86)=   "Image1(178)"
      Tab(2).Control(87)=   "Image1(177)"
      Tab(2).Control(88)=   "Image1(176)"
      Tab(2).ControlCount=   89
      TabCaption(3)   =   "4"
      TabPicture(3)   =   "frmMain.frx":0E02
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Image1(264)"
      Tab(3).Control(1)=   "Image1(265)"
      Tab(3).Control(2)=   "Image1(266)"
      Tab(3).Control(3)=   "Image1(267)"
      Tab(3).Control(4)=   "Image1(268)"
      Tab(3).Control(5)=   "Image1(269)"
      Tab(3).Control(6)=   "Image1(270)"
      Tab(3).Control(7)=   "Image1(271)"
      Tab(3).Control(8)=   "Shape1(3)"
      Tab(3).Control(9)=   "Image1(279)"
      Tab(3).Control(10)=   "Image1(278)"
      Tab(3).Control(11)=   "Image1(277)"
      Tab(3).Control(12)=   "Image1(276)"
      Tab(3).Control(13)=   "Image1(275)"
      Tab(3).Control(14)=   "Image1(274)"
      Tab(3).Control(15)=   "Image1(273)"
      Tab(3).Control(16)=   "Image1(272)"
      Tab(3).ControlCount=   17
      TabCaption(4)   =   "5"
      TabPicture(4)   =   "frmMain.frx":0E1E
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Shape1(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "6"
      TabPicture(5)   =   "frmMain.frx":0E3A
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Shape1(5)"
      Tab(5).ControlCount=   1
      Begin VB.Image Image1 
         Height          =   480
         Index           =   272
         Left            =   -74760
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   273
         Left            =   -74160
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   274
         Left            =   -73560
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   275
         Left            =   -72960
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   276
         Left            =   -72360
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   277
         Left            =   -71760
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   278
         Left            =   -71160
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   279
         Left            =   -70560
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   600
         Index           =   5
         Left            =   -75360
         Top             =   7080
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   600
         Index           =   4
         Left            =   -75360
         Top             =   7080
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   600
         Index           =   3
         Left            =   -75360
         Top             =   7080
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   600
         Index           =   2
         Left            =   -75360
         Top             =   7080
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   600
         Index           =   1
         Left            =   -75360
         Top             =   7080
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   271
         Left            =   -70560
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   270
         Left            =   -71160
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   269
         Left            =   -71760
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   268
         Left            =   -72360
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   267
         Left            =   -72960
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   266
         Left            =   -73560
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   265
         Left            =   -74160
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   264
         Left            =   -74760
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   263
         Left            =   -70560
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   262
         Left            =   -71160
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   261
         Left            =   -71760
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   260
         Left            =   -72360
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   259
         Left            =   -72960
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   258
         Left            =   -73560
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   257
         Left            =   -74160
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   256
         Left            =   -74760
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   255
         Left            =   -70560
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   254
         Left            =   -71160
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   253
         Left            =   -71760
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   252
         Left            =   -72360
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   251
         Left            =   -72960
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   250
         Left            =   -73560
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   249
         Left            =   -74160
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   248
         Left            =   -74760
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   247
         Left            =   -70560
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   246
         Left            =   -71160
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   245
         Left            =   -71760
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   244
         Left            =   -72360
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   243
         Left            =   -72960
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   242
         Left            =   -73560
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   241
         Left            =   -74160
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   240
         Left            =   -74760
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   239
         Left            =   -70560
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   238
         Left            =   -71160
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   237
         Left            =   -71760
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   236
         Left            =   -72360
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   235
         Left            =   -72960
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   234
         Left            =   -73560
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   233
         Left            =   -74160
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   232
         Left            =   -74760
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   231
         Left            =   -70560
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   230
         Left            =   -71160
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   229
         Left            =   -71760
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   228
         Left            =   -72360
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   227
         Left            =   -72960
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   226
         Left            =   -73560
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   225
         Left            =   -74160
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   224
         Left            =   -74760
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   223
         Left            =   -70560
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   222
         Left            =   -71160
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   221
         Left            =   -71760
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   220
         Left            =   -72360
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   219
         Left            =   -72960
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   218
         Left            =   -73560
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   217
         Left            =   -74160
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   216
         Left            =   -74760
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   215
         Left            =   -70560
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   214
         Left            =   -71160
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   213
         Left            =   -71760
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   212
         Left            =   -72360
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   211
         Left            =   -72960
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   210
         Left            =   -73560
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   209
         Left            =   -74160
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   208
         Left            =   -74760
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   207
         Left            =   -70560
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   206
         Left            =   -71160
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   205
         Left            =   -71760
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   204
         Left            =   -72360
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   203
         Left            =   -72960
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   202
         Left            =   -73560
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   201
         Left            =   -74160
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   200
         Left            =   -74760
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   199
         Left            =   -70560
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   198
         Left            =   -71160
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   197
         Left            =   -71760
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   196
         Left            =   -72360
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   195
         Left            =   -72960
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   194
         Left            =   -73560
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   193
         Left            =   -74160
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   192
         Left            =   -74760
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   191
         Left            =   -70560
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   190
         Left            =   -71160
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   189
         Left            =   -71760
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   188
         Left            =   -72360
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   187
         Left            =   -72960
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   186
         Left            =   -73560
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   185
         Left            =   -74160
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   184
         Left            =   -74760
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   183
         Left            =   -70560
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   182
         Left            =   -71160
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   181
         Left            =   -71760
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   180
         Left            =   -72360
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   179
         Left            =   -72960
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   178
         Left            =   -73560
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   177
         Left            =   -74160
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   176
         Left            =   -74760
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   175
         Left            =   -70560
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   174
         Left            =   -71160
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   173
         Left            =   -71760
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   172
         Left            =   -72360
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   171
         Left            =   -72960
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   170
         Left            =   -73560
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   169
         Left            =   -74160
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   168
         Left            =   -74760
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   167
         Left            =   -70560
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   166
         Left            =   -71160
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   165
         Left            =   -71760
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   164
         Left            =   -72360
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   163
         Left            =   -72960
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   162
         Left            =   -73560
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   161
         Left            =   -74160
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   160
         Left            =   -74760
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   159
         Left            =   -70560
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   158
         Left            =   -71160
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   157
         Left            =   -71760
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   156
         Left            =   -72360
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   155
         Left            =   -72960
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   154
         Left            =   -73560
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   153
         Left            =   -74160
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   152
         Left            =   -74760
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   151
         Left            =   -70560
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   150
         Left            =   -71160
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   149
         Left            =   -71760
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   148
         Left            =   -72360
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   147
         Left            =   -72960
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   146
         Left            =   -73560
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   145
         Left            =   -74160
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   144
         Left            =   -74760
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   143
         Left            =   -70560
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   142
         Left            =   -71160
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   141
         Left            =   -71760
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   140
         Left            =   -72360
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   139
         Left            =   -72960
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   138
         Left            =   -73560
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   137
         Left            =   -74160
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   136
         Left            =   -74760
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   135
         Left            =   -70560
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   134
         Left            =   -71160
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   133
         Left            =   -71760
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   132
         Left            =   -72360
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   131
         Left            =   -72960
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   130
         Left            =   -73560
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   129
         Left            =   -74160
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   128
         Left            =   -74760
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   127
         Left            =   -70560
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   126
         Left            =   -71160
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   125
         Left            =   -71760
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   124
         Left            =   -72360
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   123
         Left            =   -72960
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   122
         Left            =   -73560
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   121
         Left            =   -74160
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   120
         Left            =   -74760
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   119
         Left            =   -70560
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   118
         Left            =   -71160
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   117
         Left            =   -71760
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   116
         Left            =   -72360
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   115
         Left            =   -72960
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   114
         Left            =   -73560
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   113
         Left            =   -74160
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   112
         Left            =   -74760
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   103
         Left            =   -70560
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   102
         Left            =   -71160
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   101
         Left            =   -71760
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   100
         Left            =   -72360
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   99
         Left            =   -72960
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   98
         Left            =   -73560
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   97
         Left            =   -74160
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   96
         Left            =   -74760
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   111
         Left            =   -70560
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   110
         Left            =   -71160
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   109
         Left            =   -71760
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   108
         Left            =   -72360
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   107
         Left            =   -72960
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   106
         Left            =   -73560
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   105
         Left            =   -74160
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   104
         Left            =   -74760
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   95
         Left            =   -70560
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   94
         Left            =   -71160
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   93
         Left            =   -71760
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   92
         Left            =   -72360
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   91
         Left            =   -72960
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   90
         Left            =   -73560
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   89
         Left            =   -74160
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   88
         Left            =   -74760
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   87
         Left            =   4440
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   86
         Left            =   3840
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   85
         Left            =   3240
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   84
         Left            =   2640
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   83
         Left            =   2040
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   82
         Left            =   1440
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   81
         Left            =   840
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   80
         Left            =   240
         Top             =   6720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   79
         Left            =   4440
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   78
         Left            =   3840
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   77
         Left            =   3240
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   76
         Left            =   2640
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   75
         Left            =   2040
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   74
         Left            =   1440
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   73
         Left            =   840
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   72
         Left            =   240
         Top             =   6120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   71
         Left            =   4440
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   70
         Left            =   3840
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   69
         Left            =   3240
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   68
         Left            =   2640
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   67
         Left            =   2040
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   66
         Left            =   1440
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   65
         Left            =   840
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   64
         Left            =   240
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   63
         Left            =   4440
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   62
         Left            =   3840
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   61
         Left            =   3240
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   60
         Left            =   2640
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   59
         Left            =   2040
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   58
         Left            =   1440
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   57
         Left            =   840
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   56
         Left            =   240
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   55
         Left            =   4440
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   54
         Left            =   3840
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   53
         Left            =   3240
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   52
         Left            =   2640
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   51
         Left            =   2040
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   50
         Left            =   1440
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   49
         Left            =   840
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   48
         Left            =   240
         Top             =   4320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   47
         Left            =   4440
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   46
         Left            =   3840
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   45
         Left            =   3240
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   44
         Left            =   2640
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   43
         Left            =   2040
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   42
         Left            =   1440
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   41
         Left            =   840
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   40
         Left            =   240
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   39
         Left            =   4440
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   38
         Left            =   3840
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   37
         Left            =   3240
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   36
         Left            =   2640
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   35
         Left            =   2040
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   34
         Left            =   1440
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   33
         Left            =   840
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   32
         Left            =   240
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   31
         Left            =   4440
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   30
         Left            =   3840
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   29
         Left            =   3240
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   28
         Left            =   2640
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   27
         Left            =   2040
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   26
         Left            =   1440
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   25
         Left            =   840
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   24
         Left            =   240
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   240
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   840
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   1440
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2040
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   2640
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   5
         Left            =   3240
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   6
         Left            =   3840
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   7
         Left            =   4440
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   8
         Left            =   240
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   9
         Left            =   840
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   10
         Left            =   1440
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   11
         Left            =   2040
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   12
         Left            =   2640
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   13
         Left            =   3240
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   14
         Left            =   3840
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   15
         Left            =   4440
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   16
         Left            =   240
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   17
         Left            =   840
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   18
         Left            =   1440
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   19
         Left            =   2040
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   20
         Left            =   2640
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   21
         Left            =   3240
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   22
         Left            =   3840
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   23
         Left            =   4440
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   600
         Index           =   0
         Left            =   -360
         Top             =   7200
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Map Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   7080
      Width           =   2655
      Begin VB.TextBox txtMusic 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "20"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "14"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtMapName 
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   120
         Picture         =   "frmMain.frx":0E56
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Music File Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Width:"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Height:"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6720
      Left            =   75
      ScaleHeight     =   448
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   9
      Top             =   75
      Width           =   9600
      Begin VB.PictureBox Pic 
         Height          =   135
         Left            =   3360
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox txtBoxes 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape shpBush 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   3
         Height          =   480
         Left            =   0
         Top             =   2880
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgBush 
         Height          =   465
         Index           =   0
         Left            =   5880
         Top             =   3240
         Width           =   465
      End
      Begin VB.Image imgBomb 
         Height          =   465
         Index           =   0
         Left            =   4920
         Top             =   3240
         Width           =   465
      End
      Begin VB.Image imgArrow 
         Height          =   465
         Index           =   0
         Left            =   4200
         Top             =   3240
         Width           =   465
      End
      Begin VB.Image imgGold 
         Height          =   465
         Index           =   0
         Left            =   3480
         Top             =   3360
         Width           =   465
      End
      Begin ComctlLib.ImageList imlFloorTiles 
         Left            =   3840
         Top             =   1680
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
               Picture         =   "frmMain.frx":1298
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1EEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":2B3C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":378E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":43E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":5032
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":5C84
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":68D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":7528
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":817A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":8DCC
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":9A1E
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":A670
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":B2C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":BF14
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":CB66
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":D7B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":E40A
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":F05C
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":FCAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":10900
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":11552
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":121A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":12DF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":13A48
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1469A
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":152EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":15F3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":16790
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":173E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":18034
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":18C86
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":198D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1A52A
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1B17C
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1BDCE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Shape shpNPC 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Height          =   480
         Left            =   0
         Top             =   2400
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape shpGrid 
         BorderColor     =   &H00000000&
         Height          =   480
         Index           =   0
         Left            =   0
         Shape           =   5  'Rounded Square
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Height          =   480
         Left            =   0
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   480
         Left            =   0
         Top             =   1440
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Height          =   480
         Left            =   0
         Top             =   960
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Height          =   480
         Left            =   0
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgNPC 
         Height          =   135
         Index           =   0
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   1440
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image PlayerImage 
         Height          =   690
         Index           =   0
         Left            =   1920
         ToolTipText     =   "You"
         Top             =   1920
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "New Map"
      Height          =   375
      Left            =   11520
      TabIndex        =   15
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Flood With Tile"
      Height          =   375
      Left            =   9720
      TabIndex        =   12
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Image Image6 
      Height          =   1440
      Left            =   7800
      Picture         =   "frmMain.frx":1CA20
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   1605
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":1E095
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   465
   End
   Begin VB.Label Label8 
      Caption         =   "Right Click:"
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
      Left            =   480
      TabIndex        =   18
      Top             =   9000
      Width           =   1095
   End
   Begin VB.Label lblMap 
      Alignment       =   2  'Center
      Caption         =   "Unsaved Map"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   17
      Top             =   7560
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   295
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   294
      Left            =   600
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   293
      Left            =   1200
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   292
      Left            =   1800
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   291
      Left            =   2400
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   290
      Left            =   3000
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   289
      Left            =   3600
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   288
      Left            =   4200
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   451
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   647
      X2              =   647
      Y1              =   453
      Y2              =   0
   End
   Begin VB.Line Line4 
      BorderWidth     =   10
      X1              =   0
      X2              =   645
      Y1              =   4
      Y2              =   4
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   0
      X2              =   647
      Y1              =   455
      Y2              =   455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Open A New Map"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuMap 
      Caption         =   "&Map"
      Begin VB.Menu mnuMapOpen 
         Caption         =   "&Open Map"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuMapSave 
         Caption         =   "&Save Map"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuTestRun 
      Caption         =   "&Test Run"
      Begin VB.Menu mnuRunTest 
         Caption         =   "&Run A Test"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuStopTest 
         Caption         =   "&Stop the Test"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuAddins 
      Caption         =   "&Add-Ins"
      Begin VB.Menu mnuAddinsTileMaker 
         Caption         =   "&Tile Maker"
      End
      Begin VB.Menu mnuAddInsNPCWorkshop 
         Caption         =   "&NPC Workshop"
         Shortcut        =   ^W
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ImageSelected As Integer
Public X1, Y1 As Integer
Public MD As Boolean
Public MS As Boolean
Public MW As Boolean
Public MB As Boolean
Public MSW As Boolean
Public MN As Boolean
Public TotalMapData As String
Public MapData As String
Public MapHeight As Integer
Public MapWidth As Integer
Public MapName As String
Public MapBoxes As New Collection
Public Testing As Boolean

Private Sub Command1_Click()

    If List1.ListIndex > "-1" Then
        If List1.ListCount > 1 Then
            List1.ListIndex = List1.ListIndex - 1
            List1.RemoveItem List1.ListIndex + 1
        End If
        If List1.ListCount = 1 Then
            List1.Clear
        End If
        List1_Click
    End If
End Sub

Private Sub Command10_Click()
    If lstNPCs.ListCount = 0 Then Exit Sub
    Unload imgNPC(lstNPCs.ListIndex + 1)
    lstNPCs.RemoveItem lstNPCs.ListIndex
    If lstNPCs.ListCount = 0 Then Exit Sub
    lstNPCs.Selected(lstNPCs.ListIndex + 1) = True
End Sub

Private Sub Command11_Click()
    MS = True
End Sub

Private Sub Command12_Click()
    MW = True
    MSW = True
End Sub

Private Sub Command13_Click()
    Dim MsgResult As String
    MsgResult = MsgBox("Are you sure you want to delete all warps in this level?", vbYesNo, "Clear All Warps?")
    If MsgResult = vbNo Then Exit Sub
    List1.Clear
End Sub

Private Sub Command14_Click()
    On Error Resume Next
    lstSigns.RemoveItem lstSigns.ListIndex
    Shape5.Visible = False
End Sub

Private Sub Command15_Click()
'    Me.Hide
'    frmConvert.Show vbModal, Me
'    Me.Show
End Sub

Private Sub Command16_Click()
    Dim x As Integer
    Dim y As Integer
    Dim z As Integer
    
If shpGrid(0).Visible = False Then
    shpGrid(0).Visible = True
    Shape2.BorderColor = vbWhite
    z = 1
    Do Until x >= 20
        y = 0
        Do Until y >= 14
            Load shpGrid(z)
            shpGrid(z).Left = x * 32
            shpGrid(z).Top = y * 32
            shpGrid(z).Visible = True
            y = y + 1
            z = z + 1
        Loop
        x = x + 1
    Loop
Else
    shpGrid(0).Visible = False
    Shape2.BorderColor = vbBlack
    z = 1
    Do Until z >= (14 * 20) + 1
        Unload shpGrid(z)
        z = z + 1
    Loop
End If
End Sub

Private Sub Command18_Click()
    MB = True
End Sub

Private Sub Command19_Click()
    If lstBushes.ListCount = 0 Then Exit Sub
    Unload imgBush(lstBushes.ListIndex + 1)
    If lstBushes.ListCount = 0 Then Exit Sub
    lstBushes.RemoveItem lstBushes.ListIndex
End Sub

Private Sub Command2_Click()
    Dim i As Integer
    Dim FileName As String
    Dim r As Integer
    Dim t As Integer
    Dim Line As String
    
    Do Until i = 279
        If txtBoxes(i).Text = "" Then
            MsgBox "Please fill in all the spaces with tiles.", , "Error"
            Exit Sub
        End If
        i = i + 1
    Loop
    
    If Trim(txtMapName.Text) = "" Then
        MsgBox "Please entitle your map.", , "Error Saving"
        txtMapName.SetFocus
        Exit Sub
    End If
    
    ''now the save part.  =)
    If lblMap.Caption <> "" And lblMap.Caption <> "Unsaved Map" Then
        FileName = InputBox("Please enter a filename for the map: (ie. 'Map1.map')", "Save Map", lblMap.Caption)
    Else
        FileName = InputBox("Please enter a filename for the map: (ie. 'Map1.map')", "Save Map")
    End If
    
    If Trim(FileName) = "" Then Exit Sub
    lblMap.Caption = FileName
    If Dir$(App.Path & "\..\Maps\" & FileName) <> "" Then Kill (App.Path & "\..\Maps\" & FileName)
    
    Open App.Path & "\..\Maps\" & FileName For Output As #1
        Print #1, txtWidth.Text
        Print #1, txtHeight.Text
        Write #1, txtMapName.Text
        Print #1, txtMusic.Text
        
        t = 0
        Do Until t = 14
            Line = " "
            r = 0
            Do Until r = 20
                Line = Line & (txtBoxes(r + (t * 20)).Text) & ","
                r = r + 1
            Loop
                Line = Mid(Line, 1, Len(Line) - 1)
            Print #1, Line
            t = t + 1
        Loop
        
        Print #1, "*WARP*"
        If List1.ListCount <> 0 Then
        t = 0
        List1.Selected(t) = True
        Do Until t = List1.ListCount
            List1.Selected(t) = True
            Print #1, List1.Text
            t = t + 1
        Loop
        End If
        Print #1, "*NPC*"
            t = 0
            Do Until t = lstNPCs.ListCount
                lstNPCs.Selected(t) = True
                Print #1, lstNPCs.Text
                t = t + 1
            Loop
        Print #1, "*SIGNS*"
            t = 0
            Do Until t = lstSigns.ListCount
                lstSigns.Selected(t) = True
                Print #1, lstSigns.Text
                t = t + 1
            Loop
        Print #1, "*GOLD*"
        If lstGold.ListCount <= 0 Then GoTo SkipGold
        t = 0
        lstGold.ListIndex = 0
        Do Until t = lstGold.ListCount
            lstGold.ListIndex = t
            Print #1, lstGold.Text
            t = t + 1
        Loop
SkipGold:
        Print #1, "*ARROWS*"
        If lstArrows.ListCount <= 0 Then GoTo SkipArrows
        t = 0
        lstArrows.ListIndex = 0
        Do Until t = lstArrows.ListCount
            lstArrows.ListIndex = t
            Print #1, lstArrows.Text
            t = t + 1
        Loop
SkipArrows:
        Print #1, "*BOMBS*"
        If lstBombs.ListCount <= 0 Then GoTo SkipBombs
        t = 0
        lstBombs.ListIndex = 0
        Do Until t = lstBombs.ListCount
            lstBombs.ListIndex = t
            Print #1, lstBombs.Text
            t = t + 1
        Loop
SkipBombs:
        Print #1, "*BUSHES*"
        If lstBushes.ListCount <= 0 Then GoTo SkipBushes
        t = 0
        lstBushes.ListIndex = 0
        Do Until t = lstBushes.ListCount
            lstBushes.ListIndex = t
            Print #1, lstBushes.Text
            t = t + 1
        Loop
SkipBushes:
    Close #1
    
    
End Sub

Private Sub Command3_Click()
    MW = True
End Sub

Private Sub Command4_Click()
    Dim i As Integer
    
    X1 = 0
    Y1 = 0
    Do Until Y1 > 14
        X1 = 0
    Do Until X1 > 20
        Picture1.PaintPicture Image1(ImageSelected), (X1 - 1) * 32, (Y1 - 1) * 32
        X1 = X1 + 1
    Loop
        Y1 = Y1 + 1
    Loop
    Do Until i = 279
        txtBoxes(i).Text = ImageSelected + 1
        i = i + 1
    Loop
    
    
End Sub

Private Sub Command5_Click()
    frmLoad.Show vbModal, Me
End Sub

Private Sub Command6_Click()
    Dim i As Integer
    
    Picture1.Cls
    txtMapName.Text = ""
    txtMusic.Text = ""
    Do Until i = List1.ListCount
        List1.RemoveItem 0
    Loop
    lblMap.Caption = "Unsaved Map"
    Shape3.Visible = False
    Shape4.Visible = False
    Shape2.Visible = False
End Sub

Private Sub Command7_Click()
    If lblMap.Caption <> "" And lblMap.Caption <> "Unsaved Map" Then
        DrawMap lblMap.Caption
        Module1.Init
        Testing = True
        Command8.Enabled = True
        Command7.Enabled = False
        SSTab2.Enabled = False
        SSTab1.Enabled = False
        Player.PlayerX = (PlayerImage(0).Left / 32)
        Player.PlayerY = (PlayerImage(0).Top / 32)
        PlayerImage(0).Visible = True
        Picture1.SetFocus
    Else
        MsgBox "Please save your map first.", , "Save First"
        Exit Sub
    End If
End Sub

Private Sub Command8_Click()
    Testing = False
    Command8.Enabled = False
    Command7.Enabled = True
    SSTab2.Enabled = True
    SSTab1.Enabled = True
    PlayerImage(0).Visible = False
    PlayerImage(0).Left = 32
    PlayerImage(0).Top = 96
End Sub

Private Sub Command9_Click()
    MN = True
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim a As Integer
    
    Do Until i >= imlFloorTiles.ListImages.Count
        Set Image1(i).Picture = imlFloorTiles.ListImages(i + 1).Picture
        Image1(i).Visible = True
        i = i + 1
    Loop
    
    a = 1
    Do Until a > 279
        Load txtBoxes(a)
        a = a + 1
    Loop
    
    Me.Top = 0
    Load frmGraphics
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Shape2.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Image1_Click(Index As Integer)
    Dim i As Integer
    
    Do Until i >= imlFloorTiles.ListImages.Count
        Image1(i).BorderStyle = 0
        i = i + 1
    Loop
    
    Image1(Index).BorderStyle = 1
    
    ImageSelected = Index
    
    Shape1(SSTab1.Tab).Visible = True
    Shape1(SSTab1.Tab).Top = (Image1(Index).Top - ((Shape1(SSTab1.Tab).Height - Image1(Index).Height) / 2))
    Shape1(SSTab1.Tab).Left = (Image1(Index).Left - ((Shape1(SSTab1.Tab).Width - Image1(Index).Width) / 2))
End Sub

Private Sub List1_Click()
    Dim x2 As Integer
    Dim y2 As Integer
    Dim i As Integer
    
    i = 1
    Do
        If Mid(List1.Text, i, 1) = "," Then GoTo DoneWarpX
        x2 = x2 & Mid(List1.Text, i, 1)
        i = i + 1
    Loop
DoneWarpX:
    
    i = i + 1
    Do
        If Mid(List1.Text, i, 1) = "," Then GoTo DoneWarpY
        y2 = y2 & Mid(List1.Text, i, 1)
        i = i + 1
    Loop
DoneWarpY:
    
    Shape3.Width = 32
    Shape3.Height = 32
    Shape3.Top = (y2 - 1) * 32
    Shape3.Left = (x2 - 1) * 32
    Shape3.Visible = True
End Sub

Private Sub List1_DblClick()
    Dim x2 As Integer
    Dim y2 As Integer
    Dim Dx2, Dy2 As Integer
    Dim MapName As String
    Dim i As Integer
    
    List1_Click
    Shape4.Left = Shape3.Left
    Shape4.Top = Shape3.Top
    
    i = 1
    Do
        If Mid(List1.Text, i, 1) = "," Then GoTo DoneWarpX
        x2 = x2 & Mid(List1.Text, i, 1)
        i = i + 1
    Loop
DoneWarpX:
    
    i = i + 1
    Do
        If Mid(List1.Text, i, 1) = "," Then GoTo DoneWarpY
        y2 = y2 & Mid(List1.Text, i, 1)
        i = i + 1
    Loop
DoneWarpY:
    
    i = i + 1
    Do
        If Mid(List1.Text, i, 1) = "," Then GoTo DoneWarpMap
        MapName = MapName & Mid(List1.Text, i, 1)
        i = i + 1
    Loop
DoneWarpMap:
    
    i = i + 1
    Do
        If Mid(List1.Text, i, 1) = "," Then GoTo DoneWarpDX
        Dx2 = Dx2 & Mid(List1.Text, i, 1)
        i = i + 1
    Loop
DoneWarpDX:
    
    Do
        If i = Len(List1.Text) Then GoTo DoneWarpDY
        Dy2 = Dy2 & Mid(List1.Text, i + 1, 1)
        i = i + 1
    Loop
DoneWarpDY:
    
    
    frmWarp.Editing = True
    Load frmWarp
    frmWarp.txtMap.Text = MapName
    frmWarp.txtDX = Dx2
    frmWarp.txtDY = Dy2
    frmWarp.Show vbModal, Me
    
End Sub


Private Sub lstSigns_Click()
    Dim x2 As Integer
    Dim y2 As Integer
    Dim i As Integer
    
    i = 1
    Do
        If Mid(lstSigns.Text, i, 1) = "," Then GoTo DoneWarpX2
        x2 = x2 & Mid(lstSigns.Text, i, 1)
        i = i + 1
    Loop
DoneWarpX2:
    
    i = i + 1
    Do
        If Mid(lstSigns.Text, i, 1) = "," Then GoTo DoneWarpY2
        y2 = y2 & Mid(lstSigns.Text, i, 1)
        i = i + 1
    Loop
DoneWarpY2:
    
    Shape5.Width = 32
    Shape5.Height = 32
    Shape5.Top = (y2 - 1) * 32
    Shape5.Left = (x2 - 1) * 32
    Shape5.Visible = True
End Sub

Private Sub mnuAddInsNPCWorkshop_Click()
    Shell App.Path & "\..\NPC Workshop\NPC.exe", vbNormalFocus
End Sub

Private Sub mnuAddinsTileMaker_Click()
    Shell App.Path & "\..\Tile Utility\TileMaker.exe", vbNormalFocus
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    Command6_Click
End Sub

Private Sub mnuMapOpen_Click()
    Command5_Click
End Sub

Private Sub mnuMapSave_Click()
    Command2_Click
End Sub

Private Sub mnuRunTest_Click()
    Command7_Click
End Sub

Private Sub mnuStopTest_Click()
    Command8_Click
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Testing = False Then Exit Sub
    Dim i As Integer

    If KeyCode = vbKeyUp Then GoingUp
    If KeyCode = vbKeyDown Then GoingDown
    If KeyCode = vbKeyLeft Then GoingLeft
    If KeyCode = vbKeyRight Then GoingRight
    
    If Player.PlayerX = 19 And KeyCode = vbKeyRight Then Exit Sub
    If Player.PlayerX = 0 And KeyCode = vbKeyLeft Then Exit Sub
    If Player.PlayerY = 0 And KeyCode = vbKeyUp Then Exit Sub
    If Player.PlayerY = 13 And KeyCode = vbKeyDown Then Exit Sub
        
    If CheckForWalls(KeyCode) = False Then Exit Sub
        
        If Player.Direction.Stopped = False Then
            If KeyCode = vbKeyLeft Then
                PlayerImage(0).Left = PlayerImage(0).Left - 32
                Player.PlayerX = Player.PlayerX - 1
            End If
            If KeyCode = vbKeyRight Then
                PlayerImage(0).Left = PlayerImage(0).Left + 32
                Player.PlayerX = Player.PlayerX + 1
            End If
            If KeyCode = vbKeyUp Then
                PlayerImage(0).Top = PlayerImage(0).Top - 32
                Player.PlayerY = Player.PlayerY - 1
            End If
            If KeyCode = vbKeyDown Then
                PlayerImage(0).Top = PlayerImage(0).Top + 32
                Player.PlayerY = Player.PlayerY + 1
            End If
        End If ' end if player.direction.stopped = false
 
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Testing = False Then Exit Sub
        PlayerStopped
End Sub

Public Function CheckForWalls(ByVal key As String) As Boolean
        If Testing = False Then Exit Function

    Select Case key
    Case vbKeyUp
        If Player.Map.MapBoxes(Player.PlayerX & "," & Player.PlayerY - 1) = "unwalkable" Then
            CheckForWalls = False
            Exit Function
        Else
            CheckForWalls = True
            Exit Function
        End If
    Case vbKeyDown
        If Player.Map.MapBoxes(Player.PlayerX & "," & Player.PlayerY + 1) = "unwalkable" Then
            CheckForWalls = False
            Exit Function
        Else
            CheckForWalls = True
            Exit Function
        End If
    Case vbKeyLeft
        If Player.Map.MapBoxes(Player.PlayerX - 1 & "," & Player.PlayerY) = "unwalkable" Then
            CheckForWalls = False
            Exit Function
        Else
            CheckForWalls = True
            Exit Function
        End If
    Case vbKeyRight
        If Player.Map.MapBoxes(Player.PlayerX + 1 & "," & Player.PlayerY) = "unwalkable" Then
            CheckForWalls = False
            Exit Function
        Else
            CheckForWalls = True
            Exit Function
        End If
    Case Else
        CheckForWalls = True
        Exit Function
    End Select
End Function




Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Testing = True Then Exit Sub

Dim i As Integer
Dim SignText As String

If Button = 1 Then
    If MW = False And MB = False And MS = False And MN = False Then
    If Shape1(SSTab1.Tab).Visible = False Then Exit Sub
        Picture1.PaintPicture Image1(ImageSelected), (X1 - 1) * 32, (Y1 - 1) * 32
        txtBoxes(((X1 - 1) + ((Y1 - 1)) * 20)).Text = ImageSelected + 1
        MD = True
        Shape3.Visible = False
    ElseIf MW = True Then
        If (Y1 - 1) < 15 And (X1 - 1) < 21 Then
            Shape3.Visible = True
            Shape3.Width = 32
            Shape3.Height = 32
            Shape3.Top = (Y1 - 1) * 32
            Shape3.Left = (X1 - 1) * 32
            MD = True
        End If
    ElseIf MN = True Then
        If (Y1 - 1) < 15 And (X1 - 1) < 21 Then
            shpNPC.Visible = True
            shpNPC.Width = 32
            shpNPC.Height = 32
            shpNPC.Top = (Y1 - 1) * 32
            shpNPC.Left = (X1 - 1) * 32
            MD = True
        End If
    End If
End If
If Button = 2 Then
    If Option1.Value = True Then
        i = 1
        Do Until i > 20
            X1 = i
            Picture1.PaintPicture Image1(ImageSelected), (X1 - 1) * 32, (Y1 - 1) * 32
            If Y1 = 0 Then
                txtBoxes(((Y1 - 1) + ((X1 - 1)) * 20)).Text = ImageSelected + 1
            Else
                txtBoxes(((X1 - 1) + ((Y1 - 1)) * 20)).Text = ImageSelected + 1
            End If
            i = i + 1
        Loop
    End If
    If Option2.Value = True Then
        i = 1
        Do Until i > 14
            Y1 = i
            Picture1.PaintPicture Image1(ImageSelected), (X1 - 1) * 32, (Y1 - 1) * 32
            If X1 = 0 Then
                txtBoxes(((X1 - 1) + ((Y1 - 1)) * 20)).Text = ImageSelected + 1
            Else
                txtBoxes(((X1 - 1) + ((Y1 - 1)) * 20)).Text = ImageSelected + 1
            End If
            i = i + 1
        Loop
    End If
    If Option3.Value = True Then
        If txtBoxes(((X1 - 1) + ((Y1 - 1)) * 20)).Text = "" Then Exit Sub
        ImageSelected = txtBoxes(((X1 - 1) + ((Y1 - 1)) * 20)).Text
        ImageSelected = ImageSelected - 1
        Image1_Click ImageSelected
    End If
End If

    If MS = True Then
        Shape2.Visible = False
        Shape5.Top = (Y1 - 1) * 32
        Shape5.Left = (X1 - 1) * 32
        Shape5.Visible = True
        DoEvents
        SignText = InputBox("What will the sign say?", "Sign Text")
        If Trim(SignText) = "" Then MS = False: MD = False: MW = False: Shape2.Visible = True: Shape5.Visible = False: Exit Sub
        lstSigns.AddItem X1 & "," & Y1 & "," & SignText
        Shape5.Visible = False
        Shape2.Visible = True
        DoEvents
        MD = False
        MW = False
        MS = False
        Exit Sub
    End If
    If MB = True Then
        Load imgBush(lstBushes.ListCount + 1)
        lstBushes.AddItem X1 & "," & Y1
        imgBush(lstBushes.ListCount).Left = (X1 - 1) * 32
        imgBush(lstBushes.ListCount).Top = (Y1 - 1) * 32
        imgBush(lstBushes.ListCount).Picture = frmGraphics.imgBush.Picture
        imgBush(lstBushes.ListCount).Visible = True
    End If
    
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Testing = True Then Exit Sub

    Dim i As Integer
    Dim Temp As String
        
    x = x / 32
    y = y / 32
    If x >= 21 Then Exit Sub
    If y >= 15 Then Exit Sub
    i = 1
    
    If MN = True And MW = False And MB = False Then
        Shape2.Visible = False
        If MD = True Then
            shpNPC.Visible = True
            If X1 >= 0 And Y1 >= 0 Then
            If X1 < 21 And Y1 < 15 Then
                If shpNPC.Height >= 1 And shpNPC.Width >= 1 Then
                    shpNPC.Height = (y * 32) - shpNPC.Top
                    shpNPC.Width = (x * 32) - shpNPC.Left
                End If
            End If
            End If
        End If
    End If
    
    Do Until Temp = "." Or i >= Len(x) + 1
        Temp = Mid(x, i + 1, 1)
        i = i + 1
    Loop
    x = Mid(x, 1, i)
    i = 1
    Temp = ""
    Do Until Temp = "." Or i >= Len(y) + 1
        Temp = Mid(y, i + 1, 1)
        i = i + 1
    Loop
    y = Mid(y, 1, i)
    Me.Caption = "Map Editor [" & x + 1 & "," & y + 1 & "]"
    
    If MW = False And MN = False And MB = False Then
        If x < 20 Then
        If y < 14 Then
        Shape2.Visible = True
        Shape2.Left = x * 32
        Shape2.Top = y * 32
        End If
        End If
    ElseIf MW = True And MN = False Then
        Shape2.Visible = False
        If MD = True Then
            Shape3.Visible = True
            If X1 >= 0 And Y1 >= 0 Then
            If X1 < 21 And Y1 < 15 Then
                If Shape3.Width >= 32 Then
                If (X1 - 1) > x Then
                    Shape3.Width = Shape3.Width - 32
                End If
                If (X1 - 1) < x Then
                    Shape3.Width = Shape3.Width + 32
                End If
                Else
                    If (X1 - 1) < x Then
                    If y >= (Y1 - 1) Then
                        Shape3.Width = Shape3.Width + 32
                    End If
                    End If
                End If
                If Shape3.Height >= 32 Then
                If (Y1 - 1) > y Then
                    Shape3.Height = Shape3.Height - 32
                End If
                If (Y1 - 1) < y Then
                    Shape3.Height = Shape3.Height + 32
                End If
                Else
                    If (Y1 - 1) < y Then
                    If x >= (X1 - 1) Then
                        Shape3.Height = Shape3.Height + 32
                    End If
                    End If
                End If
            End If
            End If
        End If
    End If
    X1 = x + 1
    Y1 = y + 1
    
    ''see if it is a drag
    If MD = True And MW = False And MN = False And MB = False Then
        If X1 >= 0 And Y1 >= 0 Then
        If X1 < 21 And Y1 < 15 Then
            Picture1.PaintPicture Image1(ImageSelected), (X1 - 1) * 32, (Y1 - 1) * 32
            txtBoxes(((X1 - 1) + ((Y1 - 1)) * 20)).Text = ImageSelected + 1
        End If
        End If
    End If
    If MB = True And MW = False And MN = False And MD = False Then
        If X1 >= 0 And Y1 >= 0 Then
        If X1 < 21 And Y1 < 15 Then
            shpBush.Left = (X1 - 1) * 32
            shpBush.Top = (Y1 - 1) * 32
            shpBush.Visible = True
        End If
        End If
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Testing = True Then Exit Sub

    Dim i, a, B As Integer
    Dim MapName As String
    Dim ScriptSource, Name, GifName As String
    
    Picture1.Refresh
    If MB = True Then
        shpBush.Visible = False
        MB = False
        Exit Sub
    End If
    
    If MW = True Then
        frmWarp.Editing = False
        Load frmWarp
        If MSW = True Then
            If Shape3.Height > 32 And Shape3.Width = 32 Then
                a = 1
                MapName = InputBox("Please enter the destination map:", "Destination Map?")
                If Trim(MapName) = "" Then GoTo Done
                Do Until a > Shape3.Height / 32
                    i = Shape3.Left / 32
                    i = i + 1
                    If i = "20" Then i = 1: GoTo k
                    
                    If i = "1" Then i = 20
k:
                    List1.AddItem (Shape3.Left / 32) + 1 & "," & (Shape3.Top / 32) + a & "," & MapName & "," & i & "," & (Shape3.Top / 32) + a
                    a = a + 1
                Loop
            End If
            If Shape3.Width > 32 And Shape3.Height = 32 Then
                a = 1
                MapName = InputBox("Please enter the destination map:", "Destination Map?")
                If Trim(MapName) = "" Then GoTo Done
                Do Until a > Shape3.Width / 32
                    i = Shape3.Top / 32
                    i = i + 1
                    If i = "14" Then i = 1: GoTo r
        
                    If i = "1" Then i = 14
r:
                    List1.AddItem (Shape3.Left / 32) + a & "," & (Shape3.Top / 32) + 1 & "," & MapName & "," & (Shape3.Left / 32) + a & "," & i
                    a = a + 1
                Loop
            End If
Done:
            MW = False
            MSW = False
            MD = False
            Shape3.Visible = False
            Exit Sub
        End If
        If Shape3.Height < 32 Or Shape3.Width < 32 Then Shape3.Visible = False
        MW = False
        If Shape3.Visible = True Then
            If Shape3.Height = 32 And Shape3.Width = 32 Then
                Shape4.Top = Shape3.Top
                Shape4.Left = Shape3.Left
                Shape4.Visible = True
                frmWarp.Show vbModal, Me
                Shape4.Visible = False
            Else
                If Shape3.Height > 32 And Shape3.Width = 32 Then
                i = 0
                Do Until i >= (Shape3.Height \ 32)
                    Shape4.Left = Shape3.Left
                    Shape4.Top = Shape3.Top + (i * 32)
                    Shape4.Visible = True
                    frmWarp.Show vbModal, Me
                    Shape4.Visible = False
                    i = i + 1
                Loop
                End If
                
                If Shape3.Width > 32 And Shape3.Height = 32 Then
                i = 0
                Do Until i >= (Shape3.Width \ 32)
                    Shape4.Left = Shape3.Left + (i * 32)
                    Shape4.Top = Shape3.Top
                    Shape4.Visible = True
                    frmWarp.Show vbModal, Me
                    Shape4.Visible = False
                    i = i + 1
                Loop
                End If
                
                If Shape3.Height > 32 And Shape3.Width > 32 Then
                    i = 0
                    B = 0
                    Do Until i >= (Shape3.Width \ 32)
                        Shape4.Left = Shape3.Left + (i * 32)
                        Shape4.Top = Shape3.Top
                        Shape4.Visible = True
                        frmWarp.Show vbModal, Me
                        Shape4.Visible = False
                            a = 1
                            Do Until a >= (Shape3.Height \ 32)
                                Shape4.Left = Shape3.Left + (i * 32)
                                Shape4.Top = Shape3.Top + (a * 32)
                                Shape4.Visible = True
                                frmWarp.Show vbModal, Me
                                Shape4.Visible = False
                                a = a + 1
                            Loop
                        i = i + 1
                        'b = b + 1
                    Loop
                End If
            End If
        End If
        Shape2.Visible = True
    ElseIf MD = True And MN = True Then
        Load imgNPC(lstNPCs.ListCount + 1)
        shpNPC.Refresh
        DoEvents
        Name = InputBox("Please enter a name for the NPC:", "Name?")
        GifName = InputBox("Please enter a gifname in the images directory for the NPC: (ie. bob.gif)", "Gif To Use?")
        ScriptSource = InputBox("Please enter the filename of the NPC script to use: (ie. bob.npc)", "Script To Use?")
        If Name = "" Or GifName = "" Or ScriptSource = "" Then Unload imgNPC(lstNPCs.ListCount + 1): MD = False: MN = False: GoTo skipp
        imgNPC(lstNPCs.ListCount + 1).Visible = True
        imgNPC(lstNPCs.ListCount + 1).Left = shpNPC.Left
        imgNPC(lstNPCs.ListCount + 1).Top = shpNPC.Top
        imgNPC(lstNPCs.ListCount + 1).Height = shpNPC.Height
        imgNPC(lstNPCs.ListCount + 1).Width = shpNPC.Width
        imgNPC(lstNPCs.ListCount + 1).Picture = LoadPicture(App.Path & "\..\Images\" & GifName)
        imgNPC(lstNPCs.ListCount + 1).ToolTipText = Name
        lstNPCs.AddItem Name & "," & ScriptSource & "," & GifName & "," & imgNPC(lstNPCs.ListCount + 1).Left / 32 & "," & imgNPC(lstNPCs.ListCount + 1).Top / 32 & "," & imgNPC(lstNPCs.ListCount + 1).Height & "," & imgNPC(lstNPCs.ListCount + 1).Width
        MN = False
        MD = False
        shpNPC.Visible = False
    End If
skipp:
    MD = False
    Shape3.Visible = False
End Sub


Public Function DrawMap(FileName As String)
    Dim i As Integer
    Dim TLine As Integer
    Dim Temp As String
    Dim Temp2 As String
    Dim Temp3 As String
    Dim XName, ScriptSource, GifPath As String
    Dim XPos, YPos, Wid, Hei As Integer
    Dim WarpX, WarpY As Integer
    Dim DwarpX, DwarpY As Integer
    Dim Music As String
    Dim MapName As String
    Dim Num As Integer

    
    X1 = 0
    Y1 = 0
    
    Command6_Click

    Open App.Path & "\..\Maps\" & FileName For Input As #1
    
    lblMap.Caption = FileName
    Input #1, MapWidth
    txtWidth.Text = MapWidth
    Input #1, MapHeight
    txtHeight.Text = MapHeight
    Input #1, MapName
    txtMapName.Text = MapName
    Input #1, Music
    txtMusic.Text = Music
       
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
    
        DisplayMapLine TLine
    
        'Increment the Line Counter
        
        TLine = TLine + 1
    
    Loop

    i = 1: WarpX = "0": DwarpX = "0": WarpY = "0": DwarpY = "0": MapName = ""
        
    ClearWarps
    Do Until Temp = "*NPC*" Or EOF(1)
        Input #1, WarpX
        If WarpX = "*NPC*" Then GoTo SkipWarp
        Input #1, WarpY
        Input #1, MapName
        Input #1, DwarpX
        Input #1, DwarpY
        List1.AddItem WarpX & "," & WarpY & "," & MapName & "," & DwarpX & "," & DwarpY
    Loop
    
SkipWarp:
    
    'NPCs
    ClearNPCs
   Do Until Temp = "*SIGNS*"
        ' NPC Format: Name, ScriptFile, GIF, X, Y, Height, Width
        Input #1, Temp
        If Temp = "*SIGNS*" Then GoTo SkipNPCs
        XName = Temp
        Input #1, Temp
        ScriptSource = Temp
        Input #1, Temp
        GifPath = Temp
        Input #1, Temp
        XPos = Temp
        Input #1, Temp
        YPos = Temp
        Input #1, Temp
        Hei = Temp
        Input #1, Temp
        Wid = Temp
        Player.Map.NPCs.Name.Add Name
        Player.Map.NPCs.ScriptSource.Add ScriptSource
        Player.Map.NPCs.GifPath.Add GifPath
        Player.Map.NPCs.XPos.Add XPos
        Player.Map.NPCs.YPos.Add YPos
        Player.Map.NPCs.Height.Add Hei
        Player.Map.NPCs.Width.Add Wid
    Loop
        LoadNPCs
      
SkipNPCs:
    
    ClearSigns
    Temp = ""
    lstSigns.Clear
    Do Until Temp = "*GOLD*"
        Line Input #1, Temp
        If Temp = "*GOLD*" Then GoTo GGold
        Open App.Path & "\SignTemp.txt" For Output As #23
            Print #23, Temp
        Close #23
        Open App.Path & "\SignTemp.txt" For Input As #23
            Input #23, Temp
            Input #23, Temp2
            Input #23, Temp3
        Close #23
        Kill App.Path & "\SignTemp.txt"
        Player.Map.Signs.Add Temp3, Temp & "," & Temp2
        lstSigns.AddItem Temp & "," & Temp2 & "," & Temp3
    Loop
GGold:
        
Goldd:
    ClearGold
    Do Until Temp = "*ARROWS*"
        Line Input #1, Temp
        If Temp = "*ARROWS*" Then GoTo SkipGold
        Open App.Path & "\GoldTemp.txt" For Output As #25
            Print #25, Temp
        Close #25
        Open App.Path & "\GoldTemp.txt" For Input As #25
            Input #25, Temp
            Input #25, Temp2
            Input #25, Temp3
            Player.Map.Gold.Add Temp3, Temp & "," & Temp2
            Load imgGold(Player.Map.Gold.Count)
            imgGold(Player.Map.Gold.Count).Left = (Temp - 1) * 32
            imgGold(Player.Map.Gold.Count).Top = (Temp2 - 1) * 32
            imgGold(Player.Map.Gold.Count).Tag = Temp3
            If imgGold(Player.Map.Gold.Count).Tag < 25 Then imgGold(Player.Map.Gold.Count).Picture = frmGraphics.imgGold1.Picture
            If imgGold(Player.Map.Gold.Count).Tag >= 25 And imgGold(Player.Map.Gold.Count).Tag < 75 Then imgGold(Player.Map.Gold.Count).Picture = frmGraphics.imgGold2.Picture
            If imgGold(Player.Map.Gold.Count).Tag >= 76 Then imgGold(Player.Map.Gold.Count).Picture = frmGraphics.imgGold2.Picture
            imgGold(Player.Map.Gold.Count).Visible = True
        Close #25
        Kill App.Path & "\GoldTemp.txt"
    Loop
SkipGold:
    ClearArrows
    Do Until Temp = "*BOMBS*"
        Line Input #1, Temp
        If Temp = "*BOMBS*" Then GoTo SkipArrows
        Open App.Path & "\ArrowsTemp.txt" For Output As #24
            Print #24, Temp
        Close #24
        Open App.Path & "\ArrowsTemp.txt" For Input As #24
            Input #24, Temp
            Input #24, Temp2
            Input #24, Temp3
            'Arrows format: X, Y, Ammount
            Player.Map.Arrows.Add Temp3, Temp & "," & Temp2
            Load imgArrow(Player.Map.Arrows.Count)
            imgArrow(Player.Map.Arrows.Count).Picture = frmGraphics.imgArrow.Picture
            imgArrow(Player.Map.Arrows.Count).Tag = Temp3
            imgArrow(Player.Map.Arrows.Count).Left = (Temp + 1) * 32
            imgArrow(Player.Map.Arrows.Count).Top = (Temp2 + 1) * 32
            imgArrow(Player.Map.Arrows.Count).Visible = True
        Close #24
        Kill App.Path & "\ArrowsTemp.txt"
    Loop
SkipArrows:
    ClearBombs
    Do Until Temp = "*BUSHES*"
        Line Input #1, Temp
        If Temp = "*BUSHES*" Then GoTo SkipBombs
        Open App.Path & "\BombsTemp.txt" For Output As #24
            Print #24, Temp
        Close #24
        Open App.Path & "\BombsTemp.txt" For Input As #24
            Input #24, Temp
            Input #24, Temp2
            Input #24, Temp3
            'Bombs format: X, Y, Ammount
            Player.Map.Bombs.Add Temp3, Temp & "," & Temp2
            Load imgBomb(Player.Map.Bombs.Count)
            imgBomb(Player.Map.Bombs.Count).Picture = frmGraphics.imgBomb.Picture
            imgBomb(Player.Map.Bombs.Count).Tag = Temp3
            imgBomb(Player.Map.Bombs.Count).Left = (Temp + 1) * 32
            imgBomb(Player.Map.Bombs.Count).Top = (Temp2 + 1) * 32
            imgBomb(Player.Map.Bombs.Count).Visible = True
        Close #24
        Kill App.Path & "\BombsTemp.txt"
    Loop
SkipBombs:
    ClearBushes
    Do Until EOF(1)
        Line Input #1, Temp
        Open App.Path & "\BushesTemp.txt" For Output As #24
            Print #24, Temp
        Close #24
        Open App.Path & "\BushesTemp.txt" For Input As #24
            Input #24, Temp
            Input #24, Temp2
            Player.Map.Bushes.Add "True", Temp & "," & Temp2
            'Bushes format: X,Y
            Load imgBush(Player.Map.Bushes.Count)
            imgBomb(Player.Map.Bushes.Count).Picture = frmGraphics.imgBush.Picture
            imgBomb(Player.Map.Bushes.Count).Left = (Temp + 1) * 32
            imgBomb(Player.Map.Bushes.Count).Top = (Temp2 + 1) * 32
            imgBomb(Player.Map.Bushes.Count).Visible = True
            Close #24
            Kill App.Path & "\BushesTemp.txt"
    Loop
    Close #1
        
    
End Function

Public Function DisplayMapLine(TLine As Integer)
    Dim i As Integer
    Dim p As Integer
    Dim Num As String
    Dim x As Integer
    
    i = 1
    x = 0
    X1 = 0
    Y1 = TLine
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
    Picture1.PaintPicture Pic.Picture, x * 32, TLine * 32
    txtBoxes(X1 + (Y1 * 20)).Text = p
    i = i + 1
    x = x + 1
    Y1 = TLine
    X1 = X1 + 1
    Loop
    
    
End Function

Public Sub ClearWarps()
    On Error Resume Next

    Dim i As Integer
    
    For i = 1 To List1.ListCount
        List1.RemoveItem 1
    Next i
    
End Sub

Private Sub PlayerImage_Click(Index As Integer)
    Picture1.SetFocus
End Sub

Public Sub LoadNPC3(str As String)
    Open App.Path & "\Temp.txt" For Output As #2
        Print #2, str
    Close #2
    Open App.Path & "\Temp.txt" For Input As #2
        Input #2, str 'name
        Input #2, str 'script source
        Input #2, str 'image path
        Load imgNPC(lstNPCs.ListCount)
        imgNPC(lstNPCs.ListCount).Visible = True
        imgNPC(lstNPCs.ListCount).Picture = LoadPicture(App.Path & "\..\Images\" & str)
        Input #2, str 'X
        imgNPC(lstNPCs.ListCount).Left = str * 32
        Input #2, str 'Y
        imgNPC(lstNPCs.ListCount).Top = str * 32
        Input #2, str 'height
        imgNPC(lstNPCs.ListCount).Height = str
        Input #2, str
        imgNPC(lstNPCs.ListCount).Width = str
    Close #2
End Sub

Public Function LoadNPCs()
    On Error Resume Next
    Dim i As Integer
    
    If Player.Map.NPCs.GifPath.Count = "0" Then Exit Function
    
    For i = 1 To Player.Map.NPCs.ScriptSource.Count
        DoEvents
        Load imgNPC(i)
        imgNPC(i).Visible = True
        imgNPC(i).Left = Player.Map.NPCs.XPos(i) * 32
        imgNPC(i).Top = Player.Map.NPCs.YPos(i) * 32
        imgNPC(i).Picture = LoadPicture(App.Path & "\Images\" & Player.Map.NPCs.GifPath(i))
        imgNPC(i).Width = Player.Map.NPCs.Width(i)
        imgNPC(i).Height = Player.Map.NPCs.Height(i)
        imgNPC(i).Tag = Player.Map.NPCs.Name(i)
        imgNPC(i).ToolTipText = Player.Map.NPCs.Name(i)
    Next i
End Function

